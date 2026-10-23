"""Resumable file/page scheduler. No Qt widgets, account tokens, or raw RPC logs."""
from copy import deepcopy
import hashlib
import json
from pathlib import Path
import shutil
from threading import Event, Lock
import time

from main import file_sha256
from apply_translation import apply_translation
from pptx_pipeline.stages import AnalysisStage, PipelineContext
from pptx_pipeline.layout import PPTXLayoutModifier, LayoutPolicy
from pptx_pipeline.rendering import PowerPointRenderer
from translate_texts import PPTXTextTranslator, read_records, write_translation
from gemini_cli import EmptyModelResponseError
from .adapters.codex import CodexConnectionAdapter
from .adapters.gemini import GeminiConnectionAdapter
from .adapters.antigravity import AntigravityConnectionAdapter


class Stopped(Exception): pass
class QuotaPause(Exception): pass


class RunControl:
    def __init__(self):
        self.resume_event = Event(); self.resume_event.set()
        self.stop_event = Event()
        self.lock = Lock()
        self.active = None
    def stop(self):
        self.stop_event.set(); self.resume_event.set()
        with self.lock:
            if self.active and hasattr(self.active, 'cancel'): self.active.cancel()
    def boundary(self, paused):
        if self.stop_event.is_set(): raise Stopped()
        if not self.resume_event.is_set():
            paused()
            while not self.resume_event.wait(.15):
                if self.stop_event.is_set(): raise Stopped()
        if self.stop_event.is_set(): raise Stopped()


def is_quota_error(exc):
    for event in getattr(getattr(exc,'result',None),'events', ()):
        if event.get('method') == 'turn/completed':
            info = event.get('params',{}).get('turn',{}).get('error') or {}
            tag = info.get('codexErrorInfo')
            if tag == 'usageLimitExceeded' or tag == 'rateLimitExceeded': return True
            if isinstance(tag, dict) and any(k in tag for k in ('usageLimitExceeded','rateLimitExceeded')): return True
            if isinstance(tag, dict) and tag.get('httpConnectionFailed',{}).get('httpStatusCode') == 429: return True
    return isinstance(exc, QuotaPause)


class GuardedProvider:
    def __init__(self, underlying, control, instruction=''):
        self.underlying, self.control = underlying, control
        self.model = underlying.model
        self.instruction = instruction
    def generate(self, *args, **kwargs):
        if self.control.stop_event.is_set(): raise Stopped()
        with self.control.lock: self.control.active = self.underlying
        try:
            if self.instruction:
                args = (args[0] + '\nAdditional layout instructions (keep all text and formulas unchanged):\n' + self.instruction, *args[1:])
            result = self.underlying.generate(*args, **kwargs)
            self.model = self.underlying.model
            return result
        except Exception as exc:
            if self.control.stop_event.is_set(): raise Stopped() from exc
            if is_quota_error(exc): raise QuotaPause() from exc
            raise
        finally:
            with self.control.lock: self.control.active = None


class DesktopEngine:
    def __init__(self, settings, executables, emit, control, demo=False):
        self.settings, self.executables = deepcopy(settings), dict(executables)
        self.emit, self.control, self.demo = emit, control, demo
        self.lock = Lock()
        self.current_settings = None
        self.quota_paused = False

    def update_settings(self, settings, quota_resume=False):
        with self.lock:
            self.settings = deepcopy(settings)
            if quota_resume and self.current_settings:
                for key in ('translate','adjust'):
                    current=getattr(self.current_settings,key); pending=getattr(settings,key)
                    # Switching the connection must not mix languages/options within one PPT.
                    current.provider,current.model,current.effort=pending.provider,pending.model,pending.effort

    def boundary(self): self.control.boundary(lambda: self.emit({'type':'paused'}))

    def provider(self, node, layout=False):
        key = node.provider
        adapter_type={'codex':CodexConnectionAdapter,'gemini':GeminiConnectionAdapter,'antigravity':AntigravityConnectionAdapter}[key]
        adapter = adapter_type(self.executables.get(key))
        underlying=adapter.provider(node)
        if key=='codex':
            def notification(event):
                if event.get('method')=='account/rateLimits/updated':
                    self.emit({'type':'quota_updated','data':event.get('params',{})})
            underlying.on_event=notification
        return GuardedProvider(underlying, self.control, node.prompt if layout else '')

    def unit(self, function):
        while True:
            self.boundary()
            try: return function()
            except QuotaPause:
                self.quota_paused = True
                self.control.resume_event.clear()
                self.emit({'type':'quota_paused'})
                self.boundary()
                self.quota_paused = False

    def run_file(self, job):
        with self.lock: self.current_settings = deepcopy(self.settings)
        cfg = self.current_settings
        self.boundary()
        self.emit({'type':'file_started','job':job.id})
        if self.demo:
            for node in ('extract','translate','adjust','save'):
                self.emit({'type':'node','job':job.id,'node':node,'status':'running'})
                for page in range(1, 7):
                    self.boundary(); time.sleep(.12)
                    self.emit({'type':'progress','job':job.id,'node':node,'page':page,'pages':6})
                self.emit({'type':'node','job':job.id,'node':node,'status':'done'})
            return {'output':'','status':'done','slides':6}
        target = cfg.save.target(job.source, cfg.translate.target_language)
        if target.exists() or target.with_suffix('.report.json').exists(): raise FileExistsError(f'기존 결과를 덮어쓰지 않습니다: {target.name}')
        # Work files live separately from user output names. No auth state is serialized.
        workdir = Path(job.workdir) if job.workdir else Path(cfg.save.directory) / '.pptx-work' / job.id
        job.workdir = str(workdir)
        workdir.mkdir(parents=True, exist_ok=True)
        state_file = workdir / 'checkpoint.json'
        source_hash = file_sha256(job.source)
        # Hash configuration, never persist prompts/account information in project files.
        fingerprint = hashlib.sha256(repr(cfg).encode()).hexdigest()
        state = json.loads(state_file.read_text(encoding='utf-8')) if state_file.exists() else {
            'source_sha256':source_hash,'fingerprint':fingerprint,'translated':{},'translation_failed':[],
            'adjusted_pages':[],'review_reports':[],'applied':False}
        if state['source_sha256'] != source_hash: raise ValueError('원본이 변경되어 체크포인트를 재사용할 수 없습니다.')
        if state['fingerprint'] != fingerprint: raise ValueError('설정이 변경됐습니다. 재실행으로 새 작업을 시작하세요.')
        def save_state():
            temp = state_file.with_suffix('.tmp')
            temp.write_text(json.dumps(state,ensure_ascii=False,indent=2),encoding='utf-8')
            temp.replace(state_file)
        context = PipelineContext(job.source, workdir)
        def stage(node, status): self.emit({'type':'node','job':job.id,'node':node,'status':status})
        stage('extract','running')
        if not context.manifest.exists() or not context.texts.exists(): AnalysisStage().run(context)
        document = json.loads(context.manifest.read_text(encoding='utf-8'))
        document['translate_all'] = True
        context.manifest.write_text(json.dumps(document,ensure_ascii=False,indent=2),encoding='utf-8')
        pages = len(document['slides'])
        self.emit({'type':'analyzed','job':job.id,'pages':pages,'workdir':str(workdir)})
        stage('extract','done')
        records = read_records(context.texts)
        translations = {int(k):v for k,v in state['translated'].items()}
        stage('translate','running' if cfg.translate.enabled else 'skipped')
        for page in range(1,pages + 1):
            self.boundary()
            batch = [r for r in records if r.page == page and r.id not in translations]
            if batch:
                def translate_page():
                    options = self.current_settings.translate
                    if not options.enabled: return {r.id:r.text for r in batch}
                    return PPTXTextTranslator(self.provider(options), model=None,
                        source_language=options.source_language,target_language=options.target_language,
                        glossary=options.glossary,additional_instructions=options.prompt).translate_batch(batch)
                try: translated = self.unit(translate_page)
                except EmptyModelResponseError:
                    translated = {r.id:r.text for r in batch}
                    state['translation_failed'].append(page)
                    self.emit({'type':'warning','message':f'{page}페이지 번역 빈 응답 · 원문 유지'})
                translations.update(translated)
                state['translated'] = {str(k):v for k,v in translations.items()}
                save_state()
                write_translation(workdir / 'translated.partial.txt', [r for r in records if r.id in translations], translations)
            self.emit({'type':'progress','job':job.id,'node':'translate','page':page,'pages':pages})
        write_translation(context.translations, records, translations)
        stage('translate','done' if cfg.translate.enabled else 'skipped')
        self.boundary()
        if not state['applied']:
            if cfg.translate.enabled:
                apply_translation(job.source, context.manifest, context.translations, context.translated_pptx)
            else:
                shutil.copy2(job.source, context.translated_pptx)
            state['applied'] = True; save_state()
        current = Path(state.get('current_pptx') or context.translated_pptx)
        stage('adjust','running' if cfg.adjust.enabled else 'skipped')
        if cfg.adjust.enabled:
            for page in range(1,pages + 1):
                self.boundary()
                if page in state['adjusted_pages']: continue
                destination = workdir / f'reviewed_{page:03d}.pptx'
                def adjust_page():
                    options = self.current_settings.adjust
                    modifier = PPTXLayoutModifier(self.provider(options,layout=True), PowerPointRenderer(),
                        policy=LayoutPolicy(max_passes=options.max_passes))
                    return modifier.modify(current, destination, reference=job.source,
                        manifest=context.manifest, workdir=workdir/f'page_{page:03d}',pages=[page])
                report = self.unit(adjust_page)
                current = destination
                state['current_pptx'] = str(current)
                state['adjusted_pages'].append(page)
                state['review_reports'].append({'page':page,'status':report['status'],
                    'unresolved':report['unresolved'],'report':str(destination.with_suffix('.layout.json'))})
                save_state()
                # Keep the durable current deck, images and plans; discard duplicate PPTs.
                page_work=workdir/f'page_{page:03d}'
                for duplicate in [page_work/'preview.pptx', *page_work.glob('candidate_*.pptx')]:
                    if duplicate.resolve().parent!=page_work.resolve(): raise ValueError('잘못된 작업 파일 경로')
                    duplicate.unlink(missing_ok=True)
                report.pop('preview_pptx',None)
                destination.with_suffix('.layout.json').write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8')
                (page_work/'progress.json').write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8')
                for previous in workdir.glob('reviewed_*.pptx'):
                    if previous.resolve().parent!=workdir.resolve(): raise ValueError('잘못된 체크포인트 경로')
                    if previous!=current:
                        previous_report=previous.with_suffix('.layout.json')
                        if previous_report.exists():
                            metadata=json.loads(previous_report.read_text(encoding='utf-8'))
                            metadata.pop('output',None); metadata['intermediate_output_removed']=True
                            previous_report.write_text(json.dumps(metadata,ensure_ascii=False,indent=2),encoding='utf-8')
                            old_progress=Path(metadata['final_layout']).parent/'progress.json'
                            if workdir.resolve() in old_progress.resolve().parents:
                                old_progress.write_text(json.dumps(metadata,ensure_ascii=False,indent=2),encoding='utf-8')
                        previous.unlink()
                self.emit({'type':'progress','job':job.id,'node':'adjust','page':page,'pages':pages})
            stage('adjust','done')
        self.boundary()
        stage('save','running')
        target.parent.mkdir(parents=True, exist_ok=True)
        if target.resolve() == job.source.resolve(): raise ValueError('원본 경로는 저장 대상으로 사용할 수 없습니다.')
        # Exclusive creation protects both original and pre-existing results, including races.
        created = False
        try:
            with target.open('xb') as out:
                created = True
                with current.open('rb') as incoming: shutil.copyfileobj(incoming,out)
        except Exception:
            if created: target.unlink(missing_ok=True)
            raise
        needs_review = bool(state['translation_failed']) or any(r['status']=='needs_review' for r in state['review_reports'])
        result = {'output':str(target),'status':'needs_review' if needs_review else 'done','slides':pages,
            'translation_failed_pages':state['translation_failed'],'review_reports':state['review_reports'],
            'checkpoint':str(state_file)}
        with target.with_suffix('.report.json').open('x',encoding='utf-8') as output_report:
            json.dump(result,output_report,ensure_ascii=False,indent=2)
        state['result'] = result; save_state()
        stage('save','done')
        return result
