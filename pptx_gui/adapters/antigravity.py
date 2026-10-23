"""Antigravity is an independent provider, not an alias for official Gemini CLI."""
from datetime import datetime
from pathlib import Path
import subprocess
from gemini_cli import GeminiCLI, GeminiCLIError, CLIResult, ANSI, find_executable
from pptx_pipeline.ai import GeminiCLIProvider
from ..models import Connection,ModelInfo,QuotaBucket,utc_now
from .base import version_of,safe_error


class CancellableAntigravityCLI(GeminiCLI):
    def __init__(self,*args,**kwargs):
        super().__init__(*args,**kwargs); self.active=None
    def _execute(self,command,input_text):
        process=subprocess.Popen(command,stdin=subprocess.PIPE,stdout=subprocess.PIPE,stderr=subprocess.PIPE,
            text=True,encoding='utf-8',errors='replace',cwd=self.cwd,
            creationflags=getattr(subprocess,'CREATE_NO_WINDOW',0))
        self.active=process
        try:
            stdout,stderr=process.communicate(input_text,timeout=self.timeout)
            result=CLIResult(tuple(command),process.returncode,stdout,stderr)
            if result.returncode:
                raise GeminiCLIError('Antigravity CLI 호출 실패: '+ANSI.sub('',stderr or stdout).strip(),result)
            return result
        except subprocess.TimeoutExpired as exc:
            raise GeminiCLIError('Antigravity CLI 응답 시간 초과') from exc
        finally:
            if process.poll() is None:
                process.kill(); process.communicate()
            self.active=None
    def cancel(self):
        process=self.active
        if process and process.poll() is None:
            try: process.kill()
            except OSError: pass


class AntigravityProvider(GeminiCLIProvider):
    """Reuse validated stream parsing and successful image-view evidence checks."""
    def cancel(self): self.cli.cancel()


def parse_models(output):
    models=[]; seen=set()
    for line in ANSI.sub('',output).splitlines():
        if not line.strip(): continue
        fields=line.split('\t',1)
        if len(fields)!=2 or not fields[0].strip() or not fields[1].strip():
            raise ValueError('Antigravity 모델 목록이 예상한 slug/이름 TSV 형식이 아닙니다.')
        slug,label=(s.strip() for s in fields)
        if slug in seen: raise ValueError('Antigravity 모델 ID가 중복되었습니다.')
        seen.add(slug)
        # Only the Gemini 3.8 family has passed this app's actual image-tool probe.
        vision=slug in {'gemini-3.8-flash-low','gemini-3.8-flash-medium','gemini-3.8-flash-high'}
        models.append(ModelInfo(slug,label,(),vision,slug=='gemini-3.8-flash-medium'))
    if not models: raise ValueError('Antigravity에서 모델을 반환하지 않았습니다.')
    return models


class AntigravityConnectionAdapter:
    def __init__(self,executable=None): self.executable=executable; self.active_cli=None
    def snapshot(self):
        state=Connection('antigravity')
        try:
            state.executable=find_executable(self.executable); state.installed=True
            state.version=version_of(state.executable)
            cli=CancellableAntigravityCLI(state.executable,timeout=45,cwd=Path.cwd()); self.active_cli=cli
            result=cli._execute((cli.executable,'models'),'')
            state.models=parse_models(result.stdout)
            usage=cli.get_usage()
            state.connected=True; state.account='CLI 로컬 로그인 · 계정명 조회 미지원'
            state.buckets=[QuotaBucket(f'{q.model_group} · {q.limit_name}',q.remaining_percent,
                datetime.fromisoformat(q.reset_at.replace('Z','+00:00')).timestamp(),
                key=q.model_group+':'+q.limit_name) for q in usage.quotas]
            state.quota_status='소진' if state.exhausted else '정상'
            state.updated_at=utc_now()
        except Exception as exc:
            state.error=safe_error(exc); state.quota_status='조회 불가' if state.installed else '미연결'
        finally: self.active_cli=None
        return state
    def provider(self,settings):
        return AntigravityProvider(cli=CancellableAntigravityCLI(self.executable,timeout=300,cwd=Path.cwd()),model=settings.model)
    def test(self):
        state=self.snapshot()
        if not state.connected or not state.models: return state
        model=next((m for m in state.models if m.default),state.models[0])
        from ..models import NodeSettings
        provider=self.provider(NodeSettings(provider='antigravity',model=model.id))
        self.active_cli=provider.cli
        try: provider.generate('Reply with OK only. Do not use tools.')
        finally: self.active_cli=None
        return state
    def cancel(self):
        if self.active_cli: self.active_cli.cancel()
