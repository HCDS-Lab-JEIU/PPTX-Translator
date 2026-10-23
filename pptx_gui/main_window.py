from copy import deepcopy
from datetime import datetime
from pathlib import Path
from PySide6.QtCore import Qt, Slot, QSettings, QTimer, QUrl
from PySide6.QtGui import QAction, QDesktopServices
from PySide6.QtWidgets import QMainWindow,QWidget,QVBoxLayout,QHBoxLayout,QLabel,QPushButton,QToolBar,QSplitter,QFileDialog,QMessageBox,QSizePolicy
from .models import AppSettings,FileJob,NODES,LANGUAGES,validate_model
from .connections import ConnectionManager
from .runner import FileRunner
from .widgets.files_panel import FilesPanel
from .widgets.pipeline_panel import PipelinePanel
from .widgets.inspector import Inspector
from .widgets.progress_panel import ProgressPanel
from .widgets.connections_dialog import ConnectionsDialog


class MainWindow(QMainWindow):
    def __init__(self,demo=False,preferences=None):
        super().__init__()
        self.demo=demo; self.settings=AppSettings(); self.jobs=[]; self.selected='translate'
        self.runner=None; self.current_job=None; self.paused=False; self.pausing=False; self.closing=False
        self.preferences=preferences or QSettings('LocalPPTTools','PPTTranslator')
        self.connections=ConnectionManager(demo,self)
        self.setWindowTitle('PPT 번역 도구'+(' — 모의 데이터' if demo else ''))
        self.resize(1440,960); self.setMinimumSize(1100,760)
        self.toolbar=QToolBar(); self.toolbar.setMovable(False); self.addToolBar(self.toolbar)
        branding=QWidget(); brand=QVBoxLayout(branding); brand.setContentsMargins(0,0,15,0); brand.setSpacing(3)
        title=QLabel('PPT 번역 도구'); title.setObjectName('appTitle'); brand.addWidget(title)
        self.direction=QLabel(); self.direction.setObjectName('muted'); brand.addWidget(self.direction); self.toolbar.addWidget(branding)
        spacer=QWidget(); spacer.setSizePolicy(QSizePolicy.Policy.Expanding,QSizePolicy.Policy.Preferred); self.toolbar.addWidget(spacer)
        manage=QPushButton('AI 연결 관리'); manage.clicked.connect(self.manage_connections); self.toolbar.addWidget(manage)
        add=QPushButton('＋ 파일 추가'); add.clicked.connect(self.add_files); self.toolbar.addWidget(add)
        self.run_button=QPushButton('실행 시작'); self.run_button.setObjectName('primary'); self.run_button.clicked.connect(self.start); self.toolbar.addWidget(self.run_button)
        body=QWidget(); box=QVBoxLayout(body); box.setContentsMargins(0,0,0,0); box.setSpacing(0); self.setCentralWidget(body)
        if demo:
            banner=QLabel('  모의 데이터 모드 · 계정·모델·한도·파일은 예시이며 실제 AI 호출·PPT 저장을 하지 않습니다.')
            banner.setStyleSheet('background:#fff4d9;color:#765315;padding:9px;'); banner.setWordWrap(True); box.addWidget(banner)
        self.splitter=QSplitter(Qt.Orientation.Horizontal); box.addWidget(self.splitter,1)
        self.files=FilesPanel(); self.pipeline=PipelinePanel(); self.inspector=Inspector(self.settings,self.connections)
        for panel in (self.files,self.pipeline,self.inspector): self.splitter.addWidget(panel)
        self.splitter.setSizes([240,800,360]); self.splitter.setStretchFactor(1,1)
        self.progress=ProgressPanel(); box.addWidget(self.progress)
        self.statusBar().showMessage('준비 · 파일을 추가하세요.')
        self.status_note=QLabel('모의 데이터 · 실제 AI 호출 없음' if demo else '로컬 계정 · 원본 보존')
        self.statusBar().addPermanentWidget(self.status_note)
        self.files.add_requested.connect(self.add_files); self.files.dropped.connect(self.add_paths)
        self.files.remove_requested.connect(self.remove_job); self.files.move_requested.connect(self.move_job)
        self.files.rerun_requested.connect(self.rerun_job); self.files.open_requested.connect(self.open_result)
        self.files.selected.connect(self.select_file)
        self.pipeline.selected.connect(self.select_node)
        self.inspector.changed.connect(self.settings_changed); self.inspector.connection_requested.connect(self.manage_connections)
        self.inspector.refresh_requested.connect(self.connections.refresh)
        self.connections.changed.connect(self.connection_changed); self.connections.notice.connect(self.log)
        self.progress.pause_requested.connect(self.pause_resume); self.progress.stop_requested.connect(self.stop)
        geometry=self.preferences.value('geometry'); splitter=self.preferences.value('splitter'); logs=self.preferences.value('logs')
        if geometry: self.restoreGeometry(geometry)
        if splitter: self.splitter.restoreState(splitter)
        if logs: self.pipeline.splitter.restoreState(logs)
        if self.preferences.value('logHidden',False,type=bool): self.pipeline.toggle_log()
        if demo:
            self.jobs=[FileJob(Path('07_Java_Advanced.pptx'),slides=60,demo=True),
                       FileJob(Path('08_Inheritance.pptx'),slides=42,demo=True),
                       FileJob(Path('06_Object_Detection.pptx'),status='done',slides=58,demo=True)]
        for node in ('translate','adjust'):
            cfg=getattr(self.settings,node); validate_model(cfg,self.connections.state(cfg.provider))
        self.refresh()
        if not demo: QTimer.singleShot(0,self.connections.refresh_all)

    def job(self,identifier): return next((j for j in self.jobs if j.id==identifier),None)
    def running(self): return self.runner is not None and self.runner.isRunning()
    def log(self,message):
        from .adapters.base import safe_error
        self.pipeline.log.appendPlainText(datetime.now().strftime('%H:%M:%S')+'  '+safe_error(message))
    def refresh(self):
        self.files.render(self.jobs)
        job=self.job(self.current_job) or self.job(self.files.selected_id())
        self.pipeline.render(self.settings,self.selected,job.node_states if job else None)
        a=self.settings.translate
        self.direction.setText(f'{LANGUAGES[a.source_language]} → {LANGUAGES[a.target_language]} · 레이아웃 보존')
        completed=sum(j.status in ('done','needs_review') for j in self.jobs)
        self.progress.count.setText(f'{completed} / {len(self.jobs)} 파일')
        self.progress.bar.setValue(round(completed/max(1,len(self.jobs))*100))

    @Slot()
    def settings_changed(self):
        if self.running(): self.runner.engine.update_settings(self.settings)
        self.refresh()

    @Slot(str)
    def connection_changed(self,key):
        for node in ('translate','adjust'):
            cfg=getattr(self.settings,node)
            if cfg.provider==key:
                from copy import copy
                state=copy(self.connections.state(key))
                if node=='adjust': state.models=[m for m in state.models if m.vision]
                validate_model(cfg,state)
        if self.selected in ('translate','adjust') and getattr(self.settings,self.selected).provider==key: self.inspector.render()
        self.settings_changed()

    @Slot()
    def manage_connections(self):
        ConnectionsDialog(self.connections,auth_locked=self.running() and not self.paused,parent=self).exec()

    @Slot()
    def add_files(self):
        paths,_=QFileDialog.getOpenFileNames(self,'PPTX 파일 추가','','PowerPoint (*.pptx)')
        self.add_paths(paths)
    @Slot(list)
    def add_paths(self,paths):
        known={str(j.source.resolve()).casefold() for j in self.jobs}
        for value in paths:
            path=Path(value).resolve()
            if path.suffix.lower()!='.pptx' or not path.is_file():
                self.log(f'추가 제외 · PPTX 파일이 아님: {path.name}'); continue
            if str(path).casefold() in known:
                self.log(f'중복 파일 제외 · {path.name}'); continue
            self.jobs.append(FileJob(path,demo=self.demo)); known.add(str(path).casefold())
        self.refresh()
    @Slot(str)
    def remove_job(self,identifier):
        job=self.job(identifier)
        if job and not self.running() and job.status not in ('done','needs_review'):
            self.jobs.remove(job); self.refresh()
    @Slot(str,int)
    def move_job(self,identifier,delta):
        job=self.job(identifier)
        if not job or self.running(): return
        index=self.jobs.index(job); target=index+delta
        if 0<=target<len(self.jobs): self.jobs[index],self.jobs[target]=self.jobs[target],self.jobs[index]; self.refresh()
    @Slot(str)
    def rerun_job(self,identifier):
        job=self.job(identifier)
        if job and not self.running():
            index=self.jobs.index(job); self.jobs[index]=FileJob(job.source,demo=job.demo); self.refresh()
            self.log('새 작업으로 재실행 대기 · '+job.source.name)
    @Slot(str,bool)
    def open_result(self,identifier,folder):
        job=self.job(identifier)
        if job and job.output and Path(job.output).is_file():
            target=Path(job.output).parent if folder else Path(job.output)
            QDesktopServices.openUrl(QUrl.fromLocalFile(str(target)))
    @Slot(str)
    def select_file(self,identifier):
        job=self.job(identifier)
        if job:
            self.inspector.source=job.source
            if self.selected=='save': self.inspector.render()
            self.pipeline.render(self.settings,self.selected,job.node_states)
    @Slot(str)
    def select_node(self,node):
        self.selected=node; self.inspector.render(node); self.refresh()

    def preflight(self,jobs):
        if self.demo: return []
        for node in ('translate','adjust'):
            cfg=getattr(self.settings,node)
            if not cfg.enabled: continue
            state=self.connections.state(cfg.provider)
            if not state.connected: raise ValueError(f'{NODES[node][0]}: AI 연결 관리에서 먼저 로그인·연결 테스트를 하세요.')
            model=next((m for m in state.models if m.id==cfg.model),None)
            if not model or (node=='adjust' and not model.vision): raise ValueError(f'{NODES[node][0]}: 사용할 수 있는 모델을 선택하세요.')
            if cfg.effort and cfg.effort not in model.efforts: raise ValueError('모델의 추론 수준이 유효하지 않습니다.')
            if state.exhausted: raise ValueError('계정 한도가 소진됐습니다. 새로고침하거나 다른 연결을 선택하세요.')
        targets=[self.settings.save.target(j.source,self.settings.translate.target_language) for j in jobs]
        if len(set(targets))!=len(targets): raise ValueError('여러 파일의 출력 이름이 같습니다. {stem}을 포함한 이름 규칙을 사용하세요.')
        for target in targets:
            if target.exists() or target.with_suffix('.report.json').exists(): raise ValueError(f'기존 결과 파일이 있습니다. 저장 이름이나 폴더를 바꾸세요:\n{target}')
        return targets

    @Slot()
    def start(self):
        if self.running(): return
        jobs=[j for j in self.jobs if j.status not in ('done','needs_review')]
        if not jobs:
            QMessageBox.information(self,'파일 추가','대기 중인 PPTX 파일을 추가하세요.'); return
        try: targets=self.preflight(jobs)
        except Exception as exc:
            QMessageBox.warning(self,'실행 준비',str(exc)); return
        if targets:
            message='원본을 보존하고 다음 경로에 저장합니다.\n\n'+'\n'.join(map(str,targets[:6]))
            if len(targets)>6: message+=f'\n외 {len(targets)-6}개'
            if QMessageBox.question(self,'저장 경로 확인',message)!=QMessageBox.StandardButton.Yes: return
        executables={k:a.executable for k,a in self.connections.adapters.items() if hasattr(a,'executable')}
        self.runner=FileRunner(jobs,deepcopy(self.settings),executables,self.demo,self)
        self.runner.event.connect(self.on_event); self.runner.finished.connect(self.runner_finished)
        self.paused=self.pausing=False
        self.run_button.setEnabled(False); self.progress.set_running(True); self.progress.title.setText('전체 진행 · 실행 중')
        self.statusBar().showMessage('실행 중 · 설정 변경은 다음 파일부터 적용')
        self.connections.refresh_all(); self.runner.start()

    @Slot(object)
    def on_event(self,event):
        kind=event['type']; job=self.job(event.get('job'))
        if kind=='file_started':
            self.current_job=job.id; job.status='running'; self.log(f'{job.source.name} · 실행 시작')
        elif kind=='node':
            job.node_states[event['node']]=event['status']
            if event['status'] in ('running','done'): self.log(f'{job.source.name} · {NODES[event["node"]][0]} · {event["status"]}')
            self.progress.current.setText(f'현재 파일 · {job.source.name}    현재 단계 · {NODES[event["node"]][0]}')
        elif kind=='analyzed': job.slides=event['pages']; job.workdir=event['workdir']
        elif kind=='progress':
            self.progress.current.setText(f'{job.source.name} · {NODES[event["node"]][0]} · {event["page"]}/{event["pages"]}페이지 처리')
        elif kind=='file_done':
            job.status=event['result']['status']; job.output=event['result']['output']; job.slides=event['result']['slides']
            self.log(f'{job.source.name} · '+('추가 검토 필요' if job.status=='needs_review' else '완료'))
            self.connections.refresh_all()
        elif kind in ('file_failed','file_stopped'):
            job.status='failed' if kind=='file_failed' else 'stopped'
            for key,state in job.node_states.items():
                if state=='running': job.node_states[key]='failed' if kind=='file_failed' else 'paused'
            self.log(f'{job.source.name} · '+event.get('message','중지 · 체크포인트 보존'))
        elif kind in ('paused','quota_paused'):
            self.paused=True; self.pausing=False; self.progress.set_running(True,True)
            quota=kind=='quota_paused' or self.runner.engine.quota_paused
            note='한도 소진 · 초기화 대기 또는 연결 변경 후 재개' if quota else '일시정지 · 페이지 체크포인트 보존'
            self.statusBar().showMessage(note); self.progress.title.setText(note)
            if quota: self.connections.refresh_all()
        elif kind=='warning': self.log(event['message'])
        elif kind=='quota_updated': self.connections.update_limits(event['data'])
        self.refresh()

    @Slot()
    def pause_resume(self):
        if not self.running(): return
        if self.paused:
            self.runner.engine.update_settings(self.settings,quota_resume=self.runner.engine.quota_paused)
            if self.runner.engine.quota_paused:
                self.runner.engine.executables={k:a.executable for k,a in self.connections.adapters.items() if hasattr(a,'executable')}
            self.runner.control.resume_event.set(); self.paused=False; self.pausing=False
            self.progress.set_running(True); self.progress.title.setText('전체 진행 · 실행 중'); self.statusBar().showMessage('실행 재개')
        elif not self.pausing:
            self.runner.control.resume_event.clear(); self.pausing=True
            self.statusBar().showMessage('현재 처리 단위 완료 후 일시정지합니다.')
    @Slot()
    def stop(self):
        if self.running():
            self.runner.control.stop(); self.statusBar().showMessage('중지 요청 · 현재 렌더링 또는 파일 적용 완료를 기다립니다.')
    @Slot()
    def runner_finished(self):
        self.run_button.setEnabled(True); self.progress.set_running(False)
        self.current_job=None; self.paused=self.pausing=False
        self.progress.title.setText('전체 진행 · 실행 종료')
        self.statusBar().showMessage('실행 종료 · 완료/실패 목록과 검토 보고서를 확인하세요.')
        self.refresh()
        if self.closing: QTimer.singleShot(0,self.close)

    def closeEvent(self,event):
        if self.running() or self.connections.tasks:
            if not self.closing and self.running() and QMessageBox.question(self,'종료','작업을 중지하고 종료할까요? 체크포인트는 보존합니다.')!=QMessageBox.StandardButton.Yes:
                event.ignore(); return
            self.closing=True; self.stop(); self.connections.cancel_all(); event.ignore()
            QTimer.singleShot(300,self.close); return
        self.preferences.setValue('geometry',self.saveGeometry()); self.preferences.setValue('splitter',self.splitter.saveState())
        self.preferences.setValue('logs',self.pipeline.splitter.saveState()); self.preferences.setValue('logHidden',self.pipeline.log.isHidden())
        event.accept()
