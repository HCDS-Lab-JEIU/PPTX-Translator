"""Shared per-tool account/quota state, updated without blocking the GUI thread."""
from PySide6.QtCore import QObject, QThread, Signal, Slot, QTimer
from PySide6.QtGui import QDesktopServices
from PySide6.QtCore import QUrl
from .models import Connection, ModelInfo, QuotaBucket, utc_now
from .adapters.codex import CodexConnectionAdapter,parse_limits
from .adapters.gemini import GeminiConnectionAdapter
from .adapters.base import safe_error
from .adapters.antigravity import AntigravityConnectionAdapter


class ConnectionTask(QThread):
    result = Signal(str, object)
    failed = Signal(str, str)
    browser = Signal(str)
    def __init__(self, key, adapter, action, parent=None):
        super().__init__(parent)
        self.key, self.adapter, self.action = key, adapter, action
    def run(self):
        try:
            if self.action == 'login': self.adapter.login(self.browser.emit)
            if self.action == 'logout': self.adapter.logout()
            state = self.adapter.test() if self.action == 'test' and hasattr(self.adapter,'test') else self.adapter.snapshot()
            self.result.emit(self.key, state)
        except Exception as exc:
            self.failed.emit(self.key, safe_error(exc))


class ConnectionManager(QObject):
    changed = Signal(str)
    notice = Signal(str)
    def __init__(self, demo=False, parent=None):
        super().__init__(parent)
        self.demo = demo
        self.adapters = {'codex': CodexConnectionAdapter(), 'gemini': GeminiConnectionAdapter(), 'antigravity': AntigravityConnectionAdapter()}
        self.states = {key: Connection(key) for key in self.adapters}
        self.tasks = {}
        if demo:
            self.states['codex'] = Connection('codex', installed=True, executable='모의 실행 경로', version='모의 버전',
                connected=True, account='demo@example.test · 모의 계정', demo=True,
                models=[ModelInfo('demo-codex', '모의 모델 A', ('low','medium','high'), True, True)],
                buckets=[QuotaBucket('모의 계정 · 300분 구간',75), QuotaBucket('모의 계정 · 10080분 구간',90)],
                quota_status='정상', updated_at=utc_now())
            self.states['gemini'] = Connection('gemini', installed=True, connected=True, demo=True,
                account='demo@example.test · 모의 계정', models=[ModelInfo('demo-gemini','모의 모델 B')],
                quota_status='조회 불가', quota_error='모의 상태 · 지원 여부 미확인')
            self.states['antigravity']=Connection('antigravity',installed=True,connected=True,demo=True,
                account='모의 CLI 연결',models=[ModelInfo('demo-agy','모의 Antigravity 모델',vision=True)],
                quota_status='조회 불가',quota_error='모의 데이터 · 실제 조회 없음')
        self.timer = QTimer(self)
        self.timer.setInterval(120000)
        self.timer.timeout.connect(self.refresh_all)
        if not demo: self.timer.start()

    def state(self, key): return self.states[key]

    def update_limits(self,data):
        state=self.states['codex']
        new=parse_limits(data)
        groups={b.key.split(':',1)[0] for b in new}
        state.buckets=[b for b in state.buckets if b.key.split(':',1)[0] not in groups]+new
        state.updated_at=utc_now(); state.quota_status='소진' if state.exhausted else '정상' if state.buckets else '조회 불가'
        self.changed.emit('codex')

    @Slot()
    def refresh_all(self):
        for key in self.adapters: self.refresh(key)

    def refresh(self, key, action='refresh'):
        if self.demo:
            self.states[key].updated_at = utc_now()
            self.changed.emit(key)
            return
        if key in self.tasks: return
        self.states[key].quota_status = '조회 중'
        self.changed.emit(key)
        task = ConnectionTask(key, self.adapters[key], action, self)
        self.tasks[key] = task
        task.result.connect(self._result)
        task.failed.connect(self._failed)
        task.browser.connect(self._browser)
        task.finished.connect(self._finished)
        task.start()

    @Slot(str)
    def _browser(self, url): QDesktopServices.openUrl(QUrl(url))

    @Slot(str, object)
    def _result(self, key, state):
        old = self.states[key]
        if key == 'gemini' and old.connected and state.installed and not state.connected and not state.error.startswith('공식 Gemini CLI가 없습니다'):
            state.connected, state.account, state.models = old.connected, old.account, old.models
            state.updated_at = old.updated_at
            state.error = '마지막 연결 테스트 기준 · 계정 상태 직접 조회 미지원'
        if state.quota_status == '조회 불가' and old.buckets and state.connected and state.account==old.account:
            state.buckets, state.updated_at = old.buckets, old.updated_at
            state.quota_status = '오래된 정보'
        self.states[key] = state
        self.changed.emit(key)

    @Slot(str, str)
    def _failed(self, key, message):
        state = self.states[key]
        state.error = message
        state.quota_status = '오래된 정보' if state.buckets else '조회 불가'
        self.changed.emit(key)
        self.notice.emit(f'{key}: {message}')

    @Slot()
    def _finished(self):
        task = self.sender()
        if self.tasks.get(task.key) is task: self.tasks.pop(task.key)
        task.deleteLater()
        self.changed.emit(task.key)

    def cancel_all(self):
        self.timer.stop()
        for key in self.tasks: self.adapters[key].cancel()
