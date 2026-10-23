from contextlib import contextmanager
from pathlib import Path
from pptx_pipeline.codex_provider import StdioRPC, find_codex, CodexAppServerProvider
from ..models import Connection, ModelInfo, QuotaBucket, utc_now
from .base import safe_error, version_of


def parse_limits(data):
    groups = data.get('rateLimitsByLimitId')
    if not groups:
        legacy = data.get('rateLimits')
        groups = {legacy.get('limitId') or 'account': legacy} if legacy else {}
    buckets = []
    for key, group in groups.items():
        name = group.get('limitName') or key
        for field in ('primary', 'secondary'):
            window = group.get(field)
            if not window: continue
            used = window.get('usedPercent')
            remaining = max(0, min(100, 100 - used)) if isinstance(used, (int, float)) else None
            duration = window.get('windowDurationMins')
            label = f'{name} · {duration}분 구간' if duration is not None else f'{name} · {field}'
            buckets.append(QuotaBucket(label, remaining, window.get('resetsAt'), duration, f'{key}:{field}'))
    return buckets


class CodexConnectionAdapter:
    def __init__(self, executable=None):
        self.executable = executable
        self.active = None

    @contextmanager
    def session(self, timeout=45):
        executable = find_codex(self.executable)
        rpc = StdioRPC((executable, 'app-server', '--stdio'), cwd=Path.cwd(), timeout=timeout)
        self.active = rpc
        try:
            rpc.request('initialize', {'clientInfo': {'name': 'pptx_desktop', 'version': '1.0.0'}})
            rpc.send({'method': 'initialized', 'params': {}})
            yield rpc
        finally:
            rpc.close()
            self.active = None

    def snapshot(self):
        state = Connection('codex')
        try:
            state.executable = find_codex(self.executable)
            state.installed = True
            state.version = version_of(state.executable)
            with self.session() as rpc:
                account = rpc.request('account/read', {'refreshToken': False}).get('account')
                if not account:
                    return state
                state.connected = True
                state.account = account.get('email') or account.get('type') or '로컬 Codex 계정'
                state.account += f" · {account.get('planType')}" if account.get('planType') else ''
                models = []
                cursor = None
                while True:
                    response = rpc.request('model/list', {'cursor': cursor} if cursor else {})
                    for m in response.get('data', []):
                        efforts = tuple(e['reasoningEffort'] for e in m.get('supportedReasoningEfforts', [])
                                        if e['reasoningEffort'] in {'low','medium','high','xhigh','max'})
                        models.append(ModelInfo(m['model'], m.get('displayName') or m['model'], efforts,
                            'image' in m.get('inputModalities', []), m.get('isDefault', False)))
                    cursor = response.get('nextCursor')
                    if not cursor: break
                state.models = models
                try:
                    limits = rpc.request('account/rateLimits/read', {})
                    state.buckets = parse_limits(limits)
                    state.quota_status = '소진' if state.exhausted else '정상' if state.buckets else '조회 불가'
                    state.updated_at = utc_now()
                except Exception:
                    state.quota_status = '조회 불가'
                    state.quota_error = '현재 인증 방식에서 계정 잔량을 제공하지 않거나 조회에 실패했습니다.'
        except Exception as exc:
            state.error = safe_error(exc)
            state.quota_status = '조회 불가' if state.installed else '미연결'
        return state

    def login(self, open_url):
        with self.session(timeout=300) as rpc:
            result = rpc.request('account/login/start', {'type': 'chatgpt'})
            open_url(result['authUrl'])  # URL stays in memory and is never sent to the log.
            cursor = 0
            while True:
                if cursor == len(rpc.events): rpc.receive()
                event = rpc.events[cursor]
                cursor += 1
                params = event.get('params', {})
                if event.get('method') == 'account/login/completed' and params.get('loginId') == result['loginId']:
                    if not params.get('success'): raise RuntimeError('Codex 로그인 실패 또는 취소')
                    return

    def logout(self):
        with self.session() as rpc: rpc.request('account/logout', {})

    def provider(self, settings):
        return CodexAppServerProvider(executable=self.executable, model=settings.model,
            effort=settings.effort or None, timeout=300, trace_dir=None)

    def cancel(self):
        if self.active and self.active.process.poll() is None: self.active.process.kill()
