"""Official Google Gemini CLI, deliberately separate from the legacy agy wrapper."""
import json
from pathlib import Path
import shutil
import subprocess
from gemini_cli import EmptyModelResponseError
from pptx_pipeline.ai import AIResponse
from ..models import Connection, ModelInfo, utc_now
from .base import version_of
from cli_discovery import gemini_candidates,find_node


def find_gemini(executable=None):
    for candidate in gemini_candidates(executable):
        if candidate.stem.casefold() in {'agy','antigravity'}:
            raise FileNotFoundError('Antigravity와 공식 Gemini CLI는 서로 다른 도구입니다. gemini 실행 파일을 선택하세요.')
        if not candidate.is_file(): continue
        if candidate.suffix.lower() not in {'.cmd','.bat','.ps1','.js'}:
            return [str(candidate.resolve())]
        node=find_node()
        script=(candidate if candidate.suffix.lower()=='.js' else
                candidate.parent/'node_modules/@google/gemini-cli/dist/index.js')
        if node and script.is_file(): return [node,str(script.resolve())]
    raise FileNotFoundError('공식 Gemini CLI가 없습니다. 설치 후 실제 실행 파일 경로를 지정하세요.')


class GeminiOfficialProvider:
    def __init__(self, command, model=''):
        self.command, self.model, self.active = command, model, None

    def generate(self, prompt, *, model=None, images=()):
        if images:
            raise RuntimeError('공식 Gemini CLI의 이미지 검토 연동은 아직 검증되지 않았습니다. 조정에는 Codex를 선택하세요.')
        command = self.command + ['--prompt', 'Return only the requested answer.', '--output-format', 'json']
        if model or self.model: command += ['--model', model or self.model]
        self.active = subprocess.Popen(command, stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
            text=True, encoding='utf-8', errors='replace', creationflags=getattr(subprocess,'CREATE_NO_WINDOW',0))
        try:
            stdout, _ = self.active.communicate(prompt, timeout=300)
            if self.active.returncode: raise RuntimeError('Gemini CLI 호출 실패. 연결 관리에서 인증 및 모델을 확인하세요.')
            data = json.loads(stdout)
            if data.get('error'): raise RuntimeError('Gemini CLI가 모델 오류를 반환했습니다.')
            response = data.get('response')
            if not isinstance(response, str) or not response.strip(): raise EmptyModelResponseError('Gemini CLI 빈 응답')
            return AIResponse(response)
        finally:
            if self.active.poll() is None:
                self.active.kill()
                self.active.communicate()
            self.active = None

    def cancel(self):
        if self.active and self.active.poll() is None: self.active.kill()


class GeminiConnectionAdapter:
    def __init__(self, executable=None): self.executable = executable
    def snapshot(self):
        state = Connection('gemini', quota_status='조회 불가',
            quota_error='정밀 계정 잔량·추론 옵션 조회가 검증되지 않았습니다. /stats의 세션 토큰과 구분합니다.')
        try:
            command = find_gemini(self.executable)
            state.executable = self.executable or command[-1]
            state.installed = True
            state.version = version_of(command[0]) if len(command) == 1 else 'npm Gemini CLI · 연결 테스트 필요'
            state.error = '공식 CLI에서 로그인한 후 연결 테스트를 실행하세요.'
        except FileNotFoundError as exc: state.error = str(exc)
        return state

    def test(self):
        command = find_gemini(self.executable)
        result = subprocess.run(command + ['--prompt','Reply with OK only.', '--model','auto','--output-format','json'],
            capture_output=True, text=True, encoding='utf-8', errors='replace', timeout=60,
            creationflags=getattr(subprocess,'CREATE_NO_WINDOW',0))
        if result.returncode: raise RuntimeError('Gemini 연결 테스트 실패. 공식 CLI 인증 상태를 확인하세요.')
        data = json.loads(result.stdout)
        if data.get('error') or not data.get('response'): raise RuntimeError('Gemini 연결 테스트에서 유효한 응답이 없습니다.')
        state = self.snapshot()
        state.connected, state.account = True, '마지막 연결 테스트 기준 · 계정 이름 조회 미지원'
        # auto is a documented CLI mode explicitly verified by this successful request.
        # A stats map is usage history, not an account's accessible-model catalog.
        state.models = [ModelInfo('auto','자동 모델 선택 · 실제 연결 테스트 확인',default=True)]
        state.error = '구체적인 모델 목록·추론 옵션 조회 미지원 · 검증된 자동 모드만 표시'
        state.updated_at = utc_now()
        return state

    def provider(self, settings): return GeminiOfficialProvider(find_gemini(self.executable), settings.model)
    def cancel(self): pass
