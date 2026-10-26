from dataclasses import dataclass, field
from datetime import datetime, timezone
from pathlib import Path
import re
from string import Formatter
from uuid import uuid4

PROVIDERS = {"codex": "Codex App Server", "gemini": "Gemini CLI", "antigravity": "Antigravity CLI"}
NODES = {"extract": ("PPT 텍스트 추출", "텍스트·수식·이미지와 위치 정보 추출"),
         "translate": ("번역", "슬라이드 텍스트의 언어 변환"),
         "adjust": ("조정", "번역 결과의 배치·글자 크기 검토"),
         "save": ("PPT 저장", "원본을 보존하고 결과 파일 저장")}
LANGUAGES = {"Korean": "한국어", "English": "영어", "Japanese": "일본어", "Chinese": "중국어"}
STATES = {"ready": "○ 준비", "running": "● 실행", "done": "✓ 완료", "failed": "! 실패",
          "skipped": "— 건너뜀", "paused": "Ⅱ 일시정지", "needs_review": "△ 추가 검토", "stopped": "■ 중지"}


@dataclass
class ModelInfo:
    id: str
    label: str
    efforts: tuple[str, ...] = ()
    vision: bool = False
    default: bool = False


@dataclass
class QuotaBucket:
    name: str
    remaining: float | None = None
    resets_at: float | None = None
    window_minutes: int | None = None
    key: str = ''


@dataclass
class Connection:
    provider: str
    installed: bool = False
    executable: str = ""
    version: str = ""
    connected: bool = False
    account: str = "미연결"
    models: list[ModelInfo] = field(default_factory=list)
    buckets: list[QuotaBucket] = field(default_factory=list)
    quota_status: str = "미연결"
    quota_error: str = ""
    updated_at: datetime | None = None
    error: str = ""
    demo: bool = False

    @property
    def exhausted(self):
        return any(b.remaining is not None and b.remaining <= 0 for b in self.buckets)


@dataclass
class NodeSettings:
    provider: str = "codex"
    model: str = ""
    effort: str = ""
    enabled: bool = True
    prompt: str = ""
    source_language: str = "Korean"
    target_language: str = "English"
    glossary: str = ""
    max_passes: int = 2
    allow_text_compression: bool = True


@dataclass
class SaveSettings:
    directory: str = field(default_factory=lambda: str(Path.cwd() / "output" / "exports"))
    filename: str = "{stem}_{language}.pptx"

    def target(self, source, language):
        source = Path(source).resolve()
        try:
            for _,field,format_spec,conversion in Formatter().parse(self.filename):
                if field is not None and (field not in {'stem','language'} or format_spec or conversion):
                    raise ValueError('지원하지 않는 파일명 규칙')
            name = self.filename.format(stem=source.stem, language=LANGUAGES.get(language, language))
        except (KeyError, ValueError, IndexError) as exc:
            raise ValueError("파일명에는 {stem}, {language}만 사용할 수 있습니다.") from exc
        if not name or name in {".", ".."} or re.search(r'[<>:"/\\|?*\x00-\x1f]', name):
            raise ValueError("파일 이름에 사용할 수 없는 문자가 있습니다.")
        if not name.lower().endswith('.pptx'):
            name += '.pptx'
        if name.rstrip(' .')!=name or name.split('.')[0].upper() in {'CON','PRN','AUX','NUL',*[f'COM{i}' for i in range(1,10)],*[f'LPT{i}' for i in range(1,10)]}:
            raise ValueError('운영체제에서 사용할 수 없는 파일 이름입니다.')
        if not self.directory.strip():
            raise ValueError("저장 폴더를 지정하세요.")
        target = Path(self.directory).expanduser().resolve() / name
        if target == source:
            raise ValueError("원본 PPT를 덮어쓸 수 없습니다.")
        return target


@dataclass
class AppSettings:
    translate: NodeSettings = field(default_factory=lambda: NodeSettings(effort="medium"))
    adjust: NodeSettings = field(default_factory=lambda: NodeSettings(effort="low", max_passes=1))
    save: SaveSettings = field(default_factory=SaveSettings)


@dataclass
class FileJob:
    source: Path
    id: str = field(default_factory=lambda: uuid4().hex)
    status: str = "ready"
    slides: int | None = None
    output: str = ""
    workdir: str = ""
    demo: bool = False
    node_states: dict = field(default_factory=lambda: {n: "ready" for n in NODES})


def validate_model(settings, connection):
    """Clear invalid values on provider/model refresh; never invent capabilities."""
    models = connection.models
    selected = next((m for m in models if m.id == settings.model), None)
    if selected is None:
        selected = next((m for m in models if m.default), models[0] if models else None)
        settings.model = selected.id if selected else ""
    if not selected or settings.effort not in selected.efforts:
        settings.effort = selected.efforts[0] if selected and selected.efforts else ""
    return selected


def utc_now(): return datetime.now(timezone.utc)
