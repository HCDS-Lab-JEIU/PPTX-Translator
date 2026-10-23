"""Backend-neutral AI contract and registry for local CLI/server adapters."""

from dataclasses import dataclass
from pathlib import Path
import json
import os
from typing import Callable, Protocol, Sequence

from gemini_cli import GeminiCLI
from translate_texts import DEFAULT_MODEL


@dataclass(frozen=True)
class AIResponse:
    text: str
    images_read: tuple[Path, ...] = ()

    @property
    def stdout(self):
        """Compatibility with the original text translator's CLI result contract."""
        return self.text


class AIProvider(Protocol):
    def generate(self, prompt: str, *, model: str | None = None,
                 images: Sequence[Path] = ()) -> AIResponse: ...


def strings(value):
    if isinstance(value, str):
        yield value
    elif isinstance(value, dict):
        for item in value.values():
            yield from strings(item)
    elif isinstance(value, (tuple, list)):
        for item in value:
            yield from strings(item)


def image_read_evidence(events, images, cwd):
    """Require a successful CLI image/file-read tool event, not model self-report."""
    found = set()
    expected = {os.path.normcase(str(Path(path).resolve())): Path(path).resolve() for path in images}
    for event in events:
        step = event.get("step_update", {})
        info = step.get("tool_info", {})
        name = str(step.get("tool_name") or info.get("name") or "").lower()
        if not any(word in name for word in ("read_file", "view_file", "view_image", "read_image", "read_media")):
            continue
        if info.get("error") or step.get("state") != "DONE":
            continue
        for value in strings(info.get("parameters", {})):
            try:
                path = Path(value)
                if not path.is_absolute():
                    path = cwd / path
                key = os.path.normcase(str(path.resolve()))
            except (ValueError, OSError):
                continue
            if key in expected:
                found.add(expected[key])
    return tuple(sorted(found))


class GeminiCLIProvider:
    def __init__(self, *, cli=None, model=DEFAULT_MODEL, executable=None, timeout=300, cwd=None):
        self.cli = cli or GeminiCLI(executable, timeout=timeout, cwd=cwd)
        self.model = model

    def generate(self, prompt, *, model=None, images=()):
        images = tuple(Path(path).resolve() for path in images)
        for path in images:
            if not path.is_file():
                raise FileNotFoundError(str(path))
        if images:
            prompt = ("First use the image/file viewing tool to visually read EVERY PNG below. "
                      "Read actual pixels, not just filenames or file metadata. Do not modify any files. "
                      "Do not invoke other agents. Then answer the task as JSON only.\n"
                      "Image paths: " + json.dumps([str(path) for path in images], ensure_ascii=False) +
                      "\n\n" + prompt)
        result = self.cli.generate(prompt, model=model or self.model)
        cwd = Path(self.cli.cwd or Path.cwd())
        evidence = image_read_evidence(result.events, images, cwd) if images else ()
        if set(evidence) != set(images):
            raise RuntimeError("Gemini CLI의 이미지 열람 기록을 확인하지 못했습니다. 시각 검증 없이 수정하지 않습니다.")
        return AIResponse(result.stdout, evidence)


class ProviderRegistry:
    def __init__(self):
        from .codex_provider import CodexAppServerProvider
        self._factories: dict[str, Callable] = {"gemini-cli": GeminiCLIProvider,
                                               "codex-app-server": CodexAppServerProvider}

    def register(self, name, factory):
        if not name or not callable(factory):
            raise ValueError("provider 이름과 생성 함수가 필요합니다.")
        self._factories[name] = factory

    def create(self, name, **options):
        if name not in self._factories:
            raise ValueError(f"등록되지 않은 AI provider: {name}. 사용 가능: {', '.join(self._factories)}")
        return self._factories[name](**options)


def parse_json_response(text):
    text = text.strip()
    if text.startswith("```json\n") and text.endswith("```"):
        text = text[8:-3]
    elif text.startswith("```\n") and text.endswith("```"):
        text = text[4:-3]
    return json.loads(text)
