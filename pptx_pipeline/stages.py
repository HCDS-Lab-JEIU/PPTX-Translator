"""Component-based pipeline. Stages and AI/rendering dependencies are injectable."""

from dataclasses import dataclass, field
import json
from pathlib import Path
from typing import Callable, Protocol

import main as analyzer
from apply_translation import apply_translation
from translate_texts import PPTXTextTranslator


@dataclass
class PipelineContext:
    original: Path
    output_dir: Path = Path("output")
    progress: Callable[[str], None] | None = None
    reports: dict = field(default_factory=dict)

    def __post_init__(self):
        self.original, self.output_dir = Path(self.original).resolve(), Path(self.output_dir).resolve()
        if not self.original.is_file():
            raise FileNotFoundError(str(self.original))

    @property
    def manifest(self): return self.output_dir / "content.json"

    @property
    def texts(self): return self.output_dir / "texts.txt"

    @property
    def translations(self): return self.output_dir / "translated.txt"

    @property
    def translated_pptx(self): return self.output_dir / "translated.pptx"

    @property
    def modified_pptx(self): return self.output_dir / "modified.pptx"

    def notify(self, message):
        if self.progress:
            self.progress(message)


class PipelineStage(Protocol):
    name: str
    def run(self, context: PipelineContext) -> dict: ...


class AnalysisStage:
    name = "analyze"

    def run(self, context):
        document = analyzer.extract_pptx(context.original, context.output_dir)
        context.manifest.write_text(json.dumps(document, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
        context.texts.write_text(analyzer.format_text(document), encoding="utf-8")
        return {"pages": len(document["slides"]), "manifest": str(context.manifest), "texts": str(context.texts)}


class TranslationStage:
    name = "translate"

    def __init__(self, provider=None, *, translator=None):
        self.translator = translator or PPTXTextTranslator(provider, model=None)

    def run(self, context):
        return self.translator.translate_file(context.texts, context.translations, progress=context.progress)


class ApplicationStage:
    name = "apply"

    def __init__(self, applier=apply_translation):
        self.applier = applier

    def run(self, context):
        report = self.applier(context.original, context.manifest, context.translations, context.translated_pptx)
        context.translated_pptx.with_suffix(".report.json").write_text(
            json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
        return report


class ModificationStage:
    name = "modify"

    def __init__(self, modifier, *, pages=None):
        self.modifier, self.pages = modifier, pages

    def run(self, context):
        return self.modifier.modify(context.translated_pptx, context.modified_pptx,
            reference=context.original, manifest=context.manifest,
            workdir=context.output_dir / "layout", pages=self.pages, progress=context.progress)


class PresentationPipeline:
    def __init__(self, stages):
        self.stages = tuple(stages)
        names = [stage.name for stage in self.stages]
        if not names or len(names) != len(set(names)):
            raise ValueError("파이프라인 단계 이름은 비어 있거나 중복될 수 없습니다.")

    def run(self, context, *, start_at=None, stop_after=None):
        names = [stage.name for stage in self.stages]
        first = names.index(start_at) if start_at else 0
        last = names.index(stop_after) if stop_after else len(names) - 1
        if last < first:
            raise ValueError("종료 단계가 시작 단계보다 앞에 있습니다.")
        context.output_dir.mkdir(parents=True, exist_ok=True)
        if first and context.manifest.exists():
            document = json.loads(context.manifest.read_text(encoding="utf-8-sig"))
            if document.get("source_sha256") != analyzer.file_sha256(context.original):
                raise ValueError("재개할 추출 데이터가 현재 원본과 다릅니다. analyze부터 실행하세요.")
        state = {"source": str(context.original), "status": "running", "stages": context.reports}
        state_file = context.output_dir / "pipeline.json"
        try:
            for stage in self.stages[first:last + 1]:
                context.notify(f"[{stage.name}] 시작")
                state["active_stage"] = stage.name
                state_file.write_text(json.dumps(state, ensure_ascii=False, indent=2), encoding="utf-8")
                context.reports[stage.name] = stage.run(context)
                context.notify(f"[{stage.name}] 완료")
            state["status"] = ("needs_review" if any(report.get("status") == "needs_review"
                                for report in context.reports.values()) else "completed")
            state.pop("active_stage", None)
        except Exception as exc:
            state.update(status="failed", error=str(exc))
            result = getattr(exc, "result", None)
            if result is not None:
                trace = context.output_dir / "failed-ai-response.jsonl"
                trace.write_text(result.stdout, encoding="utf-8")
                state["diagnostic_response"] = str(trace)
            raise
        finally:
            state_file.write_text(json.dumps(state, ensure_ascii=False, indent=2), encoding="utf-8")
        return state
