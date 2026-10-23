"""PPT analysis → AI translation → original copy/application → AI/Python layout."""

import argparse
from pathlib import Path
import sys

from pptx_pipeline.ai import ProviderRegistry
from pptx_pipeline.layout import LayoutPolicy, PPTXLayoutModifier
from pptx_pipeline.rendering import PowerPointRenderer
from pptx_pipeline.stages import (PipelineContext, PresentationPipeline, AnalysisStage,
                                  TranslationStage, ApplicationStage, ModificationStage)
from translate_texts import DEFAULT_MODEL


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("source", type=Path)
    parser.add_argument("-o", "--output-dir", type=Path, default=Path("output"))
    names = ["analyze", "translate", "apply", "modify"]
    parser.add_argument("--start-at", choices=names, default="analyze")
    parser.add_argument("--stop-after", choices=names, default="modify")
    parser.add_argument("--provider", help="지정하면 번역·검토 공통 provider를 변경")
    parser.add_argument("--translation-provider")
    parser.add_argument("--layout-provider")
    parser.add_argument("--model", help="지정하면 번역·검토 공통 모델을 변경")
    parser.add_argument("--translation-model")
    parser.add_argument("--layout-model")
    parser.add_argument("--executable", help="현재 Gemini backend의 agy 실행 파일 경로")
    parser.add_argument("--codex-executable", help="Codex App Server의 codex.exe 경로")
    parser.add_argument("--layout-effort", default="low", choices=["low", "medium", "high", "xhigh"])
    parser.add_argument("--timeout", type=float, default=300)
    parser.add_argument("--max-passes", type=int, default=2, help="AI 레이아웃 적용 횟수. 마지막 재검토는 별도")
    parser.add_argument("--pages", help="Modify 검토 페이지: 1,2,12 (기본: 전체)")
    parser.add_argument("--pages-per-request", type=int, default=1)
    args = parser.parse_args(argv)
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            stream.reconfigure(encoding="utf-8")
    try:
        if names.index(args.stop_after) < names.index(args.start_at):
            raise ValueError("종료 단계가 시작 단계보다 앞에 있습니다.")
        selected = names[names.index(args.start_at):names.index(args.stop_after) + 1]
        registry = ProviderRegistry()
        def provider(name, model):
            options = dict(model=model, timeout=args.timeout, cwd=Path.cwd())
            if name == "codex-app-server":
                options.update(executable=args.codex_executable, effort=args.layout_effort,
                               trace_dir=args.output_dir / "codex-events")
            else:
                options["executable"] = args.executable
            return registry.create(name, **options)
        translation_name = args.translation_provider or args.provider or "gemini-cli"
        layout_name = args.layout_provider or args.provider or "codex-app-server"
        translation_ai = (provider(translation_name, args.translation_model or args.model or
                                   (DEFAULT_MODEL if translation_name == "gemini-cli" else None))
                          if "translate" in selected else None)
        layout_ai = (provider(layout_name, args.layout_model or args.model or
                             ("gemini-3.8-flash-low" if layout_name == "gemini-cli" else None))
                     if "modify" in selected else None)
        modifier = PPTXLayoutModifier(layout_ai, PowerPointRenderer(timeout=args.timeout),
            policy=LayoutPolicy(max_passes=args.max_passes), pages_per_request=args.pages_per_request)
        pages = [int(value.strip()) for value in args.pages.split(",")] if args.pages else None
        pipeline = PresentationPipeline([AnalysisStage(), TranslationStage(translation_ai),
            ApplicationStage(), ModificationStage(modifier, pages=pages)])
        context = PipelineContext(args.source, args.output_dir,
            progress=lambda message: print(message, file=sys.stderr, flush=True))
        print(f"결과 저장 폴더: {context.output_dir}", file=sys.stderr, flush=True)
        result = pipeline.run(context, start_at=args.start_at, stop_after=args.stop_after)
    except Exception as exc:
        print(f"파이프라인 실패: {exc}", file=sys.stderr)
        return 1
    print(f"파이프라인 상태: {result['status']}")
    if "modify" in result["stages"]:
        print(result["stages"]["modify"]["output"])
    return 2 if result["status"] == "needs_review" else 0


if __name__ == "__main__":
    raise SystemExit(main())
