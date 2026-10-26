"""Rendering is separate from AI and layout editing; native Office Math is supported."""

from dataclasses import dataclass
import json
import os
from pathlib import Path
import shutil
import subprocess
from typing import Protocol


@dataclass
class RenderResult:
    images: dict[int, Path]
    slides: dict[int, dict]
    pdf: Path | None = None


class PresentationRenderer(Protocol):
    def render(self, source: Path, directory: Path, *, pages=None) -> RenderResult: ...


class PowerPointRenderer:
    def __init__(self, *, timeout=300, width=1600, powershell=None):
        self.timeout, self.width = timeout, width
        self.powershell = powershell or shutil.which("powershell.exe")

    def _run(self, source, directory, *, pages=None, document=False):
        if os.name != "nt" or not self.powershell:
            raise RuntimeError("PowerPointRenderer는 Windows PowerPoint가 필요합니다. 다른 renderer를 주입할 수 있습니다.")
        directory = Path(directory).resolve()
        directory.mkdir(parents=True, exist_ok=True)
        job = directory / "render-job.json"
        job.write_text(json.dumps({"source": str(Path(source).resolve()), "output": str(directory),
                                   "pages": list(pages or []), "width": self.width,
                                   "document_pdf": document}), encoding="utf-8")
        script = Path(__file__).with_name("render_powerpoint.ps1")
        result = subprocess.run([self.powershell, "-NoProfile", "-NonInteractive", "-STA",
            "-ExecutionPolicy", "Bypass", "-File", str(script), "-JobPath", str(job)],
            stdout=subprocess.PIPE, stderr=subprocess.PIPE, encoding="utf-8", errors="replace",
            timeout=self.timeout, creationflags=subprocess.CREATE_NO_WINDOW)
        if result.returncode:
            raise RuntimeError("PowerPoint 렌더링 실패: " + result.stderr.strip())
        manifest = json.loads((directory / "render.json").read_text(encoding="utf-8-sig"))
        return directory, manifest

    def render_document(self, source, directory, *, boundary=None):
        """Export ALL slides to one native PDF before any page repair or AI call."""
        from .layout_pdf import pdf_page_images, pdf_runtime

        pdf_runtime()  # Detect missing Python/native dependencies before opening PowerPoint.
        directory, manifest = self._run(source, directory, document=True)
        pdf = directory / 'slides.pdf'
        if not pdf.is_file() or not pdf.stat().st_size:
            raise RuntimeError('전체 슬라이드 PDF가 생성되지 않았습니다.')
        slides = {int(slide['page']): slide for slide in manifest['slides']}
        images = pdf_page_images(pdf, directory / 'pdf-pages', expected_pages=len(slides),
                                 width=self.width, boundary=boundary)
        for page, image in images.items():
            slides[page]['image'] = image.relative_to(directory).as_posix()
        return RenderResult(images, slides, pdf)

    def render(self, source, directory, *, pages=None):
        directory, manifest = self._run(source, directory, pages=pages)
        images = {int(slide["page"]): directory / slide["image"] for slide in manifest["slides"]}
        if not images or any(not path.is_file() for path in images.values()):
            raise RuntimeError("슬라이드 렌더링 이미지가 누락되었습니다.")
        return RenderResult(images, {int(slide["page"]): slide for slide in manifest["slides"]})
