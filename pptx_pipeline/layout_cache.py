"""Reuse durable render outputs when resuming the same deck and renderer."""
import json
from pathlib import Path

from main import file_sha256
from .rendering import RenderResult


def render_cached(renderer, source, directory, pages, *, document=False, boundary=None):
    directory = Path(directory)
    key = {'source_sha256': file_sha256(source), 'pages': pages,
           'renderer': type(renderer).__name__, 'width': getattr(renderer, 'width', None),
           'mode':'document-pdf-v1' if document else 'slides'}
    cache = directory / 'render-cache.json'
    if cache.exists():
        try:
            saved = json.loads(cache.read_text(encoding='utf-8'))
            if saved['key'] == key:
                images = {int(page): Path(path) for page, path in saved['images'].items()}
                slides = {int(page): item for page, item in saved['slides'].items()}
                if set(images) == set(pages) == set(slides) and all(path.is_file() for path in images.values()):
                    pdf = Path(saved['pdf']) if saved.get('pdf') else None
                    if not document or (pdf and pdf.is_file() and saved.get('pdf_sha256') == file_sha256(pdf)):
                        return RenderResult(images, slides, pdf)
        except (ValueError, KeyError, TypeError):
            pass
    result = (renderer.render_document(source, directory, boundary=boundary) if document else
              renderer.render(source, directory, pages=pages))
    if document and (not result.pdf or not result.pdf.is_file()):
        raise RuntimeError('전체 슬라이드 PDF가 없는 결과는 일괄 예측에 사용할 수 없습니다.')
    if set(result.images) != set(pages) or set(result.slides) != set(pages):
        raise RuntimeError('렌더러가 요청한 페이지를 모두 반환하지 않았습니다.')
    cache.write_text(json.dumps({'key': key, 'images': {p: str(path.resolve()) for p, path in result.images.items()},
                                 'slides': result.slides, 'pdf':str(result.pdf.resolve()) if result.pdf else None,
                                 'pdf_sha256':file_sha256(result.pdf) if result.pdf else None}, ensure_ascii=False, indent=2), encoding='utf-8')
    return result
