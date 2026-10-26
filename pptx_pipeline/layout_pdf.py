"""Read the complete exported slide PDF for document-wide repair planning."""
from pathlib import Path
from threading import Lock

# PDFium must not be called concurrently, even for separate documents.
PDFIUM_LOCK = Lock()


def pdf_runtime():
    import pypdfium2 as pdfium
    return pdfium


def pdf_page_images(source, directory, *, expected_pages, width=1280, boundary=None):
    pdfium = pdf_runtime()

    directory = Path(directory)
    directory.mkdir(parents=True, exist_ok=True)
    boundary = boundary or (lambda: None)
    images = {}
    with PDFIUM_LOCK, pdfium.PdfDocument(str(source)) as document:
        if len(document) != expected_pages:
            raise ValueError('전체 PDF의 페이지 수가 슬라이드 수와 다릅니다.')
        for index in range(len(document)):
            boundary()
            page = document[index]
            try:
                bitmap = page.render(scale=width / page.get_width())
                try:
                    image = bitmap.to_pil()
                    try:
                        path = directory / f'slide_{index + 1:03d}.jpg'
                        with image.convert('RGB') as rgb:
                            rgb.save(path, quality=85)
                        images[index + 1] = path
                    finally:
                        image.close()
                finally:
                    bitmap.close()
            finally:
                page.close()
    return images
