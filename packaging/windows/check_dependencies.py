"""Fail the build early if PDF imports or the native renderer are unavailable."""
def main():
    try:
        import pypdfium2 as pdfium
    except ModuleNotFoundError as exc:
        raise SystemExit(f'Missing PDF dependency: {exc.name}. Run build.ps1 without -SkipInstall.') from None
    with pdfium.PdfDocument.new() as document:
        page = document.new_page(80, 60)
        try:
            bitmap = page.render()
            try:
                if (bitmap.width, bitmap.height) != (80, 60):
                    raise RuntimeError('PDF native rendering check failed.')
            finally:
                bitmap.close()
        finally:
            page.close()
    print('PDF Python/native runtime check: OK')


if __name__ == '__main__':
    main()
