"""Summarize existing review findings without additional model calls."""


def review_pages(result):
    pages = set(result.get('translation_failed_pages', []))
    pages.update(item['page'] for item in result.get('application_report', {}).get('kept_original', []))
    pages.update(item['page'] for item in result.get('preserved_unsupported_shapes', []))
    pages.update(item['page'] for item in result.get('review_reports', [])
                 if item.get('status') == 'needs_review')
    return sorted(pages)


def review_message(pages):
    return '검토 필요 페이지: ' + ', '.join(map(str, pages))
