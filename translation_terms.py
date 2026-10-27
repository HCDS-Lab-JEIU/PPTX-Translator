"""Reuse explicit English glosses without removing conditions, units or formulas."""
import re

from pptx_styles import MARKER, SPAN, plain_text

HANGUL = re.compile(r'[가-힣]')
PAIR = re.compile(r'(?P<ko>[가-힣ㄱ-ㅎㅏ-ㅣ ·ㆍ\-]+?)[ \t]*[（(][ \t]*(?P<en>[^()（）\n]+?)[ \t]*[)）]')
PAREN = re.compile(r'[ \t]*[（(][ \t]*([^()（）\n]+?)[ \t]*[)）]')
UNITS = {'v','a','w','hz','ohm','db','mv','ma','kw','khz','mhz','ghz','ms','mm','cm','kg'}


def english_target(language):
    return language.strip().lower().replace('_','-') in {'english','en','en-us','en-gb','영어'}


def english_term(text):
    text = text.strip()
    if not re.fullmatch(r"[A-Za-z][A-Za-z &'’.,/\-]*", text) or len(text) < 2:
        return False
    if text.lower() in UNITS:
        return False
    return not re.match(r'(?i)(if|when|unless|under|above|below|at|for|with|without|depending|assuming|where|see|note|example|use|only|not|no)\b', text)


def keep_ranges(text, ranges):
    """Project plain-text selections onto existing style spans; never invent a style."""
    output, position, raw = [], 0, 0
    for marker in list(MARKER.finditer(text)) + [None]:
        stop = marker.start() if marker is not None else len(text)
        piece = text[raw:stop]
        for start, end in ranges:
            left, right = max(start-position, 0), min(end-position, len(piece))
            if left < right:
                output.append(piece[left:right])
        position += len(piece)
        if marker is not None:
            output.append(marker[0])
            raw = marker.end()
    return SPAN.sub(lambda m: m[0] if m[2].strip() else m[2], ''.join(output))


def reuse_english_label(record):
    """Bypass the model only for complete, simple Korean (English) label records."""
    text = record.text
    ranges, position, found = [], 0, False
    for line in text.splitlines(keepends=True):
        content = line.rstrip('\r\n\v')
        if not content.strip():
            ranges.append((position, position + len(line)))
        else:
            match = PAIR.fullmatch(content.strip())
            if not match or not english_term(match['en']):
                return None
            korean = match['ko'].strip()
            if (len(korean) > 60 or len(korean.split()) > 5
                    or re.search(r'(한다|된다|하세요|하시오|이면|경우|동안|조건|가정|이상|이하|미만|초과)$', korean)):
                return None
            start = position + len(content) - len(content.lstrip()) + match.start('en')
            ranges.append((start, start + len(match['en'])))
            ranges.append((position + len(content), position + len(line)))
            found = True
        position += len(line)
    if not found:
        return None
    marked = record.styled_text or text
    if plain_text(marked) != text:
        return None
    return keep_ranges(marked, ranges)


def remove_duplicate_gloss(record, translated):
    """Only collapse literal duplicate English glosses already present in the source."""
    if not isinstance(translated, str) or not HANGUL.search(record.text):
        return translated
    glosses = {' '.join(m['en'].split()).casefold() for m in PAIR.finditer(record.text) if english_term(m['en'])}
    if not glosses:
        return translated
    plain = plain_text(translated)
    removals = []
    for match in PAREN.finditer(plain):
        term = match[1].strip()
        if ' '.join(term.split()).casefold() not in glosses:
            continue
        pattern = r'(?<![\w])' + r'[ \t]+'.join(re.escape(word) for word in term.split()) + r'$'
        if re.search(pattern, plain[:match.start()], re.IGNORECASE):
            removals.append((match.start(), match.end()))
    if not removals:
        return translated
    kept, start = [], 0
    for left, right in removals:
        kept.append((start,left))
        start = right
    kept.append((start,len(plain)))
    return keep_ranges(translated, kept)
