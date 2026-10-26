"""Carry run formatting through translation using semantic span markers."""

import re

from pptx_text import A, equation_map

MARKER = re.compile(r"\[\[(/?)STYLE_(\d+)\]\]")
SPAN = re.compile(r"\[\[STYLE_(\d+)\]\](.*?)\[\[/STYLE_\1\]\]", re.DOTALL)


def plain_text(text):
    return MARKER.sub("", text)


def relaxed_styles(source, translated):
    """Accept known, nonnested semantic spans without requiring every source run.

    A translated particle need not have a separate nonempty formatting span.
    Layout/paragraph locations are checked at application, not as translation failure.
    """
    expected = {match[1] for match in SPAN.finditer(source)}
    spans = list(SPAN.finditer(translated))
    identifiers = [match[1] for match in spans]
    if not spans or len(set(identifiers)) != len(identifiers) or not set(identifiers) <= expected:
        return False
    outside = re.sub(r'\[\[EQ_\d+\]\]', '', SPAN.sub('', translated))
    return not outside.strip() and all(match[2] and not MARKER.search(match[2])
        and not re.search(r'[\n\r\v]|\[\[EQ_\d+\]\]', match[2]) for match in spans)


def styled_frame(frame):
    """Return marked prose and original run templates, only for mixed formatting."""
    from lxml import etree

    tokens = equation_map(frame)
    templates = {}
    lines = []
    mixed = False
    for paragraph in frame.paragraphs:
        styles = set()
        parts = []
        previous_style = None
        for child in paragraph._p:
            if child in tokens:
                parts.append(tokens[child])
                previous_style = None
            elif child.tag == "{" + A + "}br":
                parts.append("\v")
                previous_style = None
            elif child.tag in {"{" + A + "}r", "{" + A + "}fld"}:
                text = "".join(n.text or "" for n in child.findall("{" + A + "}t"))
                if not text:
                    continue
                style = child.find("{" + A + "}rPr")
                signature = etree.tostring(style) if style is not None else b""
                styles.add(signature)
                if signature == previous_style:
                    # PowerPoint can split one word into several identical runs.
                    closing = parts[-1].rfind("[[/STYLE_")
                    parts[-1] = parts[-1][:closing] + text + parts[-1][closing:]
                    continue
                identifier = str(len(templates))
                templates[identifier] = child
                parts.append(f"[[STYLE_{identifier}]]{text}[[/STYLE_{identifier}]]")
                previous_style = signature
        mixed |= len(styles) > 1
        lines.append("".join(parts))
    return ("\n".join(lines), templates) if mixed else (None, {})


def validate_styles(source, translated):
    """Reject lost/unknown/nested spans and spans crossing structural boundaries."""
    expected = [m[1] for m in SPAN.finditer(source)]
    found = list(SPAN.finditer(translated))
    if sorted(m[1] for m in found) != sorted(expected):
        raise ValueError("번역문의 서식 표시가 누락·중복되거나 원본과 다릅니다.")
    outside = SPAN.sub("", translated)
    outside = re.sub(r"\[\[EQ_\d+\]\]", "", outside)
    if outside.strip():
        raise ValueError("번역문에 서식 구간 밖 텍스트가 있습니다.")
    for match in found:
        if not match[2] or MARKER.search(match[2]) or re.search(r"[\n\r\v]|\[\[EQ_\d+\]\]", match[2]):
            raise ValueError("번역문의 서식 구간이 비어 있거나 구조 경계를 넘습니다.")
    # Spans can move with meaning, but cannot cross paragraphs, breaks or equations.
    split = r"\[\[EQ_\d+\]\]|[\n\v]"
    before = [sorted(m[1] for m in SPAN.finditer(part)) for part in re.split(split, source)]
    after = [sorted(m[1] for m in SPAN.finditer(part)) for part in re.split(split, translated)]
    if before != after:
        raise ValueError("번역문의 서식 구간이 문단·수식 경계를 넘었습니다.")
