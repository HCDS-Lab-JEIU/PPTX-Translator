"""Read DrawingML prose while protecting native Office Math objects."""

import re

A = "http://schemas.openxmlformats.org/drawingml/2006/main"
M = "http://schemas.openxmlformats.org/officeDocument/2006/math"
MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
P = "http://schemas.openxmlformats.org/presentationml/2006/main"
EQUATION = re.compile(r"\[\[EQ_\d+\]\]")


def is_math(element):
    return any(node.tag.startswith("{" + M + "}") for node in element.iter()
               if isinstance(node.tag, str))


def equation_roots(frame):
    return [child for paragraph in frame.paragraphs for child in paragraph._p if is_math(child)]


def math_text_nodes(frame):
    return [node for root in equation_roots(frame) for node in root.iter("{" + M + "}t")]


def paragraph_text(paragraph, tokens):
    parts = []
    for child in paragraph:
        if child in tokens:
            parts.append(tokens[child])
        elif child.tag in {"{" + A + "}r", "{" + A + "}fld"}:
            parts.extend(node.text or "" for node in child.findall("{" + A + "}t"))
        elif child.tag == "{" + A + "}br":
            parts.append("\v")
    return "".join(parts)


def equation_map(frame):
    return {root: f"[[EQ_{index}]]" for index, root in enumerate(equation_roots(frame), 1)}


def frame_text(frame):
    tokens = equation_map(frame)
    return "\n".join(paragraph_text(paragraph._p, tokens) for paragraph in frame.paragraphs)


def equation_lines(text):
    return [EQUATION.findall(line) for line in text.replace("\v", "\n").split("\n")]


def shape_elements(tree):
    """Select one supported modern branch, avoiding duplicate Fallback content."""
    tags = {"{" + P + "}" + name for name in ("sp", "grpSp", "pic", "graphicFrame", "cxnSp", "contentPart")}
    for child in tree:
        if child.tag in tags:
            yield child
        elif child.tag == "{" + MC + "}AlternateContent":
            choices = child.findall("{" + MC + "}Choice")
            fallbacks = child.findall("{" + MC + "}Fallback")
            for branch in choices + fallbacks:
                selected = list(shape_elements(branch))
                if selected:
                    yield from selected
                    break
