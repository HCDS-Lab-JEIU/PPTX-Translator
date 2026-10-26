"""Recover presentation formatting without throwing away translated prose."""
import re

from pptx_styles import MARKER, SPAN, plain_text, relaxed_styles, validate_styles
from pptx_text import EQUATION, equation_lines


class TranslationValidationError(ValueError):
    def __init__(self, code):
        super().__init__(code)
        self.code = code


def without_styles(text):
    return re.sub(r'\[\[/?STYLE_[^\]]*\]\]', '', text)


def prepare_translation(record, text):
    """Keep content, permit missing style spans, and restore equation line anchors."""
    if not isinstance(text, str) or not text.strip():
        raise TranslationValidationError('empty_text')
    text = text.replace('\r\n', '\n').replace('\r', '\n').replace('\v', '\n')
    notes = []
    if record.styled_text:
        try:
            validate_styles(record.styled_text, text)
        except ValueError:
            # Grammatical source fragments can disappear in translation. Keep any
            # whitespace, but do not require an empty translated particle span.
            text = SPAN.sub(lambda match: match[2] if not match[2].strip() else match[0], text)
            if not relaxed_styles(record.styled_text, text):
                text = without_styles(text)
            notes.append('style_alignment_deferred')
    elif MARKER.search(text) or re.search(r'\[\[/?STYLE_', text):
        text = without_styles(text)
        notes.append('unexpected_style_markers_removed')
    if not plain_text(text).strip():
        raise TranslationValidationError('empty_text')
    original = record.text.replace('\v', '\n')
    if EQUATION.findall(original) != EQUATION.findall(text):
        raise TranslationValidationError('equation_marker_mismatch')
    if EQUATION.search(original) and equation_lines(original) != equation_lines(plain_text(text)):
        # Keep equations in their source paragraphs. Reflow only prose between
        # consecutive anchors; never recreate/move the native Office Math XML.
        text = without_styles(text)
        pattern = r'(\[\[EQ_\d+\]\])'
        source_parts, target_parts = re.split(pattern, original), re.split(pattern, text)
        for index in range(0, len(source_parts), 2):
            source_lines = source_parts[index].split('\n')
            target_lines = target_parts[index].split('\n')
            if len(source_lines) != len(target_lines):
                prose = re.sub(r'\s*\n\s*', ' ', target_parts[index]).strip()
                restored = [''] * len(source_lines)
                slot = next((i for i, line in enumerate(source_lines) if line.strip()), len(restored) - 1)
                restored[slot] = prose
                target_parts[index] = '\n'.join(restored)
        text = ''.join(target_parts)
        if equation_lines(original) != equation_lines(text):
            raise TranslationValidationError('equation_structure_mismatch')
        notes.extend(['equation_lines_restored', 'style_alignment_deferred'])
    if original.count('\n') != plain_text(text).count('\n'):
        notes.append('paragraph_reflow_deferred')
    return text, sorted(set(notes))
