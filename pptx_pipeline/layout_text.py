"""Validate optional prose compression without relaxing equation/style protection."""
from collections import Counter
import re

from pptx_styles import MARKER, plain_text, validate_styles
from translate_texts import TextRecord, changed_equations, changed_numbers


def validate_compression(frame, text):
    original = frame['text']
    if not frame.get('text_editable') or not isinstance(text, str) or not text.strip():
        raise ValueError('이 텍스트는 압축할 수 없습니다.')
    plain = plain_text(text)
    if len(plain) >= len(original):
        raise ValueError('압축문은 원문보다 짧아야 합니다.')
    record = TextRecord(0, 1, '', original)
    if changed_numbers(record, text) or changed_equations(record, text):
        raise ValueError('문장 압축 시 숫자·수식·수식 줄바꿈을 유지해야 합니다.')
    if plain.count('\n') != original.count('\n'):
        raise ValueError('문장 압축 시 문단 경계를 유지해야 합니다.')
    # Preserve URLs, inline code, units/acronyms, equation-like expressions, negation.
    if re.search(r'(?m)^\s*(?:def |class |import |#include|public |return |for\s*\(|if\s*\()', original):
        raise ValueError('코드가 포함된 프레임은 압축할 수 없습니다.')
    protected = (r'https?://\S+|`[^`]+`|\b[A-Z]{1,4}\b|\b(?:not|never|without|cannot)\b|'
                 r'(?<!\w)(?:[fpnumkMGTµμ]?(?:V|A|W|F|H|Ω|Hz|s|m|K)|°C|dB)(?!\w)|'
                 r'\b\w+\s*[=<>≤≥±×÷+*/^]\s*\S+')
    if Counter(re.findall(protected, original)) != Counter(re.findall(protected, plain)):
        raise ValueError('문장 압축 시 보호된 기호·단위·코드·부정 표현을 유지해야 합니다.')
    if frame.get('styled_text'):
        validate_styles(frame['styled_text'], text)
    elif MARKER.search(text):
        raise ValueError('원문에 없는 서식 표시입니다.')
