from contextlib import contextmanager
import re
import subprocess


def safe_error(error):
    text = str(error)
    # Errors are displayed, but raw protocol/account/auth data is never logged.
    text = re.sub(r'https?://\S+', '[URL 숨김]', text)
    text = re.sub(r'(?i)(bearer\s+|(?:access[_-]?token|refresh[_-]?token|api[_-]?key|authorization|password|secret)\s*[=:]\s*)[^\s,;}]+', r'\1[숨김]', text)
    text = re.sub(r'''(?i)((?:access[_-]?token|refresh[_-]?token|api[_-]?key|authorization|password|secret)["']?\s*[:=]\s*["']?)[^\s"',;}]+''',r'\1[숨김]',text)
    text = re.sub(r'\bsk-[A-Za-z0-9_-]+', '[키 숨김]', text)
    text = re.sub(r'\beyJ[A-Za-z0-9_.-]+', '[토큰 숨김]', text)
    return text[:500]


def version_of(executable):
    result = subprocess.run([executable, '--version'], capture_output=True, text=True,
                            encoding='utf-8', errors='replace', timeout=10,
                            creationflags=getattr(subprocess, 'CREATE_NO_WINDOW', 0))
    return result.stdout.strip().splitlines()[0] if result.stdout.strip() else '버전 확인 불가'


