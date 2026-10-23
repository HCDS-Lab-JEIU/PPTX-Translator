"""Gemini 호출용 클래스. 현재 backend는 Antigravity CLI(agy)이다."""

import argparse
from dataclasses import asdict, dataclass
from datetime import datetime
import json
import math
import os
from pathlib import Path
import re
import shutil
import subprocess
import sys

ANSI = re.compile(r"\x1b\[[0-?]*[ -/]*[@-~]|\x1b\][^\x07]*(?:\x07|\x1b\\)")


@dataclass(frozen=True)
class CLIResult:
    command: tuple[str, ...]
    returncode: int
    stdout: str
    stderr: str
    events: tuple[dict, ...] = ()


@dataclass(frozen=True)
class ModelQuota:
    model_group: str
    limit_name: str
    remaining_percent: float
    reset_at: str


@dataclass(frozen=True)
class UsageReport:
    quotas: tuple[ModelQuota, ...]
    raw_output: str
    stderr: str

    def to_dict(self):
        return asdict(self)


class GeminiCLIError(RuntimeError):
    def __init__(self, message, result=None):
        super().__init__(message)
        self.result = result


class EmptyModelResponseError(GeminiCLIError):
    """A successful CLI turn returned no text; callers may skip this request."""


def find_executable(executable=None):
    configured = executable or os.environ.get("ANTIGRAVITY_CLI_EXECUTABLE")
    if configured:
        found = shutil.which(str(configured))
        candidate = Path(found or configured).expanduser()
        if not candidate.is_file():
            raise FileNotFoundError(f"CLI 실행 파일을 찾을 수 없습니다: {configured}")
    else:
        found = shutil.which("agy")
        candidates = [Path(found)] if found else []
        local_app_data = os.environ.get("LOCALAPPDATA")
        if local_app_data:
            candidates.append(Path(local_app_data) / "agy/bin/agy.exe")
        candidates.append(Path.home() / ".local/bin/agy")
        candidate = next((path for path in candidates if path.is_file() and path.suffix.lower() not in {'.cmd','.bat','.ps1'}), None)
        if candidate is None:
            raise FileNotFoundError(
                "Antigravity CLI(agy)를 찾을 수 없습니다. --executable 또는 "
                "ANTIGRAVITY_CLI_EXECUTABLE로 실행 파일 경로를 지정하세요."
            )
    if candidate.suffix.lower() in {".cmd", ".bat", ".ps1"}:
        raise ValueError("셸 래퍼 대신 실제 agy 실행 파일 경로를 지정하세요.")
    return str(candidate.resolve())


def parse_usage(output):
    """agy 1.4.6의 TSV 보고서: 모델, 제한 이름, 남은 비율, 리셋 시각."""
    quotas = []
    for line in ANSI.sub("", output).splitlines():
        if not line.strip():
            continue
        fields = [field.strip() for field in line.split("\t")]
        if len(fields) != 4:
            raise ValueError("/usage 반환 형식이 예상한 4열 TSV가 아닙니다.")
        group, limit_name, percent, reset_at = fields
        match = re.fullmatch(r"(\d+(?:\.\d+)?)%", percent)
        if not group or "remaining" not in limit_name.casefold() or match is None:
            raise ValueError("/usage에서 모델별 남은 비율을 확인할 수 없습니다.")
        remaining = float(match.group(1))
        if not 0 <= remaining <= 100:
            raise ValueError("남은 비율이 0~100 범위를 벗어났습니다.")
        timestamp = datetime.fromisoformat(reset_at.replace("Z", "+00:00"))
        if timestamp.tzinfo is None:
            raise ValueError("리셋 시각에 시간대 정보가 없습니다.")
        quotas.append(ModelQuota(group, limit_name, remaining, reset_at))
    if not quotas:
        raise ValueError("/usage에서 사용량 정보를 반환하지 않았습니다.")
    return tuple(quotas)


class GeminiCLI:
    """Antigravity의 기존 로그인으로 독립적인 print-mode 명령을 호출한다.

    run(prompt): CLIResult, get_usage(): UsageReport.
    일반 prompt는 모델 호출을 발생시킨다. get_usage는 /usage만 실행한다.
    """

    def __init__(self, executable=None, *, timeout=60, cwd=None):
        if not math.isfinite(timeout) or timeout <= 0:
            raise ValueError("timeout은 0보다 큰 유한 숫자여야 합니다.")
        self.executable = find_executable(executable)
        self.timeout = timeout
        self.cwd = Path(cwd).resolve() if cwd is not None else None
        if self.cwd is not None and not self.cwd.is_dir():
            raise NotADirectoryError(str(self.cwd))

    def run(self, prompt):
        if not isinstance(prompt, str) or not prompt.strip():
            raise ValueError("prompt는 비어 있지 않은 문자열이어야 합니다.")
        command = (self.executable, "--print", prompt, "--output-format", "text")
        return self._execute(command, "")

    def generate(self, prompt, *, model=None):
        """긴 프롬프트를 stdin으로 전달하고 최종 모델 응답만 반환한다."""
        if not isinstance(prompt, str) or not prompt.strip():
            raise ValueError("prompt는 비어 있지 않은 문자열이어야 합니다.")
        command = (self.executable, "--input-format", "stream-json", "--output-format", "stream-json")
        if model:
            command += ("--model", model)
        message = json.dumps({"event": "user", "message": {"content": prompt}}, ensure_ascii=False)
        result = self._execute(command, message + "\n")
        try:
            events = [json.loads(line) for line in result.stdout.splitlines() if line.strip()]
            finals = [event["result"] for event in events if event.get("event") == "result"]
            if len(finals) != 1:
                raise ValueError("최종 응답이 정확히 하나여야 합니다.")
            final = finals[0]
            if final.get("status") != "SUCCESS":
                raise ValueError(str(final.get("error") or final.get("status")))
            response = final["response"]
            if not isinstance(response, str) or not response.strip():
                raise EmptyModelResponseError("모델이 빈 응답을 반환했습니다.", result)
        except (ValueError, KeyError, TypeError, AttributeError) as exc:
            raise GeminiCLIError(f"모델 응답을 읽을 수 없습니다: {exc}", result) from exc
        return CLIResult(result.command, result.returncode, response, result.stderr, tuple(events))

    def _execute(self, command, input_text):
        try:
            completed = subprocess.run(
                command, input=input_text, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                encoding="utf-8", errors="replace", timeout=self.timeout,
                cwd=self.cwd, shell=False, check=False,
                creationflags=subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0,
            )
        except subprocess.TimeoutExpired as exc:
            raise GeminiCLIError(f"CLI 응답 시간이 {self.timeout}초를 초과했습니다.") from exc
        except OSError as exc:
            raise GeminiCLIError(f"CLI를 실행할 수 없습니다: {exc}") from exc
        result = CLIResult(command, completed.returncode, completed.stdout, completed.stderr)
        if result.returncode != 0:
            detail = ANSI.sub("", result.stderr or result.stdout).strip()
            raise GeminiCLIError(f"CLI 실패(종료 코드 {result.returncode}): {detail}", result)
        return result

    def get_usage(self):
        result = self.run("/usage")
        try:
            quotas = parse_usage(result.stdout)
        except ValueError as exc:
            # Keep the original stdout/stderr for version/auth/network diagnostics.
            raise GeminiCLIError(f"사용량 보고서를 해석할 수 없습니다: {exc}", result) from exc
        return UsageReport(quotas, result.stdout, result.stderr)


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--executable", help="agy 실행 파일 경로 (기본: 자동 탐색)")
    parser.add_argument("--timeout", type=float, default=60, help="호출 제한 시간(초)")
    parser.add_argument("--json", action="store_true", help="사용량을 구조화된 JSON으로 출력")
    args = parser.parse_args(argv)
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            stream.reconfigure(encoding="utf-8")
    try:
        usage = GeminiCLI(args.executable, timeout=args.timeout).get_usage()
    except (GeminiCLIError, OSError, ValueError) as exc:
        print(str(exc), file=sys.stderr)
        return 1
    if args.json:
        print(json.dumps(usage.to_dict(), ensure_ascii=False, indent=2))
    else:
        sys.stdout.write(usage.raw_output)
        if not usage.raw_output.endswith("\n"):
            print()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
