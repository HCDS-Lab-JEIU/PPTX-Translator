import json
from pathlib import Path
import subprocess
import sys
from tempfile import TemporaryDirectory
import unittest
from unittest.mock import patch

from gemini_cli import GeminiCLI, GeminiCLIError, EmptyModelResponseError, find_executable, parse_usage

USAGE = (
    "Gemini Models\tWeekly Limit Remaining\t89%\t2026-10-30T04:45:54Z\n"
    "Claude and GPT models\tWeekly Limit Remaining\t100%\t2026-10-30T10:17:43Z\n"
)


class GeminiCLITests(unittest.TestCase):
    def test_parse_actual_usage_format(self):
        quotas = parse_usage("\x1b[32m" + USAGE + "\x1b[0m")
        self.assertEqual([q.remaining_percent for q in quotas], [89, 100])
        self.assertEqual(quotas[0].reset_at, "2026-10-30T04:45:54Z")
        self.assertEqual(quotas[1].model_group, "Claude and GPT models")

    def test_unknown_output_is_not_fabricated_quota(self):
        for output in ("", "Authentication required", "89% remaining", USAGE.replace("89%", "189%")):
            with self.subTest(output=output), self.assertRaises(ValueError):
                parse_usage(output)

    @patch("gemini_cli.subprocess.run")
    def test_get_usage_is_standalone_slash_command(self, run):
        run.return_value = subprocess.CompletedProcess([], 0, USAGE, "")
        cli = GeminiCLI(sys.executable)
        report = cli.get_usage()
        command = run.call_args.args[0]
        self.assertEqual(command[1:], ("--print", "/usage", "--output-format", "text"))
        self.assertFalse(run.call_args.kwargs["shell"])
        self.assertEqual(run.call_args.kwargs["input"], "")
        self.assertEqual(report.quotas[0].remaining_percent, 89)
        self.assertEqual(json.loads(json.dumps(report.to_dict()))["quotas"][0]["remaining_percent"], 89)

    @patch("gemini_cli.subprocess.run")
    def test_failure_keeps_diagnostics(self, run):
        run.return_value = subprocess.CompletedProcess([], 2, "", "Authentication required")
        with self.assertRaises(GeminiCLIError) as caught:
            GeminiCLI(sys.executable).get_usage()
        self.assertEqual(caught.exception.result.returncode, 2)
        self.assertIn("Authentication required", str(caught.exception))
        run.return_value = subprocess.CompletedProcess([], 0, "unknown format", "diagnostic")
        with self.assertRaises(GeminiCLIError) as caught:
            GeminiCLI(sys.executable).get_usage()
        self.assertEqual(caught.exception.result.stdout, "unknown format")

    @patch("gemini_cli.subprocess.run", side_effect=subprocess.TimeoutExpired("agy", 1))
    def test_timeout(self, run):
        with self.assertRaises(GeminiCLIError):
            GeminiCLI(sys.executable, timeout=1).get_usage()

    @patch('gemini_cli.subprocess.run')
    def test_empty_success_is_distinct_and_preserves_raw_events(self, run):
        raw = json.dumps({'event': 'result', 'result': {'status': 'SUCCESS', 'response': '  '}})
        run.return_value = subprocess.CompletedProcess([], 0, raw, 'diagnostic')
        with self.assertRaises(EmptyModelResponseError) as caught:
            GeminiCLI(sys.executable).generate('test')
        self.assertEqual(caught.exception.result.stdout, raw)

    def test_missing_executable_and_invalid_timeout(self):
        with TemporaryDirectory() as folder:
            with self.assertRaises(FileNotFoundError):
                find_executable(Path(folder) / "missing.exe")
        for timeout in (0, -1, float("nan"), float("inf")):
            with self.subTest(timeout=timeout), self.assertRaises(ValueError):
                GeminiCLI(sys.executable, timeout=timeout)


if __name__ == "__main__":
    unittest.main()
