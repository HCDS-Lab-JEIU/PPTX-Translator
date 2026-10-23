import json
from pathlib import Path
import subprocess
import sys
from tempfile import TemporaryDirectory
import unittest
from unittest.mock import Mock, patch

from gemini_cli import CLIResult, GeminiCLI, GeminiCLIError, EmptyModelResponseError
from translate_texts import PPTXTextTranslator, read_records, parse_translations


class TranslationTests(unittest.TestCase):
    def setUp(self):
        self.temp = TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.folder = Path(self.temp.name)
        self.source = self.folder / "texts.txt"
        self.prefixes = ["[page=1][x=1,y=2,w=3,h=4 cm]",
                         "[page=2][x=2,y=3,w=4,h=5 cm; table=8,row=1,col=2]"]
        self.source.write_text(
            self.prefixes[0] + '[' + json.dumps('전원\n"전류"', ensure_ascii=False) + ']\n' +
            self.prefixes[1] + '[' + json.dumps('저항 R = 10 Ω', ensure_ascii=False) + ']\n', encoding="utf-8")

    def test_english_text_preserves_prefixes_and_original_input(self):
        before = self.source.read_bytes()
        cli = Mock()
        cli.generate.return_value = CLIResult((), 0, json.dumps({"translations": [
            {"id": 1, "text": "Resistance R = 10 Ω"},
            {"id": 0, "text": 'Power supply\n"Current"'},
        ]}), "")
        report = PPTXTextTranslator(cli).translate_file(self.source)
        output = Path(report["output"])
        self.assertEqual(output.name, "translated.txt")
        translated = read_records(output)
        self.assertEqual([r.prefix for r in translated], self.prefixes)
        self.assertEqual([r.text for r in translated], ['Power supply\n"Current"', 'Resistance R = 10 Ω'])
        self.assertEqual(self.source.read_bytes(), before)
        self.assertEqual(report["records"], 2)
        self.assertIn("전원", cli.generate.call_args.args[0])

    def test_failed_batch_keeps_existing_translation(self):
        destination = self.folder / "translated.txt"
        destination.write_text("earlier translation", encoding="utf-8")
        cli = Mock()
        cli.generate.side_effect = [
            CLIResult((), 0, '{"translations":[{"id":0,"text":"Power supply"}]}', ""),
            CLIResult((), 0, '{"translations":[]}', ""),
        ]
        with self.assertRaises(ValueError):
            PPTXTextTranslator(cli, batch_chars=1).translate_file(self.source)
        self.assertEqual(cli.generate.call_count, 2)
        self.assertEqual(destination.read_text(encoding="utf-8"), "earlier translation")
        self.assertEqual(len(read_records(self.folder / 'translated.partial.txt')), 1)
        state = json.loads((self.folder / 'translated.progress.json').read_text(encoding='utf-8'))
        self.assertEqual(state['status'], 'failed')
        self.assertEqual(state['processed_records'], 1)

    def test_empty_batch_retries_by_page_and_keeps_failed_page_original(self):
        cli = Mock()
        cli.generate.side_effect = [EmptyModelResponseError('empty batch'),
            EmptyModelResponseError('empty page 1'),
            CLIResult((), 0, '{"translations":[{"id":1,"text":"Resistance R = 10 Ω"}]}', '')]
        report = PPTXTextTranslator(cli).translate_file(self.source)
        self.assertEqual(cli.generate.call_count, 3)
        self.assertEqual(report['status'], 'needs_review')
        self.assertEqual(report['failed_pages'], [1])
        self.assertEqual(report['failed_records'], [0])
        translated = read_records(report['output'])
        self.assertEqual(translated[0].text, read_records(self.source)[0].text)
        self.assertEqual(translated[1].text, 'Resistance R = 10 Ω')

    def test_numeric_changes_retry_only_affected_items(self):
        cli = Mock()
        cli.generate.side_effect = [
            CLIResult((), 0, json.dumps({"translations": [
                {"id": 0, "text": "Power supply"}, {"id": 1, "text": "Resistance R = ten Ω"}]}), ""),
            CLIResult((), 0, '{"translations":[{"id":1,"text":"Resistance R = 10 Ω"}]}', ""),
        ]
        report = PPTXTextTranslator(cli).translate_file(self.source)
        self.assertEqual(cli.generate.call_count, 2)
        self.assertNotIn('"id": 0', cli.generate.call_args.args[0].split("Input records:")[1])
        self.assertEqual(read_records(report["output"])[1].text, "Resistance R = 10 Ω")

    def test_persistent_numeric_changes_do_not_save_output(self):
        cli = Mock()
        cli.generate.return_value = CLIResult((), 0, json.dumps({"translations": [
            {"id": 0, "text": "Power supply"}, {"id": 1, "text": "Resistance R = ten Ω"}]}), "")
        # Retry is only for id 1.
        cli.generate.side_effect = [cli.generate.return_value,
            CLIResult((), 0, '{"translations":[{"id":1,"text":"Resistance R = ten Ω"}]}', "")]
        with self.assertRaises(ValueError):
            PPTXTextTranslator(cli).translate_file(self.source)
        self.assertFalse((self.folder / "translated.txt").exists())

    def test_rejects_missing_duplicate_unknown_and_empty_translations(self):
        records = read_records(self.source)
        for items in ([], [{"id": 0, "text": "x"}] * 2,
                      [{"id": 9, "text": "x"}], [{"id": 0, "text": ""}],
                      [{"id": False, "text": "x"}]):
            with self.subTest(items=items), self.assertRaises(ValueError):
                parse_translations(json.dumps({"translations": items}), records)

    def test_invalid_input_and_source_overwrite_rejected_before_model_call(self):
        cli = Mock()
        with self.assertRaises(ValueError):
            PPTXTextTranslator(cli).translate_file(self.source, self.source)
        self.source.write_text("malformed extractor output", encoding="utf-8")
        with self.assertRaises(ValueError):
            PPTXTextTranslator(cli).translate_file(self.source)
        cli.generate.assert_not_called()

    def test_empty_extraction_saves_empty_translation_without_model_call(self):
        self.source.write_text("", encoding="utf-8")
        cli = Mock()
        report = PPTXTextTranslator(cli).translate_file(self.source)
        self.assertEqual(report["records"], 0)
        self.assertEqual(Path(report["output"]).read_text(encoding="utf-8"), "")
        cli.generate.assert_not_called()

    @patch("gemini_cli.subprocess.run")
    def test_long_prompt_is_passed_via_stdin_and_final_response_read(self, run):
        response = '{"translations":[{"id":0,"text":"Power supply"}]}'
        run.return_value = subprocess.CompletedProcess([], 0,
            json.dumps({"event": "init"}) + "\n" + json.dumps({"event": "result", "result": {
                "status": "SUCCESS", "response": response}}) + "\n", "")
        prompt = "번역 " * 20000
        result = GeminiCLI(sys.executable).generate(prompt, model="test-model")
        self.assertEqual(result.stdout, response)
        self.assertNotIn(prompt, run.call_args.args[0])
        self.assertEqual(json.loads(run.call_args.kwargs["input"])["message"]["content"], prompt)
        self.assertIn("test-model", run.call_args.args[0])
        run.return_value = subprocess.CompletedProcess([], 0, json.dumps({"event": "result", "result": {
            "status": "ERROR", "error": "quota exceeded"}}), "")
        with self.assertRaises(GeminiCLIError):
            GeminiCLI(sys.executable).generate("translate")


if __name__ == "__main__":
    unittest.main()
