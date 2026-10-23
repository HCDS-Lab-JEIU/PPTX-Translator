import os
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest
from unittest.mock import patch
from pptx_pipeline.codex_provider import find_codex
from pptx_gui.adapters.gemini import find_gemini
from gemini_cli import find_executable
from pptx_gui.adapters.antigravity import parse_models,AntigravityConnectionAdapter
from pptx_gui.models import NodeSettings


class DiscoveryTests(unittest.TestCase):
    def setUp(self):
        self.temp=TemporaryDirectory(); self.addCleanup(self.temp.cleanup)
        self.root=Path(self.temp.name)
    def file(self,relative):
        path=self.root/relative; path.parent.mkdir(parents=True,exist_ok=True); path.write_bytes(b'fixture'); return path
    def test_codex_is_found_in_vscode_without_path(self):
        path=self.file('.vscode/extensions/openai.chatgpt-1.2-win32-x64/bin/windows-x86_64/codex.exe')
        env={'PATH':'','LOCALAPPDATA':str(self.root/'local'),'APPDATA':str(self.root/'roaming'),'CODEX_EXECUTABLE':''}
        with patch.dict(os.environ,env),patch('cli_discovery.Path.home',return_value=self.root),patch('cli_discovery.shutil.which',return_value=None):
            self.assertEqual(find_codex(),str(path.resolve()))
            with self.assertRaises(FileNotFoundError): find_codex(self.root/'invalid.exe')
    def test_standalone_codex_is_found_without_path(self):
        path=self.file('local/Programs/OpenAI/Codex/bin/codex.exe')
        with patch.dict(os.environ,{'CODEX_EXECUTABLE':'','LOCALAPPDATA':str(self.root/'local')}),patch('cli_discovery.shutil.which',return_value=None):
            self.assertEqual(find_codex(),str(path.resolve()))
    def test_gemini_npm_wrapper_resolves_to_node_and_js_without_path(self):
        self.file('roaming/npm/gemini.cmd')
        script=self.file('roaming/npm/node_modules/@google/gemini-cli/dist/index.js')
        node=self.file('programs/nodejs/node.exe')
        env={'GEMINI_CLI_EXECUTABLE':'','APPDATA':str(self.root/'roaming'),'ProgramFiles':str(self.root/'programs'),'LOCALAPPDATA':str(self.root/'local')}
        with patch.dict(os.environ,env),patch('cli_discovery.shutil.which',return_value=None),patch('cli_discovery.Path.home',return_value=self.root):
            self.assertEqual(find_gemini(),[str(node.resolve()),str(script.resolve())])
            self.assertEqual(find_gemini(str(script)),[str(node.resolve()),str(script.resolve())])
    def test_antigravity_falls_back_from_path_wrapper_to_native_install(self):
        wrapper=self.file('path/agy.cmd'); native=self.file('local/agy/bin/agy.exe')
        with patch.dict(os.environ,{'ANTIGRAVITY_CLI_EXECUTABLE':'','LOCALAPPDATA':str(self.root/'local')}),patch('gemini_cli.shutil.which',return_value=str(wrapper)):
            self.assertEqual(find_executable(),str(native.resolve()))
    def test_gemini_rejects_antigravity_executable(self):
        with self.assertRaises(FileNotFoundError): find_gemini(str(self.file('agy.exe')))
    def test_antigravity_models_are_actual_slugs_not_invented_efforts(self):
        models=parse_models('gemini-3.8-flash-medium\tGemini 3.8 Flash (Medium)\nclaude-sonnet-4-6\tClaude Sonnet\n')
        self.assertTrue(models[0].vision); self.assertTrue(models[0].default)
        self.assertEqual(models[0].efforts,())
        self.assertFalse(models[1].vision)
        for output in ('','unknown output','a\tA\na\tA\n'):
            with self.assertRaises(ValueError): parse_models(output)
    def test_antigravity_provider_uses_selected_model_and_separate_wrapper(self):
        executable=self.file('agy.exe')
        provider=AntigravityConnectionAdapter(str(executable)).provider(NodeSettings(provider='antigravity',model='actual-slug'))
        self.assertEqual(provider.model,'actual-slug')
        self.assertTrue(callable(provider.cancel))


if __name__=='__main__': unittest.main()
