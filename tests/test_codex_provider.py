import json
from pathlib import Path
import sys
from tempfile import TemporaryDirectory
import unittest
from unittest.mock import Mock

from gemini_cli import EmptyModelResponseError
from pptx_pipeline.codex_provider import CodexAppServerProvider, CodexRPCError, StdioRPC


class FakeRPC:
    def __init__(self, *, text='{"reviews":[],"changes":[]}', phase='final_answer', status='completed'):
        self.events, self.stderr, self.requests, self.sent = [], [], [], []
        self.closed = False
        self.text, self.phase, self.status = text, phase, status

    def send(self, value): self.sent.append(value)

    def request(self, method, params):
        self.requests.append((method, params))
        if method == 'initialize': return {}
        if method == 'model/list':
            return {'data': [{'model': 'available-vision', 'isDefault': True, 'inputModalities': ['text', 'image']}]}
        if method == 'thread/start': return {'thread': {'id': 'thread1'}, 'model': params['model']}
        if method == 'turn/start':
            self.events.extend([
                {'method': 'item/completed', 'params': {'threadId': 'thread1', 'turnId': 'turn1',
                    'item': {'id': 'comment', 'type': 'agentMessage', 'phase': 'commentary', 'text': 'Working'}}},
                {'method': 'item/completed', 'params': {'threadId': 'other', 'turnId': 'turn1',
                    'item': {'id': 'wrong', 'type': 'agentMessage', 'text': 'Wrong thread'}}},
                {'method': 'item/completed', 'params': {'threadId': 'thread1', 'turnId': 'turn1',
                    'item': {'id': 'reply', 'type': 'agentMessage', 'phase': self.phase, 'text': self.text}}},
                {'method': 'turn/completed', 'params': {'threadId': 'thread1',
                    'turn': {'id': 'turn1', 'status': self.status, 'error': {'message': 'denied'}}}},
            ])
            return {'turn': {'id': 'turn1'}}
        raise AssertionError(method)

    def receive(self): raise AssertionError('Events already available')
    def close(self): self.closed = True


class CodexProviderTests(unittest.TestCase):
    def setUp(self):
        self.temp = TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.folder = Path(self.temp.name)
        self.image = self.folder / 'slide.png'
        self.image.write_bytes(b'image fixture')

    def provider(self, rpc, **options):
        return CodexAppServerProvider(executable=sys.executable, cwd=self.folder,
            rpc_factory=lambda *args, **kwargs: rpc, trace_dir=self.folder / 'traces', **options)

    def test_native_image_inputs_model_discovery_and_final_message(self):
        rpc = FakeRPC()
        provider = self.provider(rpc)
        result = provider.generate('Review', images=[self.image])
        self.assertEqual(json.loads(result.text), {'reviews': [], 'changes': []})
        self.assertEqual(result.images_read, (self.image.resolve(),))
        self.assertEqual(provider.model, 'available-vision')
        thread = dict(rpc.requests)['thread/start']
        self.assertEqual(thread['sandbox'], 'read-only')
        self.assertEqual(thread['approvalPolicy'], 'never')
        turn = dict(rpc.requests)['turn/start']
        self.assertEqual(turn['effort'], 'low')
        self.assertEqual(turn['input'][1], {'type': 'localImage', 'path': str(self.image.resolve())})
        self.assertTrue(rpc.closed)
        trace = json.loads(next((self.folder / 'traces').glob('*.json')).read_text(encoding='utf-8'))
        self.assertEqual(trace['response'], result.text)

    def test_explicit_model_is_preserved_and_legacy_phase_supported(self):
        rpc = FakeRPC(phase=None)
        provider = self.provider(rpc, model='requested-model')
        self.assertTrue(provider.generate('Review').text)
        self.assertNotIn('model/list', dict(rpc.requests))
        self.assertEqual(dict(rpc.requests)['thread/start']['model'], 'requested-model')

    def test_empty_response_has_diagnostics_and_cleans_up(self):
        rpc = FakeRPC(text=' ')
        with self.assertRaises(EmptyModelResponseError) as caught:
            self.provider(rpc).generate('Review', images=[self.image])
        self.assertIn('turn/completed', caught.exception.result.stdout)
        self.assertTrue(rpc.closed)

    def test_failed_turn_is_not_skipped_as_empty(self):
        rpc = FakeRPC(status='failed')
        with self.assertRaises(CodexRPCError):
            self.provider(rpc).generate('Review')
        self.assertTrue(rpc.closed)

    def test_missing_image_rejected_before_starting_process(self):
        factory = Mock()
        provider = CodexAppServerProvider(executable=sys.executable, rpc_factory=factory)
        with self.assertRaises(FileNotFoundError):
            provider.generate('Review', images=[self.folder / 'missing.png'])
        factory.assert_not_called()

    def test_rpc_reports_early_exit_and_cleans_up(self):
        rpc = StdioRPC((sys.executable, '-c', 'pass'), cwd=self.folder, timeout=5)
        try:
            with self.assertRaises(CodexRPCError): rpc.receive()
        finally:
            rpc.close()
        self.assertIsNotNone(rpc.process.poll())

    def test_rpc_times_out_and_kills_owned_process(self):
        rpc = StdioRPC((sys.executable, '-c', 'import time; time.sleep(20)'), cwd=self.folder, timeout=0.1)
        try:
            with self.assertRaises(CodexRPCError): rpc.receive()
        finally:
            rpc.close()
        self.assertIsNotNone(rpc.process.poll())


if __name__ == '__main__':
    unittest.main()
