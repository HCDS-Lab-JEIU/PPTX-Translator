"""Full PDF precedes prediction; predictions route each page to relevant repairs."""
import json
from pathlib import Path
import unittest
from unittest.mock import patch

from pptx_pipeline.ai import AIResponse
from pptx_pipeline.layout import LayoutPolicy, PPTXLayoutModifier
from pptx_pipeline.layout_cache import render_cached
from pptx_pipeline.layout_pdf import pdf_page_images
from pptx_pipeline.rendering import PowerPointRenderer
from tests import test_layout_stages as staged_tests
from tests.test_pipeline import FakeRenderer


class PdfPlanningTests(unittest.TestCase):
    setUp = staged_tests.StagedLayoutTests.setUp
    provider = staged_tests.StagedLayoutTests.provider
    run_modifier = staged_tests.StagedLayoutTests.run_modifier

    def test_missing_pdf_runtime_is_detected_before_opening_powerpoint(self):
        renderer = PowerPointRenderer()
        with patch('pptx_pipeline.layout_pdf.pdf_runtime', side_effect=ModuleNotFoundError(name='pypdfium2')):
            with patch.object(renderer, '_run') as run, self.assertRaises(ModuleNotFoundError):
                renderer.render_document(self.source, self.root / 'pdf')
            run.assert_not_called()

    def predicted_provider(self, predictions):
        provider = self.provider()
        generate = provider.generate
        provider.pdf_requests = []
        def wrapped(prompt, *, images=()):
            if 'Screen the whole presentation' not in prompt:
                return generate(prompt, images=images)
            data = json.loads(prompt.split('Do not use tools or modify files.\n')[1])
            pdf = Path(data['pdf'])
            self.assertTrue(pdf.is_file())
            self.assertEqual(len(pdf_page_images(pdf, self.root / 'read-pdf', expected_pages=4, width=80)), 4)
            provider.pdf_requests.append(data)
            self.assertEqual(provider.requests, [])  # no repair before prediction
            return AIResponse(json.dumps({'reviews':[
                {'page':p,'needs_changes':bool(predictions[p]), 'issues':['overlap'] if predictions[p] else [],
                 'defects':[{'kind':'overlap','evidence':'New text obscures a diagram label'}] if predictions[p] else [],
                 'predicted_adjustments':predictions[p]} for p in data['pages']]}), tuple(images))
        provider.generate = wrapped
        return provider

    def test_single_pdf_prediction_routes_multiple_types_and_skips_unrelated_stages(self):
        provider = self.predicted_provider({1:['spacing','font'],2:['spacing'],3:['images'],4:['compression']})
        report, renderer = self.run_modifier(provider)
        self.assertEqual(len(provider.pdf_requests), 1)
        self.assertEqual(provider.pdf_requests[0]['pages'], [1,2,3,4])
        self.assertEqual(provider.requests, [
            ('font',False,[1]), ('font',True,[1]),
            ('spacing',False,[2]), ('spacing',True,[2]),
            ('images',False,[3]), ('images',True,[3]),
            ('compression',False,[4]), ('compression',True,[4])])
        self.assertEqual(report['page_predictions'][1], ['font','spacing'])
        self.assertEqual(report['screening']['input_format'], 'pdf_rasterized')
        self.assertEqual(renderer.requested_pages[:2], [[1,2,3,4]] * 2)
        self.assertTrue(Path(report['screening']['pdf']).is_file())

    def test_clean_pdf_never_starts_page_repairs(self):
        provider = self.predicted_provider({p:[] for p in range(1,5)})
        report, renderer = self.run_modifier(provider)
        self.assertEqual(len(provider.pdf_requests), 1)
        self.assertEqual(provider.requests, [])
        self.assertEqual(len(renderer.calls), 2)
        self.assertEqual(report['pipeline_stages'], [])
        self.assertEqual(report['status'], 'verified')

    def test_subset_repairs_still_export_and_predict_all_pdf_pages(self):
        provider = self.predicted_provider({1:[],2:[],3:[],4:['compression']})
        report = PPTXLayoutModifier(provider, FakeRenderer(), staged=True,
            policy=LayoutPolicy(max_passes=1, allow_text_compression=True)).modify(
                self.source, self.root / 'result.pptx', reference=self.source, manifest=self.manifest,
                workdir=self.root / 'work', pages=[4])
        self.assertEqual(provider.pdf_requests[0]['pages'], [1,2,3,4])
        self.assertEqual(report['requested_pages'], [4])
        self.assertEqual(report['screening']['selected_pages'], [4])
        self.assertEqual(set(report['page_predictions']), {1,2,3,4})
        self.assertEqual(provider.requests, [('compression',False,[4]), ('compression',True,[4])])

    def test_missing_prediction_does_not_silently_send_every_page_to_every_stage(self):
        provider = self.provider()
        generate = provider.generate
        def wrapped(prompt, *, images=()):
            response = generate(prompt, images=images)
            if 'Screen the whole presentation' in prompt:
                data = json.loads(response.stdout)
                for review in data['reviews']:
                    review.pop('predicted_adjustments')
                return AIResponse(json.dumps(data), tuple(images))
            return response
        provider.generate = wrapped
        report, _ = self.run_modifier(provider)
        self.assertTrue(report['screening']['failed'])
        self.assertEqual(provider.requests, [])
        self.assertEqual(report['status'], 'needs_review')

    def test_pdf_page_count_mismatch_and_missing_pdf_are_rejected(self):
        render = FakeRenderer().render_document(self.source, self.root / 'pdf')
        with self.assertRaises(ValueError):
            pdf_page_images(render.pdf, self.root / 'wrong', expected_pages=3)
        renderer = FakeRenderer()
        render.pdf = None
        with patch.object(renderer, 'render_document', return_value=render), self.assertRaises(RuntimeError):
            render_cached(renderer, self.source, self.root / 'cache', [1,2,3,4], document=True)

    def test_render_cache_recreates_a_deleted_pdf(self):
        renderer = FakeRenderer()
        first = render_cached(renderer, self.source, self.root / 'pdf', [1,2,3,4], document=True)
        first.pdf.unlink()  # disposable test fixture only
        second = render_cached(renderer, self.source, self.root / 'pdf', [1,2,3,4], document=True)
        self.assertTrue(second.pdf.is_file())
        self.assertEqual(len(renderer.calls), 2)
