"""Replaceable analysis, translation, application and layout components."""

from .ai import AIProvider, AIResponse, GeminiCLIProvider, ProviderRegistry
from .codex_provider import CodexAppServerProvider
from .rendering import PowerPointRenderer, RenderResult
from .layout import LayoutPolicy, PPTXLayoutModifier
from .stages import (PipelineContext, PresentationPipeline, AnalysisStage, TranslationStage,
                     ApplicationStage, ModificationStage)

__all__ = ["AIProvider", "AIResponse", "GeminiCLIProvider", "CodexAppServerProvider", "ProviderRegistry",
           "PowerPointRenderer", "RenderResult", "LayoutPolicy", "PPTXLayoutModifier",
           "PipelineContext", "PresentationPipeline", "AnalysisStage", "TranslationStage",
           "ApplicationStage", "ModificationStage"]
