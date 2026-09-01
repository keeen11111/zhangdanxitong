"""Stable, model-independent contracts for document-driven workbook agents."""

from .contracts import (
    BonusSourceDecision,
    Decision,
    DecisionAction,
    ExecutionPlan,
    WorkItem,
    WorkItemStatus,
    choose_bonus_source,
    ToolCall,
    ToolResult,
    ModelTrace,
    AgentEvent,
)
from .model import ModelConfig, ModelProviderError, ModelResponse, OpenAICompatibleProvider
from .orchestrator import ModelOrchestrator, OrchestrationResult, ToolRegistry
from .rules import Rule, RuleAction, RulePackage, RuleSource, resolve_rule_priority
from .rule_compiler import RuleCandidate, RuleCompilationError, compile_rule_package, validate_rule_package
from .materials import MaterialKind, MaterialText, classify_material, extract_material_text

__all__ = [
    "BonusSourceDecision",
    "Decision",
    "DecisionAction",
    "ExecutionPlan",
    "WorkItem",
    "WorkItemStatus",
    "choose_bonus_source",
    "ToolCall", "ToolResult", "ModelTrace", "AgentEvent",
    "ModelConfig", "ModelProviderError", "ModelResponse", "OpenAICompatibleProvider",
    "ModelOrchestrator", "OrchestrationResult", "ToolRegistry",
    "Rule",
    "RuleAction",
    "RulePackage",
    "RuleSource",
    "resolve_rule_priority",
    "RuleCandidate",
    "RuleCompilationError",
    "compile_rule_package",
    "validate_rule_package",
    "MaterialKind",
    "MaterialText",
    "classify_material",
    "extract_material_text",
]
