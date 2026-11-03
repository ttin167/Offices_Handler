"""
core.ba_hub — BA Enterprise Workstation Core Engine.

Comprehensive toolset for professional Business Analysts:
- story_builder: User Story & Gherkin Acceptance Criteria.
- usecase_builder: Use Case Specifications (OpenXML Word tables).
- brd_generator: Full BRD/SRS synthesis (Cover, Revision Control, RTM, NFR).
- data_dictionary: Automated schema scanner (Excel/SQL) & Source-Target Mapping.
- reconciliation: Two-way Data Reconciliation Engine (Composite Key Matching, Tolerance, Variance Excel Report).
- testcase_generator: UAT Checklist & Test Matrix Generator with Excel dropdown validations.
- templates: Industry-standard blueprints (Banking, E-commerce, WMS).
"""

from __future__ import annotations

from .story_builder import (
    UserStory,
    GherkinScenario,
    StoryBuilder,
    StoryQualityLinter,
)
from .usecase_builder import (
    UseCaseSpec,
    FlowStep,
    BranchFlow,
    UseCaseBuilder,
    UseCaseQualityLinter,
)
from .brd_generator import (
    BRDDocumentGenerator,
    BRDMetadata,
)
from .data_dictionary import (
    DataField,
    DataDictionaryScanner,
    FieldMappingEngine,
)
from .reconciliation import (
    DataReconciliationEngine,
    ReconciliationSummary,
    DiscrepancyDetail,
)
from .testcase_generator import (
    UATScenario,
    UATTestcaseGenerator,
)
from .rtm_matrix import (
    RTMItem,
    RTMCoverage,
    RTMEngine,
)
from .dbml_satsuma import (
    DBMLField,
    DBMLTable,
    DBMLEngine,
    SatsumaMappingRule,
    SatsumaMappingEngine,
)
from .scope_diff import (
    ScopeDeltaItem,
    ScopeChangeSummary,
    ScopeCreepAuditor,
)
from .sql_toolkit import (
    SQLIntegrityTestItem,
    SQLIntegrityTestSuite,
)
from .obsidian_vault import (
    ObsidianVaultExporter,
    ObsidianVaultCompiler,
)
from .templates import (
    STORY_TEMPLATES,
    USECASE_TEMPLATES,
    DATA_DICT_TEMPLATES,
    BRD_DEFAULT_OUTLINE,
)

__all__ = [
    "UserStory",
    "GherkinScenario",
    "StoryBuilder",
    "StoryQualityLinter",
    "UseCaseSpec",
    "FlowStep",
    "BranchFlow",
    "UseCaseBuilder",
    "UseCaseQualityLinter",
    "BRDDocumentGenerator",
    "BRDMetadata",
    "DataField",
    "DataDictionaryScanner",
    "FieldMappingEngine",
    "DataReconciliationEngine",
    "ReconciliationSummary",
    "DiscrepancyDetail",
    "UATScenario",
    "UATTestcaseGenerator",
    "RTMItem",
    "RTMCoverage",
    "RTMEngine",
    "DBMLField",
    "DBMLTable",
    "DBMLEngine",
    "SatsumaMappingRule",
    "SatsumaMappingEngine",
    "ScopeDeltaItem",
    "ScopeChangeSummary",
    "ScopeCreepAuditor",
    "SQLIntegrityTestItem",
    "SQLIntegrityTestSuite",
    "ObsidianVaultExporter",
    "ObsidianVaultCompiler",
    "STORY_TEMPLATES",
    "USECASE_TEMPLATES",
    "DATA_DICT_TEMPLATES",
    "BRD_DEFAULT_OUTLINE",
]

