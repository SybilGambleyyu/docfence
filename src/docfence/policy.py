"""Strict, local policy loading and safe DocFence findings."""

from __future__ import annotations

import json
from dataclasses import dataclass, replace
from pathlib import Path
from typing import Final

from docfence.errors import PolicyError
from docfence.models import DiffReport, Finding

_MAX_POLICY_BYTES: Final = 64 * 1024
_RULES: Final = {
    "no_external_relationship_changes": "DFP001",
    "no_macro_payload_changes": "DFP002",
    "no_custom_xml_changes": "DFP003",
    "require_no_unresolved_revisions": "DFP004",
    "require_no_comments": "DFP005",
    "require_no_hidden_text": "DFP006",
    "no_relationship_changes": "DFP007",
    "no_document_settings_changes": "DFP008",
    "no_unclassified_package_payload_changes": "DFP009",
    "require_track_revisions_disabled": "DFP010",
    "require_no_field_codes": "DFP011",
    "require_no_content_controls": "DFP012",
    "require_no_hidden_text_style_declarations": "DFP013",
    "require_no_hidden_paragraph_marks": "DFP014",
    "require_no_embedded_objects": "DFP015",
    "require_no_alternative_format_imports": "DFP016",
    "no_embedded_object_payload_changes": "DFP017",
    "no_alternative_format_import_changes": "DFP018",
    "no_document_property_changes": "DFP019",
    "require_no_custom_document_properties": "DFP020",
    "require_no_mail_merge": "DFP021",
    "no_mail_merge_changes": "DFP022",
}


@dataclass(frozen=True)
class Policy:
    """Validated policy switches; absent switches are disabled."""

    rules: dict[str, bool]

    def enabled(self, rule_name: str) -> bool:
        return self.rules.get(rule_name, False)


def starter_policy() -> str:
    """Return the conservative starter policy emitted by `docfence init`."""

    return """version: 1
rules:
  no_external_relationship_changes: true
  no_macro_payload_changes: true
  no_custom_xml_changes: true
  require_no_unresolved_revisions: true
  require_no_comments: true
  require_no_hidden_text: true
"""


def load_policy(path: str | Path) -> Policy:
    """Load a small strict YAML or JSON policy without content-bearing errors."""

    try:
        policy_path = Path(path)
        if policy_path.is_symlink() or not policy_path.is_file():
            raise PolicyError("policy must be a regular file")
        if policy_path.stat().st_size > _MAX_POLICY_BYTES:
            raise PolicyError("policy exceeds the size limit")
        source = policy_path.read_bytes()
    except PolicyError:
        raise
    except (OSError, TypeError, ValueError):
        raise PolicyError("policy cannot be read") from None
    if b"\x00" in source:
        raise PolicyError("policy is invalid")
    try:
        text = source.decode("utf-8-sig")
    except UnicodeDecodeError:
        raise PolicyError("policy must be UTF-8") from None

    if text.lstrip().startswith("{"):
        try:
            document = json.loads(text)
        except json.JSONDecodeError:
            raise PolicyError("policy is invalid") from None
    else:
        document = _parse_strict_yaml(text)
    return _validate_policy_document(document)


def apply_policy(report: DiffReport, policy: Policy) -> DiffReport:
    """Attach deterministic, public-safe findings to a comparison report."""

    findings = tuple(_evaluate_policy(report, policy))
    return replace(report, findings=findings)


def _parse_strict_yaml(source: str) -> dict[str, object]:
    document: dict[str, object] = {}
    rules: dict[str, bool] = {}
    in_rules = False
    for line in source.splitlines():
        without_comment = line.split("#", 1)[0].rstrip()
        if not without_comment.strip():
            continue
        if "\t" in without_comment:
            raise PolicyError("policy is invalid")
        indent = len(without_comment) - len(without_comment.lstrip(" "))
        content = without_comment[indent:]
        if any(token in content for token in ("&", "*", "[", "]", "{", "}", "|", ">")):
            raise PolicyError("policy is invalid")
        key, separator, value = content.partition(":")
        if not separator or not key or key.strip() != key:
            raise PolicyError("policy is invalid")
        if indent == 0:
            if key in document:
                raise PolicyError("policy is invalid")
            if key == "version":
                if value.strip() != "1":
                    raise PolicyError("policy version is unsupported")
                document[key] = 1
                in_rules = False
            elif key == "rules":
                if value.strip():
                    raise PolicyError("policy is invalid")
                document[key] = rules
                in_rules = True
            else:
                raise PolicyError("policy is invalid")
        elif indent == 2 and in_rules:
            if key in rules or value.strip() not in {"true", "false"}:
                raise PolicyError("policy is invalid")
            rules[key] = value.strip() == "true"
        else:
            raise PolicyError("policy is invalid")
    return document


def _validate_policy_document(document: object) -> Policy:
    if not isinstance(document, dict) or set(document) != {"version", "rules"}:
        raise PolicyError("policy is invalid")
    if document["version"] != 1 or not isinstance(document["rules"], dict):
        raise PolicyError("policy version is unsupported")
    rules: dict[str, bool] = {}
    for rule_name, enabled in document["rules"].items():
        if rule_name not in _RULES or type(enabled) is not bool:
            raise PolicyError("policy is invalid")
        rules[rule_name] = enabled
    return Policy(rules=rules)


def _evaluate_policy(report: DiffReport, policy: Policy) -> list[Finding]:
    before = report.before
    after = report.after
    before_relationships = before.relationships
    after_relationships = after.relationships
    before_external_count = before_relationships.external_count
    after_external_count = after_relationships.external_count
    before_relationship_count = before_relationships.relationship_count
    after_relationship_count = after_relationships.relationship_count
    findings: list[Finding] = []

    if (
        policy.enabled("no_external_relationship_changes")
        and before_relationships.external_signature
        != after_relationships.external_signature
    ):
        findings.append(
            _finding(
                "no_external_relationship_changes",
                "External relationships changed.",
                {
                    "before_external_relationship_count": before_external_count,
                    "after_external_relationship_count": after_external_count,
                },
            )
        )
    if (
        policy.enabled("no_macro_payload_changes")
        and before.macro_signature != after.macro_signature
    ):
        findings.append(
            _finding(
                "no_macro_payload_changes",
                "Macro payload changed.",
                {
                    "before_macro_present": before.macro_present,
                    "after_macro_present": after.macro_present,
                },
            )
        )
    if (
        policy.enabled("no_custom_xml_changes")
        and before.custom_xml_signature != after.custom_xml_signature
    ):
        findings.append(
            _finding(
                "no_custom_xml_changes",
                "Custom XML payload changed.",
                {
                    "before_custom_xml_part_count": before.custom_xml_part_count,
                    "after_custom_xml_part_count": after.custom_xml_part_count,
                },
            )
        )
    if (
        policy.enabled("no_document_property_changes")
        and before.document_properties.signature != after.document_properties.signature
    ):
        findings.append(
            _finding(
                "no_document_property_changes",
                "Document property inventory changed.",
                {
                    "before": before.document_properties.public_dict(),
                    "after": after.document_properties.public_dict(),
                },
            )
        )
    if policy.enabled("no_mail_merge_changes") and (
        before.mail_merge.signature != after.mail_merge.signature
    ):
        findings.append(
            _finding(
                "no_mail_merge_changes",
                "Mail-merge inventory changed.",
                {
                    "before": before.mail_merge.public_dict(),
                    "after": after.mail_merge.public_dict(),
                },
            )
        )
    if (
        policy.enabled("require_no_unresolved_revisions")
        and after.revisions.unresolved_count
    ):
        findings.append(
            _finding(
                "require_no_unresolved_revisions",
                "Candidate contains unresolved revision markup.",
                {"unresolved_revision_count": after.revisions.unresolved_count},
            )
        )
    if policy.enabled("require_no_comments") and after.comment_count:
        findings.append(
            _finding(
                "require_no_comments",
                "Candidate contains stored comments.",
                {"comment_count": after.comment_count},
            )
        )
    if policy.enabled("require_no_hidden_text") and after.hidden_text_run_count:
        findings.append(
            _finding(
                "require_no_hidden_text",
                "Candidate contains stored hidden-text runs.",
                {"hidden_text_run_count": after.hidden_text_run_count},
            )
        )
    if policy.enabled("require_no_hidden_text_style_declarations") and (
        after.styles.hidden_text_style_definition_count
        or after.styles.document_default_hidden_text_enabled
    ):
        findings.append(
            _finding(
                "require_no_hidden_text_style_declarations",
                (
                    "Candidate contains stored style or document-default "
                    "declarations that can hide text."
                ),
                after.styles.public_dict(),
            )
        )
    if (
        policy.enabled("require_no_hidden_paragraph_marks")
        and after.hidden_paragraph_mark_count
    ):
        findings.append(
            _finding(
                "require_no_hidden_paragraph_marks",
                "Candidate contains hidden paragraph marks.",
                {"hidden_paragraph_mark_count": after.hidden_paragraph_mark_count},
            )
        )
    if policy.enabled("require_no_embedded_objects") and any(
        after.embedded_objects.public_dict().values()
    ):
        findings.append(
            _finding(
                "require_no_embedded_objects",
                "Candidate contains stored embedded object or control evidence.",
                after.embedded_objects.public_dict(),
            )
        )
    if policy.enabled("require_no_alternative_format_imports") and (
        after.alternative_format_import_anchor_count
        or any(after.alternative_format_imports.public_dict().values())
    ):
        findings.append(
            _finding(
                "require_no_alternative_format_imports",
                "Candidate contains stored alternative-format imports.",
                {
                    "alternative_format_import_anchor_count": (
                        after.alternative_format_import_anchor_count
                    ),
                    **after.alternative_format_imports.public_dict(),
                },
            )
        )
    if (
        policy.enabled("require_no_custom_document_properties")
        and after.document_properties.custom_property_count
    ):
        findings.append(
            _finding(
                "require_no_custom_document_properties",
                "Candidate contains stored custom document properties.",
                {
                    "custom_property_part_count": (
                        after.document_properties.custom_property_part_count
                    ),
                    "custom_property_count": (
                        after.document_properties.custom_property_count
                    ),
                },
            )
        )
    if policy.enabled("require_no_mail_merge") and any(
        after.mail_merge.public_dict().values()
    ):
        findings.append(
            _finding(
                "require_no_mail_merge",
                "Candidate contains stored mail-merge configuration or data state.",
                after.mail_merge.public_dict(),
            )
        )
    if (
        policy.enabled("no_embedded_object_payload_changes")
        and before.embedded_objects.signature != after.embedded_objects.signature
    ):
        findings.append(
            _finding(
                "no_embedded_object_payload_changes",
                "Embedded object or control inventory changed.",
                {
                    "before": before.embedded_objects.public_dict(),
                    "after": after.embedded_objects.public_dict(),
                },
            )
        )
    if policy.enabled("no_alternative_format_import_changes") and (
        before.alternative_format_imports.signature
        != after.alternative_format_imports.signature
        or before.alternative_format_import_anchor_count
        != after.alternative_format_import_anchor_count
    ):
        findings.append(
            _finding(
                "no_alternative_format_import_changes",
                "Alternative-format import inventory changed.",
                {
                    "before_alternative_format_import_anchor_count": (
                        before.alternative_format_import_anchor_count
                    ),
                    "after_alternative_format_import_anchor_count": (
                        after.alternative_format_import_anchor_count
                    ),
                    "before": before.alternative_format_imports.public_dict(),
                    "after": after.alternative_format_imports.public_dict(),
                },
            )
        )
    if (
        policy.enabled("no_relationship_changes")
        and before_relationships.relationship_signature
        != after_relationships.relationship_signature
    ):
        findings.append(
            _finding(
                "no_relationship_changes",
                "Package relationships changed.",
                {
                    "before_relationship_count": before_relationship_count,
                    "after_relationship_count": after_relationship_count,
                },
            )
        )
    if (
        policy.enabled("no_document_settings_changes")
        and before.settings_signature != after.settings_signature
    ):
        findings.append(
            _finding(
                "no_document_settings_changes",
                "Document settings changed.",
                {},
            )
        )
    if (
        policy.enabled("no_unclassified_package_payload_changes")
        and before.unclassified_signature != after.unclassified_signature
    ):
        findings.append(
            _finding(
                "no_unclassified_package_payload_changes",
                "Unclassified package payload changed.",
                {
                    "before_unclassified_part_count": before.unclassified_part_count,
                    "after_unclassified_part_count": after.unclassified_part_count,
                },
            )
        )
    if (
        policy.enabled("require_track_revisions_disabled")
        and after.track_revisions_enabled
    ):
        findings.append(
            _finding(
                "require_track_revisions_disabled",
                "Candidate has Track Changes enabled.",
                {},
            )
        )
    if policy.enabled("require_no_field_codes") and after.field_code_count:
        findings.append(
            _finding(
                "require_no_field_codes",
                "Candidate contains stored field codes.",
                {"field_code_count": after.field_code_count},
            )
        )
    if policy.enabled("require_no_content_controls") and after.content_control_count:
        findings.append(
            _finding(
                "require_no_content_controls",
                "Candidate contains stored content controls.",
                {"content_control_count": after.content_control_count},
            )
        )
    return findings


def _finding(rule_name: str, message: str, details: dict[str, object]) -> Finding:
    severity = "critical" if rule_name == "no_macro_payload_changes" else "high"
    return Finding(
        rule_id=_RULES[rule_name], severity=severity, message=message, details=details
    )
