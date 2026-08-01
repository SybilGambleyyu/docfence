"""Public-safe JSON, Markdown, and SARIF renderers."""

from __future__ import annotations

import json
from typing import Literal

from docfence import __version__
from docfence.models import Change, DiffReport, DocumentSnapshot, Finding

ReportFormat = Literal["json", "markdown", "sarif"]
ProfileFormat = Literal["json", "markdown"]


def render_profile(snapshot: DocumentSnapshot, output_format: ProfileFormat) -> str:
    """Render a snapshot using only its explicit public projection."""

    payload = {"schema_version": 1, "document": snapshot.public_dict()}
    if output_format == "json":
        return _json(payload)
    if output_format == "markdown":
        return _profile_markdown(payload["document"])
    raise ValueError("unsupported profile output format")


def render_report(report: DiffReport, output_format: ReportFormat) -> str:
    """Render a diff report using only its explicit public projection."""

    if output_format == "json":
        return _json(report.public_dict())
    if output_format == "markdown":
        return _report_markdown(report)
    if output_format == "sarif":
        return _sarif(report)
    raise ValueError("unsupported report output format")


def _json(payload: object) -> str:
    return json.dumps(payload, indent=2, sort_keys=True) + "\n"


def _profile_markdown(document: dict[str, object]) -> str:
    lines = ["# DocFence profile", "", "## Document inventory", ""]
    lines.extend(_single_document_table(document))
    lines.extend(_story_kind_section(document.get("story_kind_counts", {})))
    lines.extend(_revision_section(document.get("revisions", {})))
    lines.extend(_style_section(document.get("styles", {})))
    lines.extend(_embedded_object_section(document.get("embedded_objects", {})))
    lines.extend(
        _alternative_format_import_section(
            document.get("alternative_format_imports", {})
        )
    )
    lines.extend(_document_property_section(document.get("document_properties", {})))
    lines.extend(_mail_merge_section(document.get("mail_merge", {})))
    return "\n".join(lines) + "\n"


def _report_markdown(report: DiffReport) -> str:
    before = report.before.public_dict()
    after = report.after.public_dict()
    lines = ["# DocFence comparison", "", "## Document inventory", ""]
    lines.extend(_comparison_table(before, after))
    lines.extend(_story_kind_comparison(before, after))
    lines.extend(_revision_comparison(before, after))
    lines.extend(_style_comparison(before, after))
    lines.extend(_embedded_object_comparison(before, after))
    lines.extend(_alternative_format_import_comparison(before, after))
    lines.extend(_document_property_comparison(before, after))
    lines.extend(_mail_merge_comparison(before, after))
    lines.extend(["## Changes", ""])
    if not report.changes:
        lines.append("No stored changes detected by the supported inventories.")
    else:
        for change in report.changes:
            lines.append(f"- `{change.kind}` — {change.message}")
            for key, value in sorted(change.details.items()):
                lines.append(f"  - `{key}`: {_value(value)}")
    lines.extend(["", "## Policy findings", ""])
    if not report.findings:
        lines.append("No policy findings.")
    else:
        for finding in report.findings:
            lines.append(
                f"- `{finding.rule_id}` ({finding.severity}) — {finding.message}"
            )
            for key, value in sorted(finding.details.items()):
                lines.append(f"  - `{key}`: {_value(value)}")
    return "\n".join(lines) + "\n"


def _single_document_table(document: dict[str, object]) -> list[str]:
    keys = (
        "format",
        "package_member_count",
        "story_count",
        "paragraph_count",
        "table_count",
        "text_run_count",
        "hidden_text_run_count",
        "hidden_paragraph_mark_count",
        "alternative_format_import_anchor_count",
        "field_code_count",
        "content_control_count",
        "comment_anchor_count",
        "comment_count",
        "track_revisions_enabled",
        "custom_xml_part_count",
        "macro_present",
        "unclassified_part_count",
    )
    lines = ["| Field | Value |", "| --- | ---: |"]
    lines.extend(f"| `{key}` | {_value(document[key])} |" for key in keys)
    return [*lines, ""]


def _comparison_table(before: dict[str, object], after: dict[str, object]) -> list[str]:
    keys = (
        "format",
        "package_member_count",
        "story_count",
        "paragraph_count",
        "table_count",
        "text_run_count",
        "hidden_text_run_count",
        "hidden_paragraph_mark_count",
        "alternative_format_import_anchor_count",
        "field_code_count",
        "content_control_count",
        "comment_anchor_count",
        "comment_count",
        "track_revisions_enabled",
        "custom_xml_part_count",
        "macro_present",
        "unclassified_part_count",
    )
    lines = ["| Field | Before | After |", "| --- | ---: | ---: |"]
    lines.extend(
        f"| `{key}` | {_value(before[key])} | {_value(after[key])} |" for key in keys
    )
    return [*lines, ""]


def _story_kind_section(value: object) -> list[str]:
    counts = _mapping(value)
    lines = ["## Story kinds", "", "| Kind | Count |", "| --- | ---: |"]
    lines.extend(f"| `{key}` | {_value(counts[key])} |" for key in sorted(counts))
    return [*lines, ""]


def _story_kind_comparison(
    before: dict[str, object], after: dict[str, object]
) -> list[str]:
    before_counts = _mapping(before.get("story_kind_counts", {}))
    after_counts = _mapping(after.get("story_kind_counts", {}))
    keys = sorted(set(before_counts) | set(after_counts))
    lines = ["## Story kinds", "", "| Kind | Before | After |", "| --- | ---: | ---: |"]
    lines.extend(
        _comparison_row(key, before_counts.get(key, 0), after_counts.get(key, 0))
        for key in keys
    )
    return [*lines, ""]


def _revision_section(value: object) -> list[str]:
    revisions = _mapping(value)
    lines = ["## Revision markup", "", "| Type | Count |", "| --- | ---: |"]
    lines.extend(f"| `{key}` | {_value(revisions[key])} |" for key in sorted(revisions))
    return [*lines, ""]


def _revision_comparison(
    before: dict[str, object], after: dict[str, object]
) -> list[str]:
    before_revisions = _mapping(before.get("revisions", {}))
    after_revisions = _mapping(after.get("revisions", {}))
    keys = sorted(set(before_revisions) | set(after_revisions))
    lines = [
        "## Revision markup",
        "",
        "| Type | Before | After |",
        "| --- | ---: | ---: |",
    ]
    lines.extend(
        _comparison_row(
            key,
            before_revisions.get(key, 0),
            after_revisions.get(key, 0),
        )
        for key in keys
    )
    return [*lines, ""]


def _style_section(value: object) -> list[str]:
    styles = _mapping(value)
    lines = ["## Style inventory", "", "| Field | Value |", "| --- | ---: |"]
    lines.extend(f"| `{key}` | {_value(styles[key])} |" for key in sorted(styles))
    return [*lines, ""]


def _style_comparison(before: dict[str, object], after: dict[str, object]) -> list[str]:
    before_styles = _mapping(before.get("styles", {}))
    after_styles = _mapping(after.get("styles", {}))
    keys = sorted(set(before_styles) | set(after_styles))
    lines = [
        "## Style inventory",
        "",
        "| Field | Before | After |",
        "| --- | ---: | ---: |",
    ]
    lines.extend(
        _comparison_row(
            key,
            before_styles.get(key, 0),
            after_styles.get(key, 0),
        )
        for key in keys
    )
    return [*lines, ""]


def _embedded_object_section(value: object) -> list[str]:
    return _inventory_section("Embedded object inventory", value)


def _alternative_format_import_section(value: object) -> list[str]:
    return _inventory_section("Alternative-format import inventory", value)


def _document_property_section(value: object) -> list[str]:
    return _inventory_section("Document property inventory", value)


def _mail_merge_section(value: object) -> list[str]:
    return _inventory_section("Mail-merge inventory", value)


def _inventory_section(title: str, value: object) -> list[str]:
    inventory = _mapping(value)
    lines = [f"## {title}", "", "| Field | Value |", "| --- | ---: |"]
    lines.extend(f"| `{key}` | {_value(inventory[key])} |" for key in sorted(inventory))
    return [*lines, ""]


def _embedded_object_comparison(
    before: dict[str, object], after: dict[str, object]
) -> list[str]:
    return _inventory_comparison(
        "Embedded object inventory",
        before.get("embedded_objects", {}),
        after.get("embedded_objects", {}),
    )


def _alternative_format_import_comparison(
    before: dict[str, object], after: dict[str, object]
) -> list[str]:
    return _inventory_comparison(
        "Alternative-format import inventory",
        before.get("alternative_format_imports", {}),
        after.get("alternative_format_imports", {}),
    )


def _document_property_comparison(
    before: dict[str, object], after: dict[str, object]
) -> list[str]:
    return _inventory_comparison(
        "Document property inventory",
        before.get("document_properties", {}),
        after.get("document_properties", {}),
    )


def _mail_merge_comparison(
    before: dict[str, object], after: dict[str, object]
) -> list[str]:
    return _inventory_comparison(
        "Mail-merge inventory",
        before.get("mail_merge", {}),
        after.get("mail_merge", {}),
    )


def _inventory_comparison(
    title: str, before_value: object, after_value: object
) -> list[str]:
    before_inventory = _mapping(before_value)
    after_inventory = _mapping(after_value)
    keys = sorted(set(before_inventory) | set(after_inventory))
    lines = [
        f"## {title}",
        "",
        "| Field | Before | After |",
        "| --- | ---: | ---: |",
    ]
    lines.extend(
        _comparison_row(
            key,
            before_inventory.get(key, 0),
            after_inventory.get(key, 0),
        )
        for key in keys
    )
    return [*lines, ""]


def _sarif(report: DiffReport) -> str:
    findings = list(report.findings)
    changes = list(report.changes)
    rules_by_id: dict[str, dict[str, object]] = {}
    results: list[dict[str, object]] = []
    for change in changes:
        rule_id = _change_rule_id(change)
        rules_by_id.setdefault(rule_id, _sarif_change_rule(change))
        results.append(_sarif_change_result(change, rule_id))
    for finding in findings:
        rules_by_id.setdefault(finding.rule_id, _sarif_rule(finding))
        results.append(_sarif_result(finding))
    payload = {
        "$schema": "https://json.schemastore.org/sarif-2.1.0.json",
        "version": "2.1.0",
        "runs": [
            {
                "tool": {
                    "driver": {
                        "name": "DocFence",
                        "version": __version__,
                        "informationUri": "https://github.com/SybilGambleyyu/docfence",
                        "rules": list(rules_by_id.values()),
                    }
                },
                "results": results,
            }
        ],
    }
    return _json(payload)


def _sarif_rule(finding: Finding) -> dict[str, object]:
    return {
        "id": finding.rule_id,
        "shortDescription": {"text": finding.message},
        "properties": {"severity": finding.severity},
    }


def _change_rule_id(change: Change) -> str:
    return f"DFC_{change.kind.upper()}"


def _sarif_change_rule(change: Change) -> dict[str, object]:
    return {
        "id": _change_rule_id(change),
        "shortDescription": {"text": change.message},
        "properties": {"category": "change"},
    }


def _sarif_change_result(change: Change, rule_id: str) -> dict[str, object]:
    return {
        "ruleId": rule_id,
        "level": "note",
        "message": {"text": change.message},
        "properties": {"category": "change", **change.details},
    }


def _sarif_result(finding: Finding) -> dict[str, object]:
    level = "error" if finding.severity in {"critical", "high"} else "warning"
    return {
        "ruleId": finding.rule_id,
        "level": level,
        "message": {"text": finding.message},
        "properties": finding.details,
    }


def _mapping(value: object) -> dict[str, object]:
    if isinstance(value, dict):
        return value
    return {}


def _comparison_row(key: str, before: object, after: object) -> str:
    before_value = _value(before)
    after_value = _value(after)
    return f"| `{key}` | {before_value} | {after_value} |"


def _value(value: object) -> str:
    if isinstance(value, bool):
        return "true" if value else "false"
    if isinstance(value, (int, float)):
        return str(value)
    if isinstance(value, str):
        return value.replace("|", "\\|").replace("\n", " ")
    return json.dumps(value, sort_keys=True).replace("|", "\\|").replace("\n", " ")
