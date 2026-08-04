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
    "require_no_save_through_xslt": "DFP069",
    "no_save_through_xslt_changes": "DFP070",
    "require_no_attached_custom_xml_schemas": "DFP071",
    "no_attached_custom_xml_schema_changes": "DFP072",
    "require_no_field_updates_on_open": "DFP073",
    "no_field_update_on_open_changes": "DFP074",
    "require_no_template_style_updates_on_open": "DFP075",
    "no_template_style_update_on_open_changes": "DFP076",
    "require_personal_information_removal_on_save": "DFP077",
    "no_personal_information_removal_on_save_changes": "DFP078",
    "require_no_custom_xml_data": "DFP079",
    "require_no_package_thumbnails": "DFP080",
    "no_package_thumbnail_changes": "DFP081",
    "require_no_markup_compatibility": "DFP082",
    "no_markup_compatibility_changes": "DFP083",
    "require_no_data_bindings": "DFP023",
    "no_data_binding_changes": "DFP024",
    "require_no_external_document_dependencies": "DFP025",
    "no_external_document_dependency_changes": "DFP026",
    "require_no_external_fields": "DFP027",
    "no_external_field_changes": "DFP028",
    "require_no_modern_comment_metadata": "DFP029",
    "no_modern_comment_metadata_changes": "DFP030",
    "require_no_document_tasks": "DFP031",
    "no_document_task_changes": "DFP032",
    "require_no_taskpane_web_extensions": "DFP033",
    "no_taskpane_web_extension_changes": "DFP034",
    "require_no_sensitivity_label_metadata": "DFP035",
    "no_sensitivity_label_metadata_changes": "DFP036",
    "require_no_package_digital_signatures": "DFP037",
    "no_package_digital_signature_changes": "DFP038",
    "require_no_word_protection": "DFP039",
    "no_word_protection_changes": "DFP040",
    "require_no_word_permission_ranges": "DFP041",
    "no_word_permission_range_changes": "DFP042",
    "require_no_word_document_variables": "DFP043",
    "no_word_document_variable_changes": "DFP044",
    "require_no_word_document_variable_fields": "DFP045",
    "no_word_document_variable_field_changes": "DFP046",
    "require_no_word_hyperlink_fields": "DFP047",
    "no_word_hyperlink_field_changes": "DFP048",
    "require_no_word_hyperlink_markup": "DFP049",
    "no_word_hyperlink_markup_changes": "DFP050",
    "require_no_word_drawing_hyperlinks": "DFP051",
    "no_word_drawing_hyperlink_changes": "DFP052",
    "require_no_hidden_drawing_objects": "DFP084",
    "no_drawing_object_visibility_changes": "DFP085",
    "require_no_word_vml_hyperlinks": "DFP053",
    "no_word_vml_hyperlink_changes": "DFP054",
    "require_no_word_drawing_linked_pictures": "DFP055",
    "no_word_drawing_linked_picture_changes": "DFP056",
    "require_no_word_vml_external_images": "DFP057",
    "no_word_vml_external_image_changes": "DFP058",
    "require_no_word_vml_image_hyperlinks": "DFP059",
    "no_word_vml_image_hyperlink_changes": "DFP060",
    "require_no_word_vml_linked_ole_objects": "DFP061",
    "no_word_vml_linked_ole_object_changes": "DFP062",
    "require_no_word_object_links": "DFP063",
    "no_word_object_link_changes": "DFP064",
    "require_no_word_embedded_controls": "DFP065",
    "no_word_embedded_control_changes": "DFP066",
    "require_no_word_vml_embedded_ole_objects": "DFP067",
    "no_word_vml_embedded_ole_object_changes": "DFP068",
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
    if (
        policy.enabled("no_package_thumbnail_changes")
        and before.package_thumbnails.signature != after.package_thumbnails.signature
    ):
        findings.append(
            _finding(
                "no_package_thumbnail_changes",
                "OPC package thumbnail inventory changed.",
                {
                    "before": before.package_thumbnails.public_dict(),
                    "after": after.package_thumbnails.public_dict(),
                },
            )
        )
    if (
        policy.enabled("no_markup_compatibility_changes")
        and before.markup_compatibility.signature
        != after.markup_compatibility.signature
    ):
        findings.append(
            _finding(
                "no_markup_compatibility_changes",
                "OOXML markup-compatibility inventory changed.",
                {
                    "before": before.markup_compatibility.public_dict(),
                    "after": after.markup_compatibility.public_dict(),
                },
            )
        )
    if policy.enabled("no_sensitivity_label_metadata_changes") and (
        before.sensitivity_labels.signature != after.sensitivity_labels.signature
    ):
        findings.append(
            _finding(
                "no_sensitivity_label_metadata_changes",
                "Office sensitivity-label metadata inventory changed.",
                {
                    "before": before.sensitivity_labels.public_dict(),
                    "after": after.sensitivity_labels.public_dict(),
                },
            )
        )
    if policy.enabled("no_package_digital_signature_changes") and (
        before.package_digital_signatures.signature
        != after.package_digital_signatures.signature
    ):
        findings.append(
            _finding(
                "no_package_digital_signature_changes",
                "OPC package digital-signature inventory changed.",
                {
                    "before": before.package_digital_signatures.public_dict(),
                    "after": after.package_digital_signatures.public_dict(),
                },
            )
        )
    if policy.enabled("no_word_protection_changes") and (
        before.word_protection.signature != after.word_protection.signature
    ):
        findings.append(
            _finding(
                "no_word_protection_changes",
                "Stored Word editing and write-protection inventory changed.",
                {
                    "before": before.word_protection.public_dict(),
                    "after": after.word_protection.public_dict(),
                },
            )
        )
    if policy.enabled("no_word_permission_range_changes") and (
        before.word_permission_ranges.signature
        != after.word_permission_ranges.signature
    ):
        findings.append(
            _finding(
                "no_word_permission_range_changes",
                "Stored Word editable-range permission inventory changed.",
                {
                    "before": before.word_permission_ranges.public_dict(),
                    "after": after.word_permission_ranges.public_dict(),
                },
            )
        )
    if policy.enabled("no_word_document_variable_changes") and (
        before.word_document_variables.signature
        != after.word_document_variables.signature
    ):
        findings.append(
            _finding(
                "no_word_document_variable_changes",
                "Stored Word document-variable inventory changed.",
                {
                    "before": before.word_document_variables.public_dict(),
                    "after": after.word_document_variables.public_dict(),
                },
            )
        )
    if policy.enabled("no_word_document_variable_field_changes") and (
        before.word_document_variable_fields.signature
        != after.word_document_variable_fields.signature
    ):
        findings.append(
            _finding(
                "no_word_document_variable_field_changes",
                "Stored DOCVARIABLE field-reference inventory changed.",
                {
                    "before": before.word_document_variable_fields.public_dict(),
                    "after": after.word_document_variable_fields.public_dict(),
                },
            )
        )
    if policy.enabled("no_word_hyperlink_field_changes") and (
        before.word_hyperlink_fields.signature != after.word_hyperlink_fields.signature
    ):
        findings.append(
            _finding(
                "no_word_hyperlink_field_changes",
                "Stored HYPERLINK field-reference inventory changed.",
                {
                    "before": before.word_hyperlink_fields.public_dict(),
                    "after": after.word_hyperlink_fields.public_dict(),
                },
            )
        )
    if policy.enabled("no_word_hyperlink_markup_changes") and (
        before.word_hyperlink_markup.signature != after.word_hyperlink_markup.signature
    ):
        findings.append(
            _finding(
                "no_word_hyperlink_markup_changes",
                "Stored WordprocessingML hyperlink markup inventory changed.",
                {
                    "before": before.word_hyperlink_markup.public_dict(),
                    "after": after.word_hyperlink_markup.public_dict(),
                },
            )
        )
    if policy.enabled("no_word_drawing_hyperlink_changes") and (
        before.word_drawing_hyperlinks.signature
        != after.word_drawing_hyperlinks.signature
    ):
        findings.append(
            _finding(
                "no_word_drawing_hyperlink_changes",
                "Stored DrawingML hyperlink-action inventory changed.",
                {
                    "before": before.word_drawing_hyperlinks.public_dict(),
                    "after": after.word_drawing_hyperlinks.public_dict(),
                },
            )
        )
    if policy.enabled("no_drawing_object_visibility_changes") and (
        before.word_drawing_visibility.signature
        != after.word_drawing_visibility.signature
    ):
        findings.append(
            _finding(
                "no_drawing_object_visibility_changes",
                "Stored DrawingML nonvisual visibility inventory changed.",
                {
                    "before": before.word_drawing_visibility.public_dict(),
                    "after": after.word_drawing_visibility.public_dict(),
                },
            )
        )
    if policy.enabled("no_word_drawing_linked_picture_changes") and (
        before.word_drawing_linked_pictures.signature
        != after.word_drawing_linked_pictures.signature
    ):
        findings.append(
            _finding(
                "no_word_drawing_linked_picture_changes",
                "Stored DrawingML linked-picture inventory changed.",
                {
                    "before": before.word_drawing_linked_pictures.public_dict(),
                    "after": after.word_drawing_linked_pictures.public_dict(),
                },
            )
        )
    if policy.enabled("no_word_vml_hyperlink_changes") and (
        before.word_vml_hyperlinks.signature != after.word_vml_hyperlinks.signature
    ):
        findings.append(
            _finding(
                "no_word_vml_hyperlink_changes",
                "Stored VML hyperlink markup inventory changed.",
                {
                    "before": before.word_vml_hyperlinks.public_dict(),
                    "after": after.word_vml_hyperlinks.public_dict(),
                },
            )
        )
    if policy.enabled("no_word_vml_external_image_changes") and (
        before.word_vml_external_images.signature
        != after.word_vml_external_images.signature
    ):
        findings.append(
            _finding(
                "no_word_vml_external_image_changes",
                "Stored VML external-image inventory changed.",
                {
                    "before": before.word_vml_external_images.public_dict(),
                    "after": after.word_vml_external_images.public_dict(),
                },
            )
        )
    if policy.enabled("no_word_vml_image_hyperlink_changes") and (
        before.word_vml_image_hyperlinks.signature
        != after.word_vml_image_hyperlinks.signature
    ):
        findings.append(
            _finding(
                "no_word_vml_image_hyperlink_changes",
                "Stored VML image-data hyperlink inventory changed.",
                {
                    "before": before.word_vml_image_hyperlinks.public_dict(),
                    "after": after.word_vml_image_hyperlinks.public_dict(),
                },
            )
        )
    if policy.enabled("no_word_vml_linked_ole_object_changes") and (
        before.word_vml_linked_ole_objects.signature
        != after.word_vml_linked_ole_objects.signature
    ):
        findings.append(
            _finding(
                "no_word_vml_linked_ole_object_changes",
                "Stored VML linked-OLE-object inventory changed.",
                {
                    "before": before.word_vml_linked_ole_objects.public_dict(),
                    "after": after.word_vml_linked_ole_objects.public_dict(),
                },
            )
        )
    if policy.enabled("no_word_vml_embedded_ole_object_changes") and (
        before.word_vml_embedded_ole_objects.signature
        != after.word_vml_embedded_ole_objects.signature
    ):
        findings.append(
            _finding(
                "no_word_vml_embedded_ole_object_changes",
                "Stored VML embedded-OLE-object inventory changed.",
                {
                    "before": before.word_vml_embedded_ole_objects.public_dict(),
                    "after": after.word_vml_embedded_ole_objects.public_dict(),
                },
            )
        )
    if policy.enabled("no_word_object_link_changes") and (
        before.word_object_links.signature != after.word_object_links.signature
    ):
        findings.append(
            _finding(
                "no_word_object_link_changes",
                "Stored WordprocessingML linked-object-property inventory changed.",
                {
                    "before": before.word_object_links.public_dict(),
                    "after": after.word_object_links.public_dict(),
                },
            )
        )
    if policy.enabled("no_word_embedded_control_changes") and (
        before.word_embedded_controls.signature
        != after.word_embedded_controls.signature
    ):
        findings.append(
            _finding(
                "no_word_embedded_control_changes",
                "Stored WordprocessingML embedded-control-anchor inventory changed.",
                {
                    "before": before.word_embedded_controls.public_dict(),
                    "after": after.word_embedded_controls.public_dict(),
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
    if policy.enabled("no_save_through_xslt_changes") and (
        before.save_through_xslt.signature != after.save_through_xslt.signature
    ):
        findings.append(
            _finding(
                "no_save_through_xslt_changes",
                "XSLT-on-single-XML-save inventory changed.",
                {
                    "before": before.save_through_xslt.public_dict(),
                    "after": after.save_through_xslt.public_dict(),
                },
            )
        )
    if policy.enabled("no_attached_custom_xml_schema_changes") and (
        before.attached_custom_xml_schemas.signature
        != after.attached_custom_xml_schemas.signature
    ):
        findings.append(
            _finding(
                "no_attached_custom_xml_schema_changes",
                "Attached custom XML schema inventory changed.",
                {
                    "before": before.attached_custom_xml_schemas.public_dict(),
                    "after": after.attached_custom_xml_schemas.public_dict(),
                },
            )
        )
    if policy.enabled("no_field_update_on_open_changes") and (
        before.field_updates_on_open.signature != after.field_updates_on_open.signature
    ):
        findings.append(
            _finding(
                "no_field_update_on_open_changes",
                "Automatic field-update-on-open inventory changed.",
                {
                    "before": before.field_updates_on_open.public_dict(),
                    "after": after.field_updates_on_open.public_dict(),
                },
            )
        )
    if policy.enabled("no_template_style_update_on_open_changes") and (
        before.template_style_updates_on_open.signature
        != after.template_style_updates_on_open.signature
    ):
        findings.append(
            _finding(
                "no_template_style_update_on_open_changes",
                "Automatic template-style-update-on-open inventory changed.",
                {
                    "before": before.template_style_updates_on_open.public_dict(),
                    "after": after.template_style_updates_on_open.public_dict(),
                },
            )
        )
    if policy.enabled("no_personal_information_removal_on_save_changes") and (
        before.personal_information_removal_on_save.signature
        != after.personal_information_removal_on_save.signature
    ):
        findings.append(
            _finding(
                "no_personal_information_removal_on_save_changes",
                "Personal-information-removal-on-save inventory changed.",
                {
                    "before": before.personal_information_removal_on_save.public_dict(),
                    "after": after.personal_information_removal_on_save.public_dict(),
                },
            )
        )
    if policy.enabled("no_data_binding_changes") and (
        before.data_bindings.signature != after.data_bindings.signature
    ):
        findings.append(
            _finding(
                "no_data_binding_changes",
                "Content-control data-binding inventory changed.",
                {
                    "before": before.data_bindings.public_dict(),
                    "after": after.data_bindings.public_dict(),
                },
            )
        )
    if policy.enabled("no_external_field_changes") and (
        before.external_fields.signature != after.external_fields.signature
    ):
        findings.append(
            _finding(
                "no_external_field_changes",
                "External-source Word field inventory changed.",
                {
                    "before": before.external_fields.public_dict(),
                    "after": after.external_fields.public_dict(),
                },
            )
        )
    if policy.enabled("no_modern_comment_metadata_changes") and (
        before.modern_comment_metadata.signature
        != after.modern_comment_metadata.signature
    ):
        findings.append(
            _finding(
                "no_modern_comment_metadata_changes",
                "Modern Word comment metadata inventory changed.",
                {
                    "before": before.modern_comment_metadata.public_dict(),
                    "after": after.modern_comment_metadata.public_dict(),
                },
            )
        )
    if policy.enabled("no_document_task_changes") and (
        before.document_tasks.signature != after.document_tasks.signature
    ):
        findings.append(
            _finding(
                "no_document_task_changes",
                "Word document task inventory changed.",
                {
                    "before": before.document_tasks.public_dict(),
                    "after": after.document_tasks.public_dict(),
                },
            )
        )
    if policy.enabled("no_taskpane_web_extension_changes") and (
        before.taskpane_web_extensions.signature
        != after.taskpane_web_extensions.signature
    ):
        findings.append(
            _finding(
                "no_taskpane_web_extension_changes",
                "Task-pane Office web extension inventory changed.",
                {
                    "before": before.taskpane_web_extensions.public_dict(),
                    "after": after.taskpane_web_extensions.public_dict(),
                },
            )
        )
    if policy.enabled("no_external_document_dependency_changes") and (
        before.external_document_dependencies.signature
        != after.external_document_dependencies.signature
    ):
        findings.append(
            _finding(
                "no_external_document_dependency_changes",
                "External Word document dependency inventory changed.",
                {
                    "before": before.external_document_dependencies.public_dict(),
                    "after": after.external_document_dependencies.public_dict(),
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
    if policy.enabled("require_no_custom_xml_data") and after.custom_xml_part_count:
        findings.append(
            _finding(
                "require_no_custom_xml_data",
                "Candidate contains stored custom XML package parts.",
                {"custom_xml_part_count": after.custom_xml_part_count},
            )
        )
    if policy.enabled("require_no_package_thumbnails") and (
        after.package_thumbnails.thumbnail_part_count
    ):
        findings.append(
            _finding(
                "require_no_package_thumbnails",
                "Candidate contains relationship-bound OPC package thumbnail images.",
                after.package_thumbnails.public_dict(),
            )
        )
    if (
        policy.enabled("require_no_markup_compatibility")
        and after.markup_compatibility.markup_compatibility_part_count
    ):
        findings.append(
            _finding(
                "require_no_markup_compatibility",
                "Candidate contains stored OOXML markup-compatibility markup.",
                after.markup_compatibility.public_dict(),
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
    if policy.enabled("require_no_sensitivity_label_metadata") and any(
        after.sensitivity_labels.public_dict().values()
    ):
        findings.append(
            _finding(
                "require_no_sensitivity_label_metadata",
                "Candidate contains stored Office sensitivity-label metadata.",
                after.sensitivity_labels.public_dict(),
            )
        )
    if policy.enabled("require_no_package_digital_signatures") and any(
        after.package_digital_signatures.public_dict().values()
    ):
        findings.append(
            _finding(
                "require_no_package_digital_signatures",
                "Candidate contains stored OPC package digital-signature material.",
                after.package_digital_signatures.public_dict(),
            )
        )
    if policy.enabled("require_no_word_protection") and any(
        after.word_protection.public_dict().values()
    ):
        findings.append(
            _finding(
                "require_no_word_protection",
                "Candidate contains stored Word editing or write-protection state.",
                after.word_protection.public_dict(),
            )
        )
    if policy.enabled("require_no_word_permission_ranges") and any(
        after.word_permission_ranges.public_dict().values()
    ):
        findings.append(
            _finding(
                "require_no_word_permission_ranges",
                "Candidate contains stored Word editable-range permission markup.",
                after.word_permission_ranges.public_dict(),
            )
        )
    if policy.enabled("require_no_word_document_variables") and any(
        after.word_document_variables.public_dict().values()
    ):
        findings.append(
            _finding(
                "require_no_word_document_variables",
                "Candidate contains stored Word document-variable state.",
                after.word_document_variables.public_dict(),
            )
        )
    if policy.enabled("require_no_word_document_variable_fields") and any(
        after.word_document_variable_fields.public_dict().values()
    ):
        findings.append(
            _finding(
                "require_no_word_document_variable_fields",
                "Candidate contains stored DOCVARIABLE field references.",
                after.word_document_variable_fields.public_dict(),
            )
        )
    if policy.enabled("require_no_word_hyperlink_fields") and any(
        after.word_hyperlink_fields.public_dict().values()
    ):
        findings.append(
            _finding(
                "require_no_word_hyperlink_fields",
                "Candidate contains stored HYPERLINK field references.",
                after.word_hyperlink_fields.public_dict(),
            )
        )
    if policy.enabled("require_no_word_hyperlink_markup") and any(
        after.word_hyperlink_markup.public_dict().values()
    ):
        findings.append(
            _finding(
                "require_no_word_hyperlink_markup",
                "Candidate contains stored WordprocessingML hyperlink markup.",
                after.word_hyperlink_markup.public_dict(),
            )
        )
    if policy.enabled("require_no_word_drawing_hyperlinks") and any(
        after.word_drawing_hyperlinks.public_dict().values()
    ):
        findings.append(
            _finding(
                "require_no_word_drawing_hyperlinks",
                "Candidate contains stored DrawingML hyperlink-action markup.",
                after.word_drawing_hyperlinks.public_dict(),
            )
        )
    if policy.enabled("require_no_hidden_drawing_objects") and (
        after.word_drawing_visibility.hidden_drawing_object_count
        or after.word_drawing_visibility.invalid_hidden_attribute_count
    ):
        findings.append(
            _finding(
                "require_no_hidden_drawing_objects",
                (
                    "Candidate contains stored hidden or invalid DrawingML "
                    "nonvisual visibility declarations."
                ),
                after.word_drawing_visibility.public_dict(),
            )
        )
    if policy.enabled("require_no_word_drawing_linked_pictures") and any(
        after.word_drawing_linked_pictures.public_dict().values()
    ):
        findings.append(
            _finding(
                "require_no_word_drawing_linked_pictures",
                "Candidate contains stored DrawingML linked-picture markup.",
                after.word_drawing_linked_pictures.public_dict(),
            )
        )
    if policy.enabled("require_no_word_vml_hyperlinks") and any(
        after.word_vml_hyperlinks.public_dict().values()
    ):
        findings.append(
            _finding(
                "require_no_word_vml_hyperlinks",
                "Candidate contains stored VML hyperlink markup.",
                after.word_vml_hyperlinks.public_dict(),
            )
        )
    if policy.enabled("require_no_word_vml_external_images") and any(
        after.word_vml_external_images.public_dict().values()
    ):
        findings.append(
            _finding(
                "require_no_word_vml_external_images",
                "Candidate contains stored VML external-image markup.",
                after.word_vml_external_images.public_dict(),
            )
        )
    if policy.enabled("require_no_word_vml_image_hyperlinks") and any(
        after.word_vml_image_hyperlinks.public_dict().values()
    ):
        findings.append(
            _finding(
                "require_no_word_vml_image_hyperlinks",
                "Candidate contains stored VML image-data hyperlink markup.",
                after.word_vml_image_hyperlinks.public_dict(),
            )
        )
    if policy.enabled("require_no_word_vml_linked_ole_objects") and any(
        after.word_vml_linked_ole_objects.public_dict().values()
    ):
        findings.append(
            _finding(
                "require_no_word_vml_linked_ole_objects",
                "Candidate contains stored VML linked-OLE-object markup.",
                after.word_vml_linked_ole_objects.public_dict(),
            )
        )
    if policy.enabled("require_no_word_vml_embedded_ole_objects") and any(
        after.word_vml_embedded_ole_objects.public_dict().values()
    ):
        findings.append(
            _finding(
                "require_no_word_vml_embedded_ole_objects",
                "Candidate contains stored VML embedded-OLE-object markup.",
                after.word_vml_embedded_ole_objects.public_dict(),
            )
        )
    if policy.enabled("require_no_word_object_links") and any(
        after.word_object_links.public_dict().values()
    ):
        findings.append(
            _finding(
                "require_no_word_object_links",
                (
                    "Candidate contains stored WordprocessingML "
                    "linked-object-property markup."
                ),
                after.word_object_links.public_dict(),
            )
        )
    if policy.enabled("require_no_word_embedded_controls") and any(
        after.word_embedded_controls.public_dict().values()
    ):
        findings.append(
            _finding(
                "require_no_word_embedded_controls",
                "Candidate contains stored WordprocessingML embedded-control anchors.",
                after.word_embedded_controls.public_dict(),
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
    if policy.enabled("require_no_save_through_xslt") and any(
        after.save_through_xslt.public_dict().values()
    ):
        findings.append(
            _finding(
                "require_no_save_through_xslt",
                "Candidate contains stored XSLT-on-single-XML-save configuration.",
                after.save_through_xslt.public_dict(),
            )
        )
    if policy.enabled("require_no_attached_custom_xml_schemas") and any(
        after.attached_custom_xml_schemas.public_dict().values()
    ):
        findings.append(
            _finding(
                "require_no_attached_custom_xml_schemas",
                "Candidate contains attached custom XML schema declarations.",
                after.attached_custom_xml_schemas.public_dict(),
            )
        )
    if policy.enabled("require_no_field_updates_on_open") and (
        after.field_updates_on_open.enabled_setting_count
    ):
        findings.append(
            _finding(
                "require_no_field_updates_on_open",
                "Candidate requests automatic field recalculation on open.",
                after.field_updates_on_open.public_dict(),
            )
        )
    if policy.enabled("require_no_template_style_updates_on_open") and (
        after.template_style_updates_on_open.enabled_setting_count
    ):
        findings.append(
            _finding(
                "require_no_template_style_updates_on_open",
                "Candidate enables automatic template-style updates on open.",
                after.template_style_updates_on_open.public_dict(),
            )
        )
    if policy.enabled("require_personal_information_removal_on_save") and not (
        after.personal_information_removal_on_save.enabled_setting_count
    ):
        findings.append(
            _finding(
                "require_personal_information_removal_on_save",
                "Candidate does not request personal-information removal on save.",
                after.personal_information_removal_on_save.public_dict(),
            )
        )
    if policy.enabled("require_no_data_bindings") and (
        after.data_bindings.binding_count
    ):
        findings.append(
            _finding(
                "require_no_data_bindings",
                "Candidate contains stored content-control data bindings.",
                after.data_bindings.public_dict(),
            )
        )
    if policy.enabled("require_no_external_fields") and any(
        after.external_fields.public_dict().values()
    ):
        findings.append(
            _finding(
                "require_no_external_fields",
                "Candidate contains stored external-source Word field instructions.",
                after.external_fields.public_dict(),
            )
        )
    if policy.enabled("require_no_modern_comment_metadata") and any(
        after.modern_comment_metadata.public_dict().values()
    ):
        findings.append(
            _finding(
                "require_no_modern_comment_metadata",
                "Candidate contains stored modern Word comment metadata.",
                after.modern_comment_metadata.public_dict(),
            )
        )
    if policy.enabled("require_no_document_tasks") and any(
        after.document_tasks.public_dict().values()
    ):
        findings.append(
            _finding(
                "require_no_document_tasks",
                "Candidate contains stored Word document tasks.",
                after.document_tasks.public_dict(),
            )
        )
    if policy.enabled("require_no_taskpane_web_extensions") and any(
        after.taskpane_web_extensions.public_dict().values()
    ):
        findings.append(
            _finding(
                "require_no_taskpane_web_extensions",
                "Candidate contains stored task-pane Office web extension state.",
                after.taskpane_web_extensions.public_dict(),
            )
        )
    if policy.enabled("require_no_external_document_dependencies") and any(
        after.external_document_dependencies.public_dict().values()
    ):
        findings.append(
            _finding(
                "require_no_external_document_dependencies",
                "Candidate contains stored external Word document dependencies.",
                after.external_document_dependencies.public_dict(),
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
