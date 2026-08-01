from __future__ import annotations

import json
import zipfile
from dataclasses import replace

import pytest
from conftest import CT, DOCX_MAIN_TYPE, PR, R, W, write_document

from docfence.cli import main
from docfence.diff import diff_documents
from docfence.errors import DocumentFormatError, DocumentSafetyError, PolicyError
from docfence.output import render_profile, render_report
from docfence.policy import apply_policy, load_policy, starter_policy
from docfence.snapshot import load_snapshot

_ALT_CHUNK_RELATIONSHIP_TYPE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/aFChunk"
)
_ACTIVE_X_CONTROL_BINARY_RELATIONSHIP_TYPE = (
    "http://schemas.microsoft.com/office/2006/relationships/activeXControlBinary"
)
_CONTROL_RELATIONSHIP_TYPE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/control"
)
_OLE_OBJECT_RELATIONSHIP_TYPE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject"
)
_PACKAGE_RELATIONSHIP_TYPE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/package"
)
_MAIL_MERGE_SOURCE_RELATIONSHIP_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/mailMergeSource"
_MAIL_MERGE_HEADER_SOURCE_RELATIONSHIP_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/mailMergeHeaderSource"
_MAIL_MERGE_RECIPIENT_DATA_RELATIONSHIP_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/mailMergeRecipientData"
_RECIPIENT_DATA_RELATIONSHIP_TYPE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/recipientData"
)
_STRICT_MAIL_MERGE_SOURCE_RELATIONSHIP_TYPE = (
    "http://purl.oclc.org/ooxml/officeDocument/relationships/mailMergeSource"
)
_STRICT_MAIL_MERGE_HEADER_SOURCE_RELATIONSHIP_TYPE = (
    "http://purl.oclc.org/ooxml/officeDocument/relationships/mailMergeHeaderSource"
)
_STRICT_MAIL_MERGE_RECIPIENT_DATA_RELATIONSHIP_TYPE = (
    "http://purl.oclc.org/ooxml/officeDocument/relationships/mailMergeRecipientData"
)
_CUSTOM_XML_RELATIONSHIP_TYPE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml"
)
_CUSTOM_XML_PROPERTIES_RELATIONSHIP_TYPE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps"
)
_STRICT_CUSTOM_XML_RELATIONSHIP_TYPE = (
    "http://purl.oclc.org/ooxml/officeDocument/relationships/customXml"
)
_STRICT_CUSTOM_XML_PROPERTIES_RELATIONSHIP_TYPE = (
    "http://purl.oclc.org/ooxml/officeDocument/relationships/customXmlProps"
)
_ATTACHED_TEMPLATE_RELATIONSHIP_TYPE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
    "attachedTemplate"
)
_SUBDOCUMENT_RELATIONSHIP_TYPE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/subDocument"
)
_FRAME_RELATIONSHIP_TYPE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/frame"
)
_WEB_SETTINGS_RELATIONSHIP_TYPE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings"
)
_STRICT_ATTACHED_TEMPLATE_RELATIONSHIP_TYPE = (
    "http://purl.oclc.org/ooxml/officeDocument/relationships/attachedTemplate"
)
_STRICT_SUBDOCUMENT_RELATIONSHIP_TYPE = (
    "http://purl.oclc.org/ooxml/officeDocument/relationships/subDocument"
)
_STRICT_FRAME_RELATIONSHIP_TYPE = (
    "http://purl.oclc.org/ooxml/officeDocument/relationships/frame"
)
_STRICT_DOCUMENT_SETTINGS_RELATIONSHIP_TYPE = (
    "http://purl.oclc.org/ooxml/officeDocument/relationships/settings"
)
_STRICT_WEB_SETTINGS_RELATIONSHIP_TYPE = (
    "http://purl.oclc.org/ooxml/officeDocument/relationships/webSettings"
)
_CUSTOM_XML_DATA_PROPERTIES_NAMESPACE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/customXml"
)
_STRICT_CUSTOM_XML_DATA_PROPERTIES_NAMESPACE = (
    "http://purl.oclc.org/ooxml/officeDocument/customXmlDataProps"
)
_STRICT_WORD_NAMESPACE = "http://purl.oclc.org/ooxml/wordprocessingml/main"
_STRICT_RELATIONSHIP_NAMESPACE = (
    "http://purl.oclc.org/ooxml/officeDocument/relationships"
)
_CORE_PROPERTIES_RELATIONSHIP_TYPE = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"
_EXTENDED_PROPERTIES_RELATIONSHIP_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties"
_CUSTOM_PROPERTIES_RELATIONSHIP_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties"
_STRICT_EXTENDED_PROPERTIES_RELATIONSHIP_TYPE = (
    "http://purl.oclc.org/ooxml/officeDocument/relationships/extendedProperties"
)
_STRICT_CUSTOM_PROPERTIES_RELATIONSHIP_TYPE = (
    "http://purl.oclc.org/ooxml/officeDocument/relationships/customProperties"
)
_CORE_PROPERTIES_NAMESPACE = (
    "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
)
_EXTENDED_PROPERTIES_NAMESPACE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
)
_CUSTOM_PROPERTIES_NAMESPACE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
)
_STRICT_EXTENDED_PROPERTIES_NAMESPACE = (
    "http://purl.oclc.org/ooxml/officeDocument/extendedProperties"
)
_STRICT_CUSTOM_PROPERTIES_NAMESPACE = (
    "http://purl.oclc.org/ooxml/officeDocument/customProperties"
)
_DUBLIN_CORE_NAMESPACE = "http://purl.org/dc/elements/1.1/"
_DOCUMENT_PROPERTIES_VT_NAMESPACE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"
)
_STRICT_DOCUMENT_PROPERTIES_VT_NAMESPACE = (
    "http://purl.oclc.org/ooxml/officeDocument/docPropsVTypes"
)


def test_profile_counts_review_surfaces_without_material_leaks(tmp_path) -> None:
    document = tmp_path / "candidate.docm"
    write_document(
        document,
        text="VISIBLE_DO_NOT_LEAK",
        relationship_id="rId1",
        relationship_target="https://example.invalid/EXTERNAL_DO_NOT_LEAK",
        relationship_external=True,
        revision=True,
        hidden=True,
        field=True,
        control=True,
        comment=True,
        track_revisions=True,
        custom_xml=b"CUSTOM_XML_DO_NOT_LEAK",
        macro=b"MACRO_DO_NOT_LEAK",
        unclassified=b"UNCLASSIFIED_DO_NOT_LEAK",
    )

    snapshot = load_snapshot(document)
    public = snapshot.public_dict()

    assert public["paragraph_count"] == 2
    assert public["hidden_text_run_count"] == 1
    assert public["hidden_paragraph_mark_count"] == 0
    assert public["styles"] == {
        "style_definition_count": 0,
        "hidden_text_style_definition_count": 0,
        "document_default_hidden_text_enabled": False,
    }
    assert public["embedded_objects"] == {
        "embedded_object_relationship_count": 0,
        "embedded_object_part_count": 0,
        "embedded_control_relationship_count": 0,
        "embedded_control_part_count": 0,
    }
    assert public["alternative_format_imports"] == {
        "alternative_format_import_relationship_count": 0,
        "alternative_format_import_payload_part_count": 0,
    }
    assert public["alternative_format_import_anchor_count"] == 0
    assert public["document_properties"] == {
        "core_property_part_count": 0,
        "core_property_value_count": 0,
        "extended_property_part_count": 0,
        "extended_property_value_count": 0,
        "custom_property_part_count": 0,
        "custom_property_count": 0,
    }
    assert public["field_code_count"] == 1
    assert public["content_control_count"] == 1
    assert public["comment_anchor_count"] == 1
    assert public["comment_count"] == 1
    assert public["track_revisions_enabled"] is True
    assert public["revisions"]["insertions"] == 1
    assert public["relationships"]["external_relationship_count"] == 1
    assert public["custom_xml_part_count"] == 1
    assert public["macro_present"] is True
    assert public["unclassified_part_count"] == 2

    rendered = render_profile(snapshot, "json") + render_profile(snapshot, "markdown")
    for marker in (
        "VISIBLE_DO_NOT_LEAK",
        "HIDDEN_DO_NOT_LEAK",
        "FIELD_DO_NOT_LEAK",
        "COMMENT_DO_NOT_LEAK",
        "REVIEWER_DO_NOT_LEAK",
        "EXTERNAL_DO_NOT_LEAK",
        "CUSTOM_XML_DO_NOT_LEAK",
        "MACRO_DO_NOT_LEAK",
        "UNCLASSIFIED_DO_NOT_LEAK",
    ):
        assert marker not in rendered


def test_hidden_markup_and_style_declarations_are_separate_and_private(
    tmp_path,
) -> None:
    before = tmp_path / "before.docx"
    after = tmp_path / "after.docx"
    policy_path = tmp_path / "docfence.yml"
    write_document(before)
    write_document(
        after,
        hidden=True,
        hidden_paragraph_mark=True,
        special_hidden_paragraph_mark=True,
        styles_xml=_styles_with_hidden_text_declarations(),
    )

    snapshot = load_snapshot(after)
    public = snapshot.public_dict()
    assert public["hidden_text_run_count"] == 1
    assert public["hidden_paragraph_mark_count"] == 1
    assert public["styles"] == {
        "style_definition_count": 5,
        "hidden_text_style_definition_count": 1,
        "document_default_hidden_text_enabled": True,
    }

    report = diff_documents(before, after)
    assert {
        "hidden_text_inventory_changed",
        "hidden_paragraph_mark_inventory_changed",
        "style_inventory_changed",
    } <= {change.kind for change in report.changes}
    policy_path.write_text(
        """version: 1
rules:
  require_no_hidden_text: true
  require_no_hidden_text_style_declarations: true
  require_no_hidden_paragraph_marks: true
""",
        encoding="utf-8",
    )
    gated = apply_policy(report, load_policy(policy_path))
    assert {finding.rule_id for finding in gated.findings} == {
        "DFP006",
        "DFP013",
        "DFP014",
    }

    rendered = "\n".join(
        (
            render_profile(snapshot, "json"),
            render_profile(snapshot, "markdown"),
            render_report(gated, "json"),
            render_report(gated, "markdown"),
            render_report(gated, "sarif"),
        )
    )
    sarif = json.loads(render_report(gated, "sarif"))
    assert any(
        result["ruleId"] == "DFC_STYLE_INVENTORY_CHANGED"
        for result in sarif["runs"][0]["results"]
    )
    for marker in (
        "HIDDEN_TEXT_STYLE_DO_NOT_LEAK",
        "PARAGRAPH_MARK_STYLE_DO_NOT_LEAK",
        "SPECIAL_PARAGRAPH_MARK_STYLE_DO_NOT_LEAK",
        "TRACKED_STYLE_CHANGE_DO_NOT_LEAK",
        "VISIBLE_STYLE_DO_NOT_LEAK",
    ):
        assert marker not in rendered


def test_spec_vanish_is_inventoried_only_for_a_paragraph_mark(tmp_path) -> None:
    document = tmp_path / "special-paragraph-mark.docx"
    _write_raw_package(
        document,
        f"""<w:document xmlns:w=\"{W}\"><w:body><w:p>
  <w:pPr><w:rPr><w:specVanish/></w:rPr></w:pPr>
  <w:r><w:rPr><w:specVanish/></w:rPr><w:t>SPECIAL_DO_NOT_LEAK</w:t></w:r>
  <w:r><w:rPr><w:vanish w:val=\"false\"/></w:rPr>
  <w:t>DIRECT_FALSE_DO_NOT_LEAK</w:t></w:r>
</w:p></w:body></w:document>""",
    )

    snapshot = load_snapshot(document)
    assert snapshot.hidden_text_run_count == 0
    assert snapshot.hidden_paragraph_mark_count == 1
    rendered = render_profile(snapshot, "markdown")
    assert "SPECIAL_DO_NOT_LEAK" not in rendered
    assert "DIRECT_FALSE_DO_NOT_LEAK" not in rendered


def test_embedded_objects_and_alternative_imports_are_separate_and_private(
    tmp_path,
) -> None:
    before = tmp_path / "before.docx"
    after = tmp_path / "after.docx"
    payload_changed = tmp_path / "payload-changed.docx"
    renumbered = tmp_path / "renumbered.docx"
    duplicate_anchor = tmp_path / "duplicate-anchor.docx"
    policy_path = tmp_path / "docfence.yml"
    write_document(before)
    _write_embedded_content_document(after)
    _write_embedded_content_document(
        payload_changed,
        object_payload=b"OBJECT_PAYLOAD_CHANGED_DO_NOT_LEAK",
        import_payload=b"ALT_IMPORT_CHANGED_DO_NOT_LEAK",
    )
    _write_embedded_content_document(renumbered, relationship_id_suffix="9")
    _write_embedded_content_document(duplicate_anchor, duplicate_alt_chunk_anchor=True)

    snapshot = load_snapshot(after)
    public = snapshot.public_dict()
    assert public["alternative_format_import_anchor_count"] == 1
    assert public["embedded_objects"] == {
        "embedded_object_relationship_count": 2,
        "embedded_object_part_count": 2,
        "embedded_control_relationship_count": 2,
        "embedded_control_part_count": 2,
    }
    assert public["alternative_format_imports"] == {
        "alternative_format_import_relationship_count": 1,
        "alternative_format_import_payload_part_count": 1,
    }
    assert public["unclassified_part_count"] == 1

    report = diff_documents(before, after)
    assert {
        "embedded_object_inventory_changed",
        "alternative_format_import_inventory_changed",
        "alternative_format_import_anchor_inventory_changed",
    } <= {change.kind for change in report.changes}
    payload_report = diff_documents(after, payload_changed)
    payload_kinds = {change.kind for change in payload_report.changes}
    assert {
        "embedded_object_inventory_changed",
        "alternative_format_import_inventory_changed",
    } <= payload_kinds
    assert "alternative_format_import_anchor_inventory_changed" not in payload_kinds
    assert diff_documents(after, renumbered).changes == ()
    duplicate_anchor_report = diff_documents(after, duplicate_anchor)
    duplicate_anchor_kinds = {change.kind for change in duplicate_anchor_report.changes}
    assert (
        "alternative_format_import_anchor_inventory_changed" in duplicate_anchor_kinds
    )
    assert "alternative_format_import_inventory_changed" not in duplicate_anchor_kinds
    assert "embedded_object_inventory_changed" not in duplicate_anchor_kinds

    policy_path.write_text(
        """version: 1
rules:
  require_no_embedded_objects: true
  require_no_alternative_format_imports: true
  no_embedded_object_payload_changes: true
  no_alternative_format_import_changes: true
""",
        encoding="utf-8",
    )
    gated = apply_policy(report, load_policy(policy_path))
    assert {finding.rule_id for finding in gated.findings} == {
        "DFP015",
        "DFP016",
        "DFP017",
        "DFP018",
    }
    assert {
        finding.rule_id
        for finding in apply_policy(payload_report, load_policy(policy_path)).findings
    } == {
        "DFP015",
        "DFP016",
        "DFP017",
        "DFP018",
    }
    assert {
        finding.rule_id
        for finding in apply_policy(
            diff_documents(after, renumbered), load_policy(policy_path)
        ).findings
    } == {
        "DFP015",
        "DFP016",
    }
    assert {
        finding.rule_id
        for finding in apply_policy(
            duplicate_anchor_report, load_policy(policy_path)
        ).findings
    } == {"DFP015", "DFP016", "DFP018"}

    rendered = "\n".join(
        (
            render_profile(snapshot, "json"),
            render_profile(snapshot, "markdown"),
            render_report(gated, "json"),
            render_report(gated, "markdown"),
            render_report(gated, "sarif"),
        )
    )
    sarif = json.loads(render_report(gated, "sarif"))
    assert {
        "DFC_EMBEDDED_OBJECT_INVENTORY_CHANGED",
        "DFC_ALTERNATIVE_FORMAT_IMPORT_INVENTORY_CHANGED",
        "DFP015",
        "DFP016",
        "DFP017",
        "DFP018",
    } <= {result["ruleId"] for result in sarif["runs"][0]["results"]}
    for marker in (
        "VISIBLE_DO_NOT_LEAK",
        "OBJECT_PAYLOAD_DO_NOT_LEAK",
        "PACKAGE_PAYLOAD_DO_NOT_LEAK",
        "CONTROL_PAYLOAD_DO_NOT_LEAK",
        "ALT_IMPORT_DO_NOT_LEAK",
        "OBJECT_PAYLOAD_CHANGED_DO_NOT_LEAK",
        "ALT_IMPORT_CHANGED_DO_NOT_LEAK",
    ):
        assert marker not in rendered


def test_alt_chunk_requires_a_matching_internal_import_relationship(tmp_path) -> None:
    wrong_type = tmp_path / "wrong-type.docx"
    external_target = tmp_path / "external-target.docx"
    missing_target = tmp_path / "missing-target.docx"
    _write_embedded_content_document(
        wrong_type, alt_chunk_relationship_type=_OLE_OBJECT_RELATIONSHIP_TYPE
    )
    _write_embedded_content_document(external_target, alt_chunk_target_mode="External")
    _write_embedded_content_document(missing_target, include_import_payload=False)

    with pytest.raises(DocumentFormatError):
        load_snapshot(wrong_type)
    with pytest.raises(DocumentFormatError):
        load_snapshot(external_target)
    with pytest.raises(DocumentFormatError):
        load_snapshot(missing_target)


def test_document_property_inventory_is_private_and_relationship_id_stable(
    tmp_path,
) -> None:
    before = tmp_path / "before.docx"
    after = tmp_path / "after.docx"
    value_changed = tmp_path / "value-changed.docx"
    renumbered = tmp_path / "renumbered.docx"
    unlinked = tmp_path / "unlinked.docx"
    strict_properties = tmp_path / "strict-properties.docx"
    malformed = tmp_path / "malformed.docx"
    policy_path = tmp_path / "docfence.yml"
    _write_document_property_document(before, include_properties=False)
    _write_document_property_document(after)
    _write_document_property_document(
        value_changed, core_title="CORE_TITLE_CHANGED_DO_NOT_LEAK"
    )
    _write_document_property_document(renumbered, relationship_id_suffix="9")
    _write_document_property_document(unlinked, include_property_relationships=False)
    _write_document_property_document(strict_properties, strict_property_syntax=True)
    _write_document_property_document(malformed, malformed_core=True)

    snapshot = load_snapshot(after)
    assert snapshot.public_dict()["document_properties"] == {
        "core_property_part_count": 1,
        "core_property_value_count": 2,
        "extended_property_part_count": 1,
        "extended_property_value_count": 2,
        "custom_property_part_count": 1,
        "custom_property_count": 2,
    }
    assert snapshot.unclassified_part_count == 1
    assert (
        load_snapshot(unlinked).public_dict()["document_properties"]
        == (snapshot.public_dict()["document_properties"])
    )
    assert (
        load_snapshot(strict_properties).public_dict()["document_properties"]
        == (snapshot.public_dict()["document_properties"])
    )

    report = diff_documents(before, after)
    assert "document_property_inventory_changed" in {
        change.kind for change in report.changes
    }
    value_report = diff_documents(after, value_changed)
    assert {change.kind for change in value_report.changes} == {
        "document_property_inventory_changed"
    }
    assert diff_documents(after, renumbered).changes == ()
    with pytest.raises(DocumentFormatError):
        load_snapshot(malformed)

    policy_path.write_text(
        """version: 1
rules:
  no_document_property_changes: true
  require_no_custom_document_properties: true
""",
        encoding="utf-8",
    )
    policy = load_policy(policy_path)
    assert {finding.rule_id for finding in apply_policy(report, policy).findings} == {
        "DFP019",
        "DFP020",
    }
    assert {
        finding.rule_id for finding in apply_policy(value_report, policy).findings
    } == {"DFP019", "DFP020"}
    assert {
        finding.rule_id
        for finding in apply_policy(diff_documents(after, renumbered), policy).findings
    } == {"DFP020"}

    gated = apply_policy(report, policy)
    rendered = "\n".join(
        (
            render_profile(snapshot, "json"),
            render_profile(snapshot, "markdown"),
            render_report(gated, "json"),
            render_report(gated, "markdown"),
            render_report(gated, "sarif"),
        )
    )
    sarif = json.loads(render_report(gated, "sarif"))
    assert {
        "DFC_DOCUMENT_PROPERTY_INVENTORY_CHANGED",
        "DFP019",
        "DFP020",
    } <= {result["ruleId"] for result in sarif["runs"][0]["results"]}
    for marker in (
        "CORE_CREATOR_DO_NOT_LEAK",
        "CORE_TITLE_DO_NOT_LEAK",
        "CORE_TITLE_CHANGED_DO_NOT_LEAK",
        "EXTENDED_COMPANY_DO_NOT_LEAK",
        "EXTENDED_MANAGER_DO_NOT_LEAK",
        "CUSTOM_PROPERTY_NAME_DO_NOT_LEAK",
        "CUSTOM_PROPERTY_VALUE_DO_NOT_LEAK",
        "SECOND_CUSTOM_PROPERTY_NAME_DO_NOT_LEAK",
        "SECOND_CUSTOM_PROPERTY_VALUE_DO_NOT_LEAK",
    ):
        assert marker not in rendered


def test_mail_merge_inventory_is_private_and_relationship_id_stable(tmp_path) -> None:
    before = tmp_path / "before.docx"
    after = tmp_path / "after.docx"
    query_changed = tmp_path / "query-changed.docx"
    recipient_changed = tmp_path / "recipient-changed.docx"
    renumbered = tmp_path / "renumbered.docx"
    strict_mail_merge = tmp_path / "strict-mail-merge.docx"
    recipient_data_alias = tmp_path / "recipient-data-alias.docx"
    orphaned = tmp_path / "orphaned.docx"
    wrong_source_type = tmp_path / "wrong-source-type.docx"
    internal_source = tmp_path / "internal-source.docx"
    external_recipient = tmp_path / "external-recipient.docx"
    missing_recipient = tmp_path / "missing-recipient.docx"
    policy_path = tmp_path / "docfence.yml"
    _write_mail_merge_document(before, include_mail_merge=False)
    _write_mail_merge_document(after)
    _write_mail_merge_document(
        query_changed, query="MAIL_MERGE_QUERY_CHANGED_DO_NOT_LEAK"
    )
    _write_mail_merge_document(
        recipient_changed,
        recipient_value="MAIL_MERGE_RECIPIENT_CHANGED_DO_NOT_LEAK",
    )
    _write_mail_merge_document(renumbered, relationship_id_suffix="9")
    _write_mail_merge_document(strict_mail_merge, strict_syntax=True)
    _write_mail_merge_document(
        recipient_data_alias,
        recipient_relationship_type=_RECIPIENT_DATA_RELATIONSHIP_TYPE,
    )
    _write_mail_merge_document(orphaned, include_configuration=False)
    _write_mail_merge_document(
        wrong_source_type,
        source_relationship_type=_OLE_OBJECT_RELATIONSHIP_TYPE,
    )
    _write_mail_merge_document(internal_source, source_target_mode="Internal")
    _write_mail_merge_document(external_recipient, recipient_target_mode="External")
    _write_mail_merge_document(missing_recipient, include_recipient_part=False)

    snapshot = load_snapshot(after)
    expected_inventory = {
        "mail_merge_configuration_count": 1,
        "mail_merge_data_source_relationship_count": 2,
        "mail_merge_header_source_relationship_count": 1,
        "mail_merge_recipient_data_relationship_count": 1,
        "mail_merge_recipient_data_part_count": 1,
    }
    assert snapshot.public_dict()["mail_merge"] == expected_inventory
    assert snapshot.unclassified_part_count == 1
    assert (
        load_snapshot(strict_mail_merge).public_dict()["mail_merge"]
        == expected_inventory
    )
    assert (
        load_snapshot(recipient_data_alias).public_dict()["mail_merge"]
        == expected_inventory
    )
    assert load_snapshot(orphaned).public_dict()["mail_merge"] == {
        **expected_inventory,
        "mail_merge_configuration_count": 0,
    }

    report = diff_documents(before, after)
    assert {
        "external_relationships_changed",
        "mail_merge_inventory_changed",
    } <= {change.kind for change in report.changes}
    assert {change.kind for change in diff_documents(after, query_changed).changes} == {
        "document_settings_changed",
        "mail_merge_inventory_changed",
    }
    assert {
        change.kind for change in diff_documents(after, recipient_changed).changes
    } == {"mail_merge_inventory_changed"}
    assert diff_documents(after, renumbered).changes == ()
    with pytest.raises(DocumentFormatError):
        load_snapshot(wrong_source_type)
    with pytest.raises(DocumentFormatError):
        load_snapshot(internal_source)
    with pytest.raises(DocumentFormatError):
        load_snapshot(external_recipient)
    with pytest.raises(DocumentFormatError):
        load_snapshot(missing_recipient)

    policy_path.write_text(
        """version: 1
rules:
  require_no_mail_merge: true
  no_mail_merge_changes: true
""",
        encoding="utf-8",
    )
    policy = load_policy(policy_path)
    assert {finding.rule_id for finding in apply_policy(report, policy).findings} == {
        "DFP021",
        "DFP022",
    }
    assert {
        finding.rule_id
        for finding in apply_policy(diff_documents(after, renumbered), policy).findings
    } == {"DFP021"}
    assert {
        finding.rule_id
        for finding in apply_policy(diff_documents(before, orphaned), policy).findings
    } == {"DFP021", "DFP022"}

    gated = apply_policy(report, policy)
    rendered = "\n".join(
        (
            render_profile(snapshot, "json"),
            render_profile(snapshot, "markdown"),
            render_report(gated, "json"),
            render_report(gated, "markdown"),
            render_report(gated, "sarif"),
        )
    )
    sarif = json.loads(render_report(gated, "sarif"))
    assert {
        "DFC_MAIL_MERGE_INVENTORY_CHANGED",
        "DFP021",
        "DFP022",
    } <= {result["ruleId"] for result in sarif["runs"][0]["results"]}
    for marker in (
        "MAIL_MERGE_DATA_SOURCE_DO_NOT_LEAK",
        "MAIL_MERGE_HEADER_SOURCE_DO_NOT_LEAK",
        "MAIL_MERGE_QUERY_DO_NOT_LEAK",
        "MAIL_MERGE_QUERY_CHANGED_DO_NOT_LEAK",
        "MAIL_MERGE_UDL_DO_NOT_LEAK",
        "MAIL_MERGE_TABLE_DO_NOT_LEAK",
        "MAIL_MERGE_FIELD_NAME_DO_NOT_LEAK",
        "MAIL_MERGE_FIELD_MAPPING_DO_NOT_LEAK",
        "MAIL_MERGE_RECIPIENT_DO_NOT_LEAK",
        "MAIL_MERGE_RECIPIENT_CHANGED_DO_NOT_LEAK",
    ):
        assert marker not in rendered


def test_data_binding_inventory_is_private_and_relationship_id_stable(tmp_path) -> None:
    before = tmp_path / "before.docx"
    after = tmp_path / "after.docx"
    payload_changed = tmp_path / "payload-changed.docx"
    xpath_changed = tmp_path / "xpath-changed.docx"
    renumbered = tmp_path / "renumbered.docx"
    strict_data_binding = tmp_path / "strict-data-binding.docx"
    without_store_item_id = tmp_path / "without-store-item-id.docx"
    unmatched_store_item_id = tmp_path / "unmatched-store-item-id.docx"
    multiple_bindings = tmp_path / "multiple-bindings.docx"
    out_of_scope_markup = tmp_path / "out-of-scope-markup.docx"
    policy_path = tmp_path / "docfence.yml"
    _write_data_binding_document(before, include_data_binding=False)
    _write_data_binding_document(after)
    _write_data_binding_document(
        payload_changed,
        custom_xml_value="DATA_BINDING_PAYLOAD_CHANGED_DO_NOT_LEAK",
    )
    _write_data_binding_document(
        xpath_changed,
        xpath="/private/records/changedValue",
    )
    _write_data_binding_document(renumbered, relationship_id_suffix="9")
    _write_data_binding_document(strict_data_binding, strict_syntax=True)
    _write_data_binding_document(
        without_store_item_id,
        include_store_item_id=False,
    )
    _write_data_binding_document(
        unmatched_store_item_id,
        item_store_item_id="{99999999-9999-9999-9999-999999999999}",
    )
    _write_data_binding_document(multiple_bindings, binding_count=2)
    _write_data_binding_document(
        out_of_scope_markup,
        data_binding_outside_sdt_properties=True,
    )

    snapshot = load_snapshot(after)
    expected_inventory = {
        "data_binding_count": 1,
        "data_binding_with_store_item_id_count": 1,
        "data_binding_without_store_item_id_count": 0,
        "data_binding_referenced_custom_xml_part_count": 1,
        "data_binding_unmatched_store_item_id_count": 0,
    }
    assert snapshot.public_dict()["data_bindings"] == expected_inventory
    assert load_snapshot(strict_data_binding).public_dict()["data_bindings"] == (
        expected_inventory
    )
    assert load_snapshot(without_store_item_id).public_dict()["data_bindings"] == {
        **expected_inventory,
        "data_binding_with_store_item_id_count": 0,
        "data_binding_without_store_item_id_count": 1,
        "data_binding_referenced_custom_xml_part_count": 0,
    }
    assert load_snapshot(unmatched_store_item_id).public_dict()["data_bindings"] == {
        **expected_inventory,
        "data_binding_referenced_custom_xml_part_count": 0,
        "data_binding_unmatched_store_item_id_count": 1,
    }
    assert load_snapshot(multiple_bindings).public_dict()["data_bindings"] == {
        **expected_inventory,
        "data_binding_count": 2,
        "data_binding_with_store_item_id_count": 2,
    }
    assert load_snapshot(out_of_scope_markup).public_dict()["data_bindings"] == {
        **expected_inventory,
        "data_binding_count": 0,
        "data_binding_with_store_item_id_count": 0,
        "data_binding_referenced_custom_xml_part_count": 0,
    }

    report = diff_documents(before, after)
    assert {
        "content_control_inventory_changed",
        "custom_xml_changed",
        "data_binding_inventory_changed",
    } <= {change.kind for change in report.changes}
    assert {
        change.kind for change in diff_documents(after, payload_changed).changes
    } == {"custom_xml_changed", "data_binding_inventory_changed"}
    assert "data_binding_inventory_changed" in {
        change.kind for change in diff_documents(after, xpath_changed).changes
    }
    assert diff_documents(after, renumbered).changes == ()

    policy_path.write_text(
        """version: 1
rules:
  require_no_data_bindings: true
  no_data_binding_changes: true
""",
        encoding="utf-8",
    )
    policy = load_policy(policy_path)
    assert {finding.rule_id for finding in apply_policy(report, policy).findings} == {
        "DFP023",
        "DFP024",
    }
    assert {
        finding.rule_id
        for finding in apply_policy(diff_documents(after, renumbered), policy).findings
    } == {"DFP023"}
    assert {
        finding.rule_id
        for finding in apply_policy(
            diff_documents(after, payload_changed), policy
        ).findings
    } == {"DFP023", "DFP024"}

    gated = apply_policy(report, policy)
    rendered = "\n".join(
        (
            render_profile(snapshot, "json"),
            render_profile(snapshot, "markdown"),
            render_report(gated, "json"),
            render_report(gated, "markdown"),
            render_report(gated, "sarif"),
        )
    )
    sarif = json.loads(render_report(gated, "sarif"))
    assert {
        "DFC_DATA_BINDING_INVENTORY_CHANGED",
        "DFP023",
        "DFP024",
    } <= {result["ruleId"] for result in sarif["runs"][0]["results"]}
    for marker in (
        "DATA_BINDING_VISIBLE_DO_NOT_LEAK",
        "DATA_BINDING_XPATH_DO_NOT_LEAK",
        "DATA_BINDING_PREFIXES_DO_NOT_LEAK",
        "DATA_BINDING_PAYLOAD_DO_NOT_LEAK",
        "DATA_BINDING_PAYLOAD_CHANGED_DO_NOT_LEAK",
    ):
        assert marker not in rendered


def test_data_binding_rejects_invalid_referenced_custom_xml_properties(
    tmp_path,
) -> None:
    invalid_properties_root = tmp_path / "invalid-properties-root.docx"
    external_properties_target = tmp_path / "external-properties-target.docx"
    _write_data_binding_document(
        invalid_properties_root,
        invalid_custom_xml_properties_root=True,
    )
    _write_data_binding_document(
        external_properties_target,
        custom_xml_properties_target_mode="External",
    )

    with pytest.raises(DocumentFormatError):
        load_snapshot(invalid_properties_root)
    with pytest.raises(DocumentFormatError):
        load_snapshot(external_properties_target)


def test_external_document_dependency_inventory_is_private_and_stable(tmp_path) -> None:
    before = tmp_path / "before.docx"
    after = tmp_path / "after.docx"
    target_changed = tmp_path / "target-changed.docx"
    renumbered = tmp_path / "renumbered.docx"
    strict_dependencies = tmp_path / "strict-dependencies.docx"
    orphaned = tmp_path / "orphaned.docx"
    wrong_attached_template_type = tmp_path / "wrong-attached-template-type.docx"
    internal_dependencies = tmp_path / "internal-dependencies.docx"
    wrong_frame_type = tmp_path / "wrong-frame-type.docx"
    external_web_settings = tmp_path / "external-web-settings.docx"
    policy_path = tmp_path / "docfence.yml"
    _write_external_document_dependency_document(before, include_dependencies=False)
    _write_external_document_dependency_document(after)
    _write_external_document_dependency_document(
        target_changed,
        subdocument_target="https://example.invalid/SUBDOCUMENT_CHANGED_DO_NOT_LEAK",
    )
    _write_external_document_dependency_document(renumbered, relationship_id_suffix="9")
    _write_external_document_dependency_document(
        strict_dependencies, strict_syntax=True
    )
    _write_external_document_dependency_document(orphaned, include_anchors=False)
    _write_external_document_dependency_document(
        wrong_attached_template_type,
        attached_template_relationship_type=_OLE_OBJECT_RELATIONSHIP_TYPE,
    )
    _write_external_document_dependency_document(
        internal_dependencies, dependency_target_mode="Internal"
    )
    _write_external_document_dependency_document(
        wrong_frame_type,
        frame_relationship_type=_OLE_OBJECT_RELATIONSHIP_TYPE,
    )
    _write_external_document_dependency_document(
        external_web_settings,
        web_settings_target_mode="External",
    )

    expected_inventory = {
        "attached_template_anchor_count": 1,
        "attached_template_relationship_count": 1,
        "subdocument_anchor_count": 1,
        "subdocument_relationship_count": 1,
        "frame_source_anchor_count": 1,
        "frame_relationship_count": 1,
    }
    snapshot = load_snapshot(after)
    assert snapshot.public_dict()["external_document_dependencies"] == (
        expected_inventory
    )
    assert (
        load_snapshot(strict_dependencies).public_dict()[
            "external_document_dependencies"
        ]
        == expected_inventory
    )
    assert load_snapshot(orphaned).public_dict()["external_document_dependencies"] == {
        "attached_template_anchor_count": 0,
        "attached_template_relationship_count": 1,
        "subdocument_anchor_count": 0,
        "subdocument_relationship_count": 1,
        "frame_source_anchor_count": 0,
        "frame_relationship_count": 1,
    }

    report = diff_documents(before, after)
    assert {
        "external_relationships_changed",
        "external_document_dependency_inventory_changed",
    } <= {change.kind for change in report.changes}
    assert {
        "external_relationships_changed",
        "external_document_dependency_inventory_changed",
    } <= {change.kind for change in diff_documents(after, target_changed).changes}
    assert diff_documents(after, renumbered).changes == ()
    for invalid_document in (
        wrong_attached_template_type,
        internal_dependencies,
        wrong_frame_type,
        external_web_settings,
    ):
        with pytest.raises(DocumentFormatError):
            load_snapshot(invalid_document)

    policy_path.write_text(
        """version: 1
rules:
  require_no_external_document_dependencies: true
  no_external_document_dependency_changes: true
""",
        encoding="utf-8",
    )
    policy = load_policy(policy_path)
    assert {finding.rule_id for finding in apply_policy(report, policy).findings} == {
        "DFP025",
        "DFP026",
    }
    assert {
        finding.rule_id
        for finding in apply_policy(diff_documents(after, renumbered), policy).findings
    } == {"DFP025"}
    assert {
        finding.rule_id
        for finding in apply_policy(diff_documents(before, orphaned), policy).findings
    } == {"DFP025", "DFP026"}

    gated = apply_policy(report, policy)
    rendered = "\n".join(
        (
            render_profile(snapshot, "json"),
            render_profile(snapshot, "markdown"),
            render_report(gated, "json"),
            render_report(gated, "markdown"),
            render_report(gated, "sarif"),
        )
    )
    sarif = json.loads(render_report(gated, "sarif"))
    assert {
        "DFC_EXTERNAL_DOCUMENT_DEPENDENCY_INVENTORY_CHANGED",
        "DFP025",
        "DFP026",
    } <= {result["ruleId"] for result in sarif["runs"][0]["results"]}
    for marker in (
        "ATTACHED_TEMPLATE_DO_NOT_LEAK",
        "SUBDOCUMENT_DO_NOT_LEAK",
        "SUBDOCUMENT_CHANGED_DO_NOT_LEAK",
        "FRAME_SOURCE_DO_NOT_LEAK",
        "FRAME_NAME_DO_NOT_LEAK",
    ):
        assert marker not in rendered


def test_external_field_inventory_is_private_and_semantic(tmp_path) -> None:
    before = tmp_path / "before.docx"
    after = tmp_path / "after.docx"
    target_changed = tmp_path / "target-changed.docx"
    split_complex = tmp_path / "split-complex.docx"
    strict = tmp_path / "strict.docx"
    header = tmp_path / "header.docx"
    policy_path = tmp_path / "docfence.yml"
    _write_external_field_document(before, include_external_fields=False)
    _write_external_field_document(after)
    _write_external_field_document(
        target_changed,
        target_marker="EXTERNAL_FIELD_TARGET_CHANGED_DO_NOT_LEAK",
    )
    _write_external_field_document(split_complex, split_complex_instructions=True)
    _write_external_field_document(strict, strict_syntax=True)
    _write_external_field_document(
        header,
        include_external_fields=False,
        include_header_field=True,
    )

    expected_inventory = {
        "database_field_count": 1,
        "legacy_data_field_count": 1,
        "dde_field_count": 1,
        "dde_auto_field_count": 1,
        "include_text_field_count": 2,
        "include_picture_field_count": 2,
        "link_field_count": 1,
        "referenced_document_field_count": 1,
    }
    snapshot = load_snapshot(after)
    assert snapshot.public_dict()["external_fields"] == expected_inventory
    assert load_snapshot(before).public_dict()["external_fields"] == {
        key: 0 for key in expected_inventory
    }
    assert load_snapshot(strict).public_dict()["external_fields"] == expected_inventory
    assert (
        load_snapshot(split_complex).external_fields.signature
        == snapshot.external_fields.signature
    )
    assert load_snapshot(header).public_dict()["external_fields"] == {
        **{key: 0 for key in expected_inventory},
        "link_field_count": 1,
    }

    report = diff_documents(before, after)
    assert "external_field_inventory_changed" in {
        change.kind for change in report.changes
    }
    assert "external_field_inventory_changed" in {
        change.kind for change in diff_documents(after, target_changed).changes
    }

    policy_path.write_text(
        """version: 1
rules:
  require_no_external_fields: true
  no_external_field_changes: true
""",
        encoding="utf-8",
    )
    policy = load_policy(policy_path)
    assert {finding.rule_id for finding in apply_policy(report, policy).findings} == {
        "DFP027",
        "DFP028",
    }
    assert {
        finding.rule_id
        for finding in apply_policy(
            diff_documents(after, split_complex), policy
        ).findings
    } == {"DFP027"}

    gated = apply_policy(report, policy)
    rendered = "\n".join(
        (
            render_profile(snapshot, "json"),
            render_profile(snapshot, "markdown"),
            render_report(gated, "json"),
            render_report(gated, "markdown"),
            render_report(gated, "sarif"),
        )
    )
    sarif = json.loads(render_report(gated, "sarif"))
    assert {
        "DFC_EXTERNAL_FIELD_INVENTORY_CHANGED",
        "DFP027",
        "DFP028",
    } <= {result["ruleId"] for result in sarif["runs"][0]["results"]}
    for marker in (
        "EXTERNAL_FIELD_TARGET_DO_NOT_LEAK",
        "EXTERNAL_FIELD_TARGET_CHANGED_DO_NOT_LEAK",
        "LOOSE_EXTERNAL_FIELD_DO_NOT_LEAK",
        "POST_SEPARATOR_EXTERNAL_FIELD_DO_NOT_LEAK",
        "INCOMPLETE_EXTERNAL_FIELD_DO_NOT_LEAK",
        "HEADER_EXTERNAL_FIELD_DO_NOT_LEAK",
    ):
        assert marker not in rendered


def test_external_field_inventory_handles_nested_and_resultless_complex_fields(
    tmp_path,
) -> None:
    document = tmp_path / "nested.docx"
    _write_raw_package(
        document,
        f'''<w:document xmlns:w="{W}"><w:body><w:p>
  <w:r><w:fldChar w:fldCharType="begin"/></w:r>
  <w:r><w:instrText xml:space="preserve"> IF 1 = 1 </w:instrText></w:r>
  <w:r><w:fldChar w:fldCharType="begin"/></w:r>
  <w:r><w:instrText xml:space="preserve"> INCLUDETEXT "NESTED_</w:instrText></w:r>
  <w:r><w:instrText xml:space="preserve">EXTERNAL_FIELD_</w:instrText></w:r>
  <w:r><w:instrText xml:space="preserve">DO_NOT_LEAK.docx" </w:instrText></w:r>
  <w:r><w:fldChar w:fldCharType="end"/></w:r>
  <w:r><w:instrText xml:space="preserve"> "yes" "no" </w:instrText></w:r>
  <w:r><w:fldChar w:fldCharType="separate"/></w:r>
  <w:r><w:t>FIELD_RESULT_DO_NOT_LEAK</w:t></w:r>
  <w:r><w:fldChar w:fldCharType="end"/></w:r>
  <w:r><w:fldChar w:fldCharType="begin"/></w:r>
  <w:r><w:instrText xml:space="preserve"> RD "RESULTLESS_</w:instrText></w:r>
  <w:r><w:instrText xml:space="preserve">EXTERNAL_FIELD_</w:instrText></w:r>
  <w:r><w:instrText xml:space="preserve">DO_NOT_LEAK.docx" </w:instrText></w:r>
  <w:r><w:fldChar w:fldCharType="end"/></w:r>
</w:p><w:sectPr/></w:body></w:document>''',
    )

    assert load_snapshot(document).public_dict()["external_fields"] == {
        "database_field_count": 0,
        "legacy_data_field_count": 0,
        "dde_field_count": 0,
        "dde_auto_field_count": 0,
        "include_text_field_count": 1,
        "include_picture_field_count": 0,
        "link_field_count": 0,
        "referenced_document_field_count": 1,
    }


def test_external_field_inventory_preserves_revision_variants_privately(
    tmp_path,
) -> None:
    document = tmp_path / "revision-field.docx"
    _write_raw_package(
        document,
        f'''<w:document xmlns:w="{W}"><w:body><w:p>
  <w:r><w:fldChar w:fldCharType="begin"/></w:r>
  <w:r><w:instrText xml:space="preserve"> INCLUDETEXT "</w:instrText></w:r>
  <w:del w:id="1"><w:r>
    <w:delInstrText>DELETED_EXTERNAL_FIELD_DO_NOT_LEAK</w:delInstrText>
  </w:r></w:del>
  <w:ins w:id="2"><w:r>
    <w:instrText>INSERTED_EXTERNAL_FIELD_DO_NOT_LEAK</w:instrText>
  </w:r></w:ins>
  <w:r><w:instrText>.docx" </w:instrText></w:r>
  <w:r><w:fldChar w:fldCharType="end"/></w:r>
  <w:r><w:fldChar w:fldCharType="begin"/></w:r>
  <w:del w:id="3"><w:r>
    <w:delInstrText>RD "RETIRED_</w:delInstrText>
    <w:delInstrText>EXTERNAL_FIELD_DO_NOT_LEAK.docx"</w:delInstrText>
  </w:r></w:del>
  <w:ins w:id="4"><w:r><w:instrText>DATE</w:instrText></w:r></w:ins>
  <w:r><w:fldChar w:fldCharType="end"/></w:r>
  <w:r><w:fldChar w:fldCharType="begin"/></w:r>
  <w:moveFrom w:id="6"><w:r>
    <w:delInstrText>LINK "MOVED_FROM_EXTERNAL_FIELD_DO_NOT_LEAK"</w:delInstrText>
  </w:r></w:moveFrom>
  <w:moveTo w:id="7"><w:r>
    <w:instrText>LINK "MOVED_TO_EXTERNAL_FIELD_DO_NOT_LEAK"</w:instrText>
  </w:r></w:moveTo>
  <w:r><w:fldChar w:fldCharType="end"/></w:r>
  <w:del w:id="5"><w:r>
    <w:delInstrText>INCLUDETEXT "LOOSE_DELETED_</w:delInstrText>
    <w:delInstrText>EXTERNAL_FIELD_DO_NOT_LEAK.docx"</w:delInstrText>
  </w:r></w:del>
</w:p><w:sectPr/></w:body></w:document>''',
    )

    snapshot = load_snapshot(document)
    assert snapshot.public_dict()["external_fields"] == {
        "database_field_count": 0,
        "legacy_data_field_count": 0,
        "dde_field_count": 0,
        "dde_auto_field_count": 0,
        "include_text_field_count": 2,
        "include_picture_field_count": 0,
        "link_field_count": 2,
        "referenced_document_field_count": 1,
    }
    rendered = render_profile(snapshot, "json") + render_profile(snapshot, "markdown")
    for marker in (
        "DELETED_EXTERNAL_FIELD_DO_NOT_LEAK",
        "INSERTED_EXTERNAL_FIELD_DO_NOT_LEAK",
        "RETIRED_EXTERNAL_FIELD_DO_NOT_LEAK",
        "MOVED_FROM_EXTERNAL_FIELD_DO_NOT_LEAK",
        "MOVED_TO_EXTERNAL_FIELD_DO_NOT_LEAK",
        "LOOSE_DELETED_EXTERNAL_FIELD_DO_NOT_LEAK",
    ):
        assert marker not in rendered


def test_external_document_dependency_discovers_glossary_settings(tmp_path) -> None:
    before = tmp_path / "glossary-before.docx"
    renumbered = tmp_path / "glossary-renumbered.docx"
    strict = tmp_path / "glossary-strict.docx"
    _write_glossary_attached_template_document(before)
    _write_glossary_attached_template_document(renumbered, relationship_id_suffix="9")
    _write_glossary_attached_template_document(strict, strict_syntax=True)

    expected_inventory = {
        "attached_template_anchor_count": 1,
        "attached_template_relationship_count": 1,
        "subdocument_anchor_count": 0,
        "subdocument_relationship_count": 0,
        "frame_source_anchor_count": 0,
        "frame_relationship_count": 0,
    }
    assert load_snapshot(before).public_dict()["external_document_dependencies"] == (
        expected_inventory
    )
    assert load_snapshot(strict).public_dict()["external_document_dependencies"] == (
        expected_inventory
    )
    assert diff_documents(before, renumbered).changes == ()


def test_diff_reports_supported_changes_without_document_material(tmp_path) -> None:
    before = tmp_path / "approved.docx"
    after = tmp_path / "candidate.docm"
    write_document(before, text="APPROVED_DO_NOT_LEAK")
    write_document(
        after,
        text="CANDIDATE_DO_NOT_LEAK",
        relationship_id="rId1",
        relationship_target="https://example.invalid/EXTERNAL_DO_NOT_LEAK",
        relationship_external=True,
        revision=True,
        hidden=True,
        field=True,
        control=True,
        comment=True,
        track_revisions=True,
        custom_xml=b"CUSTOM_XML_DO_NOT_LEAK",
        macro=b"MACRO_DO_NOT_LEAK",
        unclassified=b"UNCLASSIFIED_DO_NOT_LEAK",
    )

    report = diff_documents(before, after)
    kinds = {change.kind for change in report.changes}

    assert {
        "document_format_changed",
        "content_changed",
        "revision_inventory_changed",
        "hidden_text_inventory_changed",
        "field_code_inventory_changed",
        "content_control_inventory_changed",
        "comment_anchor_inventory_changed",
        "comment_inventory_changed",
        "track_revisions_setting_changed",
        "external_relationships_changed",
        "custom_xml_changed",
        "macro_payload_changed",
        "unclassified_package_payload_changed",
    } <= kinds

    rendered = "\n".join(
        (
            render_report(report, "json"),
            render_report(report, "markdown"),
            render_report(report, "sarif"),
        )
    )
    sarif = json.loads(render_report(report, "sarif"))
    assert any(
        result["ruleId"] == "DFC_CONTENT_CHANGED"
        for result in sarif["runs"][0]["results"]
    )
    for marker in (
        "APPROVED_DO_NOT_LEAK",
        "CANDIDATE_DO_NOT_LEAK",
        "HIDDEN_DO_NOT_LEAK",
        "FIELD_DO_NOT_LEAK",
        "COMMENT_DO_NOT_LEAK",
        "REVIEWER_DO_NOT_LEAK",
        "EXTERNAL_DO_NOT_LEAK",
        "CUSTOM_XML_DO_NOT_LEAK",
        "MACRO_DO_NOT_LEAK",
        "UNCLASSIFIED_DO_NOT_LEAK",
    ):
        assert marker not in rendered


def test_rsid_and_relationship_id_rewrites_are_quiet(tmp_path) -> None:
    before = tmp_path / "before.docx"
    after = tmp_path / "after.docx"
    write_document(
        before,
        text="UNCHANGED_DO_NOT_LEAK",
        rsid="11111111",
        relationship_id="rId1",
        relationship_target="media/image1.png",
        include_settings_rsids=True,
    )
    write_document(
        after,
        text="UNCHANGED_DO_NOT_LEAK",
        rsid="22222222",
        relationship_id="rId9",
        relationship_target="media/image1.png",
        include_settings_rsids=True,
    )

    report = diff_documents(before, after)

    assert report.changes == ()


def test_all_supported_story_kinds_are_inventoried_without_story_text(tmp_path) -> None:
    document = tmp_path / "stories.docx"
    extra_stories = {
        "header": "HEADER_DO_NOT_LEAK",
        "footer": "FOOTER_DO_NOT_LEAK",
        "footnote": "FOOTNOTE_DO_NOT_LEAK",
        "endnote": "ENDNOTE_DO_NOT_LEAK",
        "glossary": "GLOSSARY_DO_NOT_LEAK",
    }
    write_document(document, comment=True, extra_stories=extra_stories)

    snapshot = load_snapshot(document)
    assert snapshot.story_kind_counts == {
        "body": 1,
        "comment": 1,
        "endnote": 1,
        "footer": 1,
        "footnote": 1,
        "glossary": 1,
        "header": 1,
    }
    assert snapshot.public_dict()["paragraph_count"] == 7
    rendered = render_profile(snapshot, "markdown")
    assert "DO_NOT_LEAK" not in rendered


def test_repeated_blocks_use_bounded_summary_mode(tmp_path) -> None:
    document = tmp_path / "base.docx"
    write_document(document)
    base = load_snapshot(document)
    repeated = ("same-private-fingerprint",) * 1_200
    before_story = replace(
        base.stories[0], block_signatures=(*repeated, "before-private-fingerprint")
    )
    after_story = replace(
        base.stories[0], block_signatures=(*repeated, "after-private-fingerprint")
    )

    report = diff_documents(
        replace(base, stories=(before_story,)),
        replace(base, stories=(after_story,)),
    )

    change = next(
        change for change in report.changes if change.kind == "content_changed"
    )
    assert change.details["comparison_mode"] == "bounded_summary"


def test_policy_gates_candidate_state_and_never_leaks_markers(tmp_path) -> None:
    before = tmp_path / "before.docx"
    after = tmp_path / "after.docm"
    policy_path = tmp_path / "docfence.yml"
    write_document(before)
    write_document(
        after,
        relationship_id="rId1",
        relationship_target="https://example.invalid/EXTERNAL_DO_NOT_LEAK",
        relationship_external=True,
        revision=True,
        hidden=True,
        comment=True,
        custom_xml=b"CUSTOM_XML_DO_NOT_LEAK",
        macro=b"MACRO_DO_NOT_LEAK",
    )
    policy_path.write_text(starter_policy(), encoding="utf-8")

    report = apply_policy(diff_documents(before, after), load_policy(policy_path))
    rule_ids = {finding.rule_id for finding in report.findings}

    assert rule_ids == {"DFP001", "DFP002", "DFP003", "DFP004", "DFP005", "DFP006"}
    rendered = render_report(report, "sarif")
    assert "DO_NOT_LEAK" not in rendered
    parsed = json.loads(rendered)
    assert parsed["runs"][0]["results"]
    assert all("locations" not in result for result in parsed["runs"][0]["results"])


def test_rejects_unsafe_xml_and_unsafe_package_member_names(tmp_path) -> None:
    dtd_document = tmp_path / "dtd.docx"
    unsafe_member_document = tmp_path / "unsafe-member.docx"
    prefix_collision_document = tmp_path / "prefix-collision.docx"
    _write_raw_package(
        dtd_document,
        '<!DOCTYPE document [<!ENTITY xxe "LEAK">]>'
        f'<w:document xmlns:w="{W}"><w:body><w:p><w:r><w:t>&xxe;</w:t>'
        "</w:r></w:p></w:body></w:document>",
    )
    _write_raw_package(
        unsafe_member_document,
        f'<w:document xmlns:w="{W}"><w:body/></w:document>',
        extra_name="../not-safe.xml",
    )
    _write_raw_package(
        prefix_collision_document,
        f'<w:document xmlns:w="{W}"><w:body/></w:document>',
        extra_name="word/document.xml/child.xml",
    )

    with pytest.raises(DocumentSafetyError):
        load_snapshot(dtd_document)
    with pytest.raises(DocumentFormatError):
        load_snapshot(unsafe_member_document)
    with pytest.raises(DocumentFormatError):
        load_snapshot(prefix_collision_document)


def test_rejects_a_styles_part_with_the_wrong_word_root(tmp_path) -> None:
    document = tmp_path / "invalid-styles.docx"
    write_document(document, styles_xml=f'<w:settings xmlns:w="{W}"/>')

    with pytest.raises(DocumentFormatError):
        load_snapshot(document)


def test_policy_parser_is_strict_and_cli_returns_gate_status(tmp_path, capsys) -> None:
    policy_path = tmp_path / "docfence.yml"
    invalid_policy_path = tmp_path / "invalid.yml"
    before = tmp_path / "before.docx"
    after = tmp_path / "after.docm"
    result_path = tmp_path / "result.sarif"
    write_document(before)
    write_document(after, hidden=True)

    assert main(["init", str(policy_path)]) == 0
    assert main(["init", str(policy_path)]) == 2
    assert main(["profile", str(before), "--format", "json"]) == 0
    assert '"document"' in capsys.readouterr().out
    assert (
        main(
            [
                "check",
                str(before),
                str(after),
                "--policy",
                str(policy_path),
                "--format",
                "sarif",
                "--output",
                str(result_path),
            ]
        )
        == 1
    )
    assert "HIDDEN_DO_NOT_LEAK" not in result_path.read_text(encoding="utf-8")

    invalid_policy_path.write_text("version: 1\nrules: &unsafe\n", encoding="utf-8")
    with pytest.raises(PolicyError):
        load_policy(invalid_policy_path)


def test_malformed_paths_fail_with_safe_domain_errors(capsys) -> None:
    with pytest.raises(DocumentFormatError):
        load_snapshot("\x00.docx")
    with pytest.raises(PolicyError):
        load_policy("\x00.yml")
    assert main(["profile", "\x00.docx"]) == 2
    assert "docfence: document" in capsys.readouterr().err


def _write_raw_package(path, document_xml: str, extra_name: str | None = None) -> None:
    content_types = (
        f'<Types xmlns="{CT}"><Override PartName="/word/document.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
        "</Types>"
    )
    with zipfile.ZipFile(path, "w") as archive:
        archive.writestr("[Content_Types].xml", content_types)
        archive.writestr("word/document.xml", document_xml)
        if extra_name is not None:
            archive.writestr(extra_name, b"x")


def _write_external_field_document(
    path,
    *,
    target_marker: str = "EXTERNAL_FIELD_TARGET_DO_NOT_LEAK",
    include_external_fields: bool = True,
    split_complex_instructions: bool = False,
    strict_syntax: bool = False,
    include_header_field: bool = False,
) -> None:
    word_namespace = _STRICT_WORD_NAMESPACE if strict_syntax else W

    def simple_field(instruction: str) -> str:
        escaped_instruction = (
            instruction.replace("&", "&amp;")
            .replace('"', "&quot;")
            .replace("<", "&lt;")
        )
        return (
            f'<w:fldSimple w:instr="{escaped_instruction}"><w:r><w:t>'
            "FIELD_RESULT_DO_NOT_LEAK</w:t></w:r></w:fldSimple>"
        )

    def complex_field(instruction: str) -> str:
        chunks = [instruction]
        if split_complex_instructions:
            split_at = len(instruction) // 2
            chunks = [instruction[:split_at], instruction[split_at:]]
        instruction_markup = "".join(
            f'<w:r><w:instrText xml:space="preserve">{chunk}</w:instrText></w:r>'
            for chunk in chunks
        )
        return (
            '<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
            f"{instruction_markup}"
            '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
            "<w:r><w:t>FIELD_RESULT_DO_NOT_LEAK</w:t></w:r>"
            '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
        )

    external_fields = ""
    if include_external_fields:
        external_fields = "".join(
            (
                simple_field(f'DATABASE "{target_marker}"'),
                simple_field(f'DATA "{target_marker}"'),
                complex_field(f'DDE Excel "{target_marker}" "Sheet1!R1C1"'),
                simple_field(f'DDEAUTO Excel "{target_marker}" "Sheet1!R1C1"'),
                complex_field(f'INCLUDE "{target_marker}.docx"'),
                simple_field(f'INCLUDETEXT "{target_marker}.docx"'),
                complex_field(f'IMPORT "{target_marker}.png"'),
                simple_field(f'INCLUDEPICTURE "{target_marker}.png"'),
                complex_field(f'LINK Excel.Chart.12 "{target_marker}"'),
                simple_field(f'RD "{target_marker}.docx"'),
            )
        )

    ignored_instruction_markup = (
        "<w:r><w:instrText>"
        'INCLUDETEXT "LOOSE_EXTERNAL_FIELD_DO_NOT_LEAK.docx"'
        "</w:instrText></w:r>"
        '<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
        "<w:r><w:instrText>DATE</w:instrText></w:r>"
        '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
        "<w:r><w:instrText>"
        'INCLUDETEXT "POST_SEPARATOR_EXTERNAL_FIELD_DO_NOT_LEAK.docx"'
        "</w:instrText></w:r>"
        '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
        '<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
        "<w:r><w:instrText>"
        'INCLUDETEXT "INCOMPLETE_EXTERNAL_FIELD_DO_NOT_LEAK.docx"'
        "</w:instrText></w:r>"
    )
    header_override = (
        '<Override PartName="/word/header1.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.'
        'wordprocessingml.header+xml"/>'
        if include_header_field
        else ""
    )
    header_field = simple_field(
        'LINK Excel.Chart.12 "HEADER_EXTERNAL_FIELD_DO_NOT_LEAK"'
    )
    entries: dict[str, bytes] = {
        "[Content_Types].xml": (
            f'<Types xmlns="{CT}">'
            f'<Override PartName="/word/document.xml" ContentType="{DOCX_MAIN_TYPE}"/>'
            f"{header_override}</Types>"
        ).encode(),
        "word/document.xml": (
            f'<w:document xmlns:w="{word_namespace}"><w:body><w:p>'
            "<w:r><w:t>VISIBLE_DO_NOT_LEAK</w:t></w:r>"
            f"{external_fields}{ignored_instruction_markup}"
            "</w:p><w:sectPr/></w:body></w:document>"
        ).encode(),
    }
    if include_header_field:
        entries["word/header1.xml"] = (
            f'<w:hdr xmlns:w="{word_namespace}"><w:p>{header_field}</w:p></w:hdr>'
        ).encode()

    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for name, payload in entries.items():
            archive.writestr(name, payload)


def _styles_with_hidden_text_declarations() -> str:
    return f"""<w:styles xmlns:w=\"{W}\">
  <w:docDefaults>
    <w:rPrDefault><w:rPr><w:vanish w:val=\"true\"/></w:rPr></w:rPrDefault>
  </w:docDefaults>
  <w:style w:type=\"character\" w:styleId=\"HIDDEN_TEXT_STYLE_DO_NOT_LEAK\">
    <w:rPr><w:vanish/></w:rPr>
  </w:style>
  <w:style w:type=\"paragraph\" w:styleId=\"PARAGRAPH_MARK_STYLE_DO_NOT_LEAK\">
    <w:pPr><w:rPr><w:vanish/></w:rPr></w:pPr>
  </w:style>
  <w:style w:type=\"character\" w:styleId=\"SPECIAL_PARAGRAPH_MARK_STYLE_DO_NOT_LEAK\">
    <w:rPr><w:specVanish/></w:rPr>
  </w:style>
  <w:style w:type=\"character\" w:styleId=\"TRACKED_STYLE_CHANGE_DO_NOT_LEAK\">
    <w:rPr><w:rPrChange><w:rPr><w:vanish/></w:rPr></w:rPrChange></w:rPr>
  </w:style>
  <w:style w:type=\"character\" w:styleId=\"VISIBLE_STYLE_DO_NOT_LEAK\">
    <w:rPr><w:vanish w:val=\"false\"/></w:rPr>
  </w:style>
</w:styles>"""


def _write_embedded_content_document(
    path,
    *,
    object_payload: bytes = b"OBJECT_PAYLOAD_DO_NOT_LEAK",
    import_payload: bytes = b"ALT_IMPORT_DO_NOT_LEAK",
    relationship_id_suffix: str = "1",
    alt_chunk_relationship_type: str = _ALT_CHUNK_RELATIONSHIP_TYPE,
    alt_chunk_target_mode: str = "Internal",
    include_import_payload: bool = True,
    duplicate_alt_chunk_anchor: bool = False,
) -> None:
    object_id = f"rIdObject{relationship_id_suffix}"
    package_id = f"rIdPackage{relationship_id_suffix}"
    control_id = f"rIdControl{relationship_id_suffix}"
    binary_control_id = f"rIdControlBinary{relationship_id_suffix}"
    import_id = f"rIdImport{relationship_id_suffix}"
    import_target_mode = (
        f' TargetMode="{alt_chunk_target_mode}"'
        if alt_chunk_target_mode != "Internal"
        else ""
    )
    import_override = (
        '<Override PartName="/word/afchunk1.html" ContentType="text/html"/>'
        if include_import_payload
        else ""
    )
    alt_chunk_markup = f'<w:altChunk r:id="{import_id}"/>'
    if duplicate_alt_chunk_anchor:
        alt_chunk_markup *= 2
    entries: dict[str, bytes] = {
        "[Content_Types].xml": (
            f'<Types xmlns="{CT}">'
            f'<Override PartName="/word/document.xml" ContentType="{DOCX_MAIN_TYPE}"/>'
            f"{import_override}</Types>"
        ).encode(),
        "word/document.xml": (
            f'<w:document xmlns:w="{W}" xmlns:r="{R}"><w:body>'
            "<w:p><w:r><w:t>VISIBLE_DO_NOT_LEAK</w:t></w:r></w:p>"
            f"{alt_chunk_markup}"
            "<w:sectPr/></w:body></w:document>"
        ).encode(),
        "word/_rels/document.xml.rels": (
            f'<Relationships xmlns="{PR}">'
            f'<Relationship Id="{object_id}" Type="{_OLE_OBJECT_RELATIONSHIP_TYPE}" '
            'Target="embeddings/object1.bin"/>'
            f'<Relationship Id="{package_id}" Type="{_PACKAGE_RELATIONSHIP_TYPE}" '
            'Target="embeddings/package1.bin"/>'
            f'<Relationship Id="{control_id}" Type="{_CONTROL_RELATIONSHIP_TYPE}" '
            'Target="activeX/activeX1.xml"/>'
            f'<Relationship Id="{import_id}" Type="{alt_chunk_relationship_type}" '
            f'Target="afchunk1.html"{import_target_mode}/>'
            "</Relationships>"
        ).encode(),
        "word/embeddings/object1.bin": object_payload,
        "word/embeddings/package1.bin": b"PACKAGE_PAYLOAD_DO_NOT_LEAK",
        "word/activeX/activeX1.xml": b"<activeXControl/>",
        "word/activeX/_rels/activeX1.xml.rels": (
            f'<Relationships xmlns="{PR}">'
            f'<Relationship Id="{binary_control_id}" '
            f'Type="{_ACTIVE_X_CONTROL_BINARY_RELATIONSHIP_TYPE}" '
            'Target="activeX1.bin"/>'
            "</Relationships>"
        ).encode(),
        "word/activeX/activeX1.bin": b"CONTROL_PAYLOAD_DO_NOT_LEAK",
    }
    if include_import_payload:
        entries["word/afchunk1.html"] = import_payload

    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for name, payload in entries.items():
            archive.writestr(name, payload)


def _write_data_binding_document(
    path,
    *,
    custom_xml_value: str = "DATA_BINDING_PAYLOAD_DO_NOT_LEAK",
    xpath: str = "/private/DATA_BINDING_XPATH_DO_NOT_LEAK",
    store_item_id: str = "{11111111-1111-1111-1111-111111111111}",
    item_store_item_id: str | None = None,
    relationship_id_suffix: str = "1",
    include_data_binding: bool = True,
    binding_count: int = 1,
    include_store_item_id: bool = True,
    strict_syntax: bool = False,
    data_binding_outside_sdt_properties: bool = False,
    invalid_custom_xml_properties_root: bool = False,
    custom_xml_properties_target_mode: str = "Internal",
) -> None:
    word_namespace = _STRICT_WORD_NAMESPACE if strict_syntax else W
    custom_xml_relationship_type = (
        _STRICT_CUSTOM_XML_RELATIONSHIP_TYPE
        if strict_syntax
        else _CUSTOM_XML_RELATIONSHIP_TYPE
    )
    custom_xml_properties_relationship_type = (
        _STRICT_CUSTOM_XML_PROPERTIES_RELATIONSHIP_TYPE
        if strict_syntax
        else _CUSTOM_XML_PROPERTIES_RELATIONSHIP_TYPE
    )
    custom_xml_data_properties_namespace = (
        _STRICT_CUSTOM_XML_DATA_PROPERTIES_NAMESPACE
        if strict_syntax
        else _CUSTOM_XML_DATA_PROPERTIES_NAMESPACE
    )
    binding_store_item_id = (
        f' w:storeItemID="{store_item_id}"' if include_store_item_id else ""
    )
    data_binding = (
        f'<w:dataBinding w:xpath="{xpath}" '
        "w:prefixMappings=\"xmlns:private='DATA_BINDING_PREFIXES_DO_NOT_LEAK'\""
        f"{binding_store_item_id}/>"
    )
    run = "<w:r><w:t>DATA_BINDING_VISIBLE_DO_NOT_LEAK</w:t></w:r>"
    if include_data_binding and data_binding_outside_sdt_properties:
        paragraph = f"<w:p>{data_binding}{run}</w:p>"
    elif include_data_binding:
        structured_document_tag = (
            "<w:sdt><w:sdtPr>"
            f"{data_binding}"
            f"</w:sdtPr><w:sdtContent>{run}</w:sdtContent></w:sdt>"
        )
        paragraph = f"<w:p>{structured_document_tag * binding_count}</w:p>"
    else:
        paragraph = f"<w:p>{run}</w:p>"
    custom_xml_properties_override = (
        '<Override PartName="/customXml/itemProps1.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.'
        'customXmlProperties+xml"/>'
        if include_data_binding
        else ""
    )
    entries: dict[str, bytes] = {
        "[Content_Types].xml": (
            f'<Types xmlns="{CT}">'
            '<Default Extension="rels" '
            'ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            f'<Override PartName="/word/document.xml" ContentType="{DOCX_MAIN_TYPE}"/>'
            f"{custom_xml_properties_override}"
            "</Types>"
        ).encode(),
        "word/document.xml": (
            f'<w:document xmlns:w="{word_namespace}"><w:body>{paragraph}'
            "<w:sectPr/></w:body></w:document>"
        ).encode(),
    }
    if include_data_binding:
        item_id = item_store_item_id or store_item_id
        data_relationship_id = f"rIdCustomXml{relationship_id_suffix}"
        properties_relationship_id = f"rIdProperties{relationship_id_suffix}"
        properties_target_mode = (
            f' TargetMode="{custom_xml_properties_target_mode}"'
            if custom_xml_properties_target_mode != "Internal"
            else ""
        )
        entries["word/_rels/document.xml.rels"] = (
            f'<Relationships xmlns="{PR}">'
            f'<Relationship Id="{data_relationship_id}" '
            f'Type="{custom_xml_relationship_type}" '
            'Target="../customXml/item1.xml"/>'
            "</Relationships>"
        ).encode()
        entries["customXml/item1.xml"] = (
            f"<private><value>{custom_xml_value}</value></private>"
        ).encode()
        entries["customXml/_rels/item1.xml.rels"] = (
            f'<Relationships xmlns="{PR}">'
            f'<Relationship Id="{properties_relationship_id}" '
            f'Type="{custom_xml_properties_relationship_type}" '
            f'Target="itemProps1.xml"{properties_target_mode}/>'
            "</Relationships>"
        ).encode()
        entries["customXml/itemProps1.xml"] = (
            b"<invalid/>"
            if invalid_custom_xml_properties_root
            else (
                f'<ds:datastoreItem xmlns:ds="{custom_xml_data_properties_namespace}" '
                f'ds:itemID="{item_id}"/>'
            ).encode()
        )

    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for name, payload in entries.items():
            archive.writestr(name, payload)


def _write_external_document_dependency_document(
    path,
    *,
    relationship_id_suffix: str = "1",
    include_dependencies: bool = True,
    include_anchors: bool = True,
    strict_syntax: bool = False,
    attached_template_target: str = (
        "https://example.invalid/ATTACHED_TEMPLATE_DO_NOT_LEAK"
    ),
    subdocument_target: str = "https://example.invalid/SUBDOCUMENT_DO_NOT_LEAK",
    frame_target: str = "https://example.invalid/FRAME_SOURCE_DO_NOT_LEAK",
    attached_template_relationship_type: str = _ATTACHED_TEMPLATE_RELATIONSHIP_TYPE,
    subdocument_relationship_type: str = _SUBDOCUMENT_RELATIONSHIP_TYPE,
    frame_relationship_type: str = _FRAME_RELATIONSHIP_TYPE,
    dependency_target_mode: str = "External",
    web_settings_target_mode: str = "Internal",
) -> None:
    word_namespace = _STRICT_WORD_NAMESPACE if strict_syntax else W
    relationship_namespace = _STRICT_RELATIONSHIP_NAMESPACE if strict_syntax else R
    attached_template_type = (
        _STRICT_ATTACHED_TEMPLATE_RELATIONSHIP_TYPE
        if strict_syntax
        else attached_template_relationship_type
    )
    subdocument_type = (
        _STRICT_SUBDOCUMENT_RELATIONSHIP_TYPE
        if strict_syntax
        else subdocument_relationship_type
    )
    frame_type = (
        _STRICT_FRAME_RELATIONSHIP_TYPE if strict_syntax else frame_relationship_type
    )
    web_settings_type = (
        _STRICT_WEB_SETTINGS_RELATIONSHIP_TYPE
        if strict_syntax
        else _WEB_SETTINGS_RELATIONSHIP_TYPE
    )
    attached_template_id = f"rIdAttachedTemplate{relationship_id_suffix}"
    subdocument_id = f"rIdSubdocument{relationship_id_suffix}"
    frame_id = f"rIdFrame{relationship_id_suffix}"
    dependency_target_mode_attribute = (
        f' TargetMode="{dependency_target_mode}"'
        if dependency_target_mode != "Internal"
        else ""
    )
    web_settings_target_mode_attribute = (
        f' TargetMode="{web_settings_target_mode}"'
        if web_settings_target_mode != "Internal"
        else ""
    )
    include_dependency_anchors = include_dependencies and include_anchors
    attached_template_markup = (
        f'<w:attachedTemplate r:id="{attached_template_id}"/>'
        if include_dependency_anchors
        else ""
    )
    subdocument_markup = (
        f'<w:subDoc r:id="{subdocument_id}"/>' if include_dependency_anchors else ""
    )
    frame_source_markup = (
        f'<w:sourceFileName r:id="{frame_id}"/>' if include_dependency_anchors else ""
    )
    web_settings_override = (
        '<Override PartName="/word/webSettings.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.'
        'wordprocessingml.webSettings+xml"/>'
        if include_dependencies
        else ""
    )
    entries: dict[str, bytes] = {
        "[Content_Types].xml": (
            f'<Types xmlns="{CT}">'
            '<Default Extension="rels" '
            'ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            f'<Override PartName="/word/document.xml" ContentType="{DOCX_MAIN_TYPE}"/>'
            f"{web_settings_override}"
            "</Types>"
        ).encode(),
        "word/document.xml": (
            f'<w:document xmlns:w="{word_namespace}" '
            f'xmlns:r="{relationship_namespace}"><w:body>'
            "<w:p><w:r><w:t>VISIBLE_DO_NOT_LEAK</w:t></w:r></w:p>"
            f"{subdocument_markup}"
            "<w:sectPr/></w:body></w:document>"
        ).encode(),
    }
    if include_dependencies:
        entries["word/_rels/document.xml.rels"] = (
            f'<Relationships xmlns="{PR}">'
            f'<Relationship Id="rIdSettings{relationship_id_suffix}" '
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/'
            'relationships/settings" Target="settings.xml"/>'
            f'<Relationship Id="{subdocument_id}" Type="{subdocument_type}" '
            f'Target="{subdocument_target}"{dependency_target_mode_attribute}/>'
            f'<Relationship Id="rIdWebSettings{relationship_id_suffix}" '
            f'Type="{web_settings_type}" Target="webSettings.xml"'
            f"{web_settings_target_mode_attribute}/>"
            "</Relationships>"
        ).encode()
        entries["word/settings.xml"] = (
            f'<w:settings xmlns:w="{word_namespace}" '
            f'xmlns:r="{relationship_namespace}">{attached_template_markup}'
            "</w:settings>"
        ).encode()
        entries["word/_rels/settings.xml.rels"] = (
            f'<Relationships xmlns="{PR}">'
            f'<Relationship Id="{attached_template_id}" '
            f'Type="{attached_template_type}" Target="{attached_template_target}"'
            f"{dependency_target_mode_attribute}/>"
            "</Relationships>"
        ).encode()
        entries["word/webSettings.xml"] = (
            f'<w:webSettings xmlns:w="{word_namespace}" '
            f'xmlns:r="{relationship_namespace}"><w:frameset><w:frame>'
            '<w:name w:val="FRAME_NAME_DO_NOT_LEAK"/>'
            f"{frame_source_markup}"
            "</w:frame></w:frameset></w:webSettings>"
        ).encode()
        entries["word/_rels/webSettings.xml.rels"] = (
            f'<Relationships xmlns="{PR}">'
            f'<Relationship Id="{frame_id}" Type="{frame_type}" '
            f'Target="{frame_target}"{dependency_target_mode_attribute}/>'
            "</Relationships>"
        ).encode()

    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for name, payload in entries.items():
            archive.writestr(name, payload)


def _write_glossary_attached_template_document(
    path, *, relationship_id_suffix: str = "1", strict_syntax: bool = False
) -> None:
    attached_template_id = f"rIdGlossaryTemplate{relationship_id_suffix}"
    word_namespace = _STRICT_WORD_NAMESPACE if strict_syntax else W
    relationship_namespace = _STRICT_RELATIONSHIP_NAMESPACE if strict_syntax else R
    settings_relationship_type = (
        _STRICT_DOCUMENT_SETTINGS_RELATIONSHIP_TYPE
        if strict_syntax
        else "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings"
    )
    attached_template_relationship_type = (
        _STRICT_ATTACHED_TEMPLATE_RELATIONSHIP_TYPE
        if strict_syntax
        else _ATTACHED_TEMPLATE_RELATIONSHIP_TYPE
    )
    entries: dict[str, bytes] = {
        "[Content_Types].xml": (
            f'<Types xmlns="{CT}">'
            '<Default Extension="rels" '
            'ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            f'<Override PartName="/word/document.xml" ContentType="{DOCX_MAIN_TYPE}"/>'
            '<Override PartName="/word/glossary/document.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.'
            'wordprocessingml.document.glossaryDocument+xml"/>'
            '<Override PartName="/word/glossary/glossarySettings.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.'
            'wordprocessingml.settings+xml"/>'
            "</Types>"
        ).encode(),
        "word/document.xml": (
            f'<w:document xmlns:w="{word_namespace}"><w:body>'
            "<w:p><w:r><w:t>VISIBLE_DO_NOT_LEAK</w:t></w:r></w:p>"
            "<w:sectPr/></w:body></w:document>"
        ).encode(),
        "word/_rels/document.xml.rels": (
            f'<Relationships xmlns="{PR}">'
            '<Relationship Id="rIdGlossary" '
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/'
            'relationships/glossaryDocument" Target="glossary/document.xml"/>'
            "</Relationships>"
        ).encode(),
        "word/glossary/document.xml": (
            f'<w:glossaryDocument xmlns:w="{word_namespace}"><w:docParts><w:docPart>'
            "<w:docPartBody><w:p><w:r><w:t>GLOSSARY_DO_NOT_LEAK</w:t>"
            "</w:r></w:p></w:docPartBody></w:docPart></w:docParts>"
            "</w:glossaryDocument>"
        ).encode(),
        "word/glossary/_rels/document.xml.rels": (
            f'<Relationships xmlns="{PR}">'
            '<Relationship Id="rIdSettings" '
            f'Type="{settings_relationship_type}" '
            'Target="glossarySettings.xml"/>'
            "</Relationships>"
        ).encode(),
        "word/glossary/glossarySettings.xml": (
            f'<w:settings xmlns:w="{word_namespace}" '
            f'xmlns:r="{relationship_namespace}">'
            f'<w:attachedTemplate r:id="{attached_template_id}"/>'
            "</w:settings>"
        ).encode(),
        "word/glossary/_rels/glossarySettings.xml.rels": (
            f'<Relationships xmlns="{PR}">'
            f'<Relationship Id="{attached_template_id}" '
            f'Type="{attached_template_relationship_type}" '
            'Target="https://example.invalid/GLOSSARY_TEMPLATE_DO_NOT_LEAK" '
            'TargetMode="External"/>'
            "</Relationships>"
        ).encode(),
    }

    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for name, payload in entries.items():
            archive.writestr(name, payload)


def _write_mail_merge_document(
    path,
    *,
    query: str = "MAIL_MERGE_QUERY_DO_NOT_LEAK",
    recipient_value: str = "MAIL_MERGE_RECIPIENT_DO_NOT_LEAK",
    relationship_id_suffix: str = "1",
    include_mail_merge: bool = True,
    include_configuration: bool = True,
    include_recipient_part: bool = True,
    strict_syntax: bool = False,
    source_relationship_type: str = _MAIL_MERGE_SOURCE_RELATIONSHIP_TYPE,
    source_target_mode: str = "External",
    recipient_target_mode: str = "Internal",
    recipient_relationship_type: str = _MAIL_MERGE_RECIPIENT_DATA_RELATIONSHIP_TYPE,
) -> None:
    word_namespace = _STRICT_WORD_NAMESPACE if strict_syntax else W
    relationship_namespace = _STRICT_RELATIONSHIP_NAMESPACE if strict_syntax else R
    data_source_relationship_type = (
        _STRICT_MAIL_MERGE_SOURCE_RELATIONSHIP_TYPE
        if strict_syntax
        else source_relationship_type
    )
    header_source_relationship_type = (
        _STRICT_MAIL_MERGE_HEADER_SOURCE_RELATIONSHIP_TYPE
        if strict_syntax
        else _MAIL_MERGE_HEADER_SOURCE_RELATIONSHIP_TYPE
    )
    recipient_data_relationship_type = (
        _STRICT_MAIL_MERGE_RECIPIENT_DATA_RELATIONSHIP_TYPE
        if strict_syntax
        else recipient_relationship_type
    )
    source_id = f"rIdSource{relationship_id_suffix}"
    odso_source_id = f"rIdOdsoSource{relationship_id_suffix}"
    header_id = f"rIdHeader{relationship_id_suffix}"
    recipient_id = f"rIdRecipient{relationship_id_suffix}"
    source_target = (
        "mailMergeSource.bin"
        if source_target_mode == "Internal"
        else "file:///private/MAIL_MERGE_DATA_SOURCE_DO_NOT_LEAK"
    )
    recipient_target = "recipientData.xml"
    source_target_mode_attribute = (
        f' TargetMode="{source_target_mode}"'
        if source_target_mode != "Internal"
        else ""
    )
    recipient_target_mode_attribute = (
        f' TargetMode="{recipient_target_mode}"'
        if recipient_target_mode != "Internal"
        else ""
    )
    configuration = ""
    if include_configuration:
        configuration = f"""<w:mailMerge>
  <w:mainDocumentType w:val="formLetters"/>
  <w:dataSource r:id="{source_id}"/>
  <w:headerSource r:id="{header_id}"/>
  <w:query w:val="{query}"/>
  <w:odso>
    <w:udl w:val="MAIL_MERGE_UDL_DO_NOT_LEAK"/>
    <w:table w:val="MAIL_MERGE_TABLE_DO_NOT_LEAK"/>
    <w:src r:id="{odso_source_id}"/>
    <w:fieldMapData>
      <w:name w:val="MAIL_MERGE_FIELD_NAME_DO_NOT_LEAK"/>
      <w:mappedName w:val="MAIL_MERGE_FIELD_MAPPING_DO_NOT_LEAK"/>
    </w:fieldMapData>
    <w:recipientData r:id="{recipient_id}"/>
  </w:odso>
</w:mailMerge>"""
    entries: dict[str, bytes] = {
        "[Content_Types].xml": (
            f'<Types xmlns="{CT}">'
            f'<Override PartName="/word/document.xml" ContentType="{DOCX_MAIN_TYPE}"/>'
            "</Types>"
        ).encode(),
        "word/document.xml": (
            f'<w:document xmlns:w="{word_namespace}"><w:body>'
            "<w:p><w:r><w:t>VISIBLE_DO_NOT_LEAK</w:t></w:r></w:p>"
            "<w:sectPr/></w:body></w:document>"
        ).encode(),
    }
    if include_mail_merge:
        entries["word/settings.xml"] = (
            f'<w:settings xmlns:w="{word_namespace}" '
            f'xmlns:r="{relationship_namespace}">{configuration}</w:settings>'
        ).encode()
        entries["word/_rels/settings.xml.rels"] = (
            f'<Relationships xmlns="{PR}">'
            f'<Relationship Id="{source_id}" Type="{data_source_relationship_type}" '
            f'Target="{source_target}"{source_target_mode_attribute}/>'
            f'<Relationship Id="{odso_source_id}" '
            f'Type="{data_source_relationship_type}" '
            'Target="file:///private/MAIL_MERGE_DATA_SOURCE_DO_NOT_LEAK" '
            'TargetMode="External"/>'
            f'<Relationship Id="{header_id}" Type="{header_source_relationship_type}" '
            'Target="file:///private/MAIL_MERGE_HEADER_SOURCE_DO_NOT_LEAK" '
            'TargetMode="External"/>'
            f'<Relationship Id="{recipient_id}" '
            f'Type="{recipient_data_relationship_type}" '
            f'Target="{recipient_target}"{recipient_target_mode_attribute}/>'
            "</Relationships>"
        ).encode()
        if include_recipient_part:
            entries["word/recipientData.xml"] = (
                f'<w:recipients xmlns:w="{word_namespace}">'
                '<w:recipientData><w:active w:val="0"/>'
                f'<w:hash w:val="{recipient_value}"/>'
                "</w:recipientData></w:recipients>"
            ).encode()

    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for name, payload in entries.items():
            archive.writestr(name, payload)


def _write_document_property_document(
    path,
    *,
    core_title: str = "CORE_TITLE_DO_NOT_LEAK",
    relationship_id_suffix: str = "1",
    include_properties: bool = True,
    include_property_relationships: bool = True,
    strict_property_syntax: bool = False,
    malformed_core: bool = False,
) -> None:
    entries: dict[str, bytes] = {
        "[Content_Types].xml": (
            f'<Types xmlns="{CT}">'
            f'<Override PartName="/word/document.xml" ContentType="{DOCX_MAIN_TYPE}"/>'
            "</Types>"
        ).encode(),
        "word/document.xml": (
            f'<w:document xmlns:w="{W}"><w:body>'
            "<w:p><w:r><w:t>VISIBLE_DO_NOT_LEAK</w:t></w:r></w:p>"
            "<w:sectPr/></w:body></w:document>"
        ).encode(),
    }
    extended_relationship_type = (
        _STRICT_EXTENDED_PROPERTIES_RELATIONSHIP_TYPE
        if strict_property_syntax
        else _EXTENDED_PROPERTIES_RELATIONSHIP_TYPE
    )
    custom_relationship_type = (
        _STRICT_CUSTOM_PROPERTIES_RELATIONSHIP_TYPE
        if strict_property_syntax
        else _CUSTOM_PROPERTIES_RELATIONSHIP_TYPE
    )
    extended_properties_namespace = (
        _STRICT_EXTENDED_PROPERTIES_NAMESPACE
        if strict_property_syntax
        else _EXTENDED_PROPERTIES_NAMESPACE
    )
    custom_properties_namespace = (
        _STRICT_CUSTOM_PROPERTIES_NAMESPACE
        if strict_property_syntax
        else _CUSTOM_PROPERTIES_NAMESPACE
    )
    document_properties_vt_namespace = (
        _STRICT_DOCUMENT_PROPERTIES_VT_NAMESPACE
        if strict_property_syntax
        else _DOCUMENT_PROPERTIES_VT_NAMESPACE
    )
    if include_properties:
        if include_property_relationships:
            entries["_rels/.rels"] = (
                f'<Relationships xmlns="{PR}">'
                f'<Relationship Id="rIdCore{relationship_id_suffix}" '
                f'Type="{_CORE_PROPERTIES_RELATIONSHIP_TYPE}" '
                'Target="docProps/core.xml"/>'
                f'<Relationship Id="rIdExtended{relationship_id_suffix}" '
                f'Type="{extended_relationship_type}" '
                'Target="docProps/app.xml"/>'
                f'<Relationship Id="rIdCustom{relationship_id_suffix}" '
                f'Type="{custom_relationship_type}" '
                'Target="docProps/custom.xml"/>'
                "</Relationships>"
            ).encode()
        entries["docProps/core.xml"] = (
            b"<invalid/>"
            if malformed_core
            else (
                f'<cp:coreProperties xmlns:cp="{_CORE_PROPERTIES_NAMESPACE}" '
                f'xmlns:dc="{_DUBLIN_CORE_NAMESPACE}">'
                "<dc:creator>CORE_CREATOR_DO_NOT_LEAK</dc:creator>"
                f"<dc:title>{core_title}</dc:title>"
                "</cp:coreProperties>"
            ).encode()
        )
        entries["docProps/app.xml"] = (
            f'<Properties xmlns="{extended_properties_namespace}">'
            "<Company>EXTENDED_COMPANY_DO_NOT_LEAK</Company>"
            "<Manager>EXTENDED_MANAGER_DO_NOT_LEAK</Manager>"
            "</Properties>"
        ).encode()
        entries["docProps/custom.xml"] = (
            f'<Properties xmlns="{custom_properties_namespace}" '
            f'xmlns:vt="{document_properties_vt_namespace}">'
            '<property fmtid="FMTID" pid="2" '
            'name="CUSTOM_PROPERTY_NAME_DO_NOT_LEAK">'
            "<vt:lpwstr>CUSTOM_PROPERTY_VALUE_DO_NOT_LEAK</vt:lpwstr>"
            "</property>"
            '<property fmtid="FMTID" pid="3" '
            'name="SECOND_CUSTOM_PROPERTY_NAME_DO_NOT_LEAK">'
            "<vt:lpwstr>SECOND_CUSTOM_PROPERTY_VALUE_DO_NOT_LEAK</vt:lpwstr>"
            "</property>"
            "</Properties>"
        ).encode()

    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for name, payload in entries.items():
            archive.writestr(name, payload)
