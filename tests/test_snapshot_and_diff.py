from __future__ import annotations

import json
import posixpath
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
_DOCUMENT_SETTINGS_RELATIONSHIP_TYPE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings"
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
_STRICT_PACKAGE_RELATIONSHIP_NAMESPACE = (
    "http://purl.oclc.org/ooxml/package/relationships"
)
_HYPERLINK_RELATIONSHIP_TYPE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
)
_STRICT_HYPERLINK_RELATIONSHIP_TYPE = (
    "http://purl.oclc.org/ooxml/officeDocument/relationships/hyperlink"
)
_IMAGE_RELATIONSHIP_TYPE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
)
_STRICT_IMAGE_RELATIONSHIP_TYPE = (
    "http://purl.oclc.org/ooxml/officeDocument/relationships/image"
)
_DRAWING_NAMESPACE = "http://schemas.openxmlformats.org/drawingml/2006/main"
_STRICT_DRAWING_NAMESPACE = "http://purl.oclc.org/ooxml/drawingml/main"
_VML_NAMESPACE = "urn:schemas-microsoft-com:vml"
_WORDPROCESSING_DRAWING_NAMESPACE = (
    "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
)
_PICTURE_NAMESPACE = "http://schemas.openxmlformats.org/drawingml/2006/picture"
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
_DOTX_MAIN_TYPE = (
    "application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml"
)
_DOTM_MAIN_TYPE = "application/vnd.ms-word.template.macroEnabledTemplate.main+xml"
_WORD_2010_11_NAMESPACE = "http://schemas.microsoft.com/office/word/2010/11/wordml"
_WORD_2012_NAMESPACE = "http://schemas.microsoft.com/office/word/2012/wordml"
_WORD_2016_COMMENT_IDS_NAMESPACE = (
    "http://schemas.microsoft.com/office/word/2016/wordml/cid"
)
_WORD_2018_COMMENT_EXTENSIBLE_NAMESPACE = (
    "http://schemas.microsoft.com/office/word/2018/wordml/cex"
)
_WORD_2018_NAMESPACE = "http://schemas.microsoft.com/office/word/2018/wordml"
_COMMENT_REACTIONS_NAMESPACE = (
    "http://schemas.microsoft.com/office/comments/2020/reactions"
)
_COMMENTS_RELATIONSHIP_TYPE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
)
_PEOPLE_RELATIONSHIP_TYPE = (
    "http://schemas.microsoft.com/office/2011/relationships/people"
)
_COMMENTS_EXTENDED_RELATIONSHIP_TYPE = (
    "http://schemas.microsoft.com/office/2011/relationships/commentsExtended"
)
_COMMENTS_IDS_RELATIONSHIP_TYPE = (
    "http://schemas.microsoft.com/office/2016/09/relationships/commentsIds"
)
_COMMENTS_EXTENSIBLE_RELATIONSHIP_TYPE = (
    "http://schemas.microsoft.com/office/2018/08/relationships/commentsExtensible"
)
_DOCUMENT_TASK_NAMESPACE = (
    "http://schemas.microsoft.com/office/tasks/2019/documenttasks"
)
_DOCUMENT_TASK_RELATIONSHIP_TYPE = (
    "http://schemas.microsoft.com/office/2019/05/relationships/documenttasks"
)
_DOCUMENT_TASK_CONTENT_TYPE = "application/vnd.ms-office.documenttasks+xml"
_TASKPANE_WEB_EXTENSION_TASKPANES_NAMESPACE = (
    "http://schemas.microsoft.com/office/webextensions/taskpanes/2010/11"
)
_TASKPANE_WEB_EXTENSION_NAMESPACE = (
    "http://schemas.microsoft.com/office/webextensions/webextension/2010/11"
)
_TASKPANE_WEB_EXTENSION_RELATIONSHIP_TYPE = (
    "http://schemas.microsoft.com/office/2011/relationships/webextensiontaskpanes"
)
_WEB_EXTENSION_RELATIONSHIP_TYPE = (
    "http://schemas.microsoft.com/office/2011/relationships/webextension"
)
_TASKPANE_WEB_EXTENSION_CONTENT_TYPE = (
    "application/vnd.ms-office.webextensiontaskpanes+xml"
)
_WEB_EXTENSION_CONTENT_TYPE = "application/vnd.ms-office.webextension+xml"
_SENSITIVITY_LABEL_NAMESPACE = (
    "http://schemas.microsoft.com/office/2020/mipLabelMetadata"
)
_SENSITIVITY_LABEL_RELATIONSHIP_TYPE = (
    "http://schemas.microsoft.com/office/2020/02/relationships/classificationlabels"
)
_SENSITIVITY_LABEL_CONTENT_TYPE = (
    "application/vnd.ms-office.classificationlabels+xml"
)
_PACKAGE_DIGITAL_SIGNATURE_ORIGIN_RELATIONSHIP_TYPE = (
    "http://schemas.openxmlformats.org/package/2006/relationships/"
    "digital-signature/origin"
)
_PACKAGE_DIGITAL_SIGNATURE_RELATIONSHIP_TYPE = (
    "http://schemas.openxmlformats.org/package/2006/relationships/"
    "digital-signature/signature"
)
_PACKAGE_DIGITAL_SIGNATURE_CERTIFICATE_RELATIONSHIP_TYPE = (
    "http://schemas.openxmlformats.org/package/2006/relationships/"
    "digital-signature/certificate"
)
_PACKAGE_DIGITAL_SIGNATURE_ORIGIN_CONTENT_TYPE = (
    "application/vnd.openxmlformats-package.digital-signature-origin"
)
_PACKAGE_DIGITAL_SIGNATURE_XML_CONTENT_TYPE = (
    "application/vnd.openxmlformats-package.digital-signature-xmlsignature+xml"
)
_PACKAGE_DIGITAL_SIGNATURE_CERTIFICATE_CONTENT_TYPE = (
    "application/vnd.openxmlformats-package.digital-signature-certificate"
)
_XMLDSIG_NAMESPACE = "http://www.w3.org/2000/09/xmldsig#"
_OPC_DIGITAL_SIGNATURE_NAMESPACE = (
    "http://schemas.openxmlformats.org/package/2006/digital-signature"
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


def test_modern_comment_metadata_inventory_is_private_and_semantic(tmp_path) -> None:
    before = tmp_path / "before.docx"
    after = tmp_path / "after.docx"
    payload_changed = tmp_path / "payload-changed.docx"
    identifier_changed = tmp_path / "identifier-changed.docx"
    renumbered = tmp_path / "renumbered.docx"
    policy_path = tmp_path / "docfence.yml"
    _write_modern_comment_metadata_document(before, include_metadata=False)
    _write_modern_comment_metadata_document(after)
    _write_modern_comment_metadata_document(
        payload_changed,
        metadata_marker="MODERN_COMMENT_METADATA_CHANGED_DO_NOT_LEAK",
    )
    _write_modern_comment_metadata_document(
        identifier_changed,
        comment_identity_marker="MODERN_COMMENT_IDENTIFIERS_CHANGED_DO_NOT_LEAK",
    )
    _write_modern_comment_metadata_document(renumbered, relationship_id_suffix="9")

    expected_inventory = {
        "people_part_count": 1,
        "person_count": 2,
        "presence_info_count": 2,
        "comments_extended_part_count": 1,
        "comment_extension_count": 2,
        "threaded_comment_count": 1,
        "resolved_comment_count": 1,
        "comments_id_part_count": 1,
        "comment_id_count": 2,
        "comments_extensible_part_count": 1,
        "comment_extensible_count": 2,
        "reaction_count": 2,
        "reaction_user_count": 3,
    }
    snapshot = load_snapshot(after)
    assert snapshot.public_dict()["modern_comment_metadata"] == expected_inventory
    assert load_snapshot(before).public_dict()["modern_comment_metadata"] == {
        key: 0 for key in expected_inventory
    }
    assert snapshot.comment_count == load_snapshot(before).comment_count == 1
    assert (
        snapshot.unclassified_part_count
        == load_snapshot(before).unclassified_part_count
    )

    report = diff_documents(before, after)
    assert "modern_comment_metadata_inventory_changed" in {
        change.kind for change in report.changes
    }
    assert {
        change.kind for change in diff_documents(after, payload_changed).changes
    } == {"modern_comment_metadata_inventory_changed"}
    assert {
        change.kind for change in diff_documents(after, identifier_changed).changes
    } == {"modern_comment_metadata_inventory_changed"}
    assert diff_documents(after, renumbered).changes == ()

    policy_path.write_text(
        """version: 1
rules:
  require_no_modern_comment_metadata: true
  no_modern_comment_metadata_changes: true
""",
        encoding="utf-8",
    )
    policy = load_policy(policy_path)
    assert {finding.rule_id for finding in apply_policy(report, policy).findings} == {
        "DFP029",
        "DFP030",
    }
    assert {
        finding.rule_id
        for finding in apply_policy(
            diff_documents(after, payload_changed), policy
        ).findings
    } == {"DFP029", "DFP030"}
    assert {
        finding.rule_id
        for finding in apply_policy(diff_documents(after, renumbered), policy).findings
    } == {"DFP029"}

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
        "DFC_MODERN_COMMENT_METADATA_INVENTORY_CHANGED",
        "DFP029",
        "DFP030",
    } <= {result["ruleId"] for result in sarif["runs"][0]["results"]}
    for marker in (
        "MODERN_COMMENT_AUTHOR_DO_NOT_LEAK",
        "MODERN_COMMENT_PROVIDER_DO_NOT_LEAK",
        "MODERN_COMMENT_USER_ID_DO_NOT_LEAK",
        "MODERN_COMMENT_DURABLE_ID_DO_NOT_LEAK",
        "MODERN_COMMENT_METADATA_CHANGED_DO_NOT_LEAK",
        "MODERN_COMMENT_IDENTIFIERS_CHANGED_DO_NOT_LEAK",
        "MODERN_COMMENT_REACTION_USER_DO_NOT_LEAK",
    ):
        assert marker not in rendered


def test_modern_comment_metadata_discovers_noncanonical_parts_and_rejects_invalid_parts(
    tmp_path,
) -> None:
    relationship_only = tmp_path / "relationship-only.docx"
    content_type_only = tmp_path / "content-type-only.docx"
    conventional_unlinked = tmp_path / "conventional-unlinked.docx"
    legacy_office_15 = tmp_path / "legacy-office-15.docx"
    wrong_root = tmp_path / "wrong-root.docx"
    external_relationship = tmp_path / "external-relationship.docx"
    part_names = {
        "people": "word/reviewer-data/people-records.xml",
        "comments_extended": "word/reviewer-data/comment-threads.xml",
        "comments_ids": "word/reviewer-data/comment-identifiers.xml",
        "comments_extensible": "word/reviewer-data/comment-reactions.xml",
    }
    _write_modern_comment_metadata_document(
        relationship_only,
        part_names=part_names,
        include_metadata_content_types=False,
    )
    _write_modern_comment_metadata_document(
        content_type_only,
        part_names=part_names,
        include_metadata_relationships=False,
    )
    _write_modern_comment_metadata_document(
        conventional_unlinked,
        include_metadata_relationships=False,
        include_metadata_content_types=False,
    )
    _write_modern_comment_metadata_document(
        legacy_office_15,
        office_15_namespace=_WORD_2010_11_NAMESPACE,
    )
    _write_modern_comment_metadata_document(wrong_root, wrong_people_root=True)
    _write_modern_comment_metadata_document(
        external_relationship,
        modern_relationship_target_mode="External",
    )

    expected_inventory = {
        "people_part_count": 1,
        "person_count": 2,
        "presence_info_count": 2,
        "comments_extended_part_count": 1,
        "comment_extension_count": 2,
        "threaded_comment_count": 1,
        "resolved_comment_count": 1,
        "comments_id_part_count": 1,
        "comment_id_count": 2,
        "comments_extensible_part_count": 1,
        "comment_extensible_count": 2,
        "reaction_count": 2,
        "reaction_user_count": 3,
    }
    assert load_snapshot(relationship_only).public_dict()[
        "modern_comment_metadata"
    ] == (expected_inventory)
    assert load_snapshot(content_type_only).public_dict()[
        "modern_comment_metadata"
    ] == (expected_inventory)
    assert load_snapshot(conventional_unlinked).public_dict()[
        "modern_comment_metadata"
    ] == (expected_inventory)
    assert load_snapshot(legacy_office_15).public_dict()["modern_comment_metadata"] == (
        expected_inventory
    )
    with pytest.raises(DocumentFormatError):
        load_snapshot(wrong_root)
    with pytest.raises(DocumentFormatError):
        load_snapshot(external_relationship)


def test_document_tasks_and_taskpane_web_extensions_are_private_and_semantic(
    tmp_path,
) -> None:
    before = tmp_path / "before.docx"
    after = tmp_path / "after.docx"
    task_payload_changed = tmp_path / "task-payload-changed.docx"
    web_extension_payload_changed = tmp_path / "web-extension-payload-changed.docx"
    renumbered = tmp_path / "renumbered.docx"
    policy_path = tmp_path / "docfence.yml"
    _write_document_task_and_web_extension_document(
        before,
        include_document_tasks=False,
        include_taskpane_web_extensions=False,
    )
    _write_document_task_and_web_extension_document(after)
    _write_document_task_and_web_extension_document(
        task_payload_changed,
        task_marker="DOCUMENT_TASK_PAYLOAD_CHANGED_DO_NOT_LEAK",
    )
    _write_document_task_and_web_extension_document(
        web_extension_payload_changed,
        web_extension_marker="WEB_EXTENSION_PAYLOAD_CHANGED_DO_NOT_LEAK",
    )
    _write_document_task_and_web_extension_document(
        renumbered,
        relationship_id_suffix="9",
    )

    expected_tasks = {
        "document_task_part_count": 1,
        "task_count": 1,
        "task_history_event_count": 11,
        "task_user_reference_count": 13,
        "task_comment_anchor_count": 2,
        "assignment_event_count": 1,
        "unassignment_event_count": 1,
        "creation_event_count": 1,
        "title_change_event_count": 1,
        "schedule_change_event_count": 1,
        "progress_change_event_count": 1,
        "priority_change_event_count": 1,
        "deletion_event_count": 1,
        "restoration_event_count": 1,
        "unassign_all_event_count": 1,
        "undo_event_count": 1,
    }
    expected_web_extensions = {
        "taskpane_part_count": 1,
        "taskpane_count": 2,
        "visible_taskpane_count": 1,
        "locked_taskpane_count": 1,
        "web_extension_part_count": 2,
        "web_extension_reference_count": 3,
        "web_extension_property_count": 3,
        "web_extension_binding_count": 3,
        "auto_show_taskpane_setting_count": 1,
        "web_extension_bound_content_control_count": 2,
    }
    before_snapshot = load_snapshot(before)
    after_snapshot = load_snapshot(after)
    assert after_snapshot.public_dict()["document_tasks"] == expected_tasks
    assert after_snapshot.public_dict()["taskpane_web_extensions"] == (
        expected_web_extensions
    )
    assert before_snapshot.public_dict()["document_tasks"] == {
        key: 0 for key in expected_tasks
    }
    assert before_snapshot.public_dict()["taskpane_web_extensions"] == {
        key: 0 for key in expected_web_extensions
    }
    assert (
        after_snapshot.unclassified_part_count
        == before_snapshot.unclassified_part_count
    )

    report = diff_documents(before, after)
    assert {
        "document_task_inventory_changed",
        "taskpane_web_extension_inventory_changed",
    } <= {change.kind for change in report.changes}
    assert {
        change.kind for change in diff_documents(after, task_payload_changed).changes
    } == {"document_task_inventory_changed"}
    assert {
        change.kind
        for change in diff_documents(after, web_extension_payload_changed).changes
    } == {"taskpane_web_extension_inventory_changed"}
    assert diff_documents(after, renumbered).changes == ()

    policy_path.write_text(
        """version: 1
rules:
  require_no_document_tasks: true
  no_document_task_changes: true
  require_no_taskpane_web_extensions: true
  no_taskpane_web_extension_changes: true
""",
        encoding="utf-8",
    )
    policy = load_policy(policy_path)
    assert {finding.rule_id for finding in apply_policy(report, policy).findings} == {
        "DFP031",
        "DFP032",
        "DFP033",
        "DFP034",
    }
    assert {
        finding.rule_id
        for finding in apply_policy(
            diff_documents(after, task_payload_changed), policy
        ).findings
    } == {"DFP031", "DFP032", "DFP033"}
    assert {
        finding.rule_id
        for finding in apply_policy(
            diff_documents(after, web_extension_payload_changed), policy
        ).findings
    } == {"DFP031", "DFP033", "DFP034"}
    assert {
        finding.rule_id
        for finding in apply_policy(diff_documents(after, renumbered), policy).findings
    } == {"DFP031", "DFP033"}

    gated = apply_policy(report, policy)
    rendered = "\n".join(
        (
            render_profile(after_snapshot, "json"),
            render_profile(after_snapshot, "markdown"),
            render_report(gated, "json"),
            render_report(gated, "markdown"),
            render_report(gated, "sarif"),
        )
    )
    sarif = json.loads(render_report(gated, "sarif"))
    assert {
        "DFC_DOCUMENT_TASK_INVENTORY_CHANGED",
        "DFC_TASKPANE_WEB_EXTENSION_INVENTORY_CHANGED",
        "DFP031",
        "DFP032",
        "DFP033",
        "DFP034",
    } <= {result["ruleId"] for result in sarif["runs"][0]["results"]}
    for marker in (
        "DOCUMENT_TASK_WORKFLOW_USER_DO_NOT_LEAK",
        "DOCUMENT_TASK_WORKFLOW_TITLE_DO_NOT_LEAK",
        "DOCUMENT_TASK_PAYLOAD_CHANGED_DO_NOT_LEAK",
        "TASKPANE_WEB_EXTENSION_STORE_DO_NOT_LEAK",
        "WEB_EXTENSION_PAYLOAD_CHANGED_DO_NOT_LEAK",
    ):
        assert marker not in rendered


def test_document_tasks_and_taskpane_web_extensions_discover_parts_and_reject_invalid(
    tmp_path,
) -> None:
    relationship_only = tmp_path / "relationship-only.docx"
    content_type_only = tmp_path / "content-type-only.docx"
    conventional_unlinked = tmp_path / "conventional-unlinked.docx"
    wrong_task_root = tmp_path / "wrong-task-root.docx"
    wrong_taskpane_root = tmp_path / "wrong-taskpane-root.docx"
    wrong_web_extension_root = tmp_path / "wrong-web-extension-root.docx"
    external_task_relationship = tmp_path / "external-task-relationship.docx"
    external_taskpane_relationship = tmp_path / "external-taskpane-relationship.docx"
    external_web_extension_relationship = tmp_path / "external-web-extension.docx"
    invalid_taskpane_reference = tmp_path / "invalid-taskpane-reference.docx"
    noncanonical_part_names = (
        "word/review/task-data.xml",
        "word/review/taskpane-data.xml",
        "word/review/extension-primary.xml",
        "word/review/extension-secondary.xml",
    )
    _write_document_task_and_web_extension_document(
        relationship_only,
        part_names=noncanonical_part_names,
        include_feature_content_types=False,
    )
    _write_document_task_and_web_extension_document(
        content_type_only,
        part_names=noncanonical_part_names,
        include_feature_relationships=False,
    )
    _write_document_task_and_web_extension_document(
        conventional_unlinked,
        part_names=(
            "word/tasks",
            "word/webextensions/taskpanes",
            "word/webextensions/webextension",
            "word/webextensions/webextension2.xml",
        ),
        include_feature_relationships=False,
        include_feature_content_types=False,
    )
    _write_document_task_and_web_extension_document(
        wrong_task_root,
        wrong_document_task_root=True,
    )
    _write_document_task_and_web_extension_document(
        wrong_taskpane_root,
        wrong_taskpane_root=True,
    )
    _write_document_task_and_web_extension_document(
        wrong_web_extension_root,
        wrong_web_extension_root=True,
    )
    _write_document_task_and_web_extension_document(
        external_task_relationship,
        document_task_target_mode="External",
    )
    _write_document_task_and_web_extension_document(
        external_taskpane_relationship,
        taskpane_target_mode="External",
    )
    _write_document_task_and_web_extension_document(
        external_web_extension_relationship,
        web_extension_target_mode="External",
    )
    _write_document_task_and_web_extension_document(
        invalid_taskpane_reference,
        invalid_taskpane_reference=True,
    )

    relationship_snapshot = load_snapshot(relationship_only)
    assert relationship_snapshot.document_tasks.task_count == 1
    assert relationship_snapshot.taskpane_web_extensions.taskpane_count == 2
    assert relationship_snapshot.taskpane_web_extensions.web_extension_part_count == 2
    content_type_snapshot = load_snapshot(content_type_only)
    assert content_type_snapshot.document_tasks.document_task_part_count == 1
    assert content_type_snapshot.taskpane_web_extensions.taskpane_part_count == 1
    assert content_type_snapshot.taskpane_web_extensions.taskpane_count == 0
    assert content_type_snapshot.taskpane_web_extensions.web_extension_part_count == 2
    conventional_snapshot = load_snapshot(conventional_unlinked)
    assert conventional_snapshot.document_tasks.task_count == 1
    assert conventional_snapshot.taskpane_web_extensions.taskpane_part_count == 1
    assert conventional_snapshot.taskpane_web_extensions.web_extension_part_count == 2

    for document in (
        wrong_task_root,
        wrong_taskpane_root,
        wrong_web_extension_root,
        external_task_relationship,
        external_taskpane_relationship,
        external_web_extension_relationship,
        invalid_taskpane_reference,
    ):
        with pytest.raises(DocumentFormatError):
            load_snapshot(document)


def test_web_extension_content_control_markers_follow_created_precedence(
    tmp_path,
) -> None:
    created_false = tmp_path / "created-false.docx"
    created_true = tmp_path / "created-true.docx"
    _write_document_task_and_web_extension_document(
        created_false,
        web_extension_created_value="false",
    )
    _write_document_task_and_web_extension_document(
        created_true,
        web_extension_created_value="true",
    )

    assert (
        load_snapshot(created_false)
        .taskpane_web_extensions.web_extension_bound_content_control_count
        == 2
    )
    assert (
        load_snapshot(created_true)
        .taskpane_web_extensions.web_extension_bound_content_control_count
        == 3
    )
    assert "taskpane_web_extension_inventory_changed" in {
        change.kind for change in diff_documents(created_false, created_true).changes
    }


def test_sensitivity_label_metadata_inventory_is_private_and_semantic(tmp_path) -> None:
    before = tmp_path / "before.docx"
    after = tmp_path / "after.docx"
    label_info_changed = tmp_path / "label-info-changed.docx"
    extension_payload_changed = tmp_path / "extension-payload-changed.docx"
    legacy_property_changed = tmp_path / "legacy-property-changed.docx"
    renumbered = tmp_path / "renumbered.docx"
    policy_path = tmp_path / "docfence.yml"
    _write_sensitivity_label_document(
        before,
        include_label_info=False,
        include_legacy_properties=False,
    )
    _write_sensitivity_label_document(after)
    _write_sensitivity_label_document(
        label_info_changed,
        label_name="SENSITIVITY_LABEL_NAME_CHANGED_DO_NOT_LEAK",
    )
    _write_sensitivity_label_document(
        extension_payload_changed,
        label_extension_marker="SENSITIVITY_LABEL_EXTENSION_CHANGED_DO_NOT_LEAK",
    )
    _write_sensitivity_label_document(
        legacy_property_changed,
        legacy_label_name="SENSITIVITY_LEGACY_NAME_CHANGED_DO_NOT_LEAK",
    )
    _write_sensitivity_label_document(renumbered, relationship_id_suffix="9")

    expected_inventory = {
        "label_info_part_count": 1,
        "label_info_label_count": 2,
        "label_info_enabled_label_count": 1,
        "label_info_removed_label_count": 1,
        "label_info_extension_count": 1,
        "legacy_mip_label_count": 1,
        "legacy_mip_property_count": 3,
        "legacy_sensitivity_property_count": 1,
        "word_content_marking_property_count": 4,
    }
    before_snapshot = load_snapshot(before)
    after_snapshot = load_snapshot(after)
    assert after_snapshot.public_dict()["sensitivity_labels"] == expected_inventory
    assert before_snapshot.public_dict()["sensitivity_labels"] == {
        key: 0 for key in expected_inventory
    }
    assert (
        after_snapshot.unclassified_part_count
        == before_snapshot.unclassified_part_count
    )

    report = diff_documents(before, after)
    assert "sensitivity_label_inventory_changed" in {
        change.kind for change in report.changes
    }
    assert {
        change.kind for change in diff_documents(after, label_info_changed).changes
    } == {"sensitivity_label_inventory_changed"}
    assert {
        change.kind
        for change in diff_documents(after, extension_payload_changed).changes
    } == {"sensitivity_label_inventory_changed"}
    assert {
        "document_property_inventory_changed",
        "sensitivity_label_inventory_changed",
    } <= {
        change.kind
        for change in diff_documents(after, legacy_property_changed).changes
    }
    assert diff_documents(after, renumbered).changes == ()

    policy_path.write_text(
        """version: 1
rules:
  require_no_sensitivity_label_metadata: true
  no_sensitivity_label_metadata_changes: true
""",
        encoding="utf-8",
    )
    policy = load_policy(policy_path)
    assert {finding.rule_id for finding in apply_policy(report, policy).findings} == {
        "DFP035",
        "DFP036",
    }
    assert {
        finding.rule_id
        for finding in apply_policy(
            diff_documents(after, label_info_changed), policy
        ).findings
    } == {"DFP035", "DFP036"}
    assert {
        finding.rule_id
        for finding in apply_policy(diff_documents(after, renumbered), policy).findings
    } == {"DFP035"}

    gated = apply_policy(report, policy)
    rendered = "\n".join(
        (
            render_profile(after_snapshot, "json"),
            render_profile(after_snapshot, "markdown"),
            render_report(gated, "json"),
            render_report(gated, "markdown"),
            render_report(gated, "sarif"),
        )
    )
    sarif = json.loads(render_report(gated, "sarif"))
    assert {
        "DFC_SENSITIVITY_LABEL_INVENTORY_CHANGED",
        "DFP035",
        "DFP036",
    } <= {result["ruleId"] for result in sarif["runs"][0]["results"]}
    for marker in (
        "SENSITIVITY_LABEL_NAME_DO_NOT_LEAK",
        "SENSITIVITY_LABEL_NAME_CHANGED_DO_NOT_LEAK",
        "SENSITIVITY_LABEL_EXTENSION_DO_NOT_LEAK",
        "SENSITIVITY_LABEL_EXTENSION_CHANGED_DO_NOT_LEAK",
        "SENSITIVITY_LEGACY_NAME_DO_NOT_LEAK",
        "SENSITIVITY_LEGACY_NAME_CHANGED_DO_NOT_LEAK",
        "SENSITIVITY_MIP_CUSTOM_DO_NOT_LEAK",
        "SENSITIVITY_MARKING_TEXT_DO_NOT_LEAK",
    ):
        assert marker not in rendered


def test_sensitivity_label_metadata_discovers_parts_and_rejects_invalid_state(
    tmp_path,
) -> None:
    relationship_only = tmp_path / "relationship-only.docx"
    content_type_only = tmp_path / "content-type-only.docx"
    conventional_unlinked = tmp_path / "conventional-unlinked.docx"
    noncanonical_custom_properties = tmp_path / "noncanonical-custom-properties.docx"
    wrong_root = tmp_path / "wrong-root.docx"
    missing_attribute = tmp_path / "missing-attribute.docx"
    invalid_site_id = tmp_path / "invalid-site-id.docx"
    external_relationship = tmp_path / "external-relationship.docx"
    non_root_relationship = tmp_path / "non-root-relationship.docx"
    missing_target = tmp_path / "missing-target.docx"
    multiple_parts = tmp_path / "multiple-parts.docx"
    _write_sensitivity_label_document(
        relationship_only,
        label_info_part_name="private/label-data.xml",
        include_label_info_content_type=False,
        include_legacy_properties=False,
    )
    _write_sensitivity_label_document(
        content_type_only,
        label_info_part_name="private/label-data.xml",
        include_label_info_relationship=False,
        include_legacy_properties=False,
    )
    _write_sensitivity_label_document(
        conventional_unlinked,
        label_info_part_name="docMetadata/LabelInfo",
        include_label_info_relationship=False,
        include_label_info_content_type=False,
        include_legacy_properties=False,
    )
    _write_sensitivity_label_document(
        noncanonical_custom_properties,
        include_label_info=False,
        custom_property_part_name="private/label-properties.xml",
        include_legacy_properties=True,
    )
    _write_sensitivity_label_document(wrong_root, wrong_label_info_root=True)
    _write_sensitivity_label_document(missing_attribute, omit_label_method=True)
    _write_sensitivity_label_document(invalid_site_id, label_site_id="not-a-guid")
    _write_sensitivity_label_document(
        external_relationship,
        label_info_target_mode="External",
    )
    _write_sensitivity_label_document(
        non_root_relationship,
        label_info_relationship_source="word/document.xml",
    )
    _write_sensitivity_label_document(
        missing_target,
        include_label_info_part=False,
        include_label_info_content_type=False,
    )
    _write_sensitivity_label_document(
        multiple_parts,
        additional_label_info_part_name="docMetadata/OtherLabelInfo.xml",
    )

    for document in (
        relationship_only,
        content_type_only,
        conventional_unlinked,
    ):
        snapshot = load_snapshot(document)
        assert snapshot.sensitivity_labels.label_info_part_count == 1
        assert snapshot.sensitivity_labels.label_info_label_count == 2
    custom_snapshot = load_snapshot(noncanonical_custom_properties)
    assert custom_snapshot.sensitivity_labels.label_info_part_count == 0
    assert custom_snapshot.sensitivity_labels.legacy_mip_property_count == 3
    assert custom_snapshot.sensitivity_labels.word_content_marking_property_count == 4

    for document in (
        wrong_root,
        missing_attribute,
        invalid_site_id,
        external_relationship,
        non_root_relationship,
        missing_target,
        multiple_parts,
    ):
        with pytest.raises(DocumentFormatError):
            load_snapshot(document)


def test_package_digital_signature_inventory_is_private_and_semantic(
    tmp_path,
) -> None:
    before = tmp_path / "before.docx"
    after = tmp_path / "after.docx"
    signature_changed = tmp_path / "signature-changed.docx"
    certificate_changed = tmp_path / "certificate-changed.docx"
    policy_path = tmp_path / "docfence.yml"
    _write_package_digital_signature_document(
        before,
        include_origin=False,
        include_xml_signature=False,
    )
    _write_package_digital_signature_document(after, include_certificate_part=True)
    _write_package_digital_signature_document(
        signature_changed,
        include_certificate_part=True,
        signature_comment="PACKAGE_SIGNATURE_COMMENT_CHANGED_DO_NOT_LEAK",
    )
    _write_package_digital_signature_document(
        certificate_changed,
        include_certificate_part=True,
        certificate_payload_marker="PACKAGE_CERTIFICATE_CHANGED_DO_NOT_LEAK",
    )

    expected_inventory = {
        "signature_origin_part_count": 1,
        "xml_signature_part_count": 1,
        "certificate_part_count": 1,
        "signed_info_reference_count": 1,
        "manifest_reference_count": 1,
        "relationship_reference_count": 1,
        "inline_x509_certificate_count": 1,
        "signature_property_count": 1,
    }
    before_snapshot = load_snapshot(before)
    after_snapshot = load_snapshot(after)
    assert (
        after_snapshot.public_dict()["package_digital_signatures"]
        == expected_inventory
    )
    assert before_snapshot.public_dict()["package_digital_signatures"] == {
        key: 0 for key in expected_inventory
    }
    assert (
        after_snapshot.unclassified_part_count
        == before_snapshot.unclassified_part_count
    )

    report = diff_documents(before, after)
    assert "package_digital_signature_inventory_changed" in {
        change.kind for change in report.changes
    }
    assert {
        change.kind for change in diff_documents(after, signature_changed).changes
    } == {"package_digital_signature_inventory_changed"}
    assert {
        change.kind for change in diff_documents(after, certificate_changed).changes
    } == {"package_digital_signature_inventory_changed"}

    policy_path.write_text(
        """version: 1
rules:
  require_no_package_digital_signatures: true
  no_package_digital_signature_changes: true
""",
        encoding="utf-8",
    )
    policy = load_policy(policy_path)
    assert {finding.rule_id for finding in apply_policy(report, policy).findings} == {
        "DFP037",
        "DFP038",
    }
    assert {
        finding.rule_id
        for finding in apply_policy(
            diff_documents(after, signature_changed), policy
        ).findings
    } == {"DFP037", "DFP038"}

    gated = apply_policy(report, policy)
    rendered = "\n".join(
        (
            render_profile(after_snapshot, "json"),
            render_profile(after_snapshot, "markdown"),
            render_report(gated, "json"),
            render_report(gated, "markdown"),
            render_report(gated, "sarif"),
        )
    )
    sarif = json.loads(render_report(gated, "sarif"))
    assert {
        "DFC_PACKAGE_DIGITAL_SIGNATURE_INVENTORY_CHANGED",
        "DFP037",
        "DFP038",
    } <= {result["ruleId"] for result in sarif["runs"][0]["results"]}
    for marker in (
        "PACKAGE_SIGNATURE_VALUE_DO_NOT_LEAK",
        "PACKAGE_INLINE_X509_DO_NOT_LEAK",
        "PACKAGE_SIGNATURE_COMMENT_DO_NOT_LEAK",
        "PACKAGE_SIGNATURE_COMMENT_CHANGED_DO_NOT_LEAK",
        "PACKAGE_DIGEST_DO_NOT_LEAK",
        "PACKAGE_MANIFEST_DIGEST_DO_NOT_LEAK",
        "PACKAGE_CERTIFICATE_DO_NOT_LEAK",
        "PACKAGE_CERTIFICATE_CHANGED_DO_NOT_LEAK",
    ):
        assert marker not in rendered


def test_package_digital_signature_discovery_and_invalid_topology(tmp_path) -> None:
    noncanonical_relationship = tmp_path / "noncanonical-relationship.docx"
    content_type_only = tmp_path / "content-type-only.docx"
    conventional_origin_only = tmp_path / "conventional-origin-only.docx"
    wrong_signature_root = tmp_path / "wrong-signature-root.docx"
    missing_signature_method = tmp_path / "missing-signature-method.docx"
    external_origin_relationship = tmp_path / "external-origin-relationship.docx"
    non_root_origin_relationship = tmp_path / "non-root-origin-relationship.docx"
    missing_origin_target = tmp_path / "missing-origin-target.docx"
    duplicate_origin_relationship = tmp_path / "duplicate-origin-relationship.docx"
    external_signature_relationship = tmp_path / "external-signature-relationship.docx"
    missing_signature_target = tmp_path / "missing-signature-target.docx"
    external_certificate_relationship = (
        tmp_path / "external-certificate-relationship.docx"
    )
    wrong_origin_content_type = tmp_path / "wrong-origin-content-type.docx"
    wrong_signature_content_type = tmp_path / "wrong-signature-content-type.docx"
    wrong_certificate_content_type = tmp_path / "wrong-certificate-content-type.docx"

    _write_package_digital_signature_document(
        noncanonical_relationship,
        origin_part_name="private/origin.sigs",
        signature_part_name="private/signature.xml",
    )
    _write_package_digital_signature_document(
        content_type_only,
        include_root_relationship=False,
        include_signature_relationship=False,
    )
    _write_package_digital_signature_document(
        conventional_origin_only,
        include_xml_signature=False,
        include_root_relationship=False,
        include_origin_content_type=False,
    )
    _write_package_digital_signature_document(
        wrong_signature_root,
        wrong_signature_root=True,
    )
    _write_package_digital_signature_document(
        missing_signature_method,
        omit_signature_method=True,
    )
    _write_package_digital_signature_document(
        external_origin_relationship,
        origin_target_mode="External",
    )
    _write_package_digital_signature_document(
        non_root_origin_relationship,
        origin_relationship_source="word/document.xml",
    )
    _write_package_digital_signature_document(
        missing_origin_target,
        include_origin_part=False,
        include_origin_content_type=False,
    )
    _write_package_digital_signature_document(
        duplicate_origin_relationship,
        duplicate_origin_relationship=True,
    )
    _write_package_digital_signature_document(
        external_signature_relationship,
        signature_target_mode="External",
    )
    _write_package_digital_signature_document(
        missing_signature_target,
        include_xml_signature_part=False,
        include_xml_signature_content_type=False,
    )
    _write_package_digital_signature_document(
        external_certificate_relationship,
        include_certificate_part=True,
        certificate_target_mode="External",
    )
    _write_package_digital_signature_document(
        wrong_origin_content_type,
        origin_content_type="application/octet-stream",
    )
    _write_package_digital_signature_document(
        wrong_signature_content_type,
        xml_signature_content_type="application/xml",
    )
    _write_package_digital_signature_document(
        wrong_certificate_content_type,
        include_certificate_part=True,
        certificate_content_type="application/octet-stream",
    )

    for document in (noncanonical_relationship, content_type_only):
        snapshot = load_snapshot(document)
        assert snapshot.package_digital_signatures.signature_origin_part_count == 1
        assert snapshot.package_digital_signatures.xml_signature_part_count == 1
    origin_only_snapshot = load_snapshot(conventional_origin_only)
    assert (
        origin_only_snapshot.package_digital_signatures.signature_origin_part_count
        == 1
    )
    assert origin_only_snapshot.package_digital_signatures.xml_signature_part_count == 0

    for document in (
        wrong_signature_root,
        missing_signature_method,
        external_origin_relationship,
        non_root_origin_relationship,
        missing_origin_target,
        duplicate_origin_relationship,
        external_signature_relationship,
        missing_signature_target,
        external_certificate_relationship,
        wrong_origin_content_type,
        wrong_signature_content_type,
        wrong_certificate_content_type,
    ):
        with pytest.raises(DocumentFormatError):
            load_snapshot(document)


def test_word_protection_inventory_is_private_and_semantic(tmp_path) -> None:
    before = tmp_path / "before.docx"
    after = tmp_path / "after.docx"
    verifier_changed = tmp_path / "verifier-changed.docx"
    policy_path = tmp_path / "docfence.yml"
    _write_word_protection_document(
        before,
        include_document_protection=False,
        include_write_protection=False,
    )
    _write_word_protection_document(
        after,
        document_provider="WORD_PROTECTION_PROVIDER_DO_NOT_LEAK",
        document_algorithm_name="WORD_PROTECTION_ALGORITHM_DO_NOT_LEAK",
    )
    _write_word_protection_document(
        verifier_changed,
        document_hash="WORD_PROTECTION_HASH_CHANGED_DO_NOT_LEAK",
        write_hash="WRITE_PROTECTION_HASH_CHANGED_DO_NOT_LEAK",
    )

    expected_inventory = {
        "document_protection_count": 1,
        "document_protection_enforcement_enabled_count": 1,
        "document_protection_formatting_restricted_count": 1,
        "document_protection_read_only_count": 1,
        "document_protection_comments_count": 0,
        "document_protection_tracked_changes_count": 0,
        "document_protection_forms_count": 0,
        "document_protection_password_material_count": 1,
        "write_protection_count": 1,
        "write_protection_recommended_count": 1,
        "write_protection_password_material_count": 1,
    }
    before_snapshot = load_snapshot(before)
    after_snapshot = load_snapshot(after)
    assert after_snapshot.public_dict()["word_protection"] == expected_inventory
    assert before_snapshot.public_dict()["word_protection"] == {
        key: 0 for key in expected_inventory
    }
    assert (
        after_snapshot.unclassified_part_count
        == before_snapshot.unclassified_part_count
    )

    report = diff_documents(before, after)
    assert "word_protection_inventory_changed" in {
        change.kind for change in report.changes
    }
    assert {
        change.kind for change in diff_documents(after, verifier_changed).changes
    } == {"document_settings_changed", "word_protection_inventory_changed"}

    policy_path.write_text(
        """version: 1
rules:
  require_no_word_protection: true
  no_word_protection_changes: true
""",
        encoding="utf-8",
    )
    policy = load_policy(policy_path)
    assert {finding.rule_id for finding in apply_policy(report, policy).findings} == {
        "DFP039",
        "DFP040",
    }
    assert {
        finding.rule_id
        for finding in apply_policy(
            diff_documents(after, verifier_changed), policy
        ).findings
    } == {"DFP039", "DFP040"}

    gated = apply_policy(report, policy)
    rendered = "\n".join(
        (
            render_profile(after_snapshot, "json"),
            render_profile(after_snapshot, "markdown"),
            render_report(gated, "json"),
            render_report(gated, "markdown"),
            render_report(gated, "sarif"),
        )
    )
    sarif = json.loads(render_report(gated, "sarif"))
    assert {
        "DFC_WORD_PROTECTION_INVENTORY_CHANGED",
        "DFP039",
        "DFP040",
    } <= {result["ruleId"] for result in sarif["runs"][0]["results"]}
    for marker in (
        "WORD_PROTECTION_HASH_DO_NOT_LEAK",
        "WORD_PROTECTION_SALT_DO_NOT_LEAK",
        "WORD_PROTECTION_HASH_VALUE_DO_NOT_LEAK",
        "WORD_PROTECTION_SALT_VALUE_DO_NOT_LEAK",
        "WORD_PROTECTION_PROVIDER_DO_NOT_LEAK",
        "WORD_PROTECTION_ALGORITHM_DO_NOT_LEAK",
        "WRITE_PROTECTION_HASH_DO_NOT_LEAK",
        "WRITE_PROTECTION_SALT_DO_NOT_LEAK",
        "WORD_PROTECTION_HASH_CHANGED_DO_NOT_LEAK",
        "WRITE_PROTECTION_HASH_CHANGED_DO_NOT_LEAK",
    ):
        assert marker not in rendered


def test_word_protection_discovery_and_invalid_markup(tmp_path) -> None:
    noncanonical = tmp_path / "noncanonical.docx"
    strict_relationship = tmp_path / "strict-relationship.docx"
    glossary_linked = tmp_path / "glossary-linked.docx"
    duplicate_document_protection = tmp_path / "duplicate-document-protection.docx"
    duplicate_write_protection = tmp_path / "duplicate-write-protection.docx"
    invalid_edit = tmp_path / "invalid-edit.docx"
    invalid_boolean = tmp_path / "invalid-boolean.docx"
    unsupported_attribute = tmp_path / "unsupported-attribute.docx"
    nonleaf = tmp_path / "nonleaf.docx"
    invalid_settings_root = tmp_path / "invalid-settings-root.docx"
    external_settings = tmp_path / "external-settings.docx"
    _write_word_protection_document(
        noncanonical,
        settings_part_name="word/private/settings.xml",
        include_settings_relationship=True,
    )
    _write_word_protection_document(
        strict_relationship,
        settings_part_name="word/private/settings.xml",
        include_settings_relationship=True,
        settings_relationship_type=_STRICT_DOCUMENT_SETTINGS_RELATIONSHIP_TYPE,
        settings_word_namespace=_STRICT_WORD_NAMESPACE,
    )
    _write_word_protection_document(
        glossary_linked,
        settings_part_name="word/glossary/private/settings.xml",
        include_settings_relationship=True,
        settings_relationship_source="word/glossary/document.xml",
    )
    _write_word_protection_document(
        duplicate_document_protection,
        duplicate_document_protection=True,
    )
    _write_word_protection_document(
        duplicate_write_protection,
        duplicate_write_protection=True,
    )
    _write_word_protection_document(invalid_edit, document_edit="unrestricted")
    _write_word_protection_document(
        invalid_boolean,
        document_enforcement="probably",
    )
    _write_word_protection_document(
        unsupported_attribute,
        document_extra_attributes='w:futureProtection="1"',
    )
    _write_word_protection_document(
        nonleaf,
        document_extra_content="<w:unexpected/>",
    )
    _write_word_protection_document(
        invalid_settings_root,
        settings_root_name="notSettings",
    )
    _write_word_protection_document(
        external_settings,
        settings_part_name="word/private/settings.xml",
        include_settings_relationship=True,
        settings_target_mode="External",
        include_settings_part=False,
    )

    for document in (noncanonical, strict_relationship, glossary_linked):
        snapshot = load_snapshot(document)
        assert snapshot.word_protection.document_protection_count == 1
        assert snapshot.word_protection.write_protection_count == 1
        assert snapshot.unclassified_part_count == 1

    for document in (
        duplicate_document_protection,
        duplicate_write_protection,
        invalid_edit,
        invalid_boolean,
        unsupported_attribute,
        nonleaf,
        invalid_settings_root,
        external_settings,
    ):
        with pytest.raises(DocumentFormatError):
            load_snapshot(document)


def test_word_document_variable_inventory_is_private_and_semantic(tmp_path) -> None:
    before = tmp_path / "before.docx"
    after = tmp_path / "after.docx"
    value_changed = tmp_path / "value-changed.docx"
    policy_path = tmp_path / "docfence.yml"
    _write_word_protection_document(
        before,
        include_document_protection=False,
        include_write_protection=False,
    )
    _write_word_protection_document(
        after,
        include_document_protection=False,
        include_write_protection=False,
        settings_extra_children=(
            "<w:docVars>"
            '<w:docVar w:name="VARIABLE_NAME_DO_NOT_LEAK" '
            'w:val="VARIABLE_VALUE_DO_NOT_LEAK"/>'
            '<w:docVar w:name="SECOND_VARIABLE_NAME_DO_NOT_LEAK" '
            'w:val="SECOND_VARIABLE_VALUE_DO_NOT_LEAK"/>'
            '<w:docVar w:name="EMPTY_VARIABLE_NAME_DO_NOT_LEAK" w:val=""/>'
            "</w:docVars>"
        ),
    )
    _write_word_protection_document(
        value_changed,
        include_document_protection=False,
        include_write_protection=False,
        settings_extra_children=(
            "<w:docVars>"
            '<w:docVar w:name="VARIABLE_NAME_DO_NOT_LEAK" '
            'w:val="VARIABLE_VALUE_CHANGED_DO_NOT_LEAK"/>'
            '<w:docVar w:name="SECOND_VARIABLE_NAME_DO_NOT_LEAK" '
            'w:val="SECOND_VARIABLE_VALUE_DO_NOT_LEAK"/>'
            '<w:docVar w:name="EMPTY_VARIABLE_NAME_DO_NOT_LEAK" w:val=""/>'
            "</w:docVars>"
        ),
    )

    expected_inventory = {
        "document_variable_container_count": 1,
        "document_variable_count": 3,
        "empty_document_variable_value_count": 1,
    }
    before_snapshot = load_snapshot(before)
    after_snapshot = load_snapshot(after)
    assert (
        after_snapshot.public_dict()["word_document_variables"] == expected_inventory
    )
    assert before_snapshot.public_dict()["word_document_variables"] == {
        key: 0 for key in expected_inventory
    }
    assert (
        after_snapshot.unclassified_part_count
        == before_snapshot.unclassified_part_count
    )

    report = diff_documents(before, after)
    assert "word_document_variable_inventory_changed" in {
        change.kind for change in report.changes
    }
    assert {
        change.kind for change in diff_documents(after, value_changed).changes
    } == {"document_settings_changed", "word_document_variable_inventory_changed"}

    policy_path.write_text(
        """version: 1
rules:
  require_no_word_document_variables: true
  no_word_document_variable_changes: true
""",
        encoding="utf-8",
    )
    policy = load_policy(policy_path)
    assert {finding.rule_id for finding in apply_policy(report, policy).findings} == {
        "DFP043",
        "DFP044",
    }
    assert {
        finding.rule_id
        for finding in apply_policy(
            diff_documents(after, value_changed), policy
        ).findings
    } == {"DFP043", "DFP044"}

    gated = apply_policy(report, policy)
    rendered = "\n".join(
        (
            render_profile(after_snapshot, "json"),
            render_profile(after_snapshot, "markdown"),
            render_report(gated, "json"),
            render_report(gated, "markdown"),
            render_report(gated, "sarif"),
        )
    )
    sarif = json.loads(render_report(gated, "sarif"))
    assert {
        "DFC_WORD_DOCUMENT_VARIABLE_INVENTORY_CHANGED",
        "DFP043",
        "DFP044",
    } <= {result["ruleId"] for result in sarif["runs"][0]["results"]}
    for marker in (
        "VARIABLE_NAME_DO_NOT_LEAK",
        "VARIABLE_VALUE_DO_NOT_LEAK",
        "VARIABLE_VALUE_CHANGED_DO_NOT_LEAK",
        "SECOND_VARIABLE_NAME_DO_NOT_LEAK",
        "SECOND_VARIABLE_VALUE_DO_NOT_LEAK",
        "EMPTY_VARIABLE_NAME_DO_NOT_LEAK",
    ):
        assert marker not in rendered


def test_word_document_variable_field_inventory_is_private_and_semantic(
    tmp_path,
) -> None:
    before = tmp_path / "before.docx"
    after = tmp_path / "after.docx"
    field_changed = tmp_path / "field-changed.docx"
    policy_path = tmp_path / "docfence.yml"

    def simple_field(instruction: str) -> str:
        escaped_instruction = (
            instruction.replace("&", "&amp;")
            .replace('"', "&quot;")
            .replace("<", "&lt;")
        )
        return (
            f'<w:fldSimple w:instr="{escaped_instruction}">'
            "<w:r><w:t>FIELD_RESULT_DO_NOT_LEAK</w:t></w:r></w:fldSimple>"
        )

    def instruction_text(text: str) -> str:
        return (
            '<w:r><w:instrText xml:space="preserve">'
            f"{text}</w:instrText></w:r>"
        )

    def field_char(field_type: str) -> str:
        return f'<w:r><w:fldChar w:fldCharType="{field_type}"/></w:r>'

    simple_instruction = "DOCVARIABLE SIMPLE_VARIABLE_NAME_DO_NOT_LEAK"
    body_markup = "".join(
        (
            instruction_text("DOCVARIABLE LOOSE_FIELD_DO_NOT_LEAK"),
            simple_field(simple_instruction),
            field_char("begin"),
            instruction_text(' DOCVARIABLE "SPACED '),
            instruction_text('VARIABLE NAME DO NOT LEAK" '),
            field_char("end"),
            simple_field("DOCVARIABLE UNMATCHED_VARIABLE_NAME_DO_NOT_LEAK"),
            simple_field(
                "DOCVARIABLE COMPOUND_VARIABLE_NAME_DO_NOT_LEAK "
                "UNPARSED_ARGUMENT_DO_NOT_LEAK"
            ),
            field_char("begin"),
            instruction_text(" DOCVARIABLE NESTED_"),
            field_char("begin"),
            instruction_text(" MERGEFIELD INNER_DO_NOT_LEAK "),
            field_char("end"),
            instruction_text("VARIABLE_NAME_DO_NOT_LEAK "),
            field_char("end"),
            field_char("begin"),
            instruction_text("DOCVARIABLE UNCLOSED_FIELD_DO_NOT_LEAK"),
        )
    )
    header_markup = simple_field(
        "DOCVARIABLE HEADER_VARIABLE_NAME_DO_NOT_LEAK \\* MERGEFORMAT"
    )
    settings_markup = "".join(
        (
            "<w:docVars>",
            '<w:docVar w:name="SIMPLE_VARIABLE_NAME_DO_NOT_LEAK" '
            'w:val="SIMPLE_VARIABLE_VALUE_DO_NOT_LEAK"/>',
            '<w:docVar w:name="SPACED VARIABLE NAME DO NOT LEAK" '
            'w:val="SPACED_VARIABLE_VALUE_DO_NOT_LEAK"/>',
            '<w:docVar w:name="HEADER_VARIABLE_NAME_DO_NOT_LEAK" '
            'w:val="HEADER_VARIABLE_VALUE_DO_NOT_LEAK"/>',
            '<w:docVar w:name="NESTED_VARIABLE_NAME_DO_NOT_LEAK" '
            'w:val="NESTED_VARIABLE_VALUE_DO_NOT_LEAK"/>',
            "</w:docVars>",
        )
    )
    _write_document_variable_field_document(before)
    _write_document_variable_field_document(
        after,
        body_markup=body_markup,
        header_markup=header_markup,
        settings_markup=settings_markup,
        word_namespace=_STRICT_WORD_NAMESPACE,
    )
    _write_document_variable_field_document(
        field_changed,
        body_markup=body_markup.replace(
            simple_instruction,
            "DOCVARIABLE CHANGED_VARIABLE_NAME_DO_NOT_LEAK",
        ),
        header_markup=header_markup,
        settings_markup=settings_markup,
        word_namespace=_STRICT_WORD_NAMESPACE,
    )

    expected_inventory = {
        "document_variable_field_reference_count": 6,
        "document_variable_field_story_count": 2,
        "literal_document_variable_field_reference_count": 4,
        "nonliteral_document_variable_field_reference_count": 2,
        (
            "literal_document_variable_field_reference_"
            "matching_stored_variable_count"
        ): 3,
        (
            "literal_document_variable_field_reference_"
            "not_matching_stored_variable_count"
        ): 1,
    }
    before_snapshot = load_snapshot(before)
    after_snapshot = load_snapshot(after)
    assert (
        after_snapshot.public_dict()["word_document_variable_fields"]
        == expected_inventory
    )
    assert before_snapshot.public_dict()["word_document_variable_fields"] == {
        key: 0 for key in expected_inventory
    }
    assert (
        after_snapshot.unclassified_part_count
        == before_snapshot.unclassified_part_count
    )

    report = diff_documents(before, after)
    assert "word_document_variable_field_inventory_changed" in {
        change.kind for change in report.changes
    }
    assert "word_document_variable_field_inventory_changed" in {
        change.kind for change in diff_documents(after, field_changed).changes
    }

    policy_path.write_text(
        """version: 1
rules:
  require_no_word_document_variables: true
  no_word_document_variable_changes: true
  require_no_word_document_variable_fields: true
  no_word_document_variable_field_changes: true
""",
        encoding="utf-8",
    )
    policy = load_policy(policy_path)
    assert {finding.rule_id for finding in apply_policy(report, policy).findings} == {
        "DFP043",
        "DFP044",
        "DFP045",
        "DFP046",
    }
    assert {
        finding.rule_id
        for finding in apply_policy(
            diff_documents(after, field_changed), policy
        ).findings
    } == {"DFP043", "DFP045", "DFP046"}

    gated = apply_policy(report, policy)
    rendered = "\n".join(
        (
            render_profile(after_snapshot, "json"),
            render_profile(after_snapshot, "markdown"),
            render_report(gated, "json"),
            render_report(gated, "markdown"),
            render_report(gated, "sarif"),
        )
    )
    sarif = json.loads(render_report(gated, "sarif"))
    assert {
        "DFC_WORD_DOCUMENT_VARIABLE_FIELD_INVENTORY_CHANGED",
        "DFP043",
        "DFP044",
        "DFP045",
        "DFP046",
    } <= {result["ruleId"] for result in sarif["runs"][0]["results"]}
    for marker in (
        "SIMPLE_VARIABLE_NAME_DO_NOT_LEAK",
        "SIMPLE_VARIABLE_VALUE_DO_NOT_LEAK",
        "SPACED VARIABLE NAME DO NOT LEAK",
        "SPACED_VARIABLE_VALUE_DO_NOT_LEAK",
        "HEADER_VARIABLE_NAME_DO_NOT_LEAK",
        "HEADER_VARIABLE_VALUE_DO_NOT_LEAK",
        "NESTED_VARIABLE_NAME_DO_NOT_LEAK",
        "NESTED_VARIABLE_VALUE_DO_NOT_LEAK",
        "UNMATCHED_VARIABLE_NAME_DO_NOT_LEAK",
        "COMPOUND_VARIABLE_NAME_DO_NOT_LEAK",
        "UNPARSED_ARGUMENT_DO_NOT_LEAK",
        "LOOSE_FIELD_DO_NOT_LEAK",
        "INNER_DO_NOT_LEAK",
        "UNCLOSED_FIELD_DO_NOT_LEAK",
        "CHANGED_VARIABLE_NAME_DO_NOT_LEAK",
    ):
        assert marker not in rendered


def test_word_document_variable_fields_keep_revision_variants_and_document_scope(
    tmp_path,
) -> None:
    document = tmp_path / "scoped-fields.docx"

    def simple_field(instruction: str) -> str:
        return (
            f'<w:fldSimple w:instr="{instruction}">'
            "<w:r><w:t>FIELD_RESULT_DO_NOT_LEAK</w:t></w:r></w:fldSimple>"
        )

    def instruction_text(text: str) -> str:
        return f"<w:r><w:instrText>{text}</w:instrText></w:r>"

    def field_char(field_type: str) -> str:
        return f'<w:r><w:fldChar w:fldCharType="{field_type}"/></w:r>'

    def deleted_instruction_text(text: str) -> str:
        return (
            '<w:del w:id="1"><w:r><w:delInstrText>'
            f"{text}</w:delInstrText></w:r></w:del>"
        )

    def inserted_instruction_text(text: str) -> str:
        return (
            '<w:ins w:id="2"><w:r><w:instrText>'
            f"{text}</w:instrText></w:r></w:ins>"
        )

    body_markup = "".join(
        (
            simple_field("DOCVARIABLE MAIN_VARIABLE_NAME_DO_NOT_LEAK"),
            field_char("begin"),
            instruction_text(' DOCVARIABLE "'),
            deleted_instruction_text("DELETED_VARIABLE_NAME_DO_NOT_LEAK"),
            inserted_instruction_text("CURRENT_VARIABLE_NAME_DO_NOT_LEAK"),
            instruction_text('" '),
            field_char("end"),
            instruction_text("DOCVARIABLE LOOSE_SCOPE_FIELD_DO_NOT_LEAK"),
            field_char("begin"),
            instruction_text("DOCVARIABLE UNCLOSED_SCOPE_FIELD_DO_NOT_LEAK"),
        )
    )
    glossary_markup = "".join(
        (
            simple_field("DOCVARIABLE MAIN_VARIABLE_NAME_DO_NOT_LEAK"),
            simple_field("DOCVARIABLE GLOSSARY_VARIABLE_NAME_DO_NOT_LEAK"),
        )
    )
    _write_document_variable_field_document(
        document,
        body_markup=body_markup,
        settings_markup=(
            "<w:docVars>"
            '<w:docVar w:name="MAIN_VARIABLE_NAME_DO_NOT_LEAK" w:val="main"/>'
            '<w:docVar w:name="DELETED_VARIABLE_NAME_DO_NOT_LEAK" w:val="old"/>'
            '<w:docVar w:name="CURRENT_VARIABLE_NAME_DO_NOT_LEAK" w:val="new"/>'
            "</w:docVars>"
        ),
        glossary_markup=glossary_markup,
        glossary_settings_markup=(
            "<w:docVars>"
            '<w:docVar w:name="GLOSSARY_VARIABLE_NAME_DO_NOT_LEAK" '
            'w:val="glossary"/>'
            "</w:docVars>"
        ),
    )

    snapshot = load_snapshot(document)
    assert snapshot.public_dict()["word_document_variable_fields"] == {
        "document_variable_field_reference_count": 5,
        "document_variable_field_story_count": 2,
        "literal_document_variable_field_reference_count": 5,
        "nonliteral_document_variable_field_reference_count": 0,
        (
            "literal_document_variable_field_reference_"
            "matching_stored_variable_count"
        ): 4,
        (
            "literal_document_variable_field_reference_"
            "not_matching_stored_variable_count"
        ): 1,
    }
    rendered = render_profile(snapshot, "json") + render_profile(snapshot, "markdown")
    for marker in (
        "MAIN_VARIABLE_NAME_DO_NOT_LEAK",
        "DELETED_VARIABLE_NAME_DO_NOT_LEAK",
        "CURRENT_VARIABLE_NAME_DO_NOT_LEAK",
        "GLOSSARY_VARIABLE_NAME_DO_NOT_LEAK",
        "LOOSE_SCOPE_FIELD_DO_NOT_LEAK",
        "UNCLOSED_SCOPE_FIELD_DO_NOT_LEAK",
    ):
        assert marker not in rendered


def test_word_hyperlink_field_inventory_is_private_and_semantic(tmp_path) -> None:
    before = tmp_path / "before.docx"
    after = tmp_path / "after.docx"
    changed = tmp_path / "changed.docx"
    split_complex = tmp_path / "split-complex.docx"
    strict = tmp_path / "strict.docx"
    policy_path = tmp_path / "docfence.yml"
    literal_destination = "https://DESTINATION_DO_NOT_LEAK.invalid/path"
    changed_destination = "https://CHANGED_DESTINATION_DO_NOT_LEAK.invalid/path"

    _write_hyperlink_field_document(before, include_fields=False)
    _write_hyperlink_field_document(
        after,
        literal_destination=literal_destination,
    )
    _write_hyperlink_field_document(
        changed,
        literal_destination=changed_destination,
    )
    _write_hyperlink_field_document(
        split_complex,
        literal_destination=literal_destination,
        split_complex_instructions=True,
    )
    _write_hyperlink_field_document(
        strict,
        literal_destination=literal_destination,
        word_namespace=_STRICT_WORD_NAMESPACE,
    )

    expected_inventory = {
        "hyperlink_field_reference_count": 7,
        "hyperlink_field_story_count": 2,
        "literal_destination_hyperlink_field_count": 3,
        "literal_internal_location_only_hyperlink_field_count": 2,
        "dynamic_or_unparseable_hyperlink_field_count": 2,
    }
    before_snapshot = load_snapshot(before)
    after_snapshot = load_snapshot(after)
    assert after_snapshot.public_dict()["word_hyperlink_fields"] == expected_inventory
    assert before_snapshot.public_dict()["word_hyperlink_fields"] == {
        key: 0 for key in expected_inventory
    }
    assert load_snapshot(strict).public_dict()["word_hyperlink_fields"] == (
        expected_inventory
    )
    assert (
        load_snapshot(split_complex).word_hyperlink_fields.signature
        == after_snapshot.word_hyperlink_fields.signature
    )
    assert after_snapshot.public_dict()["external_fields"] == {
        "database_field_count": 0,
        "legacy_data_field_count": 0,
        "dde_field_count": 0,
        "dde_auto_field_count": 0,
        "include_text_field_count": 0,
        "include_picture_field_count": 0,
        "link_field_count": 0,
        "referenced_document_field_count": 0,
    }

    report = diff_documents(before, after)
    assert "word_hyperlink_field_inventory_changed" in {
        change.kind for change in report.changes
    }
    assert "word_hyperlink_field_inventory_changed" in {
        change.kind for change in diff_documents(after, changed).changes
    }

    policy_path.write_text(
        """version: 1
rules:
  require_no_word_hyperlink_fields: true
  no_word_hyperlink_field_changes: true
""",
        encoding="utf-8",
    )
    policy = load_policy(policy_path)
    assert {finding.rule_id for finding in apply_policy(report, policy).findings} == {
        "DFP047",
        "DFP048",
    }
    assert {
        finding.rule_id
        for finding in apply_policy(diff_documents(after, changed), policy).findings
    } == {"DFP047", "DFP048"}

    gated = apply_policy(report, policy)
    rendered = "\n".join(
        (
            render_profile(after_snapshot, "json"),
            render_profile(after_snapshot, "markdown"),
            render_report(gated, "json"),
            render_report(gated, "markdown"),
            render_report(gated, "sarif"),
        )
    )
    sarif = json.loads(render_report(gated, "sarif"))
    assert {
        "DFC_WORD_HYPERLINK_FIELD_INVENTORY_CHANGED",
        "DFP047",
        "DFP048",
    } <= {result["ruleId"] for result in sarif["runs"][0]["results"]}
    for marker in (
        "DESTINATION_DO_NOT_LEAK",
        "CHANGED_DESTINATION_DO_NOT_LEAK",
        "SECOND_DESTINATION_DO_NOT_LEAK",
        "INTERNAL_LOCATION_DO_NOT_LEAK",
        "SECOND_INTERNAL_LOCATION_DO_NOT_LEAK",
        "TOOLTIP_DO_NOT_LEAK",
        "TARGET_FRAME_DO_NOT_LEAK",
        "COMPOUND_DESTINATION_DO_NOT_LEAK",
        "UNPARSED_ARGUMENT_DO_NOT_LEAK",
        "DYNAMIC_HYPERLINK_DO_NOT_LEAK",
        "LOOSE_HYPERLINK_DO_NOT_LEAK",
        "POST_SEPARATOR_HYPERLINK_DO_NOT_LEAK",
        "UNCLOSED_HYPERLINK_DO_NOT_LEAK",
        "HEADER_HYPERLINK_DO_NOT_LEAK",
    ):
        assert marker not in rendered


def test_word_hyperlink_fields_keep_revision_variants_privately(tmp_path) -> None:
    document = tmp_path / "revision-fields.docx"

    def instruction_text(text: str) -> str:
        return f'<w:r><w:instrText xml:space="preserve">{text}</w:instrText></w:r>'

    def field_char(field_type: str) -> str:
        return f'<w:r><w:fldChar w:fldCharType="{field_type}"/></w:r>'

    def deleted_instruction_text(text: str) -> str:
        return (
            '<w:del w:id="1"><w:r><w:delInstrText>'
            f"{text}</w:delInstrText></w:r></w:del>"
        )

    def inserted_instruction_text(text: str) -> str:
        return (
            '<w:ins w:id="2"><w:r><w:instrText>'
            f"{text}</w:instrText></w:r></w:ins>"
        )

    body_markup = "".join(
        (
            field_char("begin"),
            instruction_text(' HYPERLINK "'),
            deleted_instruction_text("DELETED_DESTINATION_DO_NOT_LEAK"),
            inserted_instruction_text("CURRENT_DESTINATION_DO_NOT_LEAK"),
            instruction_text('" '),
            field_char("end"),
            field_char("begin"),
            deleted_instruction_text(' HYPERLINK \\l "RETIRED_LOCATION_DO_NOT_LEAK"'),
            inserted_instruction_text(" DATE "),
            field_char("end"),
            field_char("begin"),
            (
                '<w:moveFrom w:id="3"><w:r><w:delInstrText>'
                "HYPERLINK MOVED_FROM_DESTINATION_DO_NOT_LEAK"
                "</w:delInstrText></w:r></w:moveFrom>"
            ),
            (
                '<w:moveTo w:id="4"><w:r><w:instrText>'
                "HYPERLINK MOVED_TO_DESTINATION_DO_NOT_LEAK"
                "</w:instrText></w:r></w:moveTo>"
            ),
            field_char("end"),
            deleted_instruction_text("HYPERLINK LOOSE_DELETED_HYPERLINK_DO_NOT_LEAK"),
        )
    )
    _write_document_variable_field_document(document, body_markup=body_markup)

    snapshot = load_snapshot(document)
    assert snapshot.public_dict()["word_hyperlink_fields"] == {
        "hyperlink_field_reference_count": 5,
        "hyperlink_field_story_count": 1,
        "literal_destination_hyperlink_field_count": 4,
        "literal_internal_location_only_hyperlink_field_count": 1,
        "dynamic_or_unparseable_hyperlink_field_count": 0,
    }
    rendered = render_profile(snapshot, "json") + render_profile(snapshot, "markdown")
    for marker in (
        "DELETED_DESTINATION_DO_NOT_LEAK",
        "CURRENT_DESTINATION_DO_NOT_LEAK",
        "RETIRED_LOCATION_DO_NOT_LEAK",
        "MOVED_FROM_DESTINATION_DO_NOT_LEAK",
        "MOVED_TO_DESTINATION_DO_NOT_LEAK",
        "LOOSE_DELETED_HYPERLINK_DO_NOT_LEAK",
    ):
        assert marker not in rendered


def test_word_hyperlink_markup_inventory_is_private_and_semantic(tmp_path) -> None:
    before = tmp_path / "before.docx"
    after = tmp_path / "after.docx"
    changed = tmp_path / "changed.docx"
    markup_changed = tmp_path / "markup-changed.docx"
    renumbered = tmp_path / "renumbered.docx"
    orphaned_relationship = tmp_path / "orphaned-relationship.docx"
    strict = tmp_path / "strict.docx"
    policy_path = tmp_path / "docfence.yml"

    _write_hyperlink_markup_document(before, include_markup=False)
    _write_hyperlink_markup_document(after)
    _write_hyperlink_markup_document(
        changed,
        external_target="https://CHANGED_MARKUP_TARGET_DO_NOT_LEAK.invalid/path",
    )
    _write_hyperlink_markup_document(
        markup_changed,
        shadowed_anchor="CHANGED_MARKUP_SHADOWED_ANCHOR_DO_NOT_LEAK",
    )
    _write_hyperlink_markup_document(
        renumbered,
        external_relationship_id="rIdRENUMBERED_DO_NOT_LEAK",
    )
    _write_hyperlink_markup_document(
        orphaned_relationship,
        include_markup=False,
        include_orphan_hyperlink_relationship=True,
    )
    _write_hyperlink_markup_document(
        strict,
        word_namespace=_STRICT_WORD_NAMESPACE,
        relationship_attribute_namespace=_STRICT_RELATIONSHIP_NAMESPACE,
        relationship_namespace=_STRICT_PACKAGE_RELATIONSHIP_NAMESPACE,
        hyperlink_relationship_type=_STRICT_HYPERLINK_RELATIONSHIP_TYPE,
    )

    expected_inventory = {
        "hyperlink_element_count": 6,
        "hyperlink_story_count": 2,
        "relationship_backed_hyperlink_count": 4,
        "external_relationship_hyperlink_count": 2,
        "internal_relationship_hyperlink_count": 1,
        "unsupported_relationship_hyperlink_count": 1,
        "anchor_only_hyperlink_count": 1,
        "default_document_start_hyperlink_count": 1,
        "relationship_backed_anchor_attribute_count": 1,
    }
    before_snapshot = load_snapshot(before)
    after_snapshot = load_snapshot(after)
    changed_snapshot = load_snapshot(changed)
    markup_changed_snapshot = load_snapshot(markup_changed)
    renumbered_snapshot = load_snapshot(renumbered)
    assert after_snapshot.public_dict()["word_hyperlink_markup"] == expected_inventory
    assert before_snapshot.public_dict()["word_hyperlink_markup"] == {
        key: 0 for key in expected_inventory
    }
    orphaned_snapshot = load_snapshot(orphaned_relationship)
    assert orphaned_snapshot.public_dict()["relationships"] == {
        "relationship_count": 1,
        "external_relationship_count": 1,
    }
    assert orphaned_snapshot.public_dict()["word_hyperlink_markup"] == {
        key: 0 for key in expected_inventory
    }
    assert load_snapshot(strict).public_dict()["word_hyperlink_markup"] == (
        expected_inventory
    )
    assert (
        renumbered_snapshot.word_hyperlink_markup.signature
        == after_snapshot.word_hyperlink_markup.signature
    )

    report = diff_documents(before, after)
    assert "word_hyperlink_markup_inventory_changed" in {
        change.kind for change in report.changes
    }
    assert "word_hyperlink_markup_inventory_changed" in {
        change.kind for change in diff_documents(after, changed).changes
    }
    assert "word_hyperlink_markup_inventory_changed" in {
        change.kind for change in diff_documents(after, markup_changed).changes
    }
    assert "word_hyperlink_markup_inventory_changed" not in {
        change.kind for change in diff_documents(after, renumbered).changes
    }

    policy_path.write_text(
        """version: 1
rules:
  require_no_word_hyperlink_markup: true
  no_word_hyperlink_markup_changes: true
""",
        encoding="utf-8",
    )
    policy = load_policy(policy_path)
    assert {finding.rule_id for finding in apply_policy(report, policy).findings} == {
        "DFP049",
        "DFP050",
    }
    assert {
        finding.rule_id
        for finding in apply_policy(diff_documents(after, changed), policy).findings
    } == {"DFP049", "DFP050"}

    gated = apply_policy(report, policy)
    changed_gated = apply_policy(diff_documents(after, changed), policy)
    rendered = "\n".join(
        (
            render_profile(after_snapshot, "json"),
            render_profile(after_snapshot, "markdown"),
            render_report(gated, "json"),
            render_report(gated, "markdown"),
            render_report(gated, "sarif"),
            render_profile(changed_snapshot, "json"),
            render_report(changed_gated, "sarif"),
            render_profile(markup_changed_snapshot, "markdown"),
            render_profile(renumbered_snapshot, "json"),
        )
    )
    sarif = json.loads(render_report(gated, "sarif"))
    assert {
        "DFC_WORD_HYPERLINK_MARKUP_INVENTORY_CHANGED",
        "DFP049",
        "DFP050",
    } <= {result["ruleId"] for result in sarif["runs"][0]["results"]}
    for marker in (
        "MARKUP_EXTERNAL_TARGET_DO_NOT_LEAK",
        "CHANGED_MARKUP_TARGET_DO_NOT_LEAK",
        "MARKUP_INTERNAL_TARGET_DO_NOT_LEAK",
        "MARKUP_UNSUPPORTED_TARGET_DO_NOT_LEAK",
        "MARKUP_HEADER_TARGET_DO_NOT_LEAK",
        "MARKUP_SHADOWED_ANCHOR_DO_NOT_LEAK",
        "CHANGED_MARKUP_SHADOWED_ANCHOR_DO_NOT_LEAK",
        "MARKUP_ANCHOR_DO_NOT_LEAK",
        "MARKUP_DOCUMENT_LOCATION_DO_NOT_LEAK",
        "MARKUP_TOOLTIP_DO_NOT_LEAK",
        "MARKUP_TARGET_FRAME_DO_NOT_LEAK",
        "MARKUP_DISPLAY_TEXT_DO_NOT_LEAK",
        "MARKUP_HEADER_DISPLAY_DO_NOT_LEAK",
        "rIdRENUMBERED_DO_NOT_LEAK",
    ):
        assert marker not in rendered


def test_word_drawing_hyperlink_inventory_is_private_and_semantic(tmp_path) -> None:
    before = tmp_path / "before.docx"
    after = tmp_path / "after.docx"
    changed = tmp_path / "changed.docx"
    markup_changed = tmp_path / "markup-changed.docx"
    renumbered = tmp_path / "renumbered.docx"
    orphaned_relationship = tmp_path / "orphaned-relationship.docx"
    strict = tmp_path / "strict.docx"
    policy_path = tmp_path / "docfence.yml"

    _write_drawing_hyperlink_document(before, include_markup=False)
    _write_drawing_hyperlink_document(after)
    _write_drawing_hyperlink_document(
        changed,
        external_target="https://CHANGED_DRAWING_TARGET_DO_NOT_LEAK.invalid/path",
    )
    _write_drawing_hyperlink_document(
        markup_changed,
        action="CHANGED_DRAWING_ACTION_DO_NOT_LEAK",
    )
    _write_drawing_hyperlink_document(
        renumbered,
        external_relationship_id="rIdRENUMBERED_DRAWING_DO_NOT_LEAK",
    )
    _write_drawing_hyperlink_document(
        orphaned_relationship,
        include_markup=False,
        include_orphan_hyperlink_relationship=True,
    )
    _write_drawing_hyperlink_document(
        strict,
        word_namespace=_STRICT_WORD_NAMESPACE,
        drawing_namespace=_STRICT_DRAWING_NAMESPACE,
        relationship_attribute_namespace=_STRICT_RELATIONSHIP_NAMESPACE,
        relationship_namespace=_STRICT_PACKAGE_RELATIONSHIP_NAMESPACE,
        hyperlink_relationship_type=_STRICT_HYPERLINK_RELATIONSHIP_TYPE,
    )

    expected_inventory = {
        "drawing_hyperlink_reference_count": 7,
        "drawing_hyperlink_story_count": 2,
        "click_drawing_hyperlink_count": 4,
        "hover_drawing_hyperlink_count": 2,
        "mouse_over_drawing_hyperlink_count": 1,
        "external_relationship_drawing_hyperlink_count": 4,
        "internal_relationship_drawing_hyperlink_count": 1,
        "unsupported_relationship_drawing_hyperlink_count": 1,
        "missing_relationship_id_drawing_hyperlink_count": 1,
        "action_attribute_drawing_hyperlink_count": 1,
        "invalid_url_attribute_drawing_hyperlink_count": 1,
    }
    before_snapshot = load_snapshot(before)
    after_snapshot = load_snapshot(after)
    changed_snapshot = load_snapshot(changed)
    markup_changed_snapshot = load_snapshot(markup_changed)
    renumbered_snapshot = load_snapshot(renumbered)
    assert after_snapshot.public_dict()["word_drawing_hyperlinks"] == expected_inventory
    assert before_snapshot.public_dict()["word_drawing_hyperlinks"] == {
        key: 0 for key in expected_inventory
    }
    orphaned_snapshot = load_snapshot(orphaned_relationship)
    assert orphaned_snapshot.public_dict()["relationships"] == {
        "relationship_count": 1,
        "external_relationship_count": 1,
    }
    assert orphaned_snapshot.public_dict()["word_drawing_hyperlinks"] == {
        key: 0 for key in expected_inventory
    }
    assert load_snapshot(strict).public_dict()["word_drawing_hyperlinks"] == (
        expected_inventory
    )
    assert (
        renumbered_snapshot.word_drawing_hyperlinks.signature
        == after_snapshot.word_drawing_hyperlinks.signature
    )

    report = diff_documents(before, after)
    assert "word_drawing_hyperlink_inventory_changed" in {
        change.kind for change in report.changes
    }
    assert "word_drawing_hyperlink_inventory_changed" in {
        change.kind for change in diff_documents(after, changed).changes
    }
    assert "word_drawing_hyperlink_inventory_changed" in {
        change.kind for change in diff_documents(after, markup_changed).changes
    }
    assert "word_drawing_hyperlink_inventory_changed" not in {
        change.kind for change in diff_documents(after, renumbered).changes
    }

    policy_path.write_text(
        """version: 1
rules:
  require_no_word_drawing_hyperlinks: true
  no_word_drawing_hyperlink_changes: true
""",
        encoding="utf-8",
    )
    policy = load_policy(policy_path)
    assert {finding.rule_id for finding in apply_policy(report, policy).findings} == {
        "DFP051",
        "DFP052",
    }
    assert {
        finding.rule_id
        for finding in apply_policy(diff_documents(after, changed), policy).findings
    } == {"DFP051", "DFP052"}

    gated = apply_policy(report, policy)
    changed_gated = apply_policy(diff_documents(after, changed), policy)
    rendered = "\n".join(
        (
            render_profile(after_snapshot, "json"),
            render_profile(after_snapshot, "markdown"),
            render_report(gated, "json"),
            render_report(gated, "markdown"),
            render_report(gated, "sarif"),
            render_profile(changed_snapshot, "json"),
            render_report(changed_gated, "sarif"),
            render_profile(markup_changed_snapshot, "markdown"),
            render_profile(renumbered_snapshot, "json"),
        )
    )
    sarif = json.loads(render_report(gated, "sarif"))
    assert {
        "DFC_WORD_DRAWING_HYPERLINK_INVENTORY_CHANGED",
        "DFP051",
        "DFP052",
    } <= {result["ruleId"] for result in sarif["runs"][0]["results"]}
    for marker in (
        "DRAWING_EXTERNAL_TARGET_DO_NOT_LEAK",
        "CHANGED_DRAWING_TARGET_DO_NOT_LEAK",
        "DRAWING_INTERNAL_TARGET_DO_NOT_LEAK",
        "DRAWING_UNSUPPORTED_TARGET_DO_NOT_LEAK",
        "DRAWING_HEADER_TARGET_DO_NOT_LEAK",
        "DRAWING_ACTION_DO_NOT_LEAK",
        "CHANGED_DRAWING_ACTION_DO_NOT_LEAK",
        "DRAWING_INVALID_URL_DO_NOT_LEAK",
        "DRAWING_TOOLTIP_DO_NOT_LEAK",
        "DRAWING_TARGET_FRAME_DO_NOT_LEAK",
        "DRAWING_OBJECT_NAME_DO_NOT_LEAK",
        "DRAWING_PICTURE_OBJECT_DO_NOT_LEAK",
        "rIdRENUMBERED_DRAWING_DO_NOT_LEAK",
    ):
        assert marker not in rendered


def test_word_drawing_linked_picture_inventory_is_private_and_semantic(
    tmp_path,
) -> None:
    before = tmp_path / "before.docx"
    after = tmp_path / "after.docx"
    changed = tmp_path / "changed.docx"
    markup_changed = tmp_path / "markup-changed.docx"
    renumbered = tmp_path / "renumbered.docx"
    orphaned_relationship = tmp_path / "orphaned-relationship.docx"
    strict = tmp_path / "strict.docx"
    policy_path = tmp_path / "docfence.yml"

    _write_drawing_linked_picture_document(before, include_markup=False)
    _write_drawing_linked_picture_document(after)
    _write_drawing_linked_picture_document(
        changed,
        external_target=(
            "file:///CHANGED_LINKED_PICTURE_TARGET_DO_NOT_LEAK.invalid/image.png"
        ),
    )
    _write_drawing_linked_picture_document(
        markup_changed,
        compression_state="email",
    )
    _write_drawing_linked_picture_document(
        renumbered,
        external_relationship_id="rIdRENUMBERED_LINKED_PICTURE_DO_NOT_LEAK",
    )
    _write_drawing_linked_picture_document(
        orphaned_relationship,
        include_markup=False,
        include_orphan_image_relationship=True,
    )
    _write_drawing_linked_picture_document(
        strict,
        word_namespace=_STRICT_WORD_NAMESPACE,
        drawing_namespace=_STRICT_DRAWING_NAMESPACE,
        relationship_attribute_namespace=_STRICT_RELATIONSHIP_NAMESPACE,
        relationship_namespace=_STRICT_PACKAGE_RELATIONSHIP_NAMESPACE,
        image_relationship_type=_STRICT_IMAGE_RELATIONSHIP_TYPE,
    )

    expected_inventory = {
        "drawing_linked_picture_reference_count": 5,
        "drawing_linked_picture_story_count": 2,
        "external_image_relationship_drawing_linked_picture_count": 3,
        "internal_image_relationship_drawing_linked_picture_count": 1,
        "unsupported_relationship_drawing_linked_picture_count": 1,
    }
    before_snapshot = load_snapshot(before)
    after_snapshot = load_snapshot(after)
    changed_snapshot = load_snapshot(changed)
    markup_changed_snapshot = load_snapshot(markup_changed)
    renumbered_snapshot = load_snapshot(renumbered)
    assert (
        after_snapshot.public_dict()["word_drawing_linked_pictures"]
        == expected_inventory
    )
    assert after_snapshot.public_dict()["relationships"] == {
        "relationship_count": 6,
        "external_relationship_count": 4,
    }
    assert before_snapshot.public_dict()["word_drawing_linked_pictures"] == {
        key: 0 for key in expected_inventory
    }
    orphaned_snapshot = load_snapshot(orphaned_relationship)
    assert orphaned_snapshot.public_dict()["relationships"] == {
        "relationship_count": 1,
        "external_relationship_count": 1,
    }
    assert orphaned_snapshot.public_dict()["word_drawing_linked_pictures"] == {
        key: 0 for key in expected_inventory
    }
    assert (
        load_snapshot(strict).public_dict()["word_drawing_linked_pictures"]
        == expected_inventory
    )
    assert (
        renumbered_snapshot.word_drawing_linked_pictures.signature
        == after_snapshot.word_drawing_linked_pictures.signature
    )
    for inventory_name in (
        "word_hyperlink_fields",
        "word_hyperlink_markup",
        "word_drawing_hyperlinks",
        "word_vml_hyperlinks",
    ):
        assert not any(after_snapshot.public_dict()[inventory_name].values())

    report = diff_documents(before, after)
    assert "word_drawing_linked_picture_inventory_changed" in {
        change.kind for change in report.changes
    }
    assert "word_drawing_linked_picture_inventory_changed" in {
        change.kind for change in diff_documents(after, changed).changes
    }
    assert "word_drawing_linked_picture_inventory_changed" in {
        change.kind for change in diff_documents(after, markup_changed).changes
    }
    assert "word_drawing_linked_picture_inventory_changed" not in {
        change.kind for change in diff_documents(after, renumbered).changes
    }

    policy_path.write_text(
        """version: 1
rules:
  require_no_word_drawing_linked_pictures: true
  no_word_drawing_linked_picture_changes: true
""",
        encoding="utf-8",
    )
    policy = load_policy(policy_path)
    assert {finding.rule_id for finding in apply_policy(report, policy).findings} == {
        "DFP055",
        "DFP056",
    }
    assert {
        finding.rule_id
        for finding in apply_policy(diff_documents(after, changed), policy).findings
    } == {"DFP055", "DFP056"}

    gated = apply_policy(report, policy)
    changed_gated = apply_policy(diff_documents(after, changed), policy)
    rendered = "\n".join(
        (
            render_profile(after_snapshot, "json"),
            render_profile(after_snapshot, "markdown"),
            render_report(gated, "json"),
            render_report(gated, "markdown"),
            render_report(gated, "sarif"),
            render_profile(changed_snapshot, "json"),
            render_report(changed_gated, "sarif"),
            render_profile(markup_changed_snapshot, "markdown"),
            render_profile(renumbered_snapshot, "json"),
        )
    )
    sarif = json.loads(render_report(gated, "sarif"))
    assert {
        "DFC_WORD_DRAWING_LINKED_PICTURE_INVENTORY_CHANGED",
        "DFP055",
        "DFP056",
    } <= {result["ruleId"] for result in sarif["runs"][0]["results"]}
    for marker in (
        "LINKED_PICTURE_EXTERNAL_TARGET_DO_NOT_LEAK",
        "CHANGED_LINKED_PICTURE_TARGET_DO_NOT_LEAK",
        "LINKED_PICTURE_INTERNAL_TARGET_DO_NOT_LEAK",
        "LINKED_PICTURE_UNSUPPORTED_TARGET_DO_NOT_LEAK",
        "LINKED_PICTURE_DUAL_TARGET_DO_NOT_LEAK",
        "LINKED_PICTURE_EMBEDDED_TARGET_DO_NOT_LEAK",
        "LINKED_PICTURE_HEADER_TARGET_DO_NOT_LEAK",
        "LINKED_PICTURE_OBJECT_NAME_DO_NOT_LEAK",
        "rIdRENUMBERED_LINKED_PICTURE_DO_NOT_LEAK",
    ):
        assert marker not in rendered


def test_word_vml_external_image_inventory_is_private_and_semantic(tmp_path) -> None:
    before = tmp_path / "before.docx"
    after = tmp_path / "after.docx"
    changed = tmp_path / "changed.docx"
    raw_source_changed = tmp_path / "raw-source-changed.docx"
    renumbered = tmp_path / "renumbered.docx"
    orphaned_relationship = tmp_path / "orphaned-relationship.docx"
    strict = tmp_path / "strict.docx"
    policy_path = tmp_path / "docfence.yml"

    _write_vml_external_image_document(before, include_markup=False)
    _write_vml_external_image_document(after)
    _write_vml_external_image_document(
        changed,
        external_target=(
            "file:///CHANGED_VML_EXTERNAL_IMAGE_TARGET_DO_NOT_LEAK.invalid/image.png"
        ),
    )
    _write_vml_external_image_document(
        raw_source_changed,
        primary_src="VML_CHANGED_RAW_SRC_DO_NOT_LEAK.png",
    )
    _write_vml_external_image_document(
        renumbered,
        external_relationship_id="rIdRENUMBERED_VML_EXTERNAL_IMAGE_DO_NOT_LEAK",
    )
    _write_vml_external_image_document(
        orphaned_relationship,
        include_markup=False,
        include_orphan_image_relationship=True,
    )
    _write_vml_external_image_document(
        strict,
        word_namespace=_STRICT_WORD_NAMESPACE,
        relationship_attribute_namespace=_STRICT_RELATIONSHIP_NAMESPACE,
        relationship_namespace=_STRICT_PACKAGE_RELATIONSHIP_NAMESPACE,
        image_relationship_type=_STRICT_IMAGE_RELATIONSHIP_TYPE,
    )

    expected_inventory = {
        "vml_external_image_reference_count": 4,
        "vml_external_image_story_count": 2,
        "external_image_relationship_vml_external_image_count": 3,
        "unsupported_relationship_vml_external_image_count": 1,
    }
    before_snapshot = load_snapshot(before)
    after_snapshot = load_snapshot(after)
    changed_snapshot = load_snapshot(changed)
    raw_source_changed_snapshot = load_snapshot(raw_source_changed)
    renumbered_snapshot = load_snapshot(renumbered)
    assert (
        after_snapshot.public_dict()["word_vml_external_images"]
        == expected_inventory
    )
    assert after_snapshot.public_dict()["relationships"] == {
        "relationship_count": 6,
        "external_relationship_count": 5,
    }
    assert before_snapshot.public_dict()["word_vml_external_images"] == {
        key: 0 for key in expected_inventory
    }
    orphaned_snapshot = load_snapshot(orphaned_relationship)
    assert orphaned_snapshot.public_dict()["relationships"] == {
        "relationship_count": 1,
        "external_relationship_count": 1,
    }
    assert orphaned_snapshot.public_dict()["word_vml_external_images"] == {
        key: 0 for key in expected_inventory
    }
    assert (
        load_snapshot(strict).public_dict()["word_vml_external_images"]
        == expected_inventory
    )
    assert (
        raw_source_changed_snapshot.word_vml_external_images.signature
        == after_snapshot.word_vml_external_images.signature
    )
    assert (
        renumbered_snapshot.word_vml_external_images.signature
        == after_snapshot.word_vml_external_images.signature
    )
    for inventory_name in (
        "word_hyperlink_fields",
        "word_hyperlink_markup",
        "word_drawing_hyperlinks",
        "word_drawing_linked_pictures",
        "word_vml_hyperlinks",
    ):
        assert not any(after_snapshot.public_dict()[inventory_name].values())

    report = diff_documents(before, after)
    assert "word_vml_external_image_inventory_changed" in {
        change.kind for change in report.changes
    }
    assert "word_vml_external_image_inventory_changed" in {
        change.kind for change in diff_documents(after, changed).changes
    }
    assert "word_vml_external_image_inventory_changed" not in {
        change.kind for change in diff_documents(after, raw_source_changed).changes
    }
    assert "word_vml_external_image_inventory_changed" not in {
        change.kind for change in diff_documents(after, renumbered).changes
    }

    policy_path.write_text(
        """version: 1
rules:
  require_no_word_vml_external_images: true
  no_word_vml_external_image_changes: true
""",
        encoding="utf-8",
    )
    policy = load_policy(policy_path)
    assert {finding.rule_id for finding in apply_policy(report, policy).findings} == {
        "DFP057",
        "DFP058",
    }
    assert {
        finding.rule_id
        for finding in apply_policy(diff_documents(after, changed), policy).findings
    } == {"DFP057", "DFP058"}
    assert {
        finding.rule_id
        for finding in apply_policy(
            diff_documents(after, raw_source_changed), policy
        ).findings
    } == {"DFP057"}

    gated = apply_policy(report, policy)
    changed_gated = apply_policy(diff_documents(after, changed), policy)
    rendered = "\n".join(
        (
            render_profile(after_snapshot, "json"),
            render_profile(after_snapshot, "markdown"),
            render_report(gated, "json"),
            render_report(gated, "markdown"),
            render_report(gated, "sarif"),
            render_profile(changed_snapshot, "json"),
            render_report(changed_gated, "sarif"),
            render_profile(raw_source_changed_snapshot, "markdown"),
            render_profile(renumbered_snapshot, "json"),
        )
    )
    sarif = json.loads(render_report(gated, "sarif"))
    assert {
        "DFC_WORD_VML_EXTERNAL_IMAGE_INVENTORY_CHANGED",
        "DFP057",
        "DFP058",
    } <= {result["ruleId"] for result in sarif["runs"][0]["results"]}
    for marker in (
        "VML_EXTERNAL_IMAGE_TARGET_DO_NOT_LEAK",
        "CHANGED_VML_EXTERNAL_IMAGE_TARGET_DO_NOT_LEAK",
        "VML_INTERNAL_IMAGE_TARGET_DO_NOT_LEAK",
        "VML_UNSUPPORTED_IMAGE_TARGET_DO_NOT_LEAK",
        "VML_PICT_IMAGE_TARGET_DO_NOT_LEAK",
        "VML_HREF_TARGET_DO_NOT_LEAK",
        "VML_HEADER_IMAGE_TARGET_DO_NOT_LEAK",
        "VML_RAW_SRC_DO_NOT_LEAK",
        "VML_CHANGED_RAW_SRC_DO_NOT_LEAK",
        "VML_PICT_RAW_SRC_DO_NOT_LEAK",
        "VML_OFFICE_RELID_DO_NOT_LEAK",
        "rIdRENUMBERED_VML_EXTERNAL_IMAGE_DO_NOT_LEAK",
    ):
        assert marker not in rendered


def test_word_vml_hyperlink_inventory_is_private_and_semantic(tmp_path) -> None:
    before = tmp_path / "before.docx"
    after = tmp_path / "after.docx"
    href_changed = tmp_path / "href-changed.docx"
    target_changed = tmp_path / "target-changed.docx"
    policy_path = tmp_path / "docfence.yml"

    _write_vml_hyperlink_document(before, include_markup=False)
    _write_vml_hyperlink_document(after)
    _write_vml_hyperlink_document(
        href_changed,
        primary_href="https://VML_CHANGED_HREF_DO_NOT_LEAK.invalid/path",
    )
    _write_vml_hyperlink_document(
        target_changed,
        primary_target="VML_CHANGED_TARGET_FRAME_DO_NOT_LEAK",
    )

    expected_inventory = {
        "vml_hyperlink_element_count": 11,
        "vml_hyperlink_story_count": 2,
        "concrete_shape_vml_hyperlink_count": 9,
        "group_vml_hyperlink_count": 1,
        "shape_type_vml_hyperlink_count": 1,
        "target_attribute_vml_hyperlink_count": 2,
    }
    before_snapshot = load_snapshot(before)
    after_snapshot = load_snapshot(after)
    href_changed_snapshot = load_snapshot(href_changed)
    target_changed_snapshot = load_snapshot(target_changed)
    assert after_snapshot.public_dict()["word_vml_hyperlinks"] == expected_inventory
    assert before_snapshot.public_dict()["word_vml_hyperlinks"] == {
        key: 0 for key in expected_inventory
    }
    assert after_snapshot.public_dict()["relationships"] == {
        "relationship_count": 0,
        "external_relationship_count": 0,
    }
    for inventory_name in (
        "word_hyperlink_fields",
        "word_hyperlink_markup",
        "word_drawing_hyperlinks",
    ):
        assert not any(after_snapshot.public_dict()[inventory_name].values())

    report = diff_documents(before, after)
    assert "word_vml_hyperlink_inventory_changed" in {
        change.kind for change in report.changes
    }
    assert "word_vml_hyperlink_inventory_changed" in {
        change.kind for change in diff_documents(after, href_changed).changes
    }
    assert "word_vml_hyperlink_inventory_changed" in {
        change.kind for change in diff_documents(after, target_changed).changes
    }

    policy_path.write_text(
        """version: 1
rules:
  require_no_word_vml_hyperlinks: true
  no_word_vml_hyperlink_changes: true
""",
        encoding="utf-8",
    )
    policy = load_policy(policy_path)
    assert {finding.rule_id for finding in apply_policy(report, policy).findings} == {
        "DFP053",
        "DFP054",
    }
    assert {
        finding.rule_id
        for finding in apply_policy(
            diff_documents(after, href_changed), policy
        ).findings
    } == {"DFP053", "DFP054"}

    gated = apply_policy(report, policy)
    href_changed_gated = apply_policy(diff_documents(after, href_changed), policy)
    rendered = "\n".join(
        (
            render_profile(after_snapshot, "json"),
            render_profile(after_snapshot, "markdown"),
            render_report(gated, "json"),
            render_report(gated, "markdown"),
            render_report(gated, "sarif"),
            render_profile(href_changed_snapshot, "json"),
            render_report(href_changed_gated, "sarif"),
            render_profile(target_changed_snapshot, "markdown"),
        )
    )
    sarif = json.loads(render_report(gated, "sarif"))
    assert {
        "DFC_WORD_VML_HYPERLINK_INVENTORY_CHANGED",
        "DFP053",
        "DFP054",
    } <= {result["ruleId"] for result in sarif["runs"][0]["results"]}
    for marker in (
        "VML_PRIMARY_HREF_DO_NOT_LEAK",
        "VML_CHANGED_HREF_DO_NOT_LEAK",
        "VML_PRIMARY_TARGET_FRAME_DO_NOT_LEAK",
        "VML_CHANGED_TARGET_FRAME_DO_NOT_LEAK",
        "VML_PRIMARY_TITLE_DO_NOT_LEAK",
        "VML_PRIMARY_ALT_DO_NOT_LEAK",
        "VML_PRIMARY_SHAPE_ID_DO_NOT_LEAK",
        "VML_ROUNDRECT_HREF_DO_NOT_LEAK",
        "VML_RECT_HREF_DO_NOT_LEAK",
        "VML_GROUP_HREF_DO_NOT_LEAK",
        "VML_GROUP_TARGET_FRAME_DO_NOT_LEAK",
        "VML_OVAL_HREF_DO_NOT_LEAK",
        "VML_CURVE_HREF_DO_NOT_LEAK",
        "VML_POLYLINE_HREF_DO_NOT_LEAK",
        "VML_IMAGE_HREF_DO_NOT_LEAK",
        "VML_SHAPETYPE_HREF_DO_NOT_LEAK",
        "VML_HEADER_LINE_HREF_DO_NOT_LEAK",
    ):
        assert marker not in rendered


def test_word_document_variable_discovery_and_invalid_markup(tmp_path) -> None:
    noncanonical = tmp_path / "noncanonical.docx"
    strict_relationship = tmp_path / "strict-relationship.docx"
    glossary_linked = tmp_path / "glossary-linked.docx"
    empty_value = tmp_path / "empty-value.docx"
    duplicate_containers = tmp_path / "duplicate-containers.docx"
    container_attributes = tmp_path / "container-attributes.docx"
    container_text = tmp_path / "container-text.docx"
    unexpected_child = tmp_path / "unexpected-child.docx"
    missing_name = tmp_path / "missing-name.docx"
    missing_value = tmp_path / "missing-value.docx"
    empty_name = tmp_path / "empty-name.docx"
    long_name = tmp_path / "long-name.docx"
    supplementary_long_name = tmp_path / "supplementary-long-name.docx"
    long_value = tmp_path / "long-value.docx"
    unsupported_attribute = tmp_path / "unsupported-attribute.docx"
    nonleaf = tmp_path / "nonleaf.docx"
    residual_text = tmp_path / "residual-text.docx"
    invalid_settings_root = tmp_path / "invalid-settings-root.docx"
    external_settings = tmp_path / "external-settings.docx"
    valid_markup = (
        "<w:docVars>"
        '<w:docVar w:name="VARIABLE_NAME_DO_NOT_LEAK" '
        'w:val="VARIABLE_VALUE_DO_NOT_LEAK"/>'
        "</w:docVars>"
    )

    def write_document(path, markup: str, **settings_options) -> None:
        _write_word_protection_document(
            path,
            include_document_protection=False,
            include_write_protection=False,
            settings_extra_children=markup,
            **settings_options,
        )

    write_document(
        noncanonical,
        valid_markup,
        settings_part_name="word/private/settings.xml",
        include_settings_relationship=True,
    )
    write_document(
        strict_relationship,
        valid_markup,
        settings_part_name="word/private/settings.xml",
        include_settings_relationship=True,
        settings_relationship_type=_STRICT_DOCUMENT_SETTINGS_RELATIONSHIP_TYPE,
        settings_word_namespace=_STRICT_WORD_NAMESPACE,
    )
    write_document(
        glossary_linked,
        valid_markup,
        settings_part_name="word/glossary/private/settings.xml",
        include_settings_relationship=True,
        settings_relationship_source="word/glossary/document.xml",
    )
    write_document(
        empty_value,
        '<w:docVars><w:docVar w:name="empty" w:val=""/></w:docVars>',
    )
    write_document(duplicate_containers, valid_markup + valid_markup)
    write_document(
        container_attributes,
        '<w:docVars w:futureDocumentVariable="1">'
        '<w:docVar w:name="name" w:val="value"/></w:docVars>',
    )
    write_document(
        container_text,
        '<w:docVars>unexpected<w:docVar w:name="name" w:val="value"/>'
        "</w:docVars>",
    )
    write_document(unexpected_child, "<w:docVars><w:unexpected/></w:docVars>")
    write_document(
        missing_name,
        '<w:docVars><w:docVar w:val="value"/></w:docVars>',
    )
    write_document(
        missing_value,
        '<w:docVars><w:docVar w:name="name"/></w:docVars>',
    )
    write_document(
        empty_name,
        '<w:docVars><w:docVar w:name="" w:val="value"/></w:docVars>',
    )
    write_document(
        long_name,
        f'<w:docVars><w:docVar w:name="{"N" * 256}" w:val="value"/>'
        "</w:docVars>",
    )
    write_document(
        supplementary_long_name,
        f'<w:docVars><w:docVar w:name="{"🙂" * 128}" w:val="value"/>'
        "</w:docVars>",
    )
    write_document(
        long_value,
        f'<w:docVars><w:docVar w:name="name" w:val="{"V" * 65_281}"/>'
        "</w:docVars>",
    )
    write_document(
        unsupported_attribute,
        '<w:docVars><w:docVar w:name="name" w:val="value" '
        'w:futureDocumentVariable="1"/></w:docVars>',
    )
    write_document(
        nonleaf,
        '<w:docVars><w:docVar w:name="name" w:val="value">'
        "<w:unexpected/></w:docVar></w:docVars>",
    )
    write_document(
        residual_text,
        '<w:docVars><w:docVar w:name="name" w:val="value"/>'
        "unexpected</w:docVars>",
    )
    write_document(
        invalid_settings_root,
        valid_markup,
        settings_root_name="notSettings",
    )
    write_document(
        external_settings,
        "",
        settings_part_name="word/private/settings.xml",
        include_settings_relationship=True,
        settings_target_mode="External",
        include_settings_part=False,
    )

    for document in (noncanonical, strict_relationship, glossary_linked):
        snapshot = load_snapshot(document)
        assert snapshot.word_document_variables.document_variable_container_count == 1
        assert snapshot.word_document_variables.document_variable_count == 1
        assert snapshot.unclassified_part_count == 1
    empty_snapshot = load_snapshot(empty_value)
    assert empty_snapshot.word_document_variables.document_variable_count == 1
    assert (
        empty_snapshot.word_document_variables.empty_document_variable_value_count
        == 1
    )

    for document in (
        duplicate_containers,
        container_attributes,
        container_text,
        unexpected_child,
        missing_name,
        missing_value,
        empty_name,
        long_name,
        supplementary_long_name,
        long_value,
        unsupported_attribute,
        nonleaf,
        residual_text,
        invalid_settings_root,
        external_settings,
    ):
        with pytest.raises(DocumentFormatError):
            load_snapshot(document)


def test_word_permission_range_inventory_is_private_and_semantic(tmp_path) -> None:
    before = tmp_path / "before.docx"
    after = tmp_path / "after.docx"
    editor_changed = tmp_path / "editor-changed.docx"
    policy_path = tmp_path / "docfence.yml"
    _write_word_permission_range_document(before)
    _write_word_permission_range_document(
        after,
        body_markup=(
            '<w:permStart w:id="range-identity" '
            'w:ed="RANGE_EDITOR_EMAIL_DO_NOT_LEAK" '
            'w:displacedByCustomXml="next"/>'
            '<w:permEnd w:id="range-identity"/>'
            '<w:permStart w:id="range-everyone" w:edGrp="everyone"/>'
            '<w:permEnd w:id="range-everyone" '
            'w:displacedByCustomXml="prev"/>'
            '<w:permStart w:id="range-table" w:edGrp="editors" '
            'w:colFirst="1" w:colLast="3"/>'
            '<w:permEnd w:id="range-table"/>'
            '<w:permStart w:id="range-unpaired" '
            'w:ed="RANGE_SECOND_EDITOR_DO_NOT_LEAK" w:edGrp="owners"/>'
            '<w:permEnd w:id="range-lone-end"/>'
        ),
    )
    _write_word_permission_range_document(
        editor_changed,
        body_markup=(
            '<w:permStart w:id="range-identity" '
            'w:ed="RANGE_EDITOR_CHANGED_DO_NOT_LEAK" '
            'w:displacedByCustomXml="next"/>'
            '<w:permEnd w:id="range-identity"/>'
            '<w:permStart w:id="range-everyone" w:edGrp="everyone"/>'
            '<w:permEnd w:id="range-everyone" '
            'w:displacedByCustomXml="prev"/>'
            '<w:permStart w:id="range-table" w:edGrp="editors" '
            'w:colFirst="1" w:colLast="3"/>'
            '<w:permEnd w:id="range-table"/>'
            '<w:permStart w:id="range-unpaired" '
            'w:ed="RANGE_SECOND_EDITOR_DO_NOT_LEAK" w:edGrp="owners"/>'
            '<w:permEnd w:id="range-lone-end"/>'
        ),
    )

    expected_inventory = {
        "permission_range_story_count": 1,
        "permission_start_count": 4,
        "permission_end_count": 4,
        "paired_permission_range_count": 3,
        "unpaired_permission_start_count": 1,
        "unpaired_permission_end_count": 1,
        "individual_editor_assignment_count": 2,
        "editor_group_assignment_count": 3,
        "editor_group_none_count": 0,
        "editor_group_everyone_count": 1,
        "editor_group_administrators_count": 0,
        "editor_group_contributors_count": 0,
        "editor_group_editors_count": 1,
        "editor_group_owners_count": 1,
        "editor_group_current_count": 0,
        "table_column_permission_range_start_count": 1,
        "custom_xml_displaced_permission_marker_count": 2,
    }
    before_snapshot = load_snapshot(before)
    after_snapshot = load_snapshot(after)
    assert after_snapshot.public_dict()["word_permission_ranges"] == expected_inventory
    assert before_snapshot.public_dict()["word_permission_ranges"] == {
        key: 0 for key in expected_inventory
    }

    report = diff_documents(before, after)
    assert "word_permission_range_inventory_changed" in {
        change.kind for change in report.changes
    }
    assert "word_permission_range_inventory_changed" in {
        change.kind for change in diff_documents(after, editor_changed).changes
    }

    policy_path.write_text(
        """version: 1
rules:
  require_no_word_permission_ranges: true
  no_word_permission_range_changes: true
""",
        encoding="utf-8",
    )
    policy = load_policy(policy_path)
    assert {finding.rule_id for finding in apply_policy(report, policy).findings} == {
        "DFP041",
        "DFP042",
    }
    assert {
        finding.rule_id
        for finding in apply_policy(
            diff_documents(after, editor_changed), policy
        ).findings
    } == {"DFP041", "DFP042"}

    gated = apply_policy(report, policy)
    rendered = "\n".join(
        (
            render_profile(after_snapshot, "json"),
            render_profile(after_snapshot, "markdown"),
            render_report(gated, "json"),
            render_report(gated, "markdown"),
            render_report(gated, "sarif"),
        )
    )
    sarif = json.loads(render_report(gated, "sarif"))
    assert {
        "DFC_WORD_PERMISSION_RANGE_INVENTORY_CHANGED",
        "DFP041",
        "DFP042",
    } <= {result["ruleId"] for result in sarif["runs"][0]["results"]}
    for marker in (
        "RANGE_EDITOR_EMAIL_DO_NOT_LEAK",
        "RANGE_SECOND_EDITOR_DO_NOT_LEAK",
        "RANGE_EDITOR_CHANGED_DO_NOT_LEAK",
        "range-identity",
        "range-everyone",
        "range-table",
        "range-unpaired",
        "range-lone-end",
    ):
        assert marker not in rendered


def test_word_permission_range_discovery_and_invalid_markup(tmp_path) -> None:
    multi_story = tmp_path / "multi-story.docx"
    strict = tmp_path / "strict.docx"
    unmatched = tmp_path / "unmatched.docx"
    missing_identifier = tmp_path / "missing-identifier.docx"
    duplicate_start_identifier = tmp_path / "duplicate-start-identifier.docx"
    duplicate_end_identifier = tmp_path / "duplicate-end-identifier.docx"
    invalid_group = tmp_path / "invalid-group.docx"
    invalid_column = tmp_path / "invalid-column.docx"
    invalid_displacement = tmp_path / "invalid-displacement.docx"
    unsupported_attribute = tmp_path / "unsupported-attribute.docx"
    nonleaf = tmp_path / "nonleaf.docx"

    _write_word_permission_range_document(
        multi_story,
        body_markup=(
            '<w:permStart w:id="shared" w:edGrp="current"/>'
            '<w:permEnd w:id="shared"/>'
        ),
        header_markup=(
            '<w:permStart w:id="shared" w:edGrp="administrators"/>'
            '<w:permEnd w:id="shared"/>'
        ),
    )
    _write_word_permission_range_document(
        strict,
        word_namespace=_STRICT_WORD_NAMESPACE,
        body_markup=(
            '<w:permStart w:id="strict-range" w:edGrp="contributors" '
            'w:colFirst="0"/><w:permEnd w:id="strict-range"/>'
        ),
    )
    _write_word_permission_range_document(
        unmatched,
        body_markup=(
            '<w:permEnd w:id="end-before-start"/>'
            '<w:permStart w:id="end-before-start" w:edGrp="none"/>'
        ),
    )
    _write_word_permission_range_document(
        missing_identifier, body_markup='<w:permStart w:edGrp="everyone"/>'
    )
    _write_word_permission_range_document(
        duplicate_start_identifier,
        body_markup=(
            '<w:permStart w:id="same"/><w:permStart w:id="same"/>'
            '<w:permEnd w:id="same"/>'
        ),
    )
    _write_word_permission_range_document(
        duplicate_end_identifier,
        body_markup=(
            '<w:permStart w:id="same"/><w:permEnd w:id="same"/>'
            '<w:permEnd w:id="same"/>'
        ),
    )
    _write_word_permission_range_document(
        invalid_group, body_markup='<w:permStart w:id="one" w:edGrp="guests"/>'
    )
    _write_word_permission_range_document(
        invalid_column,
        body_markup='<w:permStart w:id="one" w:colFirst="-1"/>',
    )
    _write_word_permission_range_document(
        invalid_displacement,
        body_markup='<w:permEnd w:id="one" w:displacedByCustomXml="later"/>',
    )
    _write_word_permission_range_document(
        unsupported_attribute,
        body_markup='<w:permStart w:id="one" w:futurePermission="1"/>',
    )
    _write_word_permission_range_document(
        nonleaf,
        body_markup='<w:permStart w:id="one"><w:unexpected/></w:permStart>',
    )

    multi_story_snapshot = load_snapshot(multi_story)
    assert (
        multi_story_snapshot.word_permission_ranges.permission_range_story_count
        == 2
    )
    assert (
        multi_story_snapshot.word_permission_ranges.paired_permission_range_count == 2
    )
    assert multi_story_snapshot.word_permission_ranges.editor_group_current_count == 1
    assert (
        multi_story_snapshot.word_permission_ranges.editor_group_administrators_count
        == 1
    )
    strict_snapshot = load_snapshot(strict)
    assert strict_snapshot.word_permission_ranges.paired_permission_range_count == 1
    assert (
        strict_snapshot.word_permission_ranges.editor_group_contributors_count == 1
    )
    unmatched_snapshot = load_snapshot(unmatched)
    assert (
        unmatched_snapshot.word_permission_ranges.paired_permission_range_count == 0
    )
    assert (
        unmatched_snapshot.word_permission_ranges.unpaired_permission_start_count
        == 1
    )
    assert unmatched_snapshot.word_permission_ranges.unpaired_permission_end_count == 1

    for document in (
        missing_identifier,
        duplicate_start_identifier,
        duplicate_end_identifier,
        invalid_group,
        invalid_column,
        invalid_displacement,
        unsupported_attribute,
        nonleaf,
    ):
        with pytest.raises(DocumentFormatError):
            load_snapshot(document)


def test_word_templates_are_supported_as_first_class_scan_targets(tmp_path) -> None:
    template = tmp_path / "review-template.dotx"
    macro_template = tmp_path / "review-template.dotm"
    _write_modern_comment_metadata_document(
        template,
        main_content_type=_DOTX_MAIN_TYPE,
    )
    _write_modern_comment_metadata_document(
        macro_template,
        main_content_type=_DOTM_MAIN_TYPE,
        macro_payload=b"TEMPLATE_MACRO_DO_NOT_LEAK",
    )

    template_snapshot = load_snapshot(template)
    macro_template_snapshot = load_snapshot(macro_template)
    assert template_snapshot.format == "dotx"
    assert macro_template_snapshot.format == "dotm"
    assert template_snapshot.modern_comment_metadata.person_count == 2
    assert macro_template_snapshot.modern_comment_metadata.comment_extension_count == 2
    assert macro_template_snapshot.macro_present is True


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


def _write_modern_comment_metadata_document(
    path,
    *,
    include_metadata: bool = True,
    metadata_marker: str = "MODERN_COMMENT",
    comment_identity_marker: str | None = None,
    relationship_id_suffix: str = "1",
    part_names: dict[str, str] | None = None,
    include_metadata_relationships: bool = True,
    include_metadata_content_types: bool = True,
    modern_relationship_target_mode: str = "Internal",
    wrong_people_root: bool = False,
    office_15_namespace: str = _WORD_2012_NAMESPACE,
    main_content_type: str = DOCX_MAIN_TYPE,
    macro_payload: bytes | None = None,
) -> None:
    """Write a small package with all four modern Word comment metadata parts."""

    modern_part_names = part_names or {
        "people": "word/people.xml",
        "comments_extended": "word/commentsExtended.xml",
        "comments_ids": "word/commentsIds.xml",
        "comments_extensible": "word/commentsExtensible.xml",
    }
    modern_content_types = {
        "people": (
            "application/vnd.openxmlformats-officedocument.wordprocessingml.people+xml"
        ),
        "comments_extended": (
            "application/vnd.openxmlformats-officedocument.wordprocessingml."
            "commentsExtended+xml"
        ),
        "comments_ids": (
            "application/vnd.openxmlformats-officedocument.wordprocessingml."
            "commentsIds+xml"
        ),
        "comments_extensible": (
            "application/vnd.openxmlformats-officedocument.wordprocessingml."
            "commentsExtensible+xml"
        ),
    }
    modern_relationship_types = {
        "people": _PEOPLE_RELATIONSHIP_TYPE,
        "comments_extended": _COMMENTS_EXTENDED_RELATIONSHIP_TYPE,
        "comments_ids": _COMMENTS_IDS_RELATIONSHIP_TYPE,
        "comments_extensible": _COMMENTS_EXTENSIBLE_RELATIONSHIP_TYPE,
    }
    people_root_name = "notPeople" if wrong_people_root else "people"
    identity_marker = comment_identity_marker or metadata_marker
    primary_author = f"{metadata_marker}_AUTHOR_DO_NOT_LEAK"
    primary_provider = f"{metadata_marker}_PROVIDER_DO_NOT_LEAK"
    primary_user_id = f"{metadata_marker}_USER_ID_DO_NOT_LEAK"
    secondary_author = f"{metadata_marker}_SECOND_AUTHOR_DO_NOT_LEAK"
    secondary_provider = f"{metadata_marker}_SECOND_PROVIDER_DO_NOT_LEAK"
    secondary_user_id = f"{metadata_marker}_SECOND_USER_ID_DO_NOT_LEAK"
    primary_para_id = f"{identity_marker}_PARA_A_DO_NOT_LEAK"
    secondary_para_id = f"{identity_marker}_PARA_B_DO_NOT_LEAK"
    primary_durable_id = f"{identity_marker}_DURABLE_ID_DO_NOT_LEAK"
    secondary_durable_id = f"{identity_marker}_SECOND_DURABLE_ID_DO_NOT_LEAK"
    reaction_provider = f"{metadata_marker}_REACTION_PROVIDER_DO_NOT_LEAK"
    reaction_users = (
        f"{metadata_marker}_REACTION_USER_DO_NOT_LEAK",
        f"{metadata_marker}_SECOND_REACTION_USER_DO_NOT_LEAK",
        f"{metadata_marker}_THIRD_REACTION_USER_DO_NOT_LEAK",
    )
    reaction_names = (
        f"{metadata_marker}_REACTION_NAME_DO_NOT_LEAK",
        f"{metadata_marker}_SECOND_REACTION_NAME_DO_NOT_LEAK",
        f"{metadata_marker}_THIRD_REACTION_NAME_DO_NOT_LEAK",
    )
    modern_payloads = {
        "people": f'''<w15:{people_root_name}
  xmlns:w15="{office_15_namespace}">
  <w15:person w15:author="{primary_author}">
    <w15:presenceInfo
      w15:providerId="{primary_provider}"
      w15:userId="{primary_user_id}"/>
  </w15:person>
  <w15:person w15:author="{secondary_author}">
    <w15:presenceInfo
      w15:providerId="{secondary_provider}"
      w15:userId="{secondary_user_id}"/>
  </w15:person>
</w15:{people_root_name}>'''.encode(),
        "comments_extended": f'''<w15:commentsEx
  xmlns:w15="{office_15_namespace}">
  <w15:commentEx w15:paraId="{primary_para_id}" w15:done="1"/>
  <w15:commentEx
    w15:paraId="{secondary_para_id}"
    w15:paraIdParent="{primary_para_id}"
    w15:done="0"/>
</w15:commentsEx>'''.encode(),
        "comments_ids": f'''<w16cid:commentsIds
  xmlns:w16cid="{_WORD_2016_COMMENT_IDS_NAMESPACE}">
  <w16cid:commentId
    w16cid:paraId="{primary_para_id}"
    w16cid:durableId="{primary_durable_id}"/>
  <w16cid:commentId
    w16cid:paraId="{secondary_para_id}"
    w16cid:durableId="{secondary_durable_id}"/>
</w16cid:commentsIds>'''.encode(),
        "comments_extensible": f'''<w16cex:commentsExtensible
  xmlns:w16cex="{_WORD_2018_COMMENT_EXTENSIBLE_NAMESPACE}"
  xmlns:w16="{_WORD_2018_NAMESPACE}"
  xmlns:cr="{_COMMENT_REACTIONS_NAMESPACE}">
  <w16cex:commentExtensible
    w16cex:durableId="{primary_durable_id}"
    w16cex:dateUtc="2026-08-01T00:00:00Z">
    <w16cex:extLst>
      <w16:ext w16:uri="{{CE6994B0-6A32-4C9F-8C6B-6E91EDA988CE}}">
        <cr:reactions>
          <cr:reaction reactionType="1">
            <cr:reactionInfo dateUtc="2026-08-01T00:00:00Z">
              <cr:user
                userId="{reaction_users[0]}"
                userName="{reaction_names[0]}"
                userProvider="{reaction_provider}"/>
            </cr:reactionInfo>
            <cr:reactionInfo dateUtc="2026-08-01T01:00:00Z">
              <cr:user
                userId="{reaction_users[1]}"
                userName="{reaction_names[1]}"
                userProvider="{reaction_provider}"/>
            </cr:reactionInfo>
          </cr:reaction>
        </cr:reactions>
      </w16:ext>
    </w16cex:extLst>
  </w16cex:commentExtensible>
  <w16cex:commentExtensible
    w16cex:durableId="{secondary_durable_id}"
    w16cex:dateUtc="2026-08-01T02:00:00Z">
    <w16cex:extLst>
      <w16:ext w16:uri="{{CE6994B0-6A32-4C9F-8C6B-6E91EDA988CE}}">
        <cr:reactions>
          <cr:reaction reactionType="2">
            <cr:reactionInfo dateUtc="2026-08-01T02:00:00Z">
              <cr:user
                userId="{reaction_users[2]}"
                userName="{reaction_names[2]}"
                userProvider="{reaction_provider}"/>
            </cr:reactionInfo>
          </cr:reaction>
        </cr:reactions>
      </w16:ext>
    </w16cex:extLst>
  </w16cex:commentExtensible>
</w16cex:commentsExtensible>'''.encode(),
    }

    modern_overrides = (
        "".join(
            f'<Override PartName="/{modern_part_names[kind]}" '
            f'ContentType="{modern_content_types[kind]}"/>'
            for kind in sorted(modern_part_names)
        )
        if include_metadata and include_metadata_content_types
        else ""
    )
    relationships = [
        f'<Relationship Id="rIdComments{relationship_id_suffix}" '
        f'Type="{_COMMENTS_RELATIONSHIP_TYPE}" Target="comments.xml"/>'
    ]
    if include_metadata and include_metadata_relationships:
        target_mode = (
            ' TargetMode="External"'
            if modern_relationship_target_mode == "External"
            else ""
        )
        relationships.extend(
            f'<Relationship Id="rId{kind.title().replace("_", "")}'
            f'{relationship_id_suffix}" '
            f'Type="{modern_relationship_types[kind]}" '
            f'Target="{modern_part_names[kind].removeprefix("word/")}"{target_mode}/>'
            for kind in sorted(modern_part_names)
        )

    main_document_override = (
        f'<Override PartName="/word/document.xml" ContentType="{main_content_type}"/>'
    )

    entries: dict[str, bytes] = {
        "[Content_Types].xml": (
            f'<Types xmlns="{CT}"><Default Extension="rels" '
            'ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            f"{main_document_override}"
            '<Override PartName="/word/comments.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.'
            'wordprocessingml.comments+xml"/>'
            f"{modern_overrides}</Types>"
        ).encode(),
        "word/document.xml": (
            f'<w:document xmlns:w="{W}"><w:body><w:p><w:r><w:t>'
            "BODY_DO_NOT_LEAK"
            "</w:t></w:r></w:p><w:sectPr/></w:body></w:document>"
        ).encode(),
        "word/comments.xml": (
            f'<w:comments xmlns:w="{W}"><w:comment w:id="0" '
            'w:author="CLASSIC_COMMENT_AUTHOR_DO_NOT_LEAK"><w:p><w:r><w:t>'
            "CLASSIC_COMMENT_TEXT_DO_NOT_LEAK"
            "</w:t></w:r></w:p></w:comment></w:comments>"
        ).encode(),
        "word/_rels/document.xml.rels": (
            f'<Relationships xmlns="{PR}">{"".join(relationships)}</Relationships>'
        ).encode(),
    }
    if include_metadata:
        entries.update(
            {
                modern_part_names[kind]: modern_payloads[kind]
                for kind in modern_part_names
            }
        )
    if macro_payload is not None:
        entries["word/vbaProject.bin"] = macro_payload

    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for name, payload in entries.items():
            archive.writestr(name, payload)


def _write_sensitivity_label_document(
    path,
    *,
    include_label_info: bool = True,
    include_legacy_properties: bool = True,
    include_label_info_part: bool = True,
    include_label_info_relationship: bool = True,
    include_label_info_content_type: bool = True,
    include_legacy_property_relationship: bool = True,
    label_info_part_name: str = "docMetadata/LabelInfo.xml",
    additional_label_info_part_name: str | None = None,
    custom_property_part_name: str = "docProps/custom.xml",
    label_info_target_mode: str = "Internal",
    label_info_relationship_source: str = "/",
    relationship_id_suffix: str = "1",
    wrong_label_info_root: bool = False,
    omit_label_method: bool = False,
    label_site_id: str = "{A0A00000-0000-0000-0000-000000000001}",
    label_name: str = "SENSITIVITY_LABEL_NAME_DO_NOT_LEAK",
    label_extension_marker: str = "SENSITIVITY_LABEL_EXTENSION_DO_NOT_LEAK",
    legacy_label_name: str = "SENSITIVITY_LEGACY_NAME_DO_NOT_LEAK",
) -> None:
    """Write a small package carrying current and legacy sensitivity metadata."""

    def relationship_target(
        source_part: str, target_part: str, target_mode: str
    ) -> str:
        if target_mode != "Internal":
            return "https://example.invalid/SENSITIVITY_LABEL_TARGET_DO_NOT_LEAK"
        if source_part == "/":
            return target_part
        return posixpath.relpath(target_part, start=source_part.rpartition("/")[0])

    def relationship_markup(
        relationship_id: str,
        relationship_type: str,
        source_part: str,
        target_part: str,
        target_mode: str = "Internal",
    ) -> str:
        target_mode_attribute = (
            "" if target_mode == "Internal" else f' TargetMode="{target_mode}"'
        )
        return (
            f'<Relationship Id="{relationship_id}" Type="{relationship_type}" '
            f'Target="{relationship_target(source_part, target_part, target_mode)}"'
            f"{target_mode_attribute}/>"
        )

    root_relationships: list[str] = []
    document_relationships: list[str] = []
    if include_label_info and include_label_info_relationship:
        label_relationship = relationship_markup(
            f"rIdLabelInfo{relationship_id_suffix}",
            _SENSITIVITY_LABEL_RELATIONSHIP_TYPE,
            label_info_relationship_source,
            label_info_part_name,
            label_info_target_mode,
        )
        if label_info_relationship_source == "/":
            root_relationships.append(label_relationship)
        else:
            document_relationships.append(label_relationship)
    if include_legacy_properties and include_legacy_property_relationship:
        root_relationships.append(
            relationship_markup(
                f"rIdCustomProperties{relationship_id_suffix}",
                _CUSTOM_PROPERTIES_RELATIONSHIP_TYPE,
                "/",
                custom_property_part_name,
            )
        )

    label_method = "" if omit_label_method else ' method="Privileged"'
    label_root_name = "notLabelList" if wrong_label_info_root else "labelList"
    label_info_xml = (
        f'<clbl:{label_root_name} '
        f'xmlns:clbl="{_SENSITIVITY_LABEL_NAMESPACE}" '
        'xmlns:future="urn:docfence:future-label-extension">'
        '<clbl:label id="{A0A00000-0000-0000-0000-000000000010}" '
        'enabled="true" setDate="SENSITIVITY_LABEL_DATE_DO_NOT_LEAK"'
        f"{label_method}"
        f' name="{label_name}" siteId="{label_site_id}" '
        'actionId="{A0A00000-0000-0000-0000-000000000011}" '
        'contentBits="7" removed="false"/>'
        '<clbl:label id="{A0A00000-0000-0000-0000-000000000012}" '
        'enabled="0" method="" '
        'siteId="{A0A00000-0000-0000-0000-000000000013}" removed="1"/>'
        '<clbl:extLst><clbl:ext uri="{A0A00000-0000-0000-0000-000000000014}">'
        f'<future:payload value="{label_extension_marker}"/>'
        "</clbl:ext></clbl:extLst>"
        f"</clbl:{label_root_name}>"
    ).encode()
    legacy_label_guid = "a0a00000-0000-0000-0000-000000000020"
    legacy_properties_xml = (
        f'<Properties xmlns="{_CUSTOM_PROPERTIES_NAMESPACE}" '
        f'xmlns:vt="{_DOCUMENT_PROPERTIES_VT_NAMESPACE}">'
        '<property fmtid="FMTID" pid="2" '
        f'name="MSIP_Label_{legacy_label_guid}_Enabled">'
        "<vt:lpwstr>true</vt:lpwstr></property>"
        '<property fmtid="FMTID" pid="3" '
        f'name="MSIP_Label_{legacy_label_guid}_Name">'
        f"<vt:lpwstr>{legacy_label_name}</vt:lpwstr></property>"
        '<property fmtid="FMTID" pid="4" '
        f'name="MSIP_Label_{legacy_label_guid}_GeneratedBy">'
        "<vt:lpwstr>SENSITIVITY_MIP_CUSTOM_DO_NOT_LEAK</vt:lpwstr></property>"
        '<property fmtid="FMTID" pid="5" name="Sensitivity">'
        f"<vt:lpwstr>{legacy_label_guid}</vt:lpwstr></property>"
        '<property fmtid="FMTID" pid="6" '
        'name="ClassificationContentMarkingHeaderText">'
        "<vt:lpwstr>SENSITIVITY_MARKING_TEXT_DO_NOT_LEAK</vt:lpwstr></property>"
        '<property fmtid="FMTID" pid="7" '
        'name="ClassificationContentMarkingHeaderShapeIds-1">'
        "<vt:lpwstr>aa,bb</vt:lpwstr></property>"
        '<property fmtid="FMTID" pid="8" '
        'name="ClassificationContentMarkingFooterFontProps">'
        "<vt:lpwstr>#000000,12,Calibri</vt:lpwstr></property>"
        '<property fmtid="FMTID" pid="9" name="ClassificationWatermarkText">'
        "<vt:lpwstr>SENSITIVITY_WATERMARK_DO_NOT_LEAK</vt:lpwstr></property>"
        '<property fmtid="FMTID" pid="10" name="UNRELATED_DO_NOT_LEAK">'
        "<vt:lpwstr>UNRELATED_VALUE_DO_NOT_LEAK</vt:lpwstr></property>"
        "</Properties>"
    ).encode()

    label_overrides: list[str] = []
    if include_label_info and include_label_info_content_type:
        label_overrides.append(
            f'<Override PartName="/{label_info_part_name}" '
            f'ContentType="{_SENSITIVITY_LABEL_CONTENT_TYPE}"/>'
        )
        if additional_label_info_part_name is not None:
            label_overrides.append(
                f'<Override PartName="/{additional_label_info_part_name}" '
                f'ContentType="{_SENSITIVITY_LABEL_CONTENT_TYPE}"/>'
            )

    entries: dict[str, bytes] = {
        "[Content_Types].xml": (
            f'<Types xmlns="{CT}">'
            f'<Override PartName="/word/document.xml" '
            f'ContentType="{DOCX_MAIN_TYPE}"/>'
            f"{''.join(label_overrides)}"
            "</Types>"
        ).encode(),
        "word/document.xml": (
            f'<w:document xmlns:w="{W}"><w:body><w:p><w:r><w:t>'
            "VISIBLE_DO_NOT_LEAK"
            "</w:t></w:r></w:p><w:sectPr/></w:body></w:document>"
        ).encode(),
    }
    if root_relationships:
        entries["_rels/.rels"] = (
            f'<Relationships xmlns="{PR}">{"".join(root_relationships)}</Relationships>'
        ).encode()
    if document_relationships:
        entries["word/_rels/document.xml.rels"] = (
            f'<Relationships xmlns="{PR}">'
            f"{''.join(document_relationships)}</Relationships>"
        ).encode()
    if include_label_info and include_label_info_part:
        entries[label_info_part_name] = label_info_xml
    if additional_label_info_part_name is not None:
        entries[additional_label_info_part_name] = label_info_xml
    if include_legacy_properties:
        entries[custom_property_part_name] = legacy_properties_xml

    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for name, payload in entries.items():
            archive.writestr(name, payload)


def _write_package_digital_signature_document(
    path,
    *,
    include_origin: bool = True,
    include_xml_signature: bool = True,
    include_origin_part: bool = True,
    include_xml_signature_part: bool = True,
    include_certificate_part: bool = False,
    include_root_relationship: bool = True,
    include_signature_relationship: bool = True,
    include_certificate_relationship: bool = True,
    include_origin_content_type: bool = True,
    include_xml_signature_content_type: bool = True,
    include_certificate_content_type: bool = True,
    origin_content_type: str = _PACKAGE_DIGITAL_SIGNATURE_ORIGIN_CONTENT_TYPE,
    xml_signature_content_type: str = _PACKAGE_DIGITAL_SIGNATURE_XML_CONTENT_TYPE,
    certificate_content_type: str = (
        _PACKAGE_DIGITAL_SIGNATURE_CERTIFICATE_CONTENT_TYPE
    ),
    origin_part_name: str = "_xmlsignatures/origin.sigs",
    signature_part_name: str = "_xmlsignatures/sig1.xml",
    certificate_part_name: str = "_xmlsignatures/certificate1.cer",
    origin_relationship_source: str = "/",
    origin_target_mode: str = "Internal",
    signature_target_mode: str = "Internal",
    certificate_target_mode: str = "Internal",
    relationship_id_suffix: str = "1",
    duplicate_origin_relationship: bool = False,
    wrong_signature_root: bool = False,
    omit_signature_method: bool = False,
    signature_value: str = "PACKAGE_SIGNATURE_VALUE_DO_NOT_LEAK",
    inline_certificate_marker: str = "PACKAGE_INLINE_X509_DO_NOT_LEAK",
    signature_comment: str = "PACKAGE_SIGNATURE_COMMENT_DO_NOT_LEAK",
    certificate_payload_marker: str = "PACKAGE_CERTIFICATE_DO_NOT_LEAK",
) -> None:
    """Write a small structural OPC package-signature fixture.

    The fixture deliberately contains non-cryptographic marker values. DocFence
    validates only the bounded shape used to inventory signature material; it
    never treats this fixture as a verified signature.
    """

    def relationship_target(
        source_part: str, target_part: str, target_mode: str
    ) -> str:
        if target_mode != "Internal":
            return "https://example.invalid/PACKAGE_SIGNATURE_TARGET_DO_NOT_LEAK"
        if source_part == "/":
            return target_part
        return posixpath.relpath(target_part, start=source_part.rpartition("/")[0])

    def relationship_markup(
        relationship_id: str,
        relationship_type: str,
        source_part: str,
        target_part: str,
        target_mode: str = "Internal",
    ) -> str:
        target_mode_attribute = (
            "" if target_mode == "Internal" else f' TargetMode="{target_mode}"'
        )
        return (
            f'<Relationship Id="{relationship_id}" Type="{relationship_type}" '
            f'Target="{relationship_target(source_part, target_part, target_mode)}"'
            f"{target_mode_attribute}/>"
        )

    relationships_by_source: dict[str, list[str]] = {}

    def add_relationship(source_part: str, relationship: str) -> None:
        relationships_by_source.setdefault(source_part, []).append(relationship)

    if include_origin and include_root_relationship:
        add_relationship(
            origin_relationship_source,
            relationship_markup(
                f"rIdOrigin{relationship_id_suffix}",
                _PACKAGE_DIGITAL_SIGNATURE_ORIGIN_RELATIONSHIP_TYPE,
                origin_relationship_source,
                origin_part_name,
                origin_target_mode,
            ),
        )
        if duplicate_origin_relationship:
            add_relationship(
                origin_relationship_source,
                relationship_markup(
                    f"rIdOriginDuplicate{relationship_id_suffix}",
                    _PACKAGE_DIGITAL_SIGNATURE_ORIGIN_RELATIONSHIP_TYPE,
                    origin_relationship_source,
                    origin_part_name,
                    origin_target_mode,
                ),
            )
    if include_origin and include_xml_signature and include_signature_relationship:
        add_relationship(
            origin_part_name,
            relationship_markup(
                f"rIdSignature{relationship_id_suffix}",
                _PACKAGE_DIGITAL_SIGNATURE_RELATIONSHIP_TYPE,
                origin_part_name,
                signature_part_name,
                signature_target_mode,
            ),
        )
    if (
        include_xml_signature
        and include_certificate_part
        and include_certificate_relationship
    ):
        add_relationship(
            signature_part_name,
            relationship_markup(
                f"rIdCertificate{relationship_id_suffix}",
                _PACKAGE_DIGITAL_SIGNATURE_CERTIFICATE_RELATIONSHIP_TYPE,
                signature_part_name,
                certificate_part_name,
                certificate_target_mode,
            ),
        )

    signature_root_name = "NotSignature" if wrong_signature_root else "Signature"
    signature_method = (
        ""
        if omit_signature_method
        else (
            '<ds:SignatureMethod '
            'Algorithm="http://www.w3.org/2001/04/xmldsig-more#rsa-sha256"/>'
        )
    )
    signature_xml = (
        f'<ds:{signature_root_name} xmlns:ds="{_XMLDSIG_NAMESPACE}" '
        f'xmlns:opc="{_OPC_DIGITAL_SIGNATURE_NAMESPACE}" '
        'Id="idPackageSignature">'
        "<ds:SignedInfo>"
        '<ds:CanonicalizationMethod '
        'Algorithm="http://www.w3.org/TR/2001/REC-xml-c14n-20010315"/>'
        f"{signature_method}"
        '<ds:Reference URI="#idPackageObject">'
        '<ds:DigestMethod Algorithm="http://www.w3.org/2001/04/xmlenc#sha256"/>'
        "<ds:DigestValue>PACKAGE_DIGEST_DO_NOT_LEAK</ds:DigestValue>"
        "</ds:Reference>"
        "</ds:SignedInfo>"
        f"<ds:SignatureValue>{signature_value}</ds:SignatureValue>"
        "<ds:KeyInfo><ds:X509Data>"
        f"<ds:X509Certificate>{inline_certificate_marker}</ds:X509Certificate>"
        "</ds:X509Data></ds:KeyInfo>"
        '<ds:Object Id="idPackageObject"><ds:Manifest>'
        '<ds:Reference URI="/word/document.xml?ContentType='
        'application/vnd.openxmlformats-officedocument.wordprocessingml.'
        'document.main+xml">'
        '<ds:DigestMethod Algorithm="http://www.w3.org/2001/04/xmlenc#sha256"/>'
        "<ds:DigestValue>PACKAGE_MANIFEST_DIGEST_DO_NOT_LEAK</ds:DigestValue>"
        "</ds:Reference></ds:Manifest>"
        "<ds:SignatureProperties><ds:SignatureProperty "
        'Id="idSignatureDetails" Target="#idPackageSignature">'
        f"<opc:SignatureTime>{signature_comment}</opc:SignatureTime>"
        "</ds:SignatureProperty></ds:SignatureProperties>"
        '<opc:RelationshipReference SourceId="rIdDocument1"/>'
        f"</ds:Object></ds:{signature_root_name}>"
    ).encode()

    content_types: list[str] = [
        f'<Override PartName="/word/document.xml" ContentType="{DOCX_MAIN_TYPE}"/>'
    ]
    if include_origin and include_origin_content_type:
        content_types.append(
            f'<Default Extension="sigs" '
            f'ContentType="{origin_content_type}"/>'
        )
    if include_xml_signature and include_xml_signature_content_type:
        content_types.append(
            f'<Override PartName="/{signature_part_name}" '
            f'ContentType="{xml_signature_content_type}"/>'
        )
    if include_certificate_part and include_certificate_content_type:
        content_types.append(
            f'<Override PartName="/{certificate_part_name}" '
            f'ContentType="{certificate_content_type}"/>'
        )

    entries: dict[str, bytes] = {
        "[Content_Types].xml": (
            f'<Types xmlns="{CT}">{"".join(content_types)}</Types>'
        ).encode(),
        "word/document.xml": (
            f'<w:document xmlns:w="{W}"><w:body><w:p><w:r><w:t>'
            "VISIBLE_DO_NOT_LEAK"
            "</w:t></w:r></w:p><w:sectPr/></w:body></w:document>"
        ).encode(),
    }
    if include_origin and include_origin_part:
        entries[origin_part_name] = b""
    if include_xml_signature and include_xml_signature_part:
        entries[signature_part_name] = signature_xml
    if include_certificate_part:
        entries[certificate_part_name] = certificate_payload_marker.encode()

    for source_part, relationships in relationships_by_source.items():
        if source_part == "/":
            relationship_part_name = "_rels/.rels"
        else:
            directory, _, basename = source_part.rpartition("/")
            relationship_part_name = (
                f"{directory + '/' if directory else ''}_rels/{basename}.rels"
            )
        entries[relationship_part_name] = (
            f'<Relationships xmlns="{PR}">{"".join(relationships)}</Relationships>'
        ).encode()

    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for name, payload in entries.items():
            archive.writestr(name, payload)


def _write_word_protection_document(
    path,
    *,
    include_document_protection: bool = True,
    include_write_protection: bool = True,
    settings_part_name: str = "word/settings.xml",
    include_settings_relationship: bool = False,
    settings_relationship_type: str = _DOCUMENT_SETTINGS_RELATIONSHIP_TYPE,
    settings_relationship_source: str = "word/document.xml",
    settings_target_mode: str = "Internal",
    include_settings_part: bool = True,
    settings_root_name: str = "settings",
    settings_word_namespace: str = W,
    duplicate_document_protection: bool = False,
    duplicate_write_protection: bool = False,
    document_edit: str = "readOnly",
    document_enforcement: str = "1",
    document_formatting: str = "1",
    document_hash: str = "WORD_PROTECTION_HASH_DO_NOT_LEAK",
    document_salt: str = "WORD_PROTECTION_SALT_DO_NOT_LEAK",
    document_hash_value: str = "WORD_PROTECTION_HASH_VALUE_DO_NOT_LEAK",
    document_salt_value: str = "WORD_PROTECTION_SALT_VALUE_DO_NOT_LEAK",
    document_provider: str = "rsaAES",
    document_algorithm_name: str = "SHA-512",
    document_extra_attributes: str = "",
    document_extra_content: str = "",
    write_recommended: str = "1",
    write_hash: str = "WRITE_PROTECTION_HASH_DO_NOT_LEAK",
    write_salt: str = "WRITE_PROTECTION_SALT_DO_NOT_LEAK",
    settings_extra_children: str = "",
) -> None:
    """Write a structural Word protection fixture with opaque verifier markers."""

    document_protection = (
        f'<w:documentProtection w:edit="{document_edit}" '
        f'w:formatting="{document_formatting}" '
        f'w:enforcement="{document_enforcement}" '
        f'w:cryptProviderType="{document_provider}" '
        'w:cryptAlgorithmClass="hash" '
        'w:cryptAlgorithmType="typeAny" w:cryptAlgorithmSid="14" '
        'w:cryptSpinCount="100000" '
        f'w:hash="{document_hash}" w:salt="{document_salt}" '
        f'w:algorithmName="{document_algorithm_name}" '
        f'w:hashValue="{document_hash_value}" '
        f'w:saltValue="{document_salt_value}" w:spinCount="100000" '
        f"{document_extra_attributes}>{document_extra_content}</w:documentProtection>"
    )
    write_protection = (
        f'<w:writeProtection w:recommended="{write_recommended}" '
        f'w:hash="{write_hash}" w:salt="{write_salt}"/>'
    )
    settings_children = ""
    if include_document_protection:
        settings_children += document_protection
        if duplicate_document_protection:
            settings_children += document_protection
    if include_write_protection:
        settings_children += write_protection
        if duplicate_write_protection:
            settings_children += write_protection
    settings_children += settings_extra_children

    entries: dict[str, bytes] = {
        "[Content_Types].xml": (
            f'<Types xmlns="{CT}"><Default Extension="xml" '
            'ContentType="application/xml"/>'
            f'<Override PartName="/word/document.xml" '
            f'ContentType="{DOCX_MAIN_TYPE}"/></Types>'
        ).encode(),
        "word/document.xml": (
            f'<w:document xmlns:w="{W}"><w:body><w:p><w:r><w:t>'
            "VISIBLE_DO_NOT_LEAK"
            "</w:t></w:r></w:p><w:sectPr/></w:body></w:document>"
        ).encode(),
    }
    if include_settings_part:
        entries[settings_part_name] = (
            f'<w:{settings_root_name} xmlns:w="{settings_word_namespace}">'
            f"{settings_children}</w:{settings_root_name}>"
        ).encode()
    if settings_relationship_source == "word/glossary/document.xml":
        entries[settings_relationship_source] = (
            f'<w:glossaryDocument xmlns:w="{W}"><w:docParts/>'
            "</w:glossaryDocument>"
        ).encode()
    if include_settings_relationship:
        if settings_target_mode == "Internal":
            target = posixpath.relpath(
                settings_part_name,
                start=settings_relationship_source.rpartition("/")[0],
            )
            target_mode_attribute = ""
        else:
            target = "https://example.invalid/WORD_SETTINGS_TARGET_DO_NOT_LEAK"
            target_mode_attribute = f' TargetMode="{settings_target_mode}"'
        source_directory, _, source_basename = settings_relationship_source.rpartition(
            "/"
        )
        relationship_part_name = (
            f"{source_directory}/_rels/{source_basename}.rels"
        )
        entries[relationship_part_name] = (
            f'<Relationships xmlns="{PR}"><Relationship Id="rIdSettings1" '
            f'Type="{settings_relationship_type}" Target="{target}"'
            f"{target_mode_attribute}/></Relationships>"
        ).encode()

    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for name, payload in entries.items():
            archive.writestr(name, payload)


def _write_document_variable_field_document(
    path,
    *,
    body_markup: str = "",
    header_markup: str | None = None,
    settings_markup: str = "",
    glossary_markup: str | None = None,
    glossary_settings_markup: str | None = None,
    word_namespace: str = W,
) -> None:
    """Write a small package for DOCVARIABLE-field and Settings-state tests."""

    glossary_override = (
        '<Override PartName="/word/glossary/document.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.'
        'wordprocessingml.document.glossaryDocument+xml"/>'
        if glossary_markup is not None
        else ""
    )
    entries: dict[str, bytes] = {
        "[Content_Types].xml": (
            f'<Types xmlns="{CT}"><Default Extension="xml" '
            'ContentType="application/xml"/>'
            f'<Override PartName="/word/document.xml" '
            f'ContentType="{DOCX_MAIN_TYPE}"/>'
            f"{glossary_override}</Types>"
        ).encode(),
        "word/document.xml": (
            f'<w:document xmlns:w="{word_namespace}"><w:body><w:p>'
            f"{body_markup}"
            "<w:r><w:t>VISIBLE_DO_NOT_LEAK</w:t></w:r>"
            "</w:p><w:sectPr/></w:body></w:document>"
        ).encode(),
        "word/settings.xml": (
            f'<w:settings xmlns:w="{word_namespace}">'
            f"{settings_markup}</w:settings>"
        ).encode(),
    }
    if header_markup is not None:
        entries["word/header1.xml"] = (
            f'<w:hdr xmlns:w="{word_namespace}"><w:p>{header_markup}'
            "<w:r><w:t>HEADER_VISIBLE_DO_NOT_LEAK</w:t></w:r>"
            "</w:p></w:hdr>"
        ).encode()
    if glossary_markup is not None:
        entries["word/glossary/document.xml"] = (
            f'<w:glossaryDocument xmlns:w="{word_namespace}"><w:docParts>'
            "<w:docPart><w:docPartBody><w:p>"
            f"{glossary_markup}"
            "</w:p></w:docPartBody></w:docPart></w:docParts>"
            "</w:glossaryDocument>"
        ).encode()
    if glossary_settings_markup is not None:
        entries["word/glossary/_rels/document.xml.rels"] = (
            f'<Relationships xmlns="{PR}"><Relationship Id="rIdSettings" '
            f'Type="{_DOCUMENT_SETTINGS_RELATIONSHIP_TYPE}" '
            'Target="glossarySettings.xml"/></Relationships>'
        ).encode()
        entries["word/glossary/glossarySettings.xml"] = (
            f'<w:settings xmlns:w="{word_namespace}">'
            f"{glossary_settings_markup}</w:settings>"
        ).encode()

    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for name, payload in entries.items():
            archive.writestr(name, payload)


def _write_hyperlink_field_document(
    path,
    *,
    include_fields: bool = True,
    literal_destination: str = "https://DESTINATION_DO_NOT_LEAK.invalid/path",
    split_complex_instructions: bool = False,
    word_namespace: str = W,
) -> None:
    """Write a small package spanning supported HYPERLINK field encodings."""

    def simple_field(instruction: str) -> str:
        escaped_instruction = (
            instruction.replace("&", "&amp;")
            .replace('"', "&quot;")
            .replace("<", "&lt;")
        )
        return (
            f'<w:fldSimple w:instr="{escaped_instruction}">'
            "<w:r><w:t>FIELD_RESULT_DO_NOT_LEAK</w:t></w:r></w:fldSimple>"
        )

    def instruction_text(text: str) -> str:
        return (
            '<w:r><w:instrText xml:space="preserve">'
            f"{text}</w:instrText></w:r>"
        )

    def field_char(field_type: str) -> str:
        return f'<w:r><w:fldChar w:fldCharType="{field_type}"/></w:r>'

    def complex_field(instruction: str) -> str:
        chunks = [instruction]
        if split_complex_instructions:
            split_at = len(instruction) // 2
            chunks = [instruction[:split_at], instruction[split_at:]]
        return "".join(
            (
                field_char("begin"),
                *(instruction_text(chunk) for chunk in chunks),
                field_char("separate"),
                "<w:r><w:t>FIELD_RESULT_DO_NOT_LEAK</w:t></w:r>",
                field_char("end"),
            )
        )

    nested_dynamic = "".join(
        (
            field_char("begin"),
            instruction_text(" HYPERLINK "),
            field_char("begin"),
            instruction_text(" MERGEFIELD DYNAMIC_HYPERLINK_DO_NOT_LEAK "),
            field_char("end"),
            instruction_text(" "),
            field_char("end"),
        )
    )
    fields = ""
    if include_fields:
        fields = "".join(
            (
                simple_field(
                    f'HYPERLINK "{literal_destination}" '
                    '\\o "TOOLTIP_DO_NOT_LEAK"'
                ),
                complex_field(
                    ' hyperlink SECOND_DESTINATION_DO_NOT_LEAK '
                    '\\t "TARGET_FRAME_DO_NOT_LEAK"'
                ),
                simple_field(
                    'HYPERLINK \\l "INTERNAL_LOCATION_DO_NOT_LEAK" '
                    '\\o "TOOLTIP_DO_NOT_LEAK"'
                ),
                complex_field(
                    'HYPERLINK \\L SECOND_INTERNAL_LOCATION_DO_NOT_LEAK '
                    "\\* MERGEFORMAT"
                ),
                simple_field(
                    'HYPERLINK "COMPOUND_DESTINATION_DO_NOT_LEAK" '
                    "UNPARSED_ARGUMENT_DO_NOT_LEAK"
                ),
                nested_dynamic,
            )
        )
    ignored_instruction_markup = "".join(
        (
            instruction_text("HYPERLINK LOOSE_HYPERLINK_DO_NOT_LEAK"),
            field_char("begin"),
            instruction_text(" DATE "),
            field_char("separate"),
            instruction_text("HYPERLINK POST_SEPARATOR_HYPERLINK_DO_NOT_LEAK"),
            field_char("end"),
            field_char("begin"),
            instruction_text("HYPERLINK UNCLOSED_HYPERLINK_DO_NOT_LEAK"),
        )
    )
    header_field = (
        simple_field("HYPERLINK HEADER_HYPERLINK_DO_NOT_LEAK")
        if include_fields
        else ""
    )
    entries = {
        "[Content_Types].xml": (
            f'<Types xmlns="{CT}"><Default Extension="xml" '
            'ContentType="application/xml"/>'
            f'<Override PartName="/word/document.xml" '
            f'ContentType="{DOCX_MAIN_TYPE}"/></Types>'
        ).encode(),
        "word/document.xml": (
            f'<w:document xmlns:w="{word_namespace}"><w:body><w:p>'
            f"{fields}{ignored_instruction_markup}"
            "<w:r><w:t>VISIBLE_DO_NOT_LEAK</w:t></w:r>"
            "</w:p><w:sectPr/></w:body></w:document>"
        ).encode(),
        "word/header1.xml": (
            f'<w:hdr xmlns:w="{word_namespace}"><w:p>{header_field}'
            "<w:r><w:t>HEADER_VISIBLE_DO_NOT_LEAK</w:t></w:r>"
            "</w:p></w:hdr>"
        ).encode(),
    }
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for name, payload in entries.items():
            archive.writestr(name, payload)


def _write_hyperlink_markup_document(
    path,
    *,
    include_markup: bool = True,
    include_orphan_hyperlink_relationship: bool = False,
    external_target: str = "https://MARKUP_EXTERNAL_TARGET_DO_NOT_LEAK.invalid/path",
    external_relationship_id: str = "rIdExternal",
    shadowed_anchor: str = "MARKUP_SHADOWED_ANCHOR_DO_NOT_LEAK",
    word_namespace: str = W,
    relationship_attribute_namespace: str = R,
    relationship_namespace: str = PR,
    hyperlink_relationship_type: str = _HYPERLINK_RELATIONSHIP_TYPE,
) -> None:
    """Write direct WordprocessingML hyperlink elements across two stories."""

    body_markup = ""
    header_markup = ""
    relationship_entries = ""
    header_relationship_entries = ""
    if include_markup:
        body_markup = "".join(
            (
                f'<w:hyperlink r:id="{external_relationship_id}" '
                f'w:anchor="{shadowed_anchor}" '
                'w:docLocation="MARKUP_DOCUMENT_LOCATION_DO_NOT_LEAK" '
                'w:tooltip="MARKUP_TOOLTIP_DO_NOT_LEAK" '
                'w:tgtFrame="MARKUP_TARGET_FRAME_DO_NOT_LEAK" w:history="true">'
                "<w:r><w:t>MARKUP_DISPLAY_TEXT_DO_NOT_LEAK</w:t></w:r>"
                "</w:hyperlink>",
                '<w:hyperlink r:id="rIdInternal"><w:r><w:t>'
                "INTERNAL_DISPLAY_DO_NOT_LEAK</w:t></w:r></w:hyperlink>",
                '<w:hyperlink r:id="rIdUnsupported"><w:r><w:t>'
                "UNSUPPORTED_DISPLAY_DO_NOT_LEAK</w:t></w:r></w:hyperlink>",
                '<w:hyperlink w:anchor="MARKUP_ANCHOR_DO_NOT_LEAK"><w:r><w:t>'
                "ANCHOR_DISPLAY_DO_NOT_LEAK</w:t></w:r></w:hyperlink>",
                "<w:hyperlink><w:r><w:t>DEFAULT_DISPLAY_DO_NOT_LEAK</w:t></w:r>"
                "</w:hyperlink>",
            )
        )
        header_markup = (
            '<w:hyperlink r:id="rIdHeaderExternal"><w:r><w:t>'
            "MARKUP_HEADER_DISPLAY_DO_NOT_LEAK</w:t></w:r></w:hyperlink>"
        )
        relationship_entries = "".join(
            (
                f'<Relationship Id="{external_relationship_id}" '
                f'Type="{hyperlink_relationship_type}" Target="{external_target}" '
                'TargetMode="External"/>',
                '<Relationship Id="rIdInternal" '
                f'Type="{hyperlink_relationship_type}" '
                'Target="MARKUP_INTERNAL_TARGET_DO_NOT_LEAK.xml"/>',
                '<Relationship Id="rIdUnsupported" '
                'Type="urn:docfence:unsupported-relationship" '
                'Target="https://MARKUP_UNSUPPORTED_TARGET_DO_NOT_LEAK.invalid" '
                'TargetMode="External"/>',
            )
        )
        header_relationship_entries = (
            '<Relationship Id="rIdHeaderExternal" '
            f'Type="{hyperlink_relationship_type}" '
            'Target="https://MARKUP_HEADER_TARGET_DO_NOT_LEAK.invalid" '
            'TargetMode="External"/>'
        )
    elif include_orphan_hyperlink_relationship:
        relationship_entries = (
            '<Relationship Id="rIdOrphan" '
            f'Type="{hyperlink_relationship_type}" '
            'Target="https://MARKUP_ORPHAN_TARGET_DO_NOT_LEAK.invalid" '
            'TargetMode="External"/>'
        )

    entries: dict[str, bytes] = {
        "[Content_Types].xml": (
            f'<Types xmlns="{CT}"><Default Extension="xml" '
            'ContentType="application/xml"/>'
            f'<Override PartName="/word/document.xml" '
            f'ContentType="{DOCX_MAIN_TYPE}"/></Types>'
        ).encode(),
        "word/document.xml": (
            f'<w:document xmlns:w="{word_namespace}" '
            f'xmlns:r="{relationship_attribute_namespace}"><w:body><w:p>'
            f"{body_markup}"
            "<w:r><w:t>VISIBLE_DO_NOT_LEAK</w:t></w:r>"
            "</w:p><w:sectPr/></w:body></w:document>"
        ).encode(),
        "word/header1.xml": (
            f'<w:hdr xmlns:w="{word_namespace}" '
            f'xmlns:r="{relationship_attribute_namespace}"><w:p>{header_markup}'
            "<w:r><w:t>HEADER_VISIBLE_DO_NOT_LEAK</w:t></w:r>"
            "</w:p></w:hdr>"
        ).encode(),
    }
    if relationship_entries:
        entries["word/_rels/document.xml.rels"] = (
            f'<Relationships xmlns="{relationship_namespace}">'
            f"{relationship_entries}</Relationships>"
        ).encode()
    if header_relationship_entries:
        entries["word/_rels/header1.xml.rels"] = (
            f'<Relationships xmlns="{relationship_namespace}">'
            f"{header_relationship_entries}</Relationships>"
        ).encode()

    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for name, payload in entries.items():
            archive.writestr(name, payload)


def _write_drawing_hyperlink_document(
    path,
    *,
    include_markup: bool = True,
    include_orphan_hyperlink_relationship: bool = False,
    external_target: str = "https://DRAWING_EXTERNAL_TARGET_DO_NOT_LEAK.invalid/path",
    action: str = "DRAWING_ACTION_DO_NOT_LEAK",
    external_relationship_id: str = "rIdDrawingExternal",
    word_namespace: str = W,
    drawing_namespace: str = _DRAWING_NAMESPACE,
    relationship_attribute_namespace: str = R,
    relationship_namespace: str = PR,
    hyperlink_relationship_type: str = _HYPERLINK_RELATIONSHIP_TYPE,
) -> None:
    """Write DrawingML link-action markers across body and header stories."""

    body_markup = ""
    header_markup = ""
    relationship_entries = ""
    header_relationship_entries = ""
    if include_markup:
        body_markup = "".join(
            (
                "<w:drawing><wp:inline><wp:docPr id=\"1\" "
                "name=\"DRAWING_OBJECT_NAME_DO_NOT_LEAK\">"
                f'<a:hlinkClick r:id="{external_relationship_id}" '
                f'action="{action}" '
                'invalidUrl="DRAWING_INVALID_URL_DO_NOT_LEAK" '
                'tooltip="DRAWING_TOOLTIP_DO_NOT_LEAK" '
                'tgtFrame="DRAWING_TARGET_FRAME_DO_NOT_LEAK" '
                'history="true" highlightClick="false" endSnd="true"/>'
                "</wp:docPr><pic:pic><pic:nvPicPr><pic:cNvPr id=\"101\" "
                "name=\"DRAWING_PICTURE_OBJECT_DO_NOT_LEAK\">"
                f'<a:hlinkClick r:id="{external_relationship_id}"/>'
                "</pic:cNvPr></pic:nvPicPr></pic:pic>"
                "</wp:inline></w:drawing>",
                '<w:drawing><wp:inline><wp:docPr id="2" name="INTERNAL">'
                '<a:hlinkHover r:id="rIdDrawingInternal"/>'
                "</wp:docPr></wp:inline></w:drawing>",
                '<w:drawing><wp:inline><wp:docPr id="3" name="MOUSE_OVER">'
                '<a:hlinkMouseOver r:id="rIdDrawingMouseOver"/>'
                "</wp:docPr></wp:inline></w:drawing>",
                '<w:drawing><wp:inline><wp:docPr id="4" name="UNSUPPORTED">'
                '<a:hlinkClick r:id="rIdDrawingUnsupported"/>'
                "</wp:docPr></wp:inline></w:drawing>",
                '<w:drawing><wp:inline><wp:docPr id="5" name="MISSING">'
                "<a:hlinkHover/>"
                "</wp:docPr></wp:inline></w:drawing>",
            )
        )
        header_markup = (
            '<w:drawing><wp:inline><wp:docPr id="6" name="HEADER">'
            '<a:hlinkClick r:id="rIdDrawingHeaderExternal"/>'
            "</wp:docPr></wp:inline></w:drawing>"
        )
        relationship_entries = "".join(
            (
                f'<Relationship Id="{external_relationship_id}" '
                f'Type="{hyperlink_relationship_type}" Target="{external_target}" '
                'TargetMode="External"/>',
                '<Relationship Id="rIdDrawingInternal" '
                f'Type="{hyperlink_relationship_type}" '
                'Target="DRAWING_INTERNAL_TARGET_DO_NOT_LEAK.xml"/>',
                '<Relationship Id="rIdDrawingMouseOver" '
                f'Type="{hyperlink_relationship_type}" '
                'Target="https://DRAWING_MOUSE_OVER_TARGET_DO_NOT_LEAK.invalid" '
                'TargetMode="External"/>',
                '<Relationship Id="rIdDrawingUnsupported" '
                'Type="urn:docfence:unsupported-relationship" '
                'Target="https://DRAWING_UNSUPPORTED_TARGET_DO_NOT_LEAK.invalid" '
                'TargetMode="External"/>',
            )
        )
        header_relationship_entries = (
            '<Relationship Id="rIdDrawingHeaderExternal" '
            f'Type="{hyperlink_relationship_type}" '
            'Target="https://DRAWING_HEADER_TARGET_DO_NOT_LEAK.invalid" '
            'TargetMode="External"/>'
        )
    elif include_orphan_hyperlink_relationship:
        relationship_entries = (
            '<Relationship Id="rIdDrawingOrphan" '
            f'Type="{hyperlink_relationship_type}" '
            'Target="https://DRAWING_ORPHAN_TARGET_DO_NOT_LEAK.invalid" '
            'TargetMode="External"/>'
        )

    entries: dict[str, bytes] = {
        "[Content_Types].xml": (
            f'<Types xmlns="{CT}"><Default Extension="xml" '
            'ContentType="application/xml"/>'
            f'<Override PartName="/word/document.xml" '
            f'ContentType="{DOCX_MAIN_TYPE}"/></Types>'
        ).encode(),
        "word/document.xml": (
            f'<w:document xmlns:w="{word_namespace}" '
            f'xmlns:r="{relationship_attribute_namespace}" '
            f'xmlns:a="{drawing_namespace}" '
            f'xmlns:wp="{_WORDPROCESSING_DRAWING_NAMESPACE}" '
            f'xmlns:pic="{_PICTURE_NAMESPACE}"><w:body><w:p>'
            f"{body_markup}"
            "<w:r><w:t>VISIBLE_DO_NOT_LEAK</w:t></w:r>"
            "</w:p><w:sectPr/></w:body></w:document>"
        ).encode(),
        "word/header1.xml": (
            f'<w:hdr xmlns:w="{word_namespace}" '
            f'xmlns:r="{relationship_attribute_namespace}" '
            f'xmlns:a="{drawing_namespace}" '
            f'xmlns:wp="{_WORDPROCESSING_DRAWING_NAMESPACE}" '
            f'xmlns:pic="{_PICTURE_NAMESPACE}"><w:p>{header_markup}'
            "<w:r><w:t>HEADER_VISIBLE_DO_NOT_LEAK</w:t></w:r>"
            "</w:p></w:hdr>"
        ).encode(),
    }
    if relationship_entries:
        entries["word/_rels/document.xml.rels"] = (
            f'<Relationships xmlns="{relationship_namespace}">'
            f"{relationship_entries}</Relationships>"
        ).encode()
    if header_relationship_entries:
        entries["word/_rels/header1.xml.rels"] = (
            f'<Relationships xmlns="{relationship_namespace}">'
            f"{header_relationship_entries}</Relationships>"
        ).encode()

    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for name, payload in entries.items():
            archive.writestr(name, payload)


def _write_drawing_linked_picture_document(
    path,
    *,
    include_markup: bool = True,
    include_orphan_image_relationship: bool = False,
    external_target: str = (
        "file:///LINKED_PICTURE_EXTERNAL_TARGET_DO_NOT_LEAK.invalid/image.png"
    ),
    compression_state: str = "print",
    external_relationship_id: str = "rIdLinkedPictureExternal",
    word_namespace: str = W,
    drawing_namespace: str = _DRAWING_NAMESPACE,
    relationship_attribute_namespace: str = R,
    relationship_namespace: str = PR,
    image_relationship_type: str = _IMAGE_RELATIONSHIP_TYPE,
) -> None:
    """Write direct DrawingML linked-picture markers across body and header."""

    body_markup = ""
    header_markup = ""
    relationship_entries = ""
    header_relationship_entries = ""
    if include_markup:
        body_markup = "".join(
            (
                "<w:r><w:drawing><wp:inline><a:graphic><a:graphicData "
                'uri="LINKED_PICTURE_OBJECT_NAME_DO_NOT_LEAK"><pic:pic>'
                "<pic:blipFill>"
                f'<a:blip r:link="{external_relationship_id}" '
                f'cstate="{compression_state}"/>'
                "</pic:blipFill></pic:pic></a:graphicData></a:graphic>"
                "</wp:inline></w:drawing></w:r>",
                "<w:r><w:drawing><wp:inline><a:graphic><a:graphicData>"
                "<pic:pic><pic:blipFill>"
                '<a:blip r:link="rIdLinkedPictureInternal"/>'
                "</pic:blipFill></pic:pic></a:graphicData></a:graphic>"
                "</wp:inline></w:drawing></w:r>",
                "<w:r><w:drawing><wp:inline><a:graphic><a:graphicData>"
                "<pic:pic><pic:blipFill>"
                '<a:blip r:link="rIdLinkedPictureUnsupported"/>'
                "</pic:blipFill></pic:pic></a:graphicData></a:graphic>"
                "</wp:inline></w:drawing></w:r>",
                "<w:r><w:drawing><wp:inline><a:graphic><a:graphicData>"
                "<pic:pic><pic:blipFill>"
                '<a:blip r:link="rIdLinkedPictureDual" '
                'r:embed="rIdLinkedPictureEmbedded"/>'
                "</pic:blipFill></pic:pic></a:graphicData></a:graphic>"
                "</wp:inline></w:drawing></w:r>",
                "<w:r><w:drawing><wp:inline><a:graphic><a:graphicData>"
                "<pic:pic><pic:blipFill>"
                '<a:blip r:embed="rIdLinkedPictureEmbedded"/>'
                "</pic:blipFill></pic:pic></a:graphicData></a:graphic>"
                "</wp:inline></w:drawing></w:r>",
            )
        )
        header_markup = (
            "<w:r><w:drawing><wp:inline><a:graphic><a:graphicData>"
            "<pic:pic><pic:blipFill>"
            '<a:blip r:link="rIdLinkedPictureHeaderExternal"/>'
            "</pic:blipFill></pic:pic></a:graphicData></a:graphic>"
            "</wp:inline></w:drawing></w:r>"
        )
        relationship_entries = "".join(
            (
                f'<Relationship Id="{external_relationship_id}" '
                f'Type="{image_relationship_type}" Target="{external_target}" '
                'TargetMode="External"/>',
                '<Relationship Id="rIdLinkedPictureInternal" '
                f'Type="{image_relationship_type}" '
                'Target="media/LINKED_PICTURE_INTERNAL_TARGET_DO_NOT_LEAK.png"/>',
                '<Relationship Id="rIdLinkedPictureUnsupported" '
                'Type="urn:docfence:unsupported-relationship" '
                'Target="https://LINKED_PICTURE_UNSUPPORTED_TARGET_DO_NOT_LEAK.'
                'invalid/image.png" '
                'TargetMode="External"/>',
                '<Relationship Id="rIdLinkedPictureDual" '
                f'Type="{image_relationship_type}" '
                'Target="https://LINKED_PICTURE_DUAL_TARGET_DO_NOT_LEAK.'
                'invalid/image.png" '
                'TargetMode="External"/>',
                '<Relationship Id="rIdLinkedPictureEmbedded" '
                f'Type="{image_relationship_type}" '
                'Target="media/LINKED_PICTURE_EMBEDDED_TARGET_DO_NOT_LEAK.png"/>',
            )
        )
        header_relationship_entries = (
            '<Relationship Id="rIdLinkedPictureHeaderExternal" '
            f'Type="{image_relationship_type}" '
            'Target="https://LINKED_PICTURE_HEADER_TARGET_DO_NOT_LEAK.'
            'invalid/image.png" '
            'TargetMode="External"/>'
        )
    elif include_orphan_image_relationship:
        relationship_entries = (
            '<Relationship Id="rIdLinkedPictureOrphan" '
            f'Type="{image_relationship_type}" '
            'Target="https://LINKED_PICTURE_ORPHAN_TARGET_DO_NOT_LEAK.'
            'invalid/image.png" '
            'TargetMode="External"/>'
        )

    entries: dict[str, bytes] = {
        "[Content_Types].xml": (
            f'<Types xmlns="{CT}"><Default Extension="xml" '
            'ContentType="application/xml"/>'
            f'<Override PartName="/word/document.xml" '
            f'ContentType="{DOCX_MAIN_TYPE}"/></Types>'
        ).encode(),
        "word/document.xml": (
            f'<w:document xmlns:w="{word_namespace}" '
            f'xmlns:r="{relationship_attribute_namespace}" '
            f'xmlns:a="{drawing_namespace}" '
            f'xmlns:wp="{_WORDPROCESSING_DRAWING_NAMESPACE}" '
            f'xmlns:pic="{_PICTURE_NAMESPACE}"><w:body><w:p>'
            f"{body_markup}"
            "<w:r><w:t>VISIBLE_DO_NOT_LEAK</w:t></w:r>"
            "</w:p><w:sectPr/></w:body></w:document>"
        ).encode(),
        "word/header1.xml": (
            f'<w:hdr xmlns:w="{word_namespace}" '
            f'xmlns:r="{relationship_attribute_namespace}" '
            f'xmlns:a="{drawing_namespace}" '
            f'xmlns:wp="{_WORDPROCESSING_DRAWING_NAMESPACE}" '
            f'xmlns:pic="{_PICTURE_NAMESPACE}"><w:p>{header_markup}'
            "<w:r><w:t>HEADER_VISIBLE_DO_NOT_LEAK</w:t></w:r>"
            "</w:p></w:hdr>"
        ).encode(),
    }
    if relationship_entries:
        entries["word/_rels/document.xml.rels"] = (
            f'<Relationships xmlns="{relationship_namespace}">'
            f"{relationship_entries}</Relationships>"
        ).encode()
    if header_relationship_entries:
        entries["word/_rels/header1.xml.rels"] = (
            f'<Relationships xmlns="{relationship_namespace}">'
            f"{header_relationship_entries}</Relationships>"
        ).encode()

    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for name, payload in entries.items():
            archive.writestr(name, payload)


def _write_vml_external_image_document(
    path,
    *,
    include_markup: bool = True,
    include_orphan_image_relationship: bool = False,
    external_target: str = (
        "file:///VML_EXTERNAL_IMAGE_TARGET_DO_NOT_LEAK.invalid/image.png"
    ),
    primary_src: str = "VML_RAW_SRC_DO_NOT_LEAK.png",
    external_relationship_id: str = "rIdVmlExternalImage",
    word_namespace: str = W,
    relationship_attribute_namespace: str = R,
    relationship_namespace: str = PR,
    image_relationship_type: str = _IMAGE_RELATIONSHIP_TYPE,
) -> None:
    """Write direct VML image-data markers across body and header stories."""

    body_markup = ""
    header_markup = ""
    relationship_entries = ""
    header_relationship_entries = ""
    if include_markup:
        body_markup = "".join(
            (
                "<w:r><w:pict><v:shape>"
                f'<v:imagedata r:id="{external_relationship_id}" '
                f'src="{primary_src}" r:pict="rIdVmlPict" '
                'r:href="rIdVmlHref" '
                'o:relid="VML_OFFICE_RELID_DO_NOT_LEAK"/>'
                "</v:shape></w:pict></w:r>",
                "<w:r><w:pict><v:shape>"
                '<v:imagedata r:id="rIdVmlInternal"/>'
                "</v:shape></w:pict></w:r>",
                "<w:r><w:pict><v:shape>"
                '<v:imagedata r:id="rIdVmlUnsupported"/>'
                "</v:shape></w:pict></w:r>",
                "<w:r><w:pict><v:shape>"
                f'<v:imagedata r:id="{external_relationship_id}"/>'
                "</v:shape></w:pict></w:r>",
                "<w:r><w:pict><v:shape>"
                '<v:imagedata src="VML_PICT_RAW_SRC_DO_NOT_LEAK.png" '
                'r:pict="rIdVmlPict" r:href="rIdVmlHref"/>'
                "</v:shape></w:pict></w:r>",
            )
        )
        header_markup = (
            "<w:r><w:pict><v:shape>"
            '<v:imagedata r:id="rIdVmlHeaderExternal"/>'
            "</v:shape></w:pict></w:r>"
        )
        relationship_entries = "".join(
            (
                f'<Relationship Id="{external_relationship_id}" '
                f'Type="{image_relationship_type}" Target="{external_target}" '
                'TargetMode="External"/>',
                '<Relationship Id="rIdVmlInternal" '
                f'Type="{image_relationship_type}" '
                'Target="media/VML_INTERNAL_IMAGE_TARGET_DO_NOT_LEAK.png"/>',
                '<Relationship Id="rIdVmlUnsupported" '
                'Type="urn:docfence:unsupported-relationship" '
                'Target="https://VML_UNSUPPORTED_IMAGE_TARGET_DO_NOT_LEAK.'
                'invalid/image.png" TargetMode="External"/>',
                '<Relationship Id="rIdVmlPict" '
                f'Type="{image_relationship_type}" '
                'Target="https://VML_PICT_IMAGE_TARGET_DO_NOT_LEAK.invalid/image.png" '
                'TargetMode="External"/>',
                '<Relationship Id="rIdVmlHref" '
                f'Type="{_HYPERLINK_RELATIONSHIP_TYPE}" '
                'Target="https://VML_HREF_TARGET_DO_NOT_LEAK.invalid/path" '
                'TargetMode="External"/>',
            )
        )
        header_relationship_entries = (
            '<Relationship Id="rIdVmlHeaderExternal" '
            f'Type="{image_relationship_type}" '
            'Target="https://VML_HEADER_IMAGE_TARGET_DO_NOT_LEAK.invalid/image.png" '
            'TargetMode="External"/>'
        )
    elif include_orphan_image_relationship:
        relationship_entries = (
            '<Relationship Id="rIdVmlOrphan" '
            f'Type="{image_relationship_type}" '
            'Target="https://VML_ORPHAN_IMAGE_TARGET_DO_NOT_LEAK.invalid/image.png" '
            'TargetMode="External"/>'
        )

    entries: dict[str, bytes] = {
        "[Content_Types].xml": (
            f'<Types xmlns="{CT}"><Default Extension="xml" '
            'ContentType="application/xml"/>'
            f'<Override PartName="/word/document.xml" '
            f'ContentType="{DOCX_MAIN_TYPE}"/></Types>'
        ).encode(),
        "word/document.xml": (
            f'<w:document xmlns:w="{word_namespace}" '
            f'xmlns:r="{relationship_attribute_namespace}" '
            f'xmlns:v="{_VML_NAMESPACE}" '
            'xmlns:o="urn:schemas-microsoft-com:office:office">'
            f"<w:body><w:p>{body_markup}"
            "<w:r><w:t>VISIBLE_DO_NOT_LEAK</w:t></w:r>"
            "</w:p><w:sectPr/></w:body></w:document>"
        ).encode(),
        "word/header1.xml": (
            f'<w:hdr xmlns:w="{word_namespace}" '
            f'xmlns:r="{relationship_attribute_namespace}" '
            f'xmlns:v="{_VML_NAMESPACE}"><w:p>{header_markup}'
            "<w:r><w:t>HEADER_VISIBLE_DO_NOT_LEAK</w:t></w:r>"
            "</w:p></w:hdr>"
        ).encode(),
    }
    if relationship_entries:
        entries["word/_rels/document.xml.rels"] = (
            f'<Relationships xmlns="{relationship_namespace}">'
            f"{relationship_entries}</Relationships>"
        ).encode()
    if header_relationship_entries:
        entries["word/_rels/header1.xml.rels"] = (
            f'<Relationships xmlns="{relationship_namespace}">'
            f"{header_relationship_entries}</Relationships>"
        ).encode()

    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for name, payload in entries.items():
            archive.writestr(name, payload)


def _write_vml_hyperlink_document(
    path,
    *,
    include_markup: bool = True,
    primary_href: str = "https://VML_PRIMARY_HREF_DO_NOT_LEAK.invalid/path",
    primary_target: str = "VML_PRIMARY_TARGET_FRAME_DO_NOT_LEAK",
) -> None:
    """Write direct VML shape-link markers across body and header stories."""

    body_markup = ""
    header_markup = ""
    if include_markup:
        body_markup = "".join(
            (
                "<w:r><w:pict>"
                '<v:shape id="VML_PRIMARY_SHAPE_ID_DO_NOT_LEAK" '
                f'href="{primary_href}" target="{primary_target}" '
                'title="VML_PRIMARY_TITLE_DO_NOT_LEAK" '
                'alt="VML_PRIMARY_ALT_DO_NOT_LEAK"/>'
                "</w:pict></w:r>",
                "<w:r><w:pict>"
                '<v:roundrect href="https://VML_ROUNDRECT_HREF_DO_NOT_LEAK.invalid"/>'
                "</w:pict></w:r>",
                "<w:r><w:pict>"
                '<v:rect href="https://VML_RECT_HREF_DO_NOT_LEAK.invalid"/>'
                "</w:pict></w:r>",
                "<w:r><w:pict>"
                '<v:group href="https://VML_GROUP_HREF_DO_NOT_LEAK.invalid" '
                'target="VML_GROUP_TARGET_FRAME_DO_NOT_LEAK">'
                '<v:oval href="https://VML_OVAL_HREF_DO_NOT_LEAK.invalid"/>'
                "</v:group></w:pict></w:r>",
                "<w:r><w:pict>"
                '<v:arc href=""/>'
                "</w:pict></w:r>",
                "<w:r><w:pict>"
                '<v:curve href="https://VML_CURVE_HREF_DO_NOT_LEAK.invalid"/>'
                "</w:pict></w:r>",
                "<w:r><w:pict>"
                '<v:polyline href="https://VML_POLYLINE_HREF_DO_NOT_LEAK.invalid"/>'
                "</w:pict></w:r>",
                "<w:r><w:pict>"
                '<v:image href="https://VML_IMAGE_HREF_DO_NOT_LEAK.invalid"/>'
                "</w:pict></w:r>",
                "<w:r><w:pict>"
                '<v:shapetype href="https://VML_SHAPETYPE_HREF_DO_NOT_LEAK.invalid"/>'
                "</w:pict></w:r>",
            )
        )
        header_markup = (
            "<w:r><w:pict>"
            '<v:line href="https://VML_HEADER_LINE_HREF_DO_NOT_LEAK.invalid"/>'
            "</w:pict></w:r>"
        )

    entries: dict[str, bytes] = {
        "[Content_Types].xml": (
            f'<Types xmlns="{CT}"><Default Extension="xml" '
            'ContentType="application/xml"/>'
            f'<Override PartName="/word/document.xml" '
            f'ContentType="{DOCX_MAIN_TYPE}"/></Types>'
        ).encode(),
        "word/document.xml": (
            f'<w:document xmlns:w="{W}" xmlns:v="{_VML_NAMESPACE}">'
            f"<w:body><w:p>{body_markup}"
            "<w:r><w:t>VISIBLE_DO_NOT_LEAK</w:t></w:r>"
            "</w:p><w:sectPr/></w:body></w:document>"
        ).encode(),
        "word/header1.xml": (
            f'<w:hdr xmlns:w="{W}" xmlns:v="{_VML_NAMESPACE}">'
            f"<w:p>{header_markup}"
            "<w:r><w:t>HEADER_VISIBLE_DO_NOT_LEAK</w:t></w:r>"
            "</w:p></w:hdr>"
        ).encode(),
    }
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for name, payload in entries.items():
            archive.writestr(name, payload)


def _write_word_permission_range_document(
    path,
    *,
    body_markup: str = "",
    header_markup: str | None = None,
    word_namespace: str = W,
) -> None:
    """Write a small story fixture containing optional editable-range markup."""

    entries: dict[str, bytes] = {
        "[Content_Types].xml": (
            f'<Types xmlns="{CT}"><Default Extension="xml" '
            'ContentType="application/xml"/>'
            f'<Override PartName="/word/document.xml" '
            f'ContentType="{DOCX_MAIN_TYPE}"/></Types>'
        ).encode(),
        "word/document.xml": (
            f'<w:document xmlns:w="{word_namespace}"><w:body><w:p>'
            f"{body_markup}"
            "<w:r><w:t>VISIBLE_DO_NOT_LEAK</w:t></w:r>"
            "</w:p><w:sectPr/></w:body></w:document>"
        ).encode(),
    }
    if header_markup is not None:
        entries["word/header1.xml"] = (
            f'<w:hdr xmlns:w="{word_namespace}"><w:p>{header_markup}'
            "<w:r><w:t>HEADER_VISIBLE_DO_NOT_LEAK</w:t></w:r>"
            "</w:p></w:hdr>"
        ).encode()

    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for name, payload in entries.items():
            archive.writestr(name, payload)


def _write_document_task_and_web_extension_document(
    path,
    *,
    include_document_tasks: bool = True,
    include_taskpane_web_extensions: bool = True,
    task_marker: str = "DOCUMENT_TASK_WORKFLOW",
    web_extension_marker: str = "TASKPANE_WEB_EXTENSION",
    relationship_id_suffix: str = "1",
    part_names: tuple[str, str, str, str] | None = None,
    include_feature_relationships: bool = True,
    include_feature_content_types: bool = True,
    include_taskpane_references: bool | None = None,
    document_task_target_mode: str = "Internal",
    taskpane_target_mode: str = "Internal",
    web_extension_target_mode: str = "Internal",
    wrong_document_task_root: bool = False,
    wrong_taskpane_root: bool = False,
    wrong_web_extension_root: bool = False,
    invalid_taskpane_reference: bool = False,
    web_extension_created_value: str = "false",
) -> None:
    """Write a small package with document-task and task-pane add-in state."""

    (
        document_task_part_name,
        taskpane_part_name,
        primary_web_extension_part_name,
        secondary_web_extension_part_name,
    ) = part_names or (
        "word/tasks.xml",
        "word/webextensions/taskpanes.xml",
        "word/webextensions/webextension1.xml",
        "word/webextensions/webextension2.xml",
    )
    if include_taskpane_references is None:
        include_taskpane_references = include_feature_relationships

    def relationship_target(
        source_part: str, target_part: str, target_mode: str, family: str
    ) -> str:
        if target_mode == "Internal":
            return posixpath.relpath(target_part, start=source_part.rpartition("/")[0])
        return f"https://example.invalid/{family}"

    def relationship_markup(
        relationship_id: str,
        relationship_type: str,
        source_part: str,
        target_part: str,
        target_mode: str,
        family: str,
    ) -> str:
        target = relationship_target(source_part, target_part, target_mode, family)
        target_mode_attribute = (
            "" if target_mode == "Internal" else f' TargetMode="{target_mode}"'
        )
        return (
            f'<Relationship Id="{relationship_id}" Type="{relationship_type}" '
            f'Target="{target}"{target_mode_attribute}/>'
        )

    document_relationships: list[str] = []
    if include_feature_relationships and include_document_tasks:
        document_relationships.append(
            relationship_markup(
                f"rIdDocumentTasks{relationship_id_suffix}",
                _DOCUMENT_TASK_RELATIONSHIP_TYPE,
                "word/document.xml",
                document_task_part_name,
                document_task_target_mode,
                "document-tasks",
            )
        )
    if include_feature_relationships and include_taskpane_web_extensions:
        document_relationships.append(
            relationship_markup(
                f"rIdTaskpanes{relationship_id_suffix}",
                _TASKPANE_WEB_EXTENSION_RELATIONSHIP_TYPE,
                "word/document.xml",
                taskpane_part_name,
                taskpane_target_mode,
                "taskpanes",
            )
        )

    primary_web_extension_relationship_id = (
        f"rIdWebExtensionPrimary{relationship_id_suffix}"
    )
    secondary_web_extension_relationship_id = (
        f"rIdWebExtensionSecondary{relationship_id_suffix}"
    )
    taskpane_relationships: list[str] = []
    if include_feature_relationships and include_taskpane_web_extensions:
        taskpane_relationships.extend(
            (
                relationship_markup(
                    primary_web_extension_relationship_id,
                    _WEB_EXTENSION_RELATIONSHIP_TYPE,
                    taskpane_part_name,
                    primary_web_extension_part_name,
                    web_extension_target_mode,
                    "webextension-primary",
                ),
                relationship_markup(
                    secondary_web_extension_relationship_id,
                    _WEB_EXTENSION_RELATIONSHIP_TYPE,
                    taskpane_part_name,
                    secondary_web_extension_part_name,
                    web_extension_target_mode,
                    "webextension-secondary",
                ),
            )
        )

    feature_overrides: list[str] = []
    if include_feature_content_types and include_document_tasks:
        feature_overrides.append(
            f'<Override PartName="/{document_task_part_name}" '
            f'ContentType="{_DOCUMENT_TASK_CONTENT_TYPE}"/>'
        )
    if include_feature_content_types and include_taskpane_web_extensions:
        feature_overrides.extend(
            (
                f'<Override PartName="/{taskpane_part_name}" '
                f'ContentType="{_TASKPANE_WEB_EXTENSION_CONTENT_TYPE}"/>',
                f'<Override PartName="/{primary_web_extension_part_name}" '
                f'ContentType="{_WEB_EXTENSION_CONTENT_TYPE}"/>',
                f'<Override PartName="/{secondary_web_extension_part_name}" '
                f'ContentType="{_WEB_EXTENSION_CONTENT_TYPE}"/>',
            )
        )

    content_controls = ""
    if include_taskpane_web_extensions:
        content_controls = f'''<w:sdt><w:sdtPr><w:id w:val="101"/>
  <w15:webExtensionLinked/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>
  LINKED_CONTROL_DO_NOT_LEAK</w:t></w:r></w:p></w:sdtContent></w:sdt>
<w:sdt><w:sdtPr><w:id w:val="102"/>
  <w15:webExtensionCreated w:val="{web_extension_created_value}"/>
  <w15:webExtensionLinked/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>
  CREATED_PRECEDENCE_CONTROL_DO_NOT_LEAK</w:t></w:r></w:p></w:sdtContent></w:sdt>
<w:sdt><w:sdtPr><w:id w:val="103"/>
  <w15:webExtensionCreated/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>
  CREATED_CONTROL_DO_NOT_LEAK</w:t></w:r></w:p></w:sdtContent></w:sdt>'''

    def task_user(element_name: str, role: str) -> str:
        return (
            f'<t:{element_name}\n'
            f' userId="{task_marker}_{role}_DO_NOT_LEAK"\n'
            f' userName="{task_marker}_{role}_NAME_DO_NOT_LEAK"\n'
            f' userProvider="{task_marker}_PROVIDER_DO_NOT_LEAK"/>'
        )

    def task_event(
        number: int,
        action: str,
        *,
        anchor: bool = False,
    ) -> str:
        children = [task_user("Attribution", "USER")]
        if anchor:
            children.append(
                f'<t:Anchor><t:Comment\n'
                f' id="{task_marker}_EVENT_COMMENT_DO_NOT_LEAK"/>'
                "</t:Anchor>"
            )
        children.append(action)
        return (
            f'<t:Event\n'
            f' id="{{00000000-0000-0000-0000-{number:012d}}}"\n'
            f' time="2026-08-01T{number - 1:02d}:00:00Z">'
            f'{"".join(children)}</t:Event>'
        )

    task_events = "".join(
        (
            task_event(1, "<t:Create/>", anchor=True),
            task_event(2, task_user("Assign", "ASSIGNEE")),
            task_event(3, task_user("Unassign", "ASSIGNEE")),
            task_event(
                4,
                f'<t:SetTitle\n title="{task_marker}_TITLE_DO_NOT_LEAK"/>',
            ),
            task_event(
                5,
                '<t:Schedule\n startDate="2026-08-02T00:00:00Z"\n'
                ' dueDate="2026-08-03T00:00:00Z"/>',
            ),
            task_event(6, '<t:Progress\n percentComplete="50"/>'),
            task_event(7, '<t:Priority\n value="5"/>'),
            task_event(8, "<t:Delete/>"),
            task_event(9, "<t:Undelete/>"),
            task_event(10, "<t:UnassignAll/>"),
            task_event(
                11,
                '<t:Undo\n id="{00000000-0000-0000-0000-000000000010}"/>',
            ),
        )
    )
    document_task_root_name = "notTasks" if wrong_document_task_root else "Tasks"
    document_task_xml = (
        f'<t:{document_task_root_name}\n'
        f' xmlns:t="{_DOCUMENT_TASK_NAMESPACE}">\n'
        '<t:Task id="{C1F1012D-3D7D-4C88-A44C-9DEB23456789}">\n'
        "<t:Anchor><t:Comment\n"
        f' id="{task_marker}_COMMENT_DO_NOT_LEAK"/></t:Anchor>\n'
        f"<t:History>{task_events}</t:History>\n"
        f"</t:Task></t:{document_task_root_name}>"
    )

    taskpane_root_name = "notTaskpanes" if wrong_taskpane_root else "taskpanes"
    primary_reference_id = (
        "rIdMissing"
        if invalid_taskpane_reference
        else primary_web_extension_relationship_id
    )
    taskpane_markup = ""
    if include_taskpane_references:
        taskpane_markup = (
            '<wetp:taskpane\n dockstate="right"\n visibility="1"\n'
            ' width="360"\n row="0"\n locked="true">\n'
            f'<wetp:webextensionref r:id="{primary_reference_id}"/>\n'
            "</wetp:taskpane>\n"
            '<wetp:taskpane\n dockstate="left"\n visibility="0"\n'
            ' width="240"\n row="1"\n locked="false">\n'
            f'<wetp:webextension r:id="{secondary_web_extension_relationship_id}"/>\n'
            "</wetp:taskpane>"
        )
    taskpane_xml = (
        f'<wetp:{taskpane_root_name}\n'
        f' xmlns:wetp="{_TASKPANE_WEB_EXTENSION_TASKPANES_NAMESPACE}"\n'
        f' xmlns:r="{R}">{taskpane_markup}</wetp:{taskpane_root_name}>'
    )

    def web_reference(identifier: str, version: str) -> str:
        return (
            "<we:reference\n"
            f' id="{web_extension_marker}_{identifier}_DO_NOT_LEAK"\n'
            f' version="{version}"\n'
            f' store="{web_extension_marker}_STORE_DO_NOT_LEAK"\n'
            ' storeType="omex"/>'
        )

    def web_property(name: str, value: str) -> str:
        return f'<we:property\n name="{name}"\n value="{value}"/>'

    def web_binding(identifier: str, binding_type: str) -> str:
        return (
            "<we:binding\n"
            f' id="{web_extension_marker}_{identifier}_DO_NOT_LEAK"\n'
            f' type="{binding_type}"\n'
            f' appref="{web_extension_marker}_APPREF_DO_NOT_LEAK"/>'
        )

    web_extension_root_name = (
        "notWebextension" if wrong_web_extension_root else "webextension"
    )
    primary_property_name = f"{web_extension_marker}_PROPERTY_DO_NOT_LEAK"
    primary_property_value = f"{web_extension_marker}_VALUE_DO_NOT_LEAK"
    secondary_property_name = f"{web_extension_marker}_SECOND_PROPERTY_DO_NOT_LEAK"
    secondary_property_value = f"{web_extension_marker}_SECOND_VALUE_DO_NOT_LEAK"
    primary_web_extension_xml = (
        f'<we:{web_extension_root_name}\n'
        f' xmlns:we="{_TASKPANE_WEB_EXTENSION_NAMESPACE}"\n'
        f' id="{web_extension_marker}_PRIMARY_DO_NOT_LEAK">\n'
        f"{web_reference('REFERENCE', '1.0')}\n"
        "<we:alternateReferences>\n"
        f"{web_reference('ALTERNATE', '1.1')}\n"
        "</we:alternateReferences>\n"
        "<we:properties>\n"
        f"{web_property('Office.AutoShowTaskpaneWithDocument', 'true')}\n"
        f"{web_property(primary_property_name, primary_property_value)}\n"
        "</we:properties>\n"
        "<we:bindings>\n"
        f"{web_binding('BINDING_A', 'text')}\n"
        f"{web_binding('BINDING_B', 'table')}\n"
        "</we:bindings>\n"
        f"</we:{web_extension_root_name}>"
    )
    secondary_web_extension_xml = (
        f'<we:{web_extension_root_name}\n'
        f' xmlns:we="{_TASKPANE_WEB_EXTENSION_NAMESPACE}"\n'
        f' id="{web_extension_marker}_SECONDARY_DO_NOT_LEAK">\n'
        f"{web_reference('SECOND_REFERENCE', '1.0')}\n"
        "<we:properties>\n"
        f"{web_property(secondary_property_name, secondary_property_value)}\n"
        "</we:properties>\n"
        "<we:bindings>\n"
        f"{web_binding('BINDING_C', 'matrix')}\n"
        "</we:bindings>\n"
        f"</we:{web_extension_root_name}>"
    )

    entries: dict[str, bytes] = {
        "[Content_Types].xml": (
            f'<Types xmlns="{CT}"><Default Extension="rels" '
            'ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            f'<Override PartName="/word/document.xml" ContentType="{DOCX_MAIN_TYPE}"/>'
            f'{"".join(feature_overrides)}</Types>'
        ).encode(),
        "word/document.xml": (
            f'<w:document xmlns:w="{W}" xmlns:w15="{_WORD_2012_NAMESPACE}"><w:body>'
            f"{content_controls}<w:p><w:r><w:t>BODY_DO_NOT_LEAK</w:t></w:r></w:p>"
            "<w:sectPr/></w:body></w:document>"
        ).encode(),
    }
    if document_relationships:
        entries["word/_rels/document.xml.rels"] = (
            f'<Relationships xmlns="{PR}">'
            f'{"".join(document_relationships)}'
            "</Relationships>"
        ).encode()
    if include_document_tasks:
        entries[document_task_part_name] = document_task_xml.encode()
    if include_taskpane_web_extensions:
        entries[taskpane_part_name] = taskpane_xml.encode()
        entries[primary_web_extension_part_name] = primary_web_extension_xml.encode()
        entries[secondary_web_extension_part_name] = (
            secondary_web_extension_xml.encode()
        )
    if taskpane_relationships:
        taskpane_directory, _, taskpane_filename = taskpane_part_name.rpartition("/")
        entries[
            f"{taskpane_directory}/_rels/{taskpane_filename}.rels"
        ] = (
            f'<Relationships xmlns="{PR}">'
            f'{"".join(taskpane_relationships)}'
            "</Relationships>"
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
