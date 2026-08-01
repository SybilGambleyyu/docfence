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
