from __future__ import annotations

import zipfile
from pathlib import Path

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PR = "http://schemas.openxmlformats.org/package/2006/relationships"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"
HYPERLINK_TYPE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
)
DOCX_MAIN_TYPE = (
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"
)
STORY_PARTS = {
    "header": (
        "word/header1.xml",
        "hdr",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml",
    ),
    "footer": (
        "word/footer1.xml",
        "ftr",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml",
    ),
    "footnote": (
        "word/footnotes.xml",
        "footnotes",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml",
    ),
    "endnote": (
        "word/endnotes.xml",
        "endnotes",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml",
    ),
    "glossary": (
        "word/glossary/document.xml",
        "glossaryDocument",
        "application/vnd.openxmlformats-officedocument.wordprocessingml."
        "document.glossaryDocument+xml",
    ),
}


def write_document(
    path: Path,
    *,
    text: str = "Approved text",
    rsid: str = "00112233",
    relationship_id: str | None = None,
    relationship_target: str | None = None,
    relationship_external: bool = False,
    revision: bool = False,
    hidden: bool = False,
    field: bool = False,
    control: bool = False,
    comment: bool = False,
    track_revisions: bool = False,
    custom_xml: bytes | None = None,
    macro: bytes | None = None,
    unclassified: bytes | None = None,
    include_settings_rsids: bool = False,
    extra_stories: dict[str, str] | None = None,
) -> None:
    """Write a deliberately small, standards-shaped Word OOXML fixture."""

    paragraph = _paragraph(
        text,
        rsid=rsid,
        relationship_id=relationship_id,
        revision=revision,
        hidden=hidden,
        field=field,
        control=control,
        comment=comment,
    )
    entries: dict[str, bytes] = {
        "[Content_Types].xml": _content_types(
            comment=comment,
            macro=macro is not None,
            extra_story_kinds=set(extra_stories or {}),
        ).encode(),
        "word/document.xml": (
            f'<w:document xmlns:w="{W}" xmlns:r="{R}"><w:body>{paragraph}'
            "<w:sectPr/></w:body></w:document>"
        ).encode(),
    }
    if relationship_id is not None and relationship_target is not None:
        target_mode = ' TargetMode="External"' if relationship_external else ""
        entries["word/_rels/document.xml.rels"] = (
            f'<Relationships xmlns="{PR}"><Relationship Id="{relationship_id}" '
            f'Type="{HYPERLINK_TYPE}" '
            f'Target="{relationship_target}"{target_mode}/></Relationships>'
        ).encode()
    if comment:
        entries["word/comments.xml"] = (
            f'<w:comments xmlns:w="{W}"><w:comment w:id="0" '
            'w:author="REVIEWER_DO_NOT_LEAK"><w:p><w:r><w:t>'
            "COMMENT_DO_NOT_LEAK</w:t></w:r></w:p></w:comment></w:comments>"
        ).encode()
    if track_revisions or include_settings_rsids:
        track = "<w:trackRevisions/>" if track_revisions else ""
        rsids = (
            f'<w:rsids><w:rsid w:val="{rsid}"/></w:rsids>'
            if include_settings_rsids
            else ""
        )
        entries["word/settings.xml"] = (
            f'<w:settings xmlns:w="{W}">{track}{rsids}</w:settings>'
        ).encode()
    if custom_xml is not None:
        entries["customXml/item1.xml"] = custom_xml
    if macro is not None:
        entries["word/vbaProject.bin"] = macro
    if unclassified is not None:
        entries["word/styles.xml"] = unclassified
    for kind, story_text in (extra_stories or {}).items():
        part_name, root_name, _ = STORY_PARTS[kind]
        entries[part_name] = _story_xml(kind, root_name, story_text).encode()

    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for name, payload in entries.items():
            archive.writestr(name, payload)


def _content_types(*, comment: bool, macro: bool, extra_story_kinds: set[str]) -> str:
    main = (
        "application/vnd.ms-word.document.macroEnabled.main+xml"
        if macro
        else DOCX_MAIN_TYPE
    )
    comments = (
        '<Override PartName="/word/comments.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>'
        if comment
        else ""
    )
    extra_overrides = "".join(
        (
            f'<Override PartName="/{STORY_PARTS[kind][0]}" '
            f'ContentType="{STORY_PARTS[kind][2]}"/>'
        )
        for kind in sorted(extra_story_kinds)
    )
    return (
        f'<Types xmlns="{CT}"><Default Extension="rels" '
        'ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        f'<Override PartName="/word/document.xml" ContentType="{main}"/>'
        f"{comments}{extra_overrides}</Types>"
    )


def _paragraph(
    text: str,
    *,
    rsid: str,
    relationship_id: str | None,
    revision: bool,
    hidden: bool,
    field: bool,
    control: bool,
    comment: bool,
) -> str:
    run = f"<w:r><w:t>{text}</w:t></w:r>"
    if relationship_id is not None:
        run = f'<w:hyperlink r:id="{relationship_id}">{run}</w:hyperlink>'
    if revision:
        run = (
            '<w:ins w:id="7" w:author="REVIEWER_DO_NOT_LEAK" '
            f'w:date="2026-01-01T00:00:00Z">{run}</w:ins>'
        )
    if hidden:
        run += "<w:r><w:rPr><w:vanish/></w:rPr><w:t>HIDDEN_DO_NOT_LEAK</w:t></w:r>"
    if field:
        run += (
            '<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
            "<w:r><w:instrText>FIELD_DO_NOT_LEAK</w:instrText></w:r>"
        )
    if control:
        run = f"<w:sdt><w:sdtContent>{run}</w:sdtContent></w:sdt>"
    anchor = '<w:commentRangeStart w:id="0"/>' if comment else ""
    return f'<w:p w:rsidR="{rsid}">{anchor}{run}</w:p>'


def _story_xml(kind: str, root_name: str, text: str) -> str:
    paragraph = f"<w:p><w:r><w:t>{text}</w:t></w:r></w:p>"
    if kind == "footnote":
        body = f'<w:footnote w:id="1">{paragraph}</w:footnote>'
    elif kind == "endnote":
        body = f'<w:endnote w:id="1">{paragraph}</w:endnote>'
    elif kind == "glossary":
        body = (
            "<w:docParts><w:docPart><w:docPartBody>"
            f"{paragraph}</w:docPartBody></w:docPart></w:docParts>"
        )
    else:
        body = paragraph
    return f'<w:{root_name} xmlns:w="{W}">{body}</w:{root_name}>'
