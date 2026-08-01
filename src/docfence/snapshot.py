"""Bounded, non-executing snapshots of Word OOXML packages.

The snapshot deliberately keeps document material transient.  Text, relationship
targets, reviewer metadata, field instructions, custom XML, and macro bytes are
only fed into private digests; public callers receive aggregate inventories.
"""

from __future__ import annotations

import hashlib
import io
import json
import posixpath
import stat
import unicodedata
import xml.etree.ElementTree as ET
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Final

from docfence.errors import DocumentFormatError, DocumentSafetyError
from docfence.models import (
    AlternativeFormatImportInventory,
    DocumentSnapshot,
    EmbeddedObjectInventory,
    RelationshipInventory,
    RevisionInventory,
    StorySnapshot,
    StyleInventory,
)


@dataclass(frozen=True)
class PackageLimits:
    """Hard limits applied before DocFence considers a package trustworthy."""

    max_source_bytes: int = 128 * 1024 * 1024
    max_member_count: int = 4_096
    max_member_expanded_bytes: int = 64 * 1024 * 1024
    max_total_expanded_bytes: int = 512 * 1024 * 1024
    max_compression_ratio: int = 1_000
    max_xml_bytes: int = 16 * 1024 * 1024
    max_xml_elements: int = 200_000
    max_xml_depth: int = 256


DEFAULT_LIMITS: Final = PackageLimits()

_WORD_NAMESPACES: Final = frozenset(
    {
        "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        "http://purl.oclc.org/ooxml/wordprocessingml/main",
    }
)
_REL_ATTRIBUTE_NAMESPACES: Final = frozenset(
    {
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        "http://purl.oclc.org/ooxml/officeDocument/relationships",
    }
)
_PACKAGE_REL_NAMESPACES: Final = frozenset(
    {
        "http://schemas.openxmlformats.org/package/2006/relationships",
        "http://purl.oclc.org/ooxml/package/relationships",
    }
)

_REVISION_TAGS: Final = frozenset({"ins", "del", "moveFrom", "moveTo"})
_REVISION_METADATA_ATTRIBUTES: Final = frozenset(
    {"author", "date", "dateUtc", "id", "initials", "userId"}
)
_VOLATILE_ATTRIBUTE_NAMES: Final = frozenset({"editId", "paraId", "rsid", "textId"})
_FALSE_VALUES: Final = frozenset({"0", "false", "no", "off"})
_STORY_ROOT_NAMES: Final = {
    "body": "document",
    "header": "hdr",
    "footer": "ftr",
    "footnote": "footnotes",
    "endnote": "endnotes",
    "comment": "comments",
    "glossary": "glossaryDocument",
}
_EMBEDDED_OBJECT_RELATIONSHIP_TYPES: Final = frozenset(
    {
        "http://schemas.openxmlformats.org/officedocument/2006/relationships/oleobject",
        "http://schemas.openxmlformats.org/officedocument/2006/relationships/package",
        "http://purl.oclc.org/ooxml/officedocument/relationships/oleobject",
        "http://purl.oclc.org/ooxml/officedocument/relationships/package",
    }
)
_EMBEDDED_CONTROL_RELATIONSHIP_TYPES: Final = frozenset(
    {
        "http://schemas.openxmlformats.org/officedocument/2006/relationships/control",
        "http://purl.oclc.org/ooxml/officedocument/relationships/control",
        "http://schemas.microsoft.com/office/2006/relationships/activexcontrolbinary",
    }
)
_ALTERNATIVE_FORMAT_IMPORT_RELATIONSHIP_TYPES: Final = frozenset(
    {
        "http://schemas.openxmlformats.org/officedocument/2006/relationships/afchunk",
        "http://purl.oclc.org/ooxml/officedocument/relationships/afchunk",
    }
)


@dataclass(frozen=True)
class _Relationship:
    relationship_type: str
    target: str
    target_mode: str

    def canonical_value(self) -> tuple[str, str, str]:
        return (self.relationship_type, self.target_mode.casefold(), self.target)


def load_snapshot(
    document: str | Path, *, limits: PackageLimits = DEFAULT_LIMITS
) -> DocumentSnapshot:
    """Return a bounded private snapshot of a `.docx` or `.docm` file.

    No document code, fields, links, or external relationships are executed or
    followed. Expected failure messages intentionally omit filenames and package
    contents so they remain safe to use in automation logs.
    """

    try:
        path = Path(document)
        suffix = path.suffix.casefold()
    except (TypeError, ValueError):
        raise DocumentFormatError("document path is invalid") from None
    if suffix not in {".docx", ".docm"}:
        raise DocumentFormatError("only DOCX and DOCM packages are supported")

    try:
        if path.is_symlink() or not path.is_file():
            raise DocumentFormatError("document must be a regular file")
        source_size = path.stat().st_size
        if source_size > limits.max_source_bytes:
            raise DocumentSafetyError("document exceeds the source size limit")
        source = path.read_bytes()
    except DocumentFormatError:
        raise
    except DocumentSafetyError:
        raise
    except (OSError, ValueError):
        raise DocumentFormatError("document cannot be read") from None

    return _load_package(source, suffix.removeprefix("."), limits)


def _load_package(
    source: bytes, document_format: str, limits: PackageLimits
) -> DocumentSnapshot:
    try:
        with zipfile.ZipFile(io.BytesIO(source)) as archive:
            members = _validate_members(archive.infolist(), limits)
            if (
                "[Content_Types].xml" not in members
                or "word/document.xml" not in members
            ):
                raise DocumentFormatError("package is not a Word OOXML document")

            content_types = _content_type_overrides(
                _read_xml(archive, members["[Content_Types].xml"], limits), members
            )
            _validate_main_document(content_types)

            relationships, relationship_maps = _relationship_inventory(
                archive, members, limits
            )
            embedded_objects, embedded_object_parts = _embedded_object_inventory(
                archive, members, relationship_maps, limits
            )
            alternative_format_imports, alternative_format_import_parts = (
                _alternative_format_import_inventory(
                    archive, members, relationship_maps, limits
                )
            )
            styles = _styles_inventory(archive, members, relationship_maps, limits)
            stories, comment_count = _story_snapshots(
                archive, members, content_types, relationship_maps, limits
            )
            settings_enabled, settings_signature = _settings_inventory(
                archive, members, relationship_maps, limits
            )
            custom_count, custom_signature = _custom_xml_inventory(
                archive, members, limits
            )
            macro_present, macro_signature = _macro_inventory(archive, members, limits)
            unclassified_count, unclassified_signature = _unclassified_inventory(
                archive,
                members,
                stories,
                embedded_object_parts | alternative_format_import_parts,
                limits,
            )
    except (DocumentFormatError, DocumentSafetyError):
        raise
    except (
        EOFError,
        NotImplementedError,
        OSError,
        RuntimeError,
        ValueError,
        zipfile.BadZipFile,
        zipfile.LargeZipFile,
    ):
        raise DocumentFormatError("document is not a readable OOXML package") from None

    return DocumentSnapshot(
        format=document_format,
        package_member_count=len(members),
        stories=stories,
        relationships=relationships,
        styles=styles,
        embedded_objects=embedded_objects,
        alternative_format_imports=alternative_format_imports,
        track_revisions_enabled=settings_enabled,
        comment_count=comment_count,
        custom_xml_part_count=custom_count,
        custom_xml_signature=custom_signature,
        macro_present=macro_present,
        macro_signature=macro_signature,
        settings_signature=settings_signature,
        unclassified_part_count=unclassified_count,
        unclassified_signature=unclassified_signature,
    )


def _validate_members(
    infos: list[zipfile.ZipInfo], limits: PackageLimits
) -> dict[str, zipfile.ZipInfo]:
    if len(infos) > limits.max_member_count:
        raise DocumentSafetyError("package exceeds the member count limit")

    members: dict[str, zipfile.ZipInfo] = {}
    directories: set[str] = set()
    normalized_names: set[str] = set()
    total_expanded = 0
    for info in infos:
        _validate_member_name(info.filename, normalized_names)
        if info.flag_bits & 0x1:
            raise DocumentSafetyError("encrypted package members are unsupported")
        if info.compress_type not in {zipfile.ZIP_STORED, zipfile.ZIP_DEFLATED}:
            raise DocumentSafetyError("package uses an unsupported compression method")
        mode = info.external_attr >> 16
        if stat.S_ISLNK(mode):
            raise DocumentSafetyError("symbolic-link package members are unsupported")
        if info.file_size > limits.max_member_expanded_bytes:
            raise DocumentSafetyError("package member exceeds the size limit")
        if info.file_size and (
            info.compress_size == 0
            or info.file_size / max(1, info.compress_size)
            > limits.max_compression_ratio
        ):
            raise DocumentSafetyError(
                "package member exceeds the compression ratio limit"
            )
        total_expanded += info.file_size
        if total_expanded > limits.max_total_expanded_bytes:
            raise DocumentSafetyError("package exceeds the expanded size limit")
        if info.is_dir():
            directories.add(info.filename.removesuffix("/"))
            continue
        if info.filename in members:
            raise DocumentFormatError("package contains duplicate members")
        members[info.filename] = info
    _validate_member_hierarchy(members, directories)
    return members


def _validate_member_hierarchy(
    members: dict[str, zipfile.ZipInfo], directories: set[str]
) -> None:
    for member_name in members:
        if member_name in directories:
            raise DocumentFormatError("package member names are unsafe")
        components = member_name.split("/")
        for index in range(1, len(components)):
            if "/".join(components[:index]) in members:
                raise DocumentFormatError("package member names are unsafe")


def _validate_member_name(name: str, normalized_names: set[str]) -> None:
    if not name or "\x00" in name or "\\" in name or name.startswith("/"):
        raise DocumentFormatError("package member names are unsafe")
    bare_name = name[:-1] if name.endswith("/") else name
    if not bare_name:
        raise DocumentFormatError("package member names are unsafe")
    parts = bare_name.split("/")
    if any(not part or part in {".", ".."} or ":" in part for part in parts):
        raise DocumentFormatError("package member names are unsafe")
    normalized = unicodedata.normalize("NFC", name).casefold()
    if normalized in normalized_names:
        raise DocumentFormatError("package member names collide")
    normalized_names.add(normalized)


def _read_member(
    archive: zipfile.ZipFile, info: zipfile.ZipInfo, limits: PackageLimits
) -> bytes:
    try:
        payload = archive.read(info)
    except (
        EOFError,
        NotImplementedError,
        OSError,
        RuntimeError,
        ValueError,
        zipfile.BadZipFile,
        zipfile.LargeZipFile,
    ):
        raise DocumentFormatError("package member cannot be read") from None
    if (
        len(payload) != info.file_size
        or len(payload) > limits.max_member_expanded_bytes
    ):
        raise DocumentSafetyError("package member crossed the size limit")
    return payload


def _read_xml(
    archive: zipfile.ZipFile, info: zipfile.ZipInfo, limits: PackageLimits
) -> ET.Element:
    payload = _read_member(archive, info, limits)
    return _parse_xml(payload, limits)


def _parse_xml(payload: bytes, limits: PackageLimits) -> ET.Element:
    if len(payload) > limits.max_xml_bytes:
        raise DocumentSafetyError("XML member exceeds the size limit")
    lowered = payload.lower()
    if b"<!doctype" in lowered or b"<!entity" in lowered:
        raise DocumentSafetyError("DTD and entity declarations are unsupported")

    depth = 0
    elements = 0
    root: ET.Element | None = None
    try:
        for event, element in ET.iterparse(
            io.BytesIO(payload), events=("start", "end")
        ):
            if event == "start":
                elements += 1
                depth += 1
                if elements > limits.max_xml_elements:
                    raise DocumentSafetyError("XML member exceeds the element limit")
                if depth > limits.max_xml_depth:
                    raise DocumentSafetyError("XML member exceeds the depth limit")
            else:
                depth -= 1
                root = element
    except DocumentSafetyError:
        raise
    except (ET.ParseError, UnicodeDecodeError, ValueError):
        raise DocumentFormatError("package contains unparseable XML") from None
    if root is None or depth != 0:
        raise DocumentFormatError("package contains unparseable XML")
    return root


def _content_type_overrides(
    root: ET.Element, members: dict[str, zipfile.ZipInfo]
) -> dict[str, str]:
    if _local_name(root.tag) != "Types":
        raise DocumentFormatError("package content types are invalid")
    overrides: dict[str, str] = {}
    for child in root:
        if _local_name(child.tag) != "Override":
            continue
        part_name = child.attrib.get("PartName")
        content_type = child.attrib.get("ContentType")
        if not part_name or not content_type or not part_name.startswith("/"):
            raise DocumentFormatError("package content types are invalid")
        member_name = part_name[1:]
        if member_name not in members or member_name in overrides:
            raise DocumentFormatError("package content types are invalid")
        overrides[member_name] = content_type
    return overrides


def _validate_main_document(content_types: dict[str, str]) -> None:
    content_type = content_types.get("word/document.xml", "").casefold()
    if not (
        content_type.endswith("wordprocessingml.document.main+xml")
        or content_type.endswith("word.document.macroenabled.main+xml")
    ):
        raise DocumentFormatError("package main document content type is unsupported")


def _relationship_inventory(
    archive: zipfile.ZipFile,
    members: dict[str, zipfile.ZipInfo],
    limits: PackageLimits,
) -> tuple[RelationshipInventory, dict[str, dict[str, _Relationship]]]:
    records: list[tuple[str, str, str, str]] = []
    external_records: list[tuple[str, str, str, str]] = []
    maps: dict[str, dict[str, _Relationship]] = {}
    relationship_members = sorted(name for name in members if name.endswith(".rels"))

    for member_name in relationship_members:
        source_part = _relationship_source_part(member_name)
        root = _read_xml(archive, members[member_name], limits)
        namespace, local_name = _qualified_name(root.tag)
        if local_name != "Relationships" or namespace not in _PACKAGE_REL_NAMESPACES:
            raise DocumentFormatError("package relationships are invalid")
        source_map: dict[str, _Relationship] = {}
        for child in root:
            namespace, local_name = _qualified_name(child.tag)
            if local_name != "Relationship" or namespace not in _PACKAGE_REL_NAMESPACES:
                raise DocumentFormatError("package relationships are invalid")
            relationship_id = child.attrib.get("Id")
            relationship_type = child.attrib.get("Type")
            target = child.attrib.get("Target")
            target_mode = child.attrib.get("TargetMode", "Internal")
            if (
                not relationship_id
                or not relationship_type
                or target is None
                or relationship_id in source_map
            ):
                raise DocumentFormatError("package relationships are invalid")
            relationship = _Relationship(relationship_type, target, target_mode)
            source_map[relationship_id] = relationship
            record = (source_part, *relationship.canonical_value())
            records.append(record)
            if target_mode.casefold() == "external":
                external_records.append(record)
        if source_part in maps:
            raise DocumentFormatError("package relationships are invalid")
        maps[source_part] = source_map

    return (
        RelationshipInventory(
            relationship_count=len(records),
            external_count=len(external_records),
            relationship_signature=_digest_records(records),
            external_signature=_digest_records(external_records),
        ),
        maps,
    )


def _embedded_object_inventory(
    archive: zipfile.ZipFile,
    members: dict[str, zipfile.ZipInfo],
    relationship_maps: dict[str, dict[str, _Relationship]],
    limits: PackageLimits,
) -> tuple[EmbeddedObjectInventory, frozenset[str]]:
    object_records: list[tuple[str, ...]] = []
    control_records: list[tuple[str, ...]] = []
    object_parts = set(_part_names_under(members, "word/embeddings/"))
    control_parts = set(_part_names_under(members, "word/activeX/"))
    object_targets: set[str] = set()

    for source_part, relationships in sorted(relationship_maps.items()):
        for relationship in relationships.values():
            relationship_type = relationship.relationship_type.casefold()
            if relationship_type in _EMBEDDED_OBJECT_RELATIONSHIP_TYPES:
                object_records.append(
                    ("embedded_object", source_part, *relationship.canonical_value())
                )
                target = _internal_relationship_target(
                    source_part, relationship, members
                )
                if target is not None:
                    object_targets.add(target)
                    object_parts.add(target)
            elif relationship_type in _EMBEDDED_CONTROL_RELATIONSHIP_TYPES:
                control_records.append(
                    ("embedded_control", source_part, *relationship.canonical_value())
                )
                target = _internal_relationship_target(
                    source_part, relationship, members
                )
                if target is not None:
                    control_parts.add(target)

    object_parts.difference_update(control_parts - object_targets)
    payload_parts = frozenset(object_parts | control_parts)
    return (
        EmbeddedObjectInventory(
            object_relationship_count=len(object_records),
            object_part_count=len(object_parts),
            control_relationship_count=len(control_records),
            control_part_count=len(control_parts),
            signature=_payload_inventory_signature(
                [*object_records, *control_records],
                payload_parts,
                archive,
                members,
                limits,
            ),
        ),
        payload_parts,
    )


def _alternative_format_import_inventory(
    archive: zipfile.ZipFile,
    members: dict[str, zipfile.ZipInfo],
    relationship_maps: dict[str, dict[str, _Relationship]],
    limits: PackageLimits,
) -> tuple[AlternativeFormatImportInventory, frozenset[str]]:
    records: list[tuple[str, ...]] = []
    payload_parts: set[str] = set()
    for source_part, relationships in sorted(relationship_maps.items()):
        for relationship in relationships.values():
            if (
                relationship.relationship_type.casefold()
                not in _ALTERNATIVE_FORMAT_IMPORT_RELATIONSHIP_TYPES
            ):
                continue
            records.append(
                (
                    "alternative_format_import",
                    source_part,
                    *relationship.canonical_value(),
                )
            )
            target = _internal_relationship_target(source_part, relationship, members)
            if target is not None:
                payload_parts.add(target)

    part_names = frozenset(payload_parts)
    return (
        AlternativeFormatImportInventory(
            relationship_count=len(records),
            payload_part_count=len(part_names),
            signature=_payload_inventory_signature(
                records, part_names, archive, members, limits
            ),
        ),
        part_names,
    )


def _internal_relationship_target(
    source_part: str,
    relationship: _Relationship,
    members: dict[str, zipfile.ZipInfo],
) -> str | None:
    target_mode = relationship.target_mode.casefold()
    if target_mode == "external":
        return None
    if target_mode != "internal":
        raise DocumentFormatError("package relationship target mode is invalid")
    return _resolve_internal_relationship_target(
        source_part, relationship.target, members
    )


def _resolve_internal_relationship_target(
    source_part: str,
    target: str,
    members: dict[str, zipfile.ZipInfo],
) -> str:
    if (
        not target
        or target.startswith("/")
        or any(token in target for token in ("\\", ":", "?", "#", "\x00"))
    ):
        raise DocumentFormatError("package relationship target is invalid")
    base = "" if source_part == "/" else source_part.rpartition("/")[0]
    resolved = posixpath.normpath(posixpath.join(base, target))
    if (
        resolved in {"", ".", ".."}
        or resolved.startswith("../")
        or resolved.startswith("/")
        or resolved not in members
    ):
        raise DocumentFormatError("package relationship target is unavailable")
    return resolved


def _part_names_under(members: dict[str, zipfile.ZipInfo], prefix: str) -> list[str]:
    return sorted(
        name
        for name in members
        if name.startswith(prefix)
        and "/_rels/" not in name
        and not name.endswith(".rels")
    )


def _payload_inventory_signature(
    relationship_records: list[tuple[str, ...]],
    part_names: frozenset[str],
    archive: zipfile.ZipFile,
    members: dict[str, zipfile.ZipInfo],
    limits: PackageLimits,
) -> str:
    records = [*relationship_records]
    records.extend(
        (
            "payload_part",
            part_name,
            _digest_bytes(_read_member(archive, members[part_name], limits)),
        )
        for part_name in sorted(part_names)
    )
    encoded = json.dumps(
        sorted(records), ensure_ascii=False, separators=(",", ":")
    ).encode("utf-8")
    return _digest_bytes(encoded)


def _relationship_source_part(member_name: str) -> str:
    if member_name == "_rels/.rels":
        return "/"
    parts = member_name.split("/")
    if len(parts) < 3 or parts[-2] != "_rels" or not parts[-1].endswith(".rels"):
        raise DocumentFormatError("package relationships are invalid")
    basename = parts[-1].removesuffix(".rels")
    if not basename:
        raise DocumentFormatError("package relationships are invalid")
    return "/".join([*parts[:-2], basename])


def _digest_records(records: list[tuple[str, str, str, str]]) -> str:
    encoded = json.dumps(
        sorted(records), ensure_ascii=False, separators=(",", ":")
    ).encode("utf-8")
    return hashlib.sha256(encoded).hexdigest()


def _story_snapshots(
    archive: zipfile.ZipFile,
    members: dict[str, zipfile.ZipInfo],
    content_types: dict[str, str],
    relationship_maps: dict[str, dict[str, _Relationship]],
    limits: PackageLimits,
) -> tuple[tuple[StorySnapshot, ...], int]:
    stories: list[StorySnapshot] = []
    comment_count = 0
    for part_key, kind in _discover_story_parts(members, content_types):
        root = _read_xml(archive, members[part_key], limits)
        story = _snapshot_story(
            root, part_key, kind, relationship_maps.get(part_key, {})
        )
        stories.append(story)
        if kind == "comment":
            comment_count += _count_word_elements(root, "comment")
    if not any(story.kind == "body" for story in stories):
        raise DocumentFormatError("package does not contain a document story")
    return tuple(stories), comment_count


def _discover_story_parts(
    members: dict[str, zipfile.ZipInfo], content_types: dict[str, str]
) -> list[tuple[str, str]]:
    candidates: dict[str, str] = {"word/document.xml": "body"}
    for part_key, content_type in content_types.items():
        kind = _story_kind_from_content_type(content_type)
        if kind is not None:
            candidates[part_key] = kind

    for part_key in members:
        fallback_kind = _story_kind_from_path(part_key)
        if fallback_kind is not None:
            candidates.setdefault(part_key, fallback_kind)

    return sorted(
        (part_key, kind) for part_key, kind in candidates.items() if part_key in members
    )


def _story_kind_from_content_type(content_type: str) -> str | None:
    lowered = content_type.casefold()
    suffixes = (
        ("wordprocessingml.document.main+xml", "body"),
        ("word.document.macroenabled.main+xml", "body"),
        ("wordprocessingml.header+xml", "header"),
        ("wordprocessingml.footer+xml", "footer"),
        ("wordprocessingml.footnotes+xml", "footnote"),
        ("wordprocessingml.endnotes+xml", "endnote"),
        ("wordprocessingml.comments+xml", "comment"),
        ("wordprocessingml.document.glossarydocument+xml", "glossary"),
    )
    for suffix, kind in suffixes:
        if lowered.endswith(suffix):
            return kind
    return None


def _story_kind_from_path(part_key: str) -> str | None:
    if part_key == "word/document.xml":
        return "body"
    if part_key.startswith("word/header") and part_key.endswith(".xml"):
        return "header"
    if part_key.startswith("word/footer") and part_key.endswith(".xml"):
        return "footer"
    exact_paths = {
        "word/footnotes.xml": "footnote",
        "word/endnotes.xml": "endnote",
        "word/comments.xml": "comment",
        "word/glossary/document.xml": "glossary",
    }
    return exact_paths.get(part_key)


def _snapshot_story(
    root: ET.Element,
    part_key: str,
    kind: str,
    relationships: dict[str, _Relationship],
) -> StorySnapshot:
    if not _is_word_element(root, _STORY_ROOT_NAMES[kind]):
        raise DocumentFormatError("document story is invalid")
    context = _story_context(root, kind)
    block_signatures = tuple(
        _fingerprint_element(block, relationships)
        for block in _top_level_blocks(context)
    )

    paragraphs = 0
    tables = 0
    text_runs = 0
    hidden_text_runs = 0
    hidden_paragraph_marks = 0
    alternative_format_import_anchors = 0
    simple_fields = 0
    field_begins = 0
    loose_instructions = 0
    content_controls = 0
    comment_anchors = 0
    insertions = 0
    deletions = 0
    move_from = 0
    move_to = 0
    property_changes = 0

    for element in root.iter():
        namespace, local_name = _qualified_name(element.tag)
        if namespace not in _WORD_NAMESPACES:
            continue
        if local_name == "p":
            paragraphs += 1
            if _paragraph_mark_is_hidden(element):
                hidden_paragraph_marks += 1
        elif local_name == "tbl":
            tables += 1
        elif local_name == "r":
            text_runs += 1
            if _run_is_hidden(element):
                hidden_text_runs += 1
        elif local_name == "fldSimple":
            simple_fields += 1
        elif local_name == "fldChar" and _field_char_is_begin(element):
            field_begins += 1
        elif local_name == "instrText":
            loose_instructions += 1
        elif local_name == "sdt":
            content_controls += 1
        elif local_name == "commentRangeStart":
            comment_anchors += 1
        elif local_name == "altChunk":
            _validate_alternative_format_import_anchor(element, relationships)
            alternative_format_import_anchors += 1

        if local_name == "ins":
            insertions += 1
        elif local_name == "del":
            deletions += 1
        elif local_name == "moveFrom":
            move_from += 1
        elif local_name == "moveTo":
            move_to += 1
        elif local_name.endswith("PrChange"):
            property_changes += 1

    established_fields = simple_fields + field_begins
    field_code_count = established_fields or loose_instructions
    revisions = RevisionInventory(
        insertions=insertions,
        deletions=deletions,
        move_from=move_from,
        move_to=move_to,
        property_changes=property_changes,
    )
    return StorySnapshot(
        part_key=part_key,
        kind=kind,
        block_signatures=block_signatures,
        structure_signature=_fingerprint_element(root, relationships),
        paragraph_count=paragraphs,
        table_count=tables,
        text_run_count=text_runs,
        hidden_text_run_count=hidden_text_runs,
        hidden_paragraph_mark_count=hidden_paragraph_marks,
        alternative_format_import_anchor_count=alternative_format_import_anchors,
        field_code_count=field_code_count,
        content_control_count=content_controls,
        comment_anchor_count=comment_anchors,
        revisions=revisions,
    )


def _story_context(root: ET.Element, kind: str) -> ET.Element:
    if kind != "body":
        return root
    for child in root:
        if _is_word_element(child, "body"):
            return child
    raise DocumentFormatError("document story is invalid")


def _top_level_blocks(context: ET.Element) -> list[ET.Element]:
    blocks: list[ET.Element] = []

    def visit(element: ET.Element) -> None:
        for child in element:
            if _is_word_element(child, "p") or _is_word_element(child, "tbl"):
                blocks.append(child)
            else:
                visit(child)

    visit(context)
    return blocks


def _run_is_hidden(run: ET.Element) -> bool:
    for child in run:
        if not _is_word_element(child, "rPr"):
            continue
        return _vanish_state(child) is True
    return False


def _paragraph_mark_is_hidden(paragraph: ET.Element) -> bool:
    paragraph_properties = _first_word_child(paragraph, "pPr")
    if paragraph_properties is None:
        return False
    run_properties = _first_word_child(paragraph_properties, "rPr")
    if run_properties is None:
        return False
    return _paragraph_mark_hidden_state(run_properties)


def _paragraph_mark_hidden_state(properties: ET.Element) -> bool:
    return _spec_vanish_is_enabled(properties) or _vanish_state(properties) is True


def _spec_vanish_is_enabled(properties: ET.Element) -> bool:
    return any(
        _is_word_element(property_element, "specVanish")
        and _is_enabled(property_element)
        for property_element in properties
    )


def _vanish_state(properties: ET.Element) -> bool | None:
    vanish: bool | None = None
    for property_element in properties:
        if _is_word_element(property_element, "vanish"):
            vanish = _is_enabled(property_element)
    return vanish


def _first_word_child(element: ET.Element, local_name: str) -> ET.Element | None:
    return next(
        (child for child in element if _is_word_element(child, local_name)),
        None,
    )


def _validate_alternative_format_import_anchor(
    element: ET.Element, relationships: dict[str, _Relationship]
) -> None:
    relationship_id = _relationship_id_value(element)
    relationship = relationships.get(relationship_id) if relationship_id else None
    if (
        relationship is None
        or relationship.relationship_type.casefold()
        not in _ALTERNATIVE_FORMAT_IMPORT_RELATIONSHIP_TYPES
        or relationship.target_mode.casefold() != "internal"
    ):
        raise DocumentFormatError("alternative-format import markup is invalid")


def _relationship_id_value(element: ET.Element) -> str | None:
    for attribute, value in element.attrib.items():
        namespace, local_name = _qualified_name(attribute)
        if namespace in _REL_ATTRIBUTE_NAMESPACES and local_name == "id":
            return value
    return None


def _field_char_is_begin(element: ET.Element) -> bool:
    for attribute, value in element.attrib.items():
        if _local_name(attribute) == "fldCharType":
            return value.casefold() == "begin"
    return False


def _is_enabled(element: ET.Element) -> bool:
    for attribute, value in element.attrib.items():
        if _local_name(attribute) == "val":
            return value.strip().casefold() not in _FALSE_VALUES
    return True


def _count_word_elements(root: ET.Element, local_name: str) -> int:
    return sum(1 for element in root.iter() if _is_word_element(element, local_name))


def _settings_inventory(
    archive: zipfile.ZipFile,
    members: dict[str, zipfile.ZipInfo],
    relationship_maps: dict[str, dict[str, _Relationship]],
    limits: PackageLimits,
) -> tuple[bool, str]:
    settings_member = members.get("word/settings.xml")
    if settings_member is None:
        return False, _digest_bytes(b"settings-absent")
    root = _read_xml(archive, settings_member, limits)
    if not _is_word_element(root, "settings"):
        raise DocumentFormatError("document settings are invalid")
    enabled = any(
        _is_word_element(element, "trackRevisions") and _is_enabled(element)
        for element in root.iter()
    )
    return (
        enabled,
        _fingerprint_element(
            root, relationship_maps.get("word/settings.xml", {}), ignore_rsids=True
        ),
    )


def _styles_inventory(
    archive: zipfile.ZipFile,
    members: dict[str, zipfile.ZipInfo],
    relationship_maps: dict[str, dict[str, _Relationship]],
    limits: PackageLimits,
) -> StyleInventory:
    styles_member = members.get("word/styles.xml")
    if styles_member is None:
        return StyleInventory(
            style_definition_count=0,
            hidden_text_style_definition_count=0,
            document_default_hidden_text_enabled=False,
            signature=_digest_bytes(b"styles-absent"),
        )
    root = _read_xml(archive, styles_member, limits)
    if not _is_word_element(root, "styles"):
        raise DocumentFormatError("document styles are invalid")
    style_definitions = [
        element for element in root if _is_word_element(element, "style")
    ]
    hidden_definitions = sum(
        1
        for definition in style_definitions
        if _style_definition_can_hide_text(definition)
    )
    return StyleInventory(
        style_definition_count=len(style_definitions),
        hidden_text_style_definition_count=hidden_definitions,
        document_default_hidden_text_enabled=_document_default_hides_text(root),
        signature=_fingerprint_element(
            root, relationship_maps.get("word/styles.xml", {})
        ),
    )


def _style_definition_can_hide_text(definition: ET.Element) -> bool:
    return _contains_hidden_text_run_properties(definition)


def _contains_hidden_text_run_properties(element: ET.Element) -> bool:
    """Find current text-run properties, excluding mark and revision branches."""

    for child in element:
        if _is_word_element(child, "pPr") or _is_word_element(child, "rPrChange"):
            continue
        if _is_word_element(child, "rPr"):
            if _vanish_state(child) is True:
                return True
            continue
        if _contains_hidden_text_run_properties(child):
            return True
    return False


def _document_default_hides_text(root: ET.Element) -> bool:
    document_defaults = _first_word_child(root, "docDefaults")
    if document_defaults is None:
        return False
    run_defaults = _first_word_child(document_defaults, "rPrDefault")
    if run_defaults is None:
        return False
    run_properties = _first_word_child(run_defaults, "rPr")
    if run_properties is None:
        return False
    return _vanish_state(run_properties) is True


def _custom_xml_inventory(
    archive: zipfile.ZipFile,
    members: dict[str, zipfile.ZipInfo],
    limits: PackageLimits,
) -> tuple[int, str]:
    custom_parts = _custom_xml_part_names(members)
    return len(custom_parts), _digest_member_payloads(
        archive, members, custom_parts, limits
    )


def _macro_inventory(
    archive: zipfile.ZipFile,
    members: dict[str, zipfile.ZipInfo],
    limits: PackageLimits,
) -> tuple[bool, str]:
    macro_parts = _macro_part_names(members)
    return bool(macro_parts), _digest_member_payloads(
        archive, members, macro_parts, limits
    )


def _unclassified_inventory(
    archive: zipfile.ZipFile,
    members: dict[str, zipfile.ZipInfo],
    stories: tuple[StorySnapshot, ...],
    specialized_part_names: frozenset[str],
    limits: PackageLimits,
) -> tuple[int, str]:
    known_parts = {
        "word/settings.xml",
        "word/styles.xml",
        *(story.part_key for story in stories),
        *_custom_xml_part_names(members),
        *_macro_part_names(members),
    }
    known_parts.update(specialized_part_names)
    known_parts.update(name for name in members if name.endswith(".rels"))
    unclassified_parts = sorted(name for name in members if name not in known_parts)
    return (
        len(unclassified_parts),
        _digest_member_payloads(archive, members, unclassified_parts, limits),
    )


def _custom_xml_part_names(members: dict[str, zipfile.ZipInfo]) -> list[str]:
    return sorted(
        name
        for name in members
        if name.startswith("customXml/")
        and "/_rels/" not in name
        and not name.endswith(".rels")
    )


def _macro_part_names(members: dict[str, zipfile.ZipInfo]) -> list[str]:
    return sorted(
        name
        for name in members
        if name.startswith("word/")
        and Path(name).name.casefold().startswith(("vbaproject", "vbadata"))
    )


def _digest_member_payloads(
    archive: zipfile.ZipFile,
    members: dict[str, zipfile.ZipInfo],
    part_names: list[str],
    limits: PackageLimits,
) -> str:
    records = [
        (part_name, _digest_bytes(_read_member(archive, members[part_name], limits)))
        for part_name in part_names
    ]
    encoded = json.dumps(records, ensure_ascii=False, separators=(",", ":")).encode(
        "utf-8"
    )
    return _digest_bytes(encoded)


def _fingerprint_element(
    element: ET.Element,
    relationships: dict[str, _Relationship],
    *,
    ignore_rsids: bool = False,
) -> str:
    canonical = _canonical_element(element, relationships, ignore_rsids=ignore_rsids)
    encoded = json.dumps(canonical, ensure_ascii=False, separators=(",", ":")).encode(
        "utf-8"
    )
    return _digest_bytes(encoded)


def _canonical_element(
    element: ET.Element,
    relationships: dict[str, _Relationship],
    *,
    ignore_rsids: bool,
) -> list[object]:
    namespace, local_name = _qualified_name(element.tag)
    attributes: list[tuple[str, object]] = []
    for attribute, value in sorted(element.attrib.items()):
        if _ignore_attribute(namespace, local_name, attribute):
            continue
        attribute_namespace, _ = _qualified_name(attribute)
        if attribute_namespace in _REL_ATTRIBUTE_NAMESPACES:
            relationship = relationships.get(value)
            if relationship is None:
                raise DocumentFormatError(
                    "document references an unavailable relationship"
                )
            normalized_value: object = ["relationship", *relationship.canonical_value()]
        else:
            normalized_value = value
        attributes.append((attribute, normalized_value))

    children: list[list[object]] = []
    for child in element:
        if ignore_rsids and _is_word_element(child, "rsids"):
            continue
        children.append(
            _canonical_element(child, relationships, ignore_rsids=ignore_rsids)
        )
    return [
        element.tag,
        attributes,
        element.text or "",
        element.tail or "",
        children,
    ]


def _ignore_attribute(
    element_namespace: str,
    element_local_name: str,
    attribute: str,
) -> bool:
    attribute_namespace, attribute_local_name = _qualified_name(attribute)
    if (
        attribute_namespace in _WORD_NAMESPACES
        and attribute_local_name.casefold().startswith("rsid")
    ):
        return True
    if attribute_local_name in _VOLATILE_ATTRIBUTE_NAMES:
        return True
    return (
        element_namespace in _WORD_NAMESPACES
        and (
            element_local_name in _REVISION_TAGS
            or element_local_name.endswith("PrChange")
        )
        and attribute_local_name in _REVISION_METADATA_ATTRIBUTES
    )


def _digest_bytes(payload: bytes) -> str:
    return hashlib.sha256(payload).hexdigest()


def _qualified_name(value: str) -> tuple[str, str]:
    if value.startswith("{"):
        namespace, separator, local_name = value[1:].partition("}")
        if separator:
            return namespace, local_name
    return "", value


def _local_name(value: str) -> str:
    return _qualified_name(value)[1]


def _is_word_element(element: ET.Element, local_name: str) -> bool:
    namespace, actual_name = _qualified_name(element.tag)
    return namespace in _WORD_NAMESPACES and actual_name == local_name
