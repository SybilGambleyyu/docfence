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
    DataBindingInventory,
    DocumentPropertyInventory,
    DocumentSnapshot,
    EmbeddedObjectInventory,
    MailMergeInventory,
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
_DOCUMENT_PROPERTY_RELATIONSHIP_TYPES: Final = {
    (
        "http://schemas.openxmlformats.org/package/2006/relationships/metadata/"
        "core-properties"
    ): "core",
    (
        "http://purl.oclc.org/ooxml/package/relationships/metadata/core-properties"
    ): "core",
    (
        "http://schemas.openxmlformats.org/officedocument/2006/relationships/"
        "extended-properties"
    ): "extended",
    (
        "http://purl.oclc.org/ooxml/officedocument/relationships/extendedproperties"
    ): "extended",
    (
        "http://purl.oclc.org/ooxml/officedocument/relationships/extended-properties"
    ): "extended",
    (
        "http://schemas.openxmlformats.org/officedocument/2006/relationships/"
        "custom-properties"
    ): "custom",
    (
        "http://purl.oclc.org/ooxml/officedocument/relationships/customproperties"
    ): "custom",
    (
        "http://purl.oclc.org/ooxml/officedocument/relationships/custom-properties"
    ): "custom",
}
_DOCUMENT_PROPERTY_FALLBACK_PARTS: Final = {
    "docProps/core.xml": "core",
    "docProps/app.xml": "extended",
    "docProps/custom.xml": "custom",
}
_DOCUMENT_PROPERTY_ROOTS: Final = {
    "core": (
        frozenset(
            {
                "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
                "http://purl.oclc.org/ooxml/package/metadata/core-properties",
            }
        ),
        "coreProperties",
    ),
    "extended": (
        frozenset(
            {
                "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties",
                "http://purl.oclc.org/ooxml/officeDocument/extendedProperties",
                "http://purl.oclc.org/ooxml/officeDocument/extended-properties",
            }
        ),
        "Properties",
    ),
    "custom": (
        frozenset(
            {
                "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties",
                "http://purl.oclc.org/ooxml/officeDocument/customProperties",
                "http://purl.oclc.org/ooxml/officeDocument/custom-properties",
            }
        ),
        "Properties",
    ),
}
_MAIL_MERGE_DATA_SOURCE_RELATIONSHIP_TYPES: Final = frozenset(
    {
        "http://schemas.openxmlformats.org/officedocument/2006/relationships/"
        "mailmergesource",
        "http://purl.oclc.org/ooxml/officedocument/relationships/mailmergesource",
    }
)
_MAIL_MERGE_HEADER_SOURCE_RELATIONSHIP_TYPES: Final = frozenset(
    {
        "http://schemas.openxmlformats.org/officedocument/2006/relationships/"
        "mailmergeheadersource",
        "http://purl.oclc.org/ooxml/officedocument/relationships/mailmergeheadersource",
    }
)
_MAIL_MERGE_RECIPIENT_DATA_RELATIONSHIP_TYPES: Final = frozenset(
    {
        "http://schemas.openxmlformats.org/officedocument/2006/relationships/"
        "recipientdata",
        "http://schemas.openxmlformats.org/officedocument/2006/relationships/"
        "mailmergerecipientdata",
        "http://purl.oclc.org/ooxml/officedocument/relationships/recipientdata",
        "http://purl.oclc.org/ooxml/officedocument/relationships/"
        "mailmergerecipientdata",
    }
)
_CUSTOM_XML_PROPERTIES_RELATIONSHIP_TYPES: Final = frozenset(
    {
        "http://schemas.openxmlformats.org/officedocument/2006/relationships/"
        "customxmlprops",
        "http://purl.oclc.org/ooxml/officedocument/relationships/customxmlprops",
    }
)
_CUSTOM_XML_DATA_PROPERTIES_NAMESPACES: Final = frozenset(
    {
        "http://schemas.openxmlformats.org/officeDocument/2006/customXmlDataProps",
        "http://purl.oclc.org/ooxml/officeDocument/customXmlDataProps",
        "http://schemas.openxmlformats.org/officeDocument/2006/customXml",
        "http://purl.oclc.org/ooxml/officeDocument/customXml",
    }
)


@dataclass(frozen=True)
class _Relationship:
    relationship_type: str
    target: str
    target_mode: str

    def canonical_value(self) -> tuple[str, str, str]:
        return (self.relationship_type, self.target_mode.casefold(), self.target)


@dataclass(frozen=True)
class _DataBindingReference:
    """One private standard Word data-binding declaration in a story."""

    story_part: str
    store_item_id: str | None
    signature: str


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
            document_properties, document_property_parts = _document_property_inventory(
                archive, members, relationship_maps, limits
            )
            mail_merge, mail_merge_parts = _mail_merge_inventory(
                archive, members, relationship_maps, limits
            )
            styles = _styles_inventory(archive, members, relationship_maps, limits)
            stories, comment_count, data_binding_references = _story_snapshots(
                archive, members, content_types, relationship_maps, limits
            )
            data_bindings = _data_binding_inventory(
                archive,
                members,
                data_binding_references,
                relationship_maps,
                limits,
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
                (
                    embedded_object_parts
                    | alternative_format_import_parts
                    | document_property_parts
                    | mail_merge_parts
                ),
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
        document_properties=document_properties,
        mail_merge=mail_merge,
        data_bindings=data_bindings,
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


def _document_property_inventory(
    archive: zipfile.ZipFile,
    members: dict[str, zipfile.ZipInfo],
    relationship_maps: dict[str, dict[str, _Relationship]],
    limits: PackageLimits,
) -> tuple[DocumentPropertyInventory, frozenset[str]]:
    part_names_by_kind = {"core": set(), "extended": set(), "custom": set()}
    part_kinds: dict[str, str] = {}
    records: list[tuple[str, ...]] = []

    def add_part(kind: str, part_name: str) -> None:
        existing_kind = part_kinds.get(part_name)
        if existing_kind is not None and existing_kind != kind:
            raise DocumentFormatError("document property parts are inconsistent")
        part_kinds[part_name] = kind
        part_names_by_kind[kind].add(part_name)

    for part_name, kind in _DOCUMENT_PROPERTY_FALLBACK_PARTS.items():
        if part_name in members:
            add_part(kind, part_name)

    for source_part, relationships in sorted(relationship_maps.items()):
        for relationship in relationships.values():
            kind = _DOCUMENT_PROPERTY_RELATIONSHIP_TYPES.get(
                relationship.relationship_type.casefold()
            )
            if kind is None:
                continue
            records.append(
                (
                    "document_property_relationship",
                    kind,
                    source_part,
                    *relationship.canonical_value(),
                )
            )
            target = _internal_relationship_target(source_part, relationship, members)
            if target is not None:
                add_part(kind, target)

    value_counts = {"core": 0, "extended": 0, "custom": 0}
    for kind, part_names in part_names_by_kind.items():
        for part_name in sorted(part_names):
            root = _read_xml(archive, members[part_name], limits)
            _validate_document_property_root(root, kind)
            if kind == "custom":
                value_counts[kind] += _custom_property_count(root)
            else:
                value_counts[kind] += _document_property_value_count(root)
            records.append(
                (
                    "document_property_part",
                    kind,
                    part_name,
                    _fingerprint_element(root, relationship_maps.get(part_name, {})),
                )
            )

    return (
        DocumentPropertyInventory(
            core_property_part_count=len(part_names_by_kind["core"]),
            core_property_value_count=value_counts["core"],
            extended_property_part_count=len(part_names_by_kind["extended"]),
            extended_property_value_count=value_counts["extended"],
            custom_property_part_count=len(part_names_by_kind["custom"]),
            custom_property_count=value_counts["custom"],
            signature=_digest_records(records),
        ),
        frozenset(part_kinds),
    )


def _validate_document_property_root(root: ET.Element, kind: str) -> None:
    namespaces, expected_local_name = _DOCUMENT_PROPERTY_ROOTS[kind]
    namespace, local_name = _qualified_name(root.tag)
    if namespace not in namespaces or local_name != expected_local_name:
        raise DocumentFormatError("document property part is invalid")


def _document_property_value_count(root: ET.Element) -> int:
    return sum(_element_has_text_value(child) for child in root)


def _custom_property_count(root: ET.Element) -> int:
    return sum(_local_name(child.tag) == "property" for child in root)


def _element_has_text_value(element: ET.Element) -> bool:
    return any((value or "").strip() for value in element.itertext())


def _mail_merge_inventory(
    archive: zipfile.ZipFile,
    members: dict[str, zipfile.ZipInfo],
    relationship_maps: dict[str, dict[str, _Relationship]],
    limits: PackageLimits,
) -> tuple[MailMergeInventory, frozenset[str]]:
    """Inventory mail-merge state without exposing source or recipient data."""

    records: list[tuple[str, ...]] = []
    recipient_data_parts: set[str] = set()
    data_source_relationship_count = 0
    header_source_relationship_count = 0
    recipient_data_relationship_count = 0
    settings_relationships = relationship_maps.get("word/settings.xml", {})

    for relationship in sorted(
        settings_relationships.values(), key=lambda value: value.canonical_value()
    ):
        relationship_type = relationship.relationship_type.casefold()
        if relationship_type in _MAIL_MERGE_DATA_SOURCE_RELATIONSHIP_TYPES:
            if relationship.target_mode.casefold() != "external":
                raise DocumentFormatError("mail merge relationships are invalid")
            data_source_relationship_count += 1
            records.append(("mail_merge_data_source", *relationship.canonical_value()))
        elif relationship_type in _MAIL_MERGE_HEADER_SOURCE_RELATIONSHIP_TYPES:
            if relationship.target_mode.casefold() != "external":
                raise DocumentFormatError("mail merge relationships are invalid")
            header_source_relationship_count += 1
            records.append(
                ("mail_merge_header_source", *relationship.canonical_value())
            )
        elif relationship_type in _MAIL_MERGE_RECIPIENT_DATA_RELATIONSHIP_TYPES:
            target = _internal_relationship_target(
                "word/settings.xml", relationship, members
            )
            if target is None:
                raise DocumentFormatError("mail merge relationships are invalid")
            recipient_data_relationship_count += 1
            recipient_data_parts.add(target)
            records.append(
                ("mail_merge_recipient_data", *relationship.canonical_value())
            )

    configuration_count = 0
    settings_member = members.get("word/settings.xml")
    if settings_member is not None:
        root = _read_xml(archive, settings_member, limits)
        if not _is_word_element(root, "settings"):
            raise DocumentFormatError("document settings are invalid")
        configurations = [
            child for child in root if _is_word_element(child, "mailMerge")
        ]
        configuration_count = len(configurations)
        for configuration in configurations:
            _validate_mail_merge_configuration(
                configuration, settings_relationships, members
            )
            records.append(
                (
                    "mail_merge_configuration",
                    _fingerprint_element(configuration, settings_relationships),
                )
            )

    part_names = frozenset(recipient_data_parts)
    return (
        MailMergeInventory(
            configuration_count=configuration_count,
            data_source_relationship_count=data_source_relationship_count,
            header_source_relationship_count=header_source_relationship_count,
            recipient_data_relationship_count=recipient_data_relationship_count,
            recipient_data_part_count=len(part_names),
            signature=_payload_inventory_signature(
                records, part_names, archive, members, limits
            ),
        ),
        part_names,
    )


def _validate_mail_merge_configuration(
    configuration: ET.Element,
    relationships: dict[str, _Relationship],
    members: dict[str, zipfile.ZipInfo],
) -> None:
    for child in configuration:
        if _is_word_element(child, "dataSource"):
            _validate_mail_merge_relationship(
                child,
                relationships,
                members,
                _MAIL_MERGE_DATA_SOURCE_RELATIONSHIP_TYPES,
                "external",
            )
        elif _is_word_element(child, "headerSource"):
            _validate_mail_merge_relationship(
                child,
                relationships,
                members,
                _MAIL_MERGE_HEADER_SOURCE_RELATIONSHIP_TYPES,
                "external",
            )
        elif _is_word_element(child, "odso"):
            for odso_child in child:
                if _is_word_element(odso_child, "src"):
                    _validate_mail_merge_relationship(
                        odso_child,
                        relationships,
                        members,
                        _MAIL_MERGE_DATA_SOURCE_RELATIONSHIP_TYPES,
                        "external",
                    )
                elif _is_word_element(odso_child, "recipientData"):
                    _validate_mail_merge_relationship(
                        odso_child,
                        relationships,
                        members,
                        _MAIL_MERGE_RECIPIENT_DATA_RELATIONSHIP_TYPES,
                        "internal",
                    )


def _validate_mail_merge_relationship(
    element: ET.Element,
    relationships: dict[str, _Relationship],
    members: dict[str, zipfile.ZipInfo],
    expected_types: frozenset[str],
    expected_target_mode: str,
) -> None:
    relationship_id = _relationship_id_value(element)
    relationship = relationships.get(relationship_id) if relationship_id else None
    if (
        relationship is None
        or relationship.relationship_type.casefold() not in expected_types
        or relationship.target_mode.casefold() != expected_target_mode
    ):
        raise DocumentFormatError("mail merge markup is invalid")
    if expected_target_mode == "internal":
        _internal_relationship_target("word/settings.xml", relationship, members)


def _data_binding_inventory(
    archive: zipfile.ZipFile,
    members: dict[str, zipfile.ZipInfo],
    references: tuple[_DataBindingReference, ...],
    relationship_maps: dict[str, dict[str, _Relationship]],
    limits: PackageLimits,
) -> DataBindingInventory:
    """Inventory standard content-control XML mappings without exposing data.

    A data binding can omit its store item identifier, in which case Word may
    choose a matching custom XML part itself.  DocFence records that fact but
    does not evaluate the binding XPath or infer the chosen data part.
    """

    if not references:
        return DataBindingInventory(
            binding_count=0,
            binding_with_store_item_id_count=0,
            binding_without_store_item_id_count=0,
            referenced_custom_xml_part_count=0,
            unmatched_store_item_id_count=0,
            signature=_payload_inventory_signature(
                [], frozenset(), archive, members, limits
            ),
        )

    store_items = _custom_xml_store_items(archive, members, relationship_maps, limits)
    records: list[tuple[str, ...]] = []
    referenced_data_parts: set[str] = set()
    referenced_inventory_parts: set[str] = set()
    with_store_item_id_count = 0
    without_store_item_id_count = 0
    unmatched_store_item_id_count = 0

    for reference in references:
        if reference.store_item_id is None:
            without_store_item_id_count += 1
            records.append(
                (
                    "data_binding",
                    reference.story_part,
                    reference.signature,
                    "without_store_item_id",
                )
            )
            continue

        with_store_item_id_count += 1
        matches = store_items.get(_normalize_store_item_id(reference.store_item_id), ())
        if not matches:
            unmatched_store_item_id_count += 1
            records.append(
                (
                    "data_binding",
                    reference.story_part,
                    reference.signature,
                    "unmatched_store_item_id",
                )
            )
            continue

        records.append(
            (
                "data_binding",
                reference.story_part,
                reference.signature,
                "matched_store_item_id",
            )
        )
        for data_part, properties_part in matches:
            referenced_data_parts.add(data_part)
            referenced_inventory_parts.update((data_part, properties_part))

    part_names = frozenset(referenced_inventory_parts)
    return DataBindingInventory(
        binding_count=len(references),
        binding_with_store_item_id_count=with_store_item_id_count,
        binding_without_store_item_id_count=without_store_item_id_count,
        referenced_custom_xml_part_count=len(referenced_data_parts),
        unmatched_store_item_id_count=unmatched_store_item_id_count,
        signature=_payload_inventory_signature(
            records, part_names, archive, members, limits
        ),
    )


def _custom_xml_store_items(
    archive: zipfile.ZipFile,
    members: dict[str, zipfile.ZipInfo],
    relationship_maps: dict[str, dict[str, _Relationship]],
    limits: PackageLimits,
) -> dict[str, tuple[tuple[str, str], ...]]:
    """Return private custom-XML storage ID associations discovered in-package."""

    by_store_item_id: dict[str, list[tuple[str, str]]] = {}
    for data_part, relationships in sorted(relationship_maps.items()):
        if not _is_custom_xml_data_part(data_part, members):
            continue
        for relationship in relationships.values():
            if (
                relationship.relationship_type.casefold()
                not in _CUSTOM_XML_PROPERTIES_RELATIONSHIP_TYPES
            ):
                continue
            if relationship.target_mode.casefold() != "internal":
                raise DocumentFormatError(
                    "custom XML properties relationships are invalid"
                )
            properties_part = _internal_relationship_target(
                data_part, relationship, members
            )
            if properties_part is None:
                raise DocumentFormatError(
                    "custom XML properties relationships are invalid"
                )
            root = _read_xml(archive, members[properties_part], limits)
            store_item_id = _custom_xml_properties_store_item_id(root)
            if store_item_id is None:
                raise DocumentFormatError("custom XML properties parts are invalid")
            normalized_store_item_id = _normalize_store_item_id(store_item_id)
            if normalized_store_item_id in by_store_item_id:
                raise DocumentFormatError("custom XML storage identifiers are invalid")
            by_store_item_id[normalized_store_item_id] = [(data_part, properties_part)]

    return {
        store_item_id: tuple(sorted(matches))
        for store_item_id, matches in by_store_item_id.items()
    }


def _is_custom_xml_data_part(
    part_name: str, members: dict[str, zipfile.ZipInfo]
) -> bool:
    return (
        part_name in members
        and part_name.startswith("customXml/")
        and "/_rels/" not in part_name
        and not part_name.endswith(".rels")
    )


def _custom_xml_properties_store_item_id(root: ET.Element) -> str | None:
    namespace, local_name = _qualified_name(root.tag)
    if (
        namespace not in _CUSTOM_XML_DATA_PROPERTIES_NAMESPACES
        or local_name != "datastoreItem"
    ):
        return None
    for attribute, value in root.attrib.items():
        attribute_namespace, local_name = _qualified_name(attribute)
        if (
            attribute_namespace in _CUSTOM_XML_DATA_PROPERTIES_NAMESPACES
            and local_name == "itemID"
            and value.strip()
        ):
            return value.strip()
    return None


def _normalize_store_item_id(value: str) -> str:
    return value.strip().casefold()


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


def _digest_records(records: list[tuple[str, ...]]) -> str:
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
) -> tuple[tuple[StorySnapshot, ...], int, tuple[_DataBindingReference, ...]]:
    stories: list[StorySnapshot] = []
    data_binding_references: list[_DataBindingReference] = []
    comment_count = 0
    for part_key, kind in _discover_story_parts(members, content_types):
        root = _read_xml(archive, members[part_key], limits)
        story, story_data_binding_references = _snapshot_story(
            root, part_key, kind, relationship_maps.get(part_key, {})
        )
        stories.append(story)
        data_binding_references.extend(story_data_binding_references)
        if kind == "comment":
            comment_count += _count_word_elements(root, "comment")
    if not any(story.kind == "body" for story in stories):
        raise DocumentFormatError("package does not contain a document story")
    return tuple(stories), comment_count, tuple(data_binding_references)


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
) -> tuple[StorySnapshot, tuple[_DataBindingReference, ...]]:
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
    return (
        StorySnapshot(
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
        ),
        _data_binding_references(root, part_key, relationships),
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


def _data_binding_references(
    root: ET.Element,
    story_part: str,
    relationships: dict[str, _Relationship],
) -> tuple[_DataBindingReference, ...]:
    """Find standard direct ``w:sdtPr/w:dataBinding`` declarations only."""

    references: list[_DataBindingReference] = []
    for properties in root.iter():
        if not _is_word_element(properties, "sdtPr"):
            continue
        for child in properties:
            if not _is_word_element(child, "dataBinding"):
                continue
            store_item_id = _word_attribute_value(child, "storeItemID")
            normalized_store_item_id = store_item_id.strip() if store_item_id else ""
            references.append(
                _DataBindingReference(
                    story_part=story_part,
                    store_item_id=normalized_store_item_id or None,
                    signature=_fingerprint_element(child, relationships),
                )
            )
    return tuple(references)


def _word_attribute_value(element: ET.Element, local_name: str) -> str | None:
    for attribute, value in element.attrib.items():
        namespace, attribute_local_name = _qualified_name(attribute)
        if namespace in _WORD_NAMESPACES and attribute_local_name == local_name:
            return value
    return None


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
