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
import re
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
    AttachedCustomXmlSchemaInventory,
    ContentControlLockInventory,
    DataBindingInventory,
    DocumentPropertyInventory,
    DocumentSnapshot,
    DocumentTaskInventory,
    EmbeddedObjectInventory,
    ExternalDocumentDependencyInventory,
    ExternalFieldInventory,
    FieldUpdateOnOpenInventory,
    MailMergeInventory,
    MarkupCompatibilityInventory,
    ModernCommentMetadataInventory,
    PackageDigitalSignatureInventory,
    PackageSignatureCoverageInventory,
    PackageThumbnailInventory,
    PersonalInformationRemovalOnSaveInventory,
    RelationshipInventory,
    RevisionInventory,
    SaveFormsDataInventory,
    SavePreviewPictureInventory,
    SaveThroughXsltInventory,
    SensitivityLabelInventory,
    StorySnapshot,
    StyleInventory,
    TaskpaneWebExtensionInventory,
    TemplateStyleUpdateOnOpenInventory,
    WordDocumentVariableFieldInventory,
    WordDocumentVariableInventory,
    WordDrawingHyperlinkInventory,
    WordDrawingLinkedPictureInventory,
    WordDrawingVisibilityInventory,
    WordEmbeddedControlInventory,
    WordHyperlinkFieldInventory,
    WordHyperlinkMarkupInventory,
    WordObjectLinkInventory,
    WordPermissionRangeInventory,
    WordProtectionInventory,
    WordVmlEmbeddedOleObjectInventory,
    WordVmlExternalImageInventory,
    WordVmlHyperlinkInventory,
    WordVmlImageHyperlinkInventory,
    WordVmlLinkedOleObjectInventory,
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
_DRAWING_NAMESPACES: Final = frozenset(
    {
        "http://schemas.openxmlformats.org/drawingml/2006/main",
        "http://purl.oclc.org/ooxml/drawingml/main",
    }
)
_WORD_DRAWING_VISIBILITY_ELEMENTS: Final = frozenset(
    {
        (
            "http://schemas.openxmlformats.org/drawingml/2006/main",
            "cNvPr",
        ),
        ("http://purl.oclc.org/ooxml/drawingml/main", "cNvPr"),
        (
            "http://schemas.openxmlformats.org/drawingml/2006/picture",
            "cNvPr",
        ),
        ("http://purl.oclc.org/ooxml/drawingml/picture", "cNvPr"),
        (
            "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
            "docPr",
        ),
        (
            "http://purl.oclc.org/ooxml/drawingml/wordprocessingDrawing",
            "docPr",
        ),
        ("http://schemas.microsoft.com/office/word/2010/wordml", "cNvPr"),
        (
            "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup",
            "cNvPr",
        ),
        (
            "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
            "cNvPr",
        ),
    }
)
_VML_NAMESPACE: Final = "urn:schemas-microsoft-com:vml"
_OFFICE_VML_NAMESPACE: Final = "urn:schemas-microsoft-com:office:office"
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
_ON_OFF_TRUE_VALUES: Final = frozenset({"1", "on", "true"})
_ON_OFF_FALSE_VALUES: Final = frozenset({"0", "false", "off"})
_ON_OFF_VALUES: Final = _ON_OFF_TRUE_VALUES | _ON_OFF_FALSE_VALUES
_XML_BOOLEAN_TRUE_VALUES: Final = frozenset({"1", "true"})
_XML_BOOLEAN_FALSE_VALUES: Final = frozenset({"0", "false"})
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
_OLE_OBJECT_RELATIONSHIP_TYPES: Final = frozenset(
    {
        "http://schemas.openxmlformats.org/officedocument/2006/relationships/oleobject",
        "http://purl.oclc.org/ooxml/officedocument/relationships/oleobject",
    }
)
_CONTROL_RELATIONSHIP_TYPES: Final = frozenset(
    {
        "http://schemas.openxmlformats.org/officedocument/2006/relationships/control",
        "http://purl.oclc.org/ooxml/officedocument/relationships/control",
    }
)
_EMBEDDED_CONTROL_RELATIONSHIP_TYPES: Final = _CONTROL_RELATIONSHIP_TYPES | frozenset(
    {
        "http://schemas.microsoft.com/office/2006/relationships/activexcontrolbinary",
    }
)
_ALTERNATIVE_FORMAT_IMPORT_RELATIONSHIP_TYPES: Final = frozenset(
    {
        "http://schemas.openxmlformats.org/officedocument/2006/relationships/afchunk",
        "http://purl.oclc.org/ooxml/officedocument/relationships/afchunk",
    }
)
_HYPERLINK_RELATIONSHIP_TYPES: Final = frozenset(
    {
        "http://schemas.openxmlformats.org/officedocument/2006/relationships/hyperlink",
        "http://purl.oclc.org/ooxml/officedocument/relationships/hyperlink",
    }
)
_IMAGE_RELATIONSHIP_TYPES: Final = frozenset(
    {
        "http://schemas.openxmlformats.org/officedocument/2006/relationships/image",
        "http://purl.oclc.org/ooxml/officedocument/relationships/image",
    }
)
_DRAWING_HYPERLINK_REFERENCE_KINDS: Final = {
    "hlinkClick": "click",
    "hlinkHover": "hover",
    "hlinkMouseOver": "mouse_over",
}
_VML_HYPERLINK_REFERENCE_KINDS: Final = {
    "arc": "concrete_shape",
    "curve": "concrete_shape",
    "image": "concrete_shape",
    "line": "concrete_shape",
    "oval": "concrete_shape",
    "polyline": "concrete_shape",
    "rect": "concrete_shape",
    "roundrect": "concrete_shape",
    "shape": "concrete_shape",
    "group": "group",
    "shapetype": "shape_type",
}
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
_PACKAGE_THUMBNAIL_RELATIONSHIP_TYPES: Final = frozenset(
    {
        (
            "http://schemas.openxmlformats.org/package/2006/relationships/"
            "metadata/thumbnail"
        ),
        ("http://purl.oclc.org/ooxml/officedocument/relationships/metadata/thumbnail"),
    }
)
_MARKUP_COMPATIBILITY_NAMESPACE: Final = (
    "http://schemas.openxmlformats.org/markup-compatibility/2006"
)
_MARKUP_COMPATIBILITY_MARKER: Final = b"markup-compatibility/2006"
_MARKUP_COMPATIBILITY_ELEMENT_NAMES: Final = frozenset(
    {"AlternateContent", "Choice", "Fallback"}
)
_MARKUP_COMPATIBILITY_ATTRIBUTE_NAMES: Final = frozenset(
    {
        "Ignorable",
        "MustUnderstand",
        "ProcessContent",
        "PreserveElements",
        "PreserveAttributes",
    }
)
_SENSITIVITY_LABEL_NAMESPACE: Final = (
    "http://schemas.microsoft.com/office/2020/mipLabelMetadata"
)
_SENSITIVITY_LABEL_PART_CONTENT_TYPES: Final = frozenset(
    {"application/vnd.ms-office.classificationlabels+xml"}
)
_SENSITIVITY_LABEL_RELATIONSHIP_TYPE: Final = (
    "http://schemas.microsoft.com/office/2020/02/relationships/classificationlabels"
)
_SENSITIVITY_LABEL_FALLBACK_PART_NAMES: Final = frozenset(
    {"docmetadata/labelinfo", "docmetadata/labelinfo.xml"}
)
_PACKAGE_DIGITAL_SIGNATURE_ORIGIN_CONTENT_TYPE: Final = (
    "application/vnd.openxmlformats-package.digital-signature-origin"
)
_PACKAGE_DIGITAL_SIGNATURE_XML_CONTENT_TYPE: Final = (
    "application/vnd.openxmlformats-package.digital-signature-xmlsignature+xml"
)
_PACKAGE_DIGITAL_SIGNATURE_CERTIFICATE_CONTENT_TYPE: Final = (
    "application/vnd.openxmlformats-package.digital-signature-certificate"
)
_PACKAGE_DIGITAL_SIGNATURE_ORIGIN_RELATIONSHIP_TYPE: Final = (
    "http://schemas.openxmlformats.org/package/2006/relationships/"
    "digital-signature/origin"
)
_PACKAGE_DIGITAL_SIGNATURE_RELATIONSHIP_TYPE: Final = (
    "http://schemas.openxmlformats.org/package/2006/relationships/"
    "digital-signature/signature"
)
_PACKAGE_DIGITAL_SIGNATURE_CERTIFICATE_RELATIONSHIP_TYPE: Final = (
    "http://schemas.openxmlformats.org/package/2006/relationships/"
    "digital-signature/certificate"
)
_PACKAGE_DIGITAL_SIGNATURE_ORIGIN_FALLBACK_PART_NAMES: Final = frozenset(
    {"_xmlsignatures/origin.sigs"}
)
_XMLDSIG_NAMESPACE: Final = "http://www.w3.org/2000/09/xmldsig#"
_OPC_DIGITAL_SIGNATURE_NAMESPACE: Final = (
    "http://schemas.openxmlformats.org/package/2006/digital-signature"
)
_OPC_PACKAGE_SPECIFIC_OBJECT_ID: Final = "idPackageObject"
_OPC_SIGNATURE_TIME_PROPERTY_ID: Final = "idSignatureTime"
_OPC_SIGNATURE_TIME_VALUE_PATTERNS: Final = {
    "YYYY": re.compile(r"[0-9]{4}"),
    "YYYY-MM": re.compile(r"[0-9]{4}-(?:0[1-9]|1[0-2])"),
    "YYYY-MM-DD": re.compile(r"[0-9]{4}-(?:0[1-9]|1[0-2])-(?:0[1-9]|[12][0-9]|3[01])"),
    "YYYY-MM-DDThh:mmTZD": re.compile(
        r"[0-9]{4}-(?:0[1-9]|1[0-2])-(?:0[1-9]|[12][0-9]|3[01])"
        r"T(?:[01][0-9]|2[0-3]):[0-5][0-9]"
        r"(?:Z|[+-](?:[01][0-9]|2[0-3]):[0-5][0-9])"
    ),
    "YYYY-MM-DDThh:mm:ssTZD": re.compile(
        r"[0-9]{4}-(?:0[1-9]|1[0-2])-(?:0[1-9]|[12][0-9]|3[01])"
        r"T(?:[01][0-9]|2[0-3]):[0-5][0-9]:[0-5][0-9]"
        r"(?:Z|[+-](?:[01][0-9]|2[0-3]):[0-5][0-9])"
    ),
    "YYYY-MM-DDThh:mm:ss.sTZD": re.compile(
        r"[0-9]{4}-(?:0[1-9]|1[0-2])-(?:0[1-9]|[12][0-9]|3[01])"
        r"T(?:[01][0-9]|2[0-3]):[0-5][0-9]:[0-5][0-9]\.[0-9]"
        r"(?:Z|[+-](?:[01][0-9]|2[0-3]):[0-5][0-9])"
    ),
}
_OPC_RELATIONSHIP_TRANSFORM_ALGORITHM: Final = (
    "http://schemas.openxmlformats.org/package/2006/RelationshipTransform"
)
_XML_CANONICALIZATION_TRANSFORM_ALGORITHMS: Final = frozenset(
    {
        "http://www.w3.org/TR/2001/REC-xml-c14n-20010315",
        "http://www.w3.org/TR/2001/REC-xml-c14n-20010315#WithComments",
    }
)
_OPC_SUPPORTED_MANIFEST_TRANSFORM_ALGORITHMS: Final = (
    _XML_CANONICALIZATION_TRANSFORM_ALGORITHMS
    | frozenset({_OPC_RELATIONSHIP_TRANSFORM_ALGORITHM})
)
_PACKAGE_RELATIONSHIP_CONTENT_TYPE: Final = (
    "application/vnd.openxmlformats-package.relationships+xml"
)
_ROOT_OFFICE_DOCUMENT_RELATIONSHIP_TYPES: Final = frozenset(
    {
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
        "officeDocument",
        "http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument",
    }
)
_DOCUMENT_PROTECTION_EDIT_VALUES: Final = frozenset(
    {"none", "readOnly", "comments", "trackedChanges", "forms"}
)
_WORD_PROTECTION_COMMON_ATTRIBUTE_NAMES: Final = frozenset(
    {
        "cryptProviderType",
        "cryptAlgorithmClass",
        "cryptAlgorithmType",
        "cryptAlgorithmSid",
        "cryptSpinCount",
        "cryptProvider",
        "algIdExt",
        "algIdExtSource",
        "cryptProviderTypeExt",
        "cryptProviderTypeExtSource",
        "hash",
        "salt",
        "algorithmName",
        "hashValue",
        "saltValue",
        "spinCount",
    }
)
_DOCUMENT_PROTECTION_ATTRIBUTE_NAMES: Final = (
    _WORD_PROTECTION_COMMON_ATTRIBUTE_NAMES
    | frozenset({"edit", "formatting", "enforcement"})
)
_WRITE_PROTECTION_ATTRIBUTE_NAMES: Final = (
    _WORD_PROTECTION_COMMON_ATTRIBUTE_NAMES | frozenset({"recommended"})
)
_WORD_PROTECTION_PASSWORD_MATERIAL_ATTRIBUTE_NAMES: Final = frozenset(
    {"hash", "salt", "hashValue", "saltValue"}
)
_WORD_PROTECTION_TRUE_VALUES: Final = frozenset({"1", "true", "on"})
_WORD_PROTECTION_FALSE_VALUES: Final = frozenset({"0", "false", "off"})
_WORD_DOCUMENT_VARIABLE_ATTRIBUTE_NAMES: Final = frozenset({"name", "val"})
_WORD_DOCUMENT_VARIABLE_NAME_MAX_UTF16_CODE_UNITS: Final = 255
_WORD_DOCUMENT_VARIABLE_VALUE_MAX_UTF16_CODE_UNITS: Final = 65_280
_WORD_PERMISSION_START_ATTRIBUTE_NAMES: Final = frozenset(
    {"id", "ed", "edGrp", "colFirst", "colLast", "displacedByCustomXml"}
)
_WORD_PERMISSION_END_ATTRIBUTE_NAMES: Final = frozenset({"id", "displacedByCustomXml"})
_WORD_PERMISSION_EDITOR_GROUPS: Final = frozenset(
    {
        "none",
        "everyone",
        "administrators",
        "contributors",
        "editors",
        "owners",
        "current",
    }
)
_WORD_PERMISSION_DISPLACED_BY_CUSTOM_XML_VALUES: Final = frozenset({"next", "prev"})
_WORD_PERMISSION_COLUMN_VALUE: Final = re.compile(r"^[0-9]+$")
_SENSITIVITY_LABEL_GUID: Final = re.compile(
    r"^\{[0-9A-Fa-f]{8}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-"
    r"[0-9A-Fa-f]{4}-[0-9A-Fa-f]{12}\}$"
)
_LEGACY_MIP_LABEL_PROPERTY_NAME: Final = re.compile(
    r"^msip_label_(?P<label_id>\{?[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-"
    r"[0-9a-f]{4}-[0-9a-f]{12}\}?)_(?P<attribute>.+)$",
    re.IGNORECASE,
)
_WORD_SENSITIVITY_CONTENT_MARKING_PROPERTY_NAME: Final = re.compile(
    r"^(?:classificationcontentmarking(?:header(?:fontprops|text|shapeids(?:-"
    r"[0-9a-f]+)?)|footer(?:fontprops|text|shapeids(?:-[0-9a-f]+)?))|"
    r"classificationwatermark(?:fontprops|text|shapeids(?:-[0-9a-f]+)?))$",
    re.IGNORECASE,
)
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
_SAVE_THROUGH_XSLT_RELATIONSHIP_TYPES: Final = frozenset(
    {
        "http://schemas.openxmlformats.org/officedocument/2006/relationships/transform",
        "http://purl.oclc.org/ooxml/officedocument/relationships/transform",
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
_ATTACHED_TEMPLATE_RELATIONSHIP_TYPES: Final = frozenset(
    {
        "http://schemas.openxmlformats.org/officedocument/2006/relationships/"
        "attachedtemplate",
        "http://purl.oclc.org/ooxml/officedocument/relationships/attachedtemplate",
    }
)
_SUBDOCUMENT_RELATIONSHIP_TYPES: Final = frozenset(
    {
        "http://schemas.openxmlformats.org/officedocument/2006/relationships/"
        "subdocument",
        "http://purl.oclc.org/ooxml/officedocument/relationships/subdocument",
    }
)
_FRAME_RELATIONSHIP_TYPES: Final = frozenset(
    {
        "http://schemas.openxmlformats.org/officedocument/2006/relationships/frame",
        "http://purl.oclc.org/ooxml/officedocument/relationships/frame",
    }
)
_DOCUMENT_SETTINGS_RELATIONSHIP_TYPES: Final = frozenset(
    {
        "http://schemas.openxmlformats.org/officedocument/2006/relationships/settings",
        "http://purl.oclc.org/ooxml/officedocument/relationships/settings",
    }
)
_WEB_SETTINGS_RELATIONSHIP_TYPES: Final = frozenset(
    {
        "http://schemas.openxmlformats.org/officedocument/2006/relationships/"
        "websettings",
        "http://purl.oclc.org/ooxml/officedocument/relationships/websettings",
    }
)
_SETTINGS_SOURCE_PARTS: Final = frozenset(
    {"word/document.xml", "word/glossary/document.xml"}
)
_EXTERNAL_FIELD_CATEGORY_BY_KEYWORD: Final = {
    "database": "database",
    "data": "legacy_data",
    "dde": "dde",
    "ddeauto": "dde_auto",
    "include": "include_text",
    "includetext": "include_text",
    "includepicture": "include_picture",
    "import": "include_picture",
    "link": "link",
    "rd": "referenced_document",
}
_FIELD_INSTRUCTION_VARIANTS: Final = ("current", "deleted")
_CURRENT_FIELD_INSTRUCTION_VARIANTS: Final = frozenset({"current"})
_DELETED_FIELD_INSTRUCTION_VARIANTS: Final = frozenset({"deleted"})
_CURRENT_FIELD_REVISION_TAGS: Final = frozenset({"ins", "moveTo"})
_DELETED_FIELD_REVISION_TAGS: Final = frozenset({"del", "moveFrom"})
_SUPPORTED_MAIN_DOCUMENT_CONTENT_TYPE_SUFFIXES: Final = frozenset(
    {
        "wordprocessingml.document.main+xml",
        "word.document.macroenabled.main+xml",
        "wordprocessingml.template.main+xml",
        "word.template.macroenabledtemplate.main+xml",
    }
)
_MODERN_COMMENT_PART_CONTENT_TYPES: Final = {
    (
        "application/vnd.openxmlformats-officedocument.wordprocessingml.people+xml"
    ): "people",
    (
        "application/vnd.openxmlformats-officedocument.wordprocessingml."
        "commentsextended+xml"
    ): "comments_extended",
    (
        "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsids+xml"
    ): "comments_ids",
    (
        "application/vnd.openxmlformats-officedocument.wordprocessingml."
        "commentsextensible+xml"
    ): "comments_extensible",
}
_MODERN_COMMENT_PART_RELATIONSHIP_TYPES: Final = {
    "http://schemas.microsoft.com/office/2011/relationships/people": "people",
    (
        "http://schemas.microsoft.com/office/2011/relationships/commentsextended"
    ): "comments_extended",
    (
        "http://schemas.microsoft.com/office/2016/09/relationships/commentsids"
    ): "comments_ids",
    (
        "http://schemas.microsoft.com/office/2018/08/relationships/commentsextensible"
    ): "comments_extensible",
}
_MODERN_COMMENT_FALLBACK_PARTS: Final = {
    "word/people.xml": "people",
    "word/commentsExtended.xml": "comments_extended",
    "word/commentsIds.xml": "comments_ids",
    "word/commentsExtensible.xml": "comments_extensible",
}
_MODERN_COMMENT_ROOTS: Final = {
    "people": (
        frozenset(
            {
                "http://schemas.microsoft.com/office/word/2010/11/wordml",
                "http://schemas.microsoft.com/office/word/2012/wordml",
            }
        ),
        "people",
    ),
    ("comments_extended"): (
        frozenset(
            {
                "http://schemas.microsoft.com/office/word/2010/11/wordml",
                "http://schemas.microsoft.com/office/word/2012/wordml",
            }
        ),
        "commentsEx",
    ),
    ("comments_ids"): (
        frozenset({"http://schemas.microsoft.com/office/word/2016/wordml/cid"}),
        "commentsIds",
    ),
    ("comments_extensible"): (
        frozenset({"http://schemas.microsoft.com/office/word/2018/wordml/cex"}),
        "commentsExtensible",
    ),
}
_MODERN_COMMENT_REACTIONS_NAMESPACE: Final = (
    "http://schemas.microsoft.com/office/comments/2020/reactions"
)
_DOCUMENT_TASK_NAMESPACE: Final = (
    "http://schemas.microsoft.com/office/tasks/2019/documenttasks"
)
_DOCUMENT_TASK_CONTENT_TYPE: Final = "application/vnd.ms-office.documenttasks+xml"
_DOCUMENT_TASK_RELATIONSHIP_TYPES: Final = frozenset(
    {"http://schemas.microsoft.com/office/2019/05/relationships/documenttasks"}
)
_DOCUMENT_TASK_FALLBACK_PARTS: Final = frozenset({"word/tasks.xml", "word/tasks"})
_DOCUMENT_TASK_EVENT_COUNT_KEYS: Final = {
    "Assign": "assignment",
    "Unassign": "unassignment",
    "Create": "creation",
    "SetTitle": "title_change",
    "Schedule": "schedule_change",
    "Progress": "progress_change",
    "Priority": "priority_change",
    "Delete": "deletion",
    "Undelete": "restoration",
    "UnassignAll": "unassign_all",
    "Undo": "undo",
}
_DOCUMENT_TASK_USER_REFERENCE_NAMES: Final = frozenset(
    {"Attribution", "Assign", "Unassign"}
)
_TASKPANE_WEB_EXTENSION_PART_CONTENT_TYPES: Final = {
    "application/vnd.ms-office.webextensiontaskpanes+xml": "taskpanes",
    "application/vnd.ms-office.webextension+xml": "web_extension",
}
_TASKPANE_WEB_EXTENSION_PART_RELATIONSHIP_TYPES: Final = {
    (
        "http://schemas.microsoft.com/office/2011/relationships/webextensiontaskpanes"
    ): "taskpanes",
    "http://schemas.microsoft.com/office/2011/relationships/webextension": (
        "web_extension"
    ),
}
_TASKPANE_WEB_EXTENSION_FALLBACK_PARTS: Final = {
    "word/webextensions/taskpanes.xml": "taskpanes",
    "word/webextensions/taskpanes": "taskpanes",
}
_TASKPANE_WEB_EXTENSION_TASKPANES_NAMESPACE: Final = (
    "http://schemas.microsoft.com/office/webextensions/taskpanes/2010/11"
)
_TASKPANE_WEB_EXTENSION_NAMESPACE: Final = (
    "http://schemas.microsoft.com/office/webextensions/webextension/2010/11"
)
_WORD_2012_NAMESPACE: Final = "http://schemas.microsoft.com/office/word/2012/wordml"
_TASKPANE_REFERENCE_TAGS: Final = frozenset(
    {
        (_TASKPANE_WEB_EXTENSION_TASKPANES_NAMESPACE, "webextensionref"),
        (_TASKPANE_WEB_EXTENSION_TASKPANES_NAMESPACE, "webextension"),
    }
)
_AUTO_SHOW_TASKPANE_PROPERTY_NAME: Final = "Office.AutoShowTaskpaneWithDocument"
_CONTENT_CONTROL_LOCK_STATE_BY_VALUE: Final = {
    "unlocked": "unlocked",
    "sdtLocked": "sdt_locked",
    "contentLocked": "content_locked",
    "sdtContentLocked": "sdt_content_locked",
}


@dataclass(frozen=True)
class _Relationship:
    relationship_type: str
    target: str
    target_mode: str

    def canonical_value(self) -> tuple[str, str, str]:
        return (self.relationship_type, self.target_mode.casefold(), self.target)


@dataclass(frozen=True)
class _ManifestReferenceResolution:
    """Private static result for one package-object manifest reference."""

    covered_part_name: str | None
    covered_relationship_ids: frozenset[tuple[str, str]]
    unresolved_reference_count: int
    unsupported_reference_count: int


@dataclass(frozen=True)
class _DeclaredPackageSignatureCoverage:
    """Private union contributed by one XML signature part."""

    has_declared_package_coverage: bool
    covered_part_names: frozenset[str]
    covered_relationship_ids: frozenset[tuple[str, str]]
    unresolved_reference_count: int
    unsupported_reference_count: int
    records: tuple[tuple[str, ...], ...]


@dataclass(frozen=True)
class _DataBindingReference:
    """One private standard Word data-binding declaration in a story."""

    story_part: str
    store_item_id: str | None
    signature: str


@dataclass(frozen=True)
class _ContentControlLockReference:
    """One private direct ``w:sdtPr/w:lock`` state in a Word story."""

    story_part: str
    ordinal: int
    state: str


@dataclass(frozen=True)
class _ExternalFieldReference:
    """One private external-source Word field instruction in a story."""

    story_part: str
    category: str
    instruction_signature: str


@dataclass(frozen=True)
class _FieldInstructionReference:
    """One complete field instruction retained only during story traversal."""

    story_part: str
    instruction: str
    contains_nested_instruction_field: bool


@dataclass(frozen=True)
class _DocumentVariableFieldReference:
    """One private ``DOCVARIABLE`` instruction with optional literal name."""

    story_part: str
    document_scope: str
    instruction_signature: str
    literal_name: str | None


@dataclass(frozen=True)
class _HyperlinkFieldReference:
    """One private ``HYPERLINK`` instruction and lexical target class."""

    story_part: str
    instruction_signature: str
    classification: str


@dataclass(frozen=True)
class _HyperlinkMarkupReference:
    """One private direct ``w:hyperlink`` element and target class."""

    story_part: str
    markup_signature: str
    classification: str
    relationship_backed_anchor_attribute: bool


@dataclass(frozen=True)
class _DrawingHyperlinkReference:
    """One private direct DrawingML hyperlink-action marker in a Word story."""

    story_part: str
    markup_signature: str
    kind: str
    classification: str
    action_attribute_present: bool
    invalid_url_attribute_present: bool


@dataclass(frozen=True)
class _DrawingVisibilityReference:
    """One private direct DrawingML nonvisual visibility declaration."""

    story_part: str
    namespace: str
    local_name: str
    state: str


@dataclass(frozen=True)
class _DrawingLinkedPictureReference:
    """One private direct DrawingML ``a:blip/@r:link`` marker in a Word story."""

    story_part: str
    markup_signature: str
    classification: str


@dataclass(frozen=True)
class _VmlHyperlinkReference:
    """One private direct legacy VML href marker in a Word story."""

    story_part: str
    markup_signature: str
    kind: str
    target_attribute_present: bool


@dataclass(frozen=True)
class _VmlExternalImageReference:
    """One private VML ``imagedata/@r:id`` external-image marker."""

    story_part: str
    marker_signature: str
    classification: str


@dataclass(frozen=True)
class _VmlImageHyperlinkReference:
    """One private VML ``imagedata/@r:href`` relationship marker."""

    story_part: str
    marker_signature: str
    classification: str


@dataclass(frozen=True)
class _VmlLinkedOleObjectReference:
    """One private direct legacy VML Office linked-OLE marker."""

    story_part: str
    markup_signature: str
    update_mode_classification: str
    relationship_classification: str


@dataclass(frozen=True)
class _VmlEmbeddedOleObjectReference:
    """One private direct legacy VML Office embedded-OLE marker."""

    story_part: str
    markup_signature: str
    relationship_classification: str


@dataclass(frozen=True)
class _WordObjectLinkReference:
    """One private direct WordprocessingML linked-object-property marker."""

    story_part: str
    markup_signature: str
    update_mode_classification: str
    relationship_classification: str


@dataclass(frozen=True)
class _WordEmbeddedControlReference:
    """One private direct WordprocessingML embedded-control anchor."""

    story_part: str
    markup_signature: str
    parent_kind: str
    relationship_classification: str


@dataclass(frozen=True)
class _StoredDocumentVariable:
    """A validated stored variable, private until association is complete."""

    document_scopes: frozenset[str]
    name: str


@dataclass(frozen=True)
class _WebExtensionControlReference:
    """One private marker for a Word content control bound to an add-in."""

    story_part: str
    marker_kind: str
    signature: str


@dataclass
class _ComplexFieldState:
    """Transient state for one complex field while its story is traversed."""

    simple_field_depth: int
    instruction_chunks: dict[str, list[str]]
    nested_instruction_field_variants: set[str]
    accepts_instructions: bool = True


@dataclass(frozen=True)
class _WordPermissionMarker:
    """A validated range-permission boundary retained only while loading."""

    ordinal: int
    identifier: str


def load_snapshot(
    document: str | Path, *, limits: PackageLimits = DEFAULT_LIMITS
) -> DocumentSnapshot:
    """Return a bounded private snapshot of a Word OOXML document or template.

    No document code, fields, links, or external relationships are executed or
    followed. Expected failure messages intentionally omit filenames and package
    contents so they remain safe to use in automation logs.
    """

    try:
        path = Path(document)
        suffix = path.suffix.casefold()
    except (TypeError, ValueError):
        raise DocumentFormatError("document path is invalid") from None
    if suffix not in {".docx", ".docm", ".dotx", ".dotm"}:
        raise DocumentFormatError(
            "only DOCX, DOCM, DOTX, and DOTM packages are supported"
        )

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
            package_thumbnails, package_thumbnail_parts = _package_thumbnail_inventory(
                archive,
                members,
                content_types,
                relationship_maps,
                limits,
            )
            markup_compatibility = _markup_compatibility_inventory(
                archive,
                members,
                relationship_maps,
                limits,
            )
            sensitivity_labels, sensitivity_label_parts = _sensitivity_label_inventory(
                archive, members, content_types, relationship_maps, limits
            )
            (
                package_digital_signatures,
                package_digital_signature_parts,
            ) = _package_digital_signature_inventory(
                archive, members, content_types, relationship_maps, limits
            )
            package_signature_coverage = _package_signature_coverage_inventory(
                archive,
                members,
                content_types,
                relationship_maps,
                package_digital_signature_parts,
                limits,
            )
            (
                settings_enabled,
                settings_signature,
                document_settings_parts,
                document_settings_part_scopes,
            ) = _settings_inventory(archive, members, relationship_maps, limits)
            word_protection = _word_protection_inventory(
                archive,
                members,
                relationship_maps,
                document_settings_parts,
                limits,
            )
            (
                word_document_variables,
                stored_document_variables,
            ) = _word_document_variable_inventory(
                archive,
                members,
                relationship_maps,
                document_settings_parts,
                document_settings_part_scopes,
                limits,
            )
            word_permission_ranges = _word_permission_range_inventory(
                archive,
                members,
                content_types,
                relationship_maps,
                limits,
            )
            mail_merge, mail_merge_parts = _mail_merge_inventory(
                archive, members, relationship_maps, limits
            )
            save_through_xslt = _save_through_xslt_inventory(
                archive,
                members,
                relationship_maps,
                document_settings_parts,
                limits,
            )
            attached_custom_xml_schemas = _attached_custom_xml_schema_inventory(
                archive,
                members,
                relationship_maps,
                document_settings_parts,
                limits,
            )
            field_updates_on_open = _field_update_on_open_inventory(
                archive,
                members,
                document_settings_parts,
                limits,
            )
            template_style_updates_on_open = _template_style_update_on_open_inventory(
                archive,
                members,
                document_settings_parts,
                limits,
            )
            personal_information_removal_on_save = (
                _personal_information_removal_on_save_inventory(
                    archive,
                    members,
                    document_settings_parts,
                    limits,
                )
            )
            save_forms_data = _save_forms_data_inventory(
                archive,
                members,
                document_settings_parts,
                limits,
            )
            save_preview_picture = _save_preview_picture_inventory(
                archive,
                members,
                document_settings_parts,
                limits,
            )
            styles = _styles_inventory(archive, members, relationship_maps, limits)
            (
                stories,
                comment_count,
                content_control_lock_references,
                data_binding_references,
                external_field_references,
                document_variable_field_references,
                hyperlink_field_references,
                hyperlink_markup_references,
                drawing_hyperlink_references,
                drawing_visibility_references,
                drawing_linked_picture_references,
                vml_hyperlink_references,
                vml_external_image_references,
                vml_image_hyperlink_references,
                vml_linked_ole_object_references,
                vml_embedded_ole_object_references,
                word_object_link_references,
                word_embedded_control_references,
                web_extension_control_references,
            ) = _story_snapshots(
                archive, members, content_types, relationship_maps, limits
            )
            content_control_locks = _content_control_lock_inventory(
                content_control_lock_references
            )
            data_bindings = _data_binding_inventory(
                archive,
                members,
                data_binding_references,
                relationship_maps,
                limits,
            )
            external_fields = _external_field_inventory(external_field_references)
            word_document_variable_fields = _word_document_variable_field_inventory(
                document_variable_field_references,
                stored_document_variables,
            )
            word_hyperlink_fields = _word_hyperlink_field_inventory(
                hyperlink_field_references
            )
            word_hyperlink_markup = _word_hyperlink_markup_inventory(
                hyperlink_markup_references
            )
            word_drawing_hyperlinks = _word_drawing_hyperlink_inventory(
                drawing_hyperlink_references
            )
            word_drawing_visibility = _word_drawing_visibility_inventory(
                drawing_visibility_references
            )
            word_drawing_linked_pictures = _word_drawing_linked_picture_inventory(
                drawing_linked_picture_references
            )
            word_vml_hyperlinks = _word_vml_hyperlink_inventory(
                vml_hyperlink_references
            )
            word_vml_external_images = _word_vml_external_image_inventory(
                vml_external_image_references
            )
            word_vml_image_hyperlinks = _word_vml_image_hyperlink_inventory(
                vml_image_hyperlink_references
            )
            word_vml_linked_ole_objects = _word_vml_linked_ole_object_inventory(
                vml_linked_ole_object_references
            )
            word_vml_embedded_ole_objects = _word_vml_embedded_ole_object_inventory(
                vml_embedded_ole_object_references
            )
            word_object_links = _word_object_link_inventory(word_object_link_references)
            word_embedded_controls = _word_embedded_control_inventory(
                word_embedded_control_references
            )
            modern_comment_metadata, modern_comment_metadata_parts = (
                _modern_comment_metadata_inventory(
                    archive,
                    members,
                    content_types,
                    relationship_maps,
                    limits,
                )
            )
            document_tasks, document_task_parts = _document_task_inventory(
                archive,
                members,
                content_types,
                relationship_maps,
                limits,
            )
            taskpane_web_extensions, taskpane_web_extension_parts = (
                _taskpane_web_extension_inventory(
                    archive,
                    members,
                    content_types,
                    relationship_maps,
                    web_extension_control_references,
                    limits,
                )
            )
            (
                external_document_dependencies,
                external_document_dependency_parts,
            ) = _external_document_dependency_inventory(
                archive,
                members,
                relationship_maps,
                limits,
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
                    | package_thumbnail_parts
                    | sensitivity_label_parts
                    | package_digital_signature_parts
                    | document_settings_parts
                    | mail_merge_parts
                    | modern_comment_metadata_parts
                    | document_task_parts
                    | taskpane_web_extension_parts
                    | external_document_dependency_parts
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
        package_thumbnails=package_thumbnails,
        markup_compatibility=markup_compatibility,
        sensitivity_labels=sensitivity_labels,
        package_digital_signatures=package_digital_signatures,
        package_signature_coverage=package_signature_coverage,
        word_protection=word_protection,
        word_document_variables=word_document_variables,
        word_document_variable_fields=word_document_variable_fields,
        word_hyperlink_fields=word_hyperlink_fields,
        word_hyperlink_markup=word_hyperlink_markup,
        word_drawing_hyperlinks=word_drawing_hyperlinks,
        word_drawing_visibility=word_drawing_visibility,
        word_drawing_linked_pictures=word_drawing_linked_pictures,
        word_vml_hyperlinks=word_vml_hyperlinks,
        word_vml_external_images=word_vml_external_images,
        word_vml_image_hyperlinks=word_vml_image_hyperlinks,
        word_vml_linked_ole_objects=word_vml_linked_ole_objects,
        word_vml_embedded_ole_objects=word_vml_embedded_ole_objects,
        word_object_links=word_object_links,
        word_embedded_controls=word_embedded_controls,
        word_permission_ranges=word_permission_ranges,
        mail_merge=mail_merge,
        save_through_xslt=save_through_xslt,
        attached_custom_xml_schemas=attached_custom_xml_schemas,
        field_updates_on_open=field_updates_on_open,
        template_style_updates_on_open=template_style_updates_on_open,
        personal_information_removal_on_save=personal_information_removal_on_save,
        save_forms_data=save_forms_data,
        save_preview_picture=save_preview_picture,
        content_control_locks=content_control_locks,
        data_bindings=data_bindings,
        external_fields=external_fields,
        modern_comment_metadata=modern_comment_metadata,
        document_tasks=document_tasks,
        taskpane_web_extensions=taskpane_web_extensions,
        external_document_dependencies=external_document_dependencies,
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
    defaults: dict[str, str] = {}
    for child in root:
        local_name = _local_name(child.tag)
        if local_name == "Default":
            extension = child.attrib.get("Extension")
            content_type = child.attrib.get("ContentType")
            normalized_extension = (extension or "").casefold()
            if (
                not normalized_extension
                or "." in normalized_extension
                or not content_type
                or normalized_extension in defaults
            ):
                raise DocumentFormatError("package content types are invalid")
            defaults[normalized_extension] = content_type
            continue
        if local_name != "Override":
            raise DocumentFormatError("package content types are invalid")
        part_name = child.attrib.get("PartName")
        content_type = child.attrib.get("ContentType")
        if not part_name or not content_type or not part_name.startswith("/"):
            raise DocumentFormatError("package content types are invalid")
        member_name = part_name[1:]
        if member_name not in members or member_name in overrides:
            raise DocumentFormatError("package content types are invalid")
        overrides[member_name] = content_type
    for member_name in members:
        if member_name in overrides:
            continue
        extension = Path(member_name).suffix.removeprefix(".").casefold()
        content_type = defaults.get(extension)
        if content_type is not None:
            overrides[member_name] = content_type
    return overrides


def _validate_main_document(content_types: dict[str, str]) -> None:
    content_type = content_types.get("word/document.xml", "").casefold()
    if not any(
        content_type.endswith(suffix)
        for suffix in _SUPPORTED_MAIN_DOCUMENT_CONTENT_TYPE_SUFFIXES
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


def _package_thumbnail_inventory(
    archive: zipfile.ZipFile,
    members: dict[str, zipfile.ZipInfo],
    content_types: dict[str, str],
    relationship_maps: dict[str, dict[str, _Relationship]],
    limits: PackageLimits,
) -> tuple[PackageThumbnailInventory, frozenset[str]]:
    """Inventory only standard relationship-bound OPC thumbnail image parts.

    Thumbnail image bytes remain opaque.  A conventional filename alone is not
    enough: the package must declare one exact thumbnail relationship per
    source, the relationship must be internal, and the target must have an
    image content type and no relationship part of its own.
    """

    records: list[tuple[str, ...]] = []
    part_names: set[str] = set()
    relationship_counts: dict[str, int] = {}
    for source_part, relationships in sorted(relationship_maps.items()):
        for relationship in relationships.values():
            if (
                relationship.relationship_type.casefold()
                not in _PACKAGE_THUMBNAIL_RELATIONSHIP_TYPES
            ):
                continue
            if source_part != "/" and source_part not in members:
                raise DocumentFormatError("package thumbnail relationship is invalid")
            relationship_counts[source_part] = (
                relationship_counts.get(source_part, 0) + 1
            )
            if relationship_counts[source_part] > 1:
                raise DocumentFormatError("package thumbnail relationship is invalid")
            if relationship.target_mode.casefold() != "internal":
                raise DocumentFormatError("package thumbnail relationship is invalid")
            target = _internal_relationship_target(source_part, relationship, members)
            if target is None:
                raise DocumentFormatError("package thumbnail relationship is invalid")
            content_type = content_types.get(target, "").casefold()
            if not content_type.startswith("image/") or target in relationship_maps:
                raise DocumentFormatError("package thumbnail part is invalid")
            records.append(
                (
                    "package_thumbnail",
                    source_part,
                    content_type,
                    *relationship.canonical_value(),
                )
            )
            part_names.add(target)

    return (
        PackageThumbnailInventory(
            thumbnail_relationship_count=len(records),
            thumbnail_part_count=len(part_names),
            signature=_payload_inventory_signature(
                records, frozenset(part_names), archive, members, limits
            ),
        ),
        frozenset(part_names),
    )


def _markup_compatibility_inventory(
    archive: zipfile.ZipFile,
    members: dict[str, zipfile.ZipInfo],
    relationship_maps: dict[str, dict[str, _Relationship]],
    limits: PackageLimits,
) -> MarkupCompatibilityInventory:
    """Inventory OOXML Markup Compatibility markup without preprocessing it.

    The inventory covers stored XML members under ``word/`` only. It records
    aggregate MCE branch and compatibility-rule evidence without selecting an
    AlternateContent branch, resolving feature prefixes, or rendering content.
    """

    counts = {
        "alternate_content": 0,
        "choice": 0,
        "fallback": 0,
        "choice_requires_prefix": 0,
        "ignorable_prefix": 0,
        "must_understand_prefix": 0,
        "process_content_name": 0,
        "preserve_element_name": 0,
        "preserve_attribute_name": 0,
    }
    records: list[tuple[str, ...]] = []
    part_count = 0

    for part_name in _markup_compatibility_candidate_parts(members):
        payload = _read_member(archive, members[part_name], limits)
        if _MARKUP_COMPATIBILITY_MARKER not in payload:
            continue
        root = _parse_xml(payload, limits)
        relationships = relationship_maps.get(part_name, {})
        part_has_markup_compatibility = False

        for element in root.iter():
            namespace, local_name = _qualified_name(element.tag)
            if (
                namespace == _MARKUP_COMPATIBILITY_NAMESPACE
                and local_name in _MARKUP_COMPATIBILITY_ELEMENT_NAMES
            ):
                part_has_markup_compatibility = True
                records.append(
                    (
                        "markup_compatibility_element",
                        part_name,
                        local_name,
                        _fingerprint_element(element, relationships),
                    )
                )
                if local_name == "AlternateContent":
                    counts["alternate_content"] += 1
                elif local_name == "Choice":
                    counts["choice"] += 1
                    counts["choice_requires_prefix"] += _token_count(
                        _unqualified_attribute_value(element, "Requires")
                    )
                else:
                    counts["fallback"] += 1

            for attribute, value in element.attrib.items():
                attribute_namespace, attribute_local_name = _qualified_name(attribute)
                if (
                    attribute_namespace != _MARKUP_COMPATIBILITY_NAMESPACE
                    or attribute_local_name not in _MARKUP_COMPATIBILITY_ATTRIBUTE_NAMES
                ):
                    continue
                part_has_markup_compatibility = True
                records.append(
                    (
                        "markup_compatibility_attribute",
                        part_name,
                        element.tag,
                        attribute_local_name,
                        value,
                    )
                )
                if attribute_local_name == "Ignorable":
                    counts["ignorable_prefix"] += _token_count(value)
                elif attribute_local_name == "MustUnderstand":
                    counts["must_understand_prefix"] += _token_count(value)
                elif attribute_local_name == "ProcessContent":
                    counts["process_content_name"] += _token_count(value)
                elif attribute_local_name == "PreserveElements":
                    counts["preserve_element_name"] += _token_count(value)
                else:
                    counts["preserve_attribute_name"] += _token_count(value)

        if part_has_markup_compatibility:
            part_count += 1

    return MarkupCompatibilityInventory(
        markup_compatibility_part_count=part_count,
        alternate_content_count=counts["alternate_content"],
        choice_count=counts["choice"],
        fallback_count=counts["fallback"],
        choice_requires_prefix_count=counts["choice_requires_prefix"],
        ignorable_prefix_count=counts["ignorable_prefix"],
        must_understand_prefix_count=counts["must_understand_prefix"],
        process_content_name_count=counts["process_content_name"],
        preserve_element_name_count=counts["preserve_element_name"],
        preserve_attribute_name_count=counts["preserve_attribute_name"],
        signature=_digest_records(records),
    )


def _markup_compatibility_candidate_parts(
    members: dict[str, zipfile.ZipInfo],
) -> list[str]:
    return sorted(
        part_name
        for part_name in members
        if part_name.startswith("word/")
        and "/_rels/" not in part_name
        and part_name.casefold().endswith(".xml")
    )


def _token_count(value: str | None) -> int:
    return len(value.split()) if value is not None else 0


def _sensitivity_label_inventory(
    archive: zipfile.ZipFile,
    members: dict[str, zipfile.ZipInfo],
    content_types: dict[str, str],
    relationship_maps: dict[str, dict[str, _Relationship]],
    limits: PackageLimits,
) -> tuple[SensitivityLabelInventory, frozenset[str]]:
    """Inventory stored Office sensitivity-label metadata without revealing it.

    Current Office files can store label state in an Office 2021 LabelInfo part,
    while older and coauthoring-compatible flows retain MIP key/value metadata
    in custom document properties.  Tenant IDs, label IDs, labels names, dates,
    action IDs, arbitrary MIP extension attributes, and content-marking text
    remain in the private signature.
    """

    records: list[tuple[str, ...]] = []
    label_info_parts: set[str] = set()

    def add_label_info_part(part_name: str) -> None:
        label_info_parts.add(part_name)
        if len(label_info_parts) > 1:
            raise DocumentFormatError("sensitivity label information parts are invalid")

    for part_name in members:
        if _is_conventional_sensitivity_label_part_name(part_name):
            add_label_info_part(part_name)
    for part_name, content_type in content_types.items():
        if content_type.casefold() in _SENSITIVITY_LABEL_PART_CONTENT_TYPES:
            add_label_info_part(part_name)

    for source_part, relationships in sorted(relationship_maps.items()):
        for relationship in sorted(
            relationships.values(), key=lambda value: value.canonical_value()
        ):
            if (
                relationship.relationship_type.casefold()
                != _SENSITIVITY_LABEL_RELATIONSHIP_TYPE
            ):
                continue
            if source_part != "/" or relationship.target_mode.casefold() != "internal":
                raise DocumentFormatError("sensitivity label relationships are invalid")
            target = _internal_relationship_target(source_part, relationship, members)
            if target is None:
                raise DocumentFormatError("sensitivity label relationships are invalid")
            add_label_info_part(target)
            records.append(
                (
                    "sensitivity_label_relationship",
                    source_part,
                    *relationship.canonical_value(),
                )
            )

    counts = {
        "label_info_label": 0,
        "label_info_enabled_label": 0,
        "label_info_removed_label": 0,
        "label_info_extension": 0,
        "legacy_mip_property": 0,
        "legacy_sensitivity_property": 0,
        "word_content_marking_property": 0,
    }
    for part_name in sorted(label_info_parts):
        root = _read_xml(archive, members[part_name], limits)
        labels, extension_count = _validate_sensitivity_label_root(root)
        records.append(
            (
                "sensitivity_label_info_part",
                part_name,
                _fingerprint_element(
                    root,
                    relationship_maps.get(part_name, {}),
                    preserve_volatile_attributes=True,
                ),
            )
        )
        counts["label_info_label"] += len(labels)
        counts["label_info_extension"] += extension_count
        for label in labels:
            enabled = _sensitivity_label_boolean(label.attrib["enabled"])
            removed = _sensitivity_label_boolean(label.attrib["removed"])
            if removed:
                counts["label_info_removed_label"] += 1
            elif enabled:
                counts["label_info_enabled_label"] += 1

    legacy_mip_label_ids: set[str] = set()
    for part_name in _custom_document_property_part_names(members, relationship_maps):
        root = _read_xml(archive, members[part_name], limits)
        _validate_document_property_root(root, "custom")
        relationships = relationship_maps.get(part_name, {})
        for property_element in root:
            if _local_name(property_element.tag) != "property":
                continue
            property_name = property_element.attrib.get("name")
            if property_name is None:
                continue
            category = _sensitivity_label_custom_property_category(property_name)
            if category is None:
                continue
            records.append(
                (
                    "sensitivity_label_custom_property",
                    category,
                    part_name,
                    _fingerprint_element(property_element, relationships),
                )
            )
            if category == "legacy_mip":
                counts["legacy_mip_property"] += 1
                match = _LEGACY_MIP_LABEL_PROPERTY_NAME.fullmatch(property_name)
                if match is None:
                    raise DocumentFormatError("sensitivity label metadata is invalid")
                legacy_mip_label_ids.add(match.group("label_id").strip("{}").casefold())
            elif category == "legacy_sensitivity":
                counts["legacy_sensitivity_property"] += 1
            else:
                counts["word_content_marking_property"] += 1

    return (
        SensitivityLabelInventory(
            label_info_part_count=len(label_info_parts),
            label_info_label_count=counts["label_info_label"],
            label_info_enabled_label_count=counts["label_info_enabled_label"],
            label_info_removed_label_count=counts["label_info_removed_label"],
            label_info_extension_count=counts["label_info_extension"],
            legacy_mip_label_count=len(legacy_mip_label_ids),
            legacy_mip_property_count=counts["legacy_mip_property"],
            legacy_sensitivity_property_count=counts["legacy_sensitivity_property"],
            word_content_marking_property_count=counts["word_content_marking_property"],
            signature=_digest_records(records),
        ),
        frozenset(label_info_parts),
    )


def _is_conventional_sensitivity_label_part_name(part_name: str) -> bool:
    return part_name.casefold() in _SENSITIVITY_LABEL_FALLBACK_PART_NAMES


def _validate_sensitivity_label_root(
    root: ET.Element,
) -> tuple[list[ET.Element], int]:
    if _qualified_name(root.tag) != (_SENSITIVITY_LABEL_NAMESPACE, "labelList"):
        raise DocumentFormatError("sensitivity label information part is invalid")

    labels: list[ET.Element] = []
    extension_count = 0
    saw_extension_list = False
    for child in root:
        qualified_name = _qualified_name(child.tag)
        if qualified_name == (_SENSITIVITY_LABEL_NAMESPACE, "label"):
            if saw_extension_list:
                raise DocumentFormatError(
                    "sensitivity label information part is invalid"
                )
            _validate_sensitivity_label(child)
            labels.append(child)
        elif qualified_name == (_SENSITIVITY_LABEL_NAMESPACE, "extLst"):
            if saw_extension_list:
                raise DocumentFormatError(
                    "sensitivity label information part is invalid"
                )
            saw_extension_list = True
            for extension in child:
                if (
                    _qualified_name(extension.tag)
                    != (
                        _SENSITIVITY_LABEL_NAMESPACE,
                        "ext",
                    )
                    or extension.attrib.get("uri") is None
                ):
                    raise DocumentFormatError(
                        "sensitivity label information part is invalid"
                    )
                extension_count += 1
        else:
            raise DocumentFormatError("sensitivity label information part is invalid")
    return labels, extension_count


def _validate_sensitivity_label(label: ET.Element) -> None:
    required_attributes = ("id", "enabled", "method", "siteId", "removed")
    if any(attribute not in label.attrib for attribute in required_attributes) or list(
        label
    ):
        raise DocumentFormatError("sensitivity label information part is invalid")
    if _SENSITIVITY_LABEL_GUID.fullmatch(label.attrib["siteId"]) is None:
        raise DocumentFormatError("sensitivity label information part is invalid")
    _sensitivity_label_boolean(label.attrib["enabled"])
    _sensitivity_label_boolean(label.attrib["removed"])
    content_bits = label.attrib.get("contentBits")
    if content_bits is not None:
        try:
            content_bits_value = int(content_bits)
        except ValueError:
            raise DocumentFormatError(
                "sensitivity label information part is invalid"
            ) from None
        if content_bits_value < 0 or content_bits_value > 0xFFFFFFFF:
            raise DocumentFormatError("sensitivity label information part is invalid")


def _sensitivity_label_boolean(value: str) -> bool:
    normalized = value.casefold()
    if normalized in {"1", "true"}:
        return True
    if normalized in {"0", "false"}:
        return False
    raise DocumentFormatError("sensitivity label information part is invalid")


def _custom_document_property_part_names(
    members: dict[str, zipfile.ZipInfo],
    relationship_maps: dict[str, dict[str, _Relationship]],
) -> list[str]:
    part_names = {
        part_name
        for part_name, kind in _DOCUMENT_PROPERTY_FALLBACK_PARTS.items()
        if kind == "custom" and part_name in members
    }
    for source_part, relationships in relationship_maps.items():
        for relationship in relationships.values():
            if (
                _DOCUMENT_PROPERTY_RELATIONSHIP_TYPES.get(
                    relationship.relationship_type.casefold()
                )
                != "custom"
            ):
                continue
            target = _internal_relationship_target(source_part, relationship, members)
            if target is not None:
                part_names.add(target)
    return sorted(part_names)


def _sensitivity_label_custom_property_category(property_name: str) -> str | None:
    if _LEGACY_MIP_LABEL_PROPERTY_NAME.fullmatch(property_name) is not None:
        return "legacy_mip"
    if property_name.casefold() == "sensitivity":
        return "legacy_sensitivity"
    if _WORD_SENSITIVITY_CONTENT_MARKING_PROPERTY_NAME.fullmatch(property_name):
        return "word_content_marking"
    return None


def _package_digital_signature_inventory(
    archive: zipfile.ZipFile,
    members: dict[str, zipfile.ZipInfo],
    content_types: dict[str, str],
    relationship_maps: dict[str, dict[str, _Relationship]],
    limits: PackageLimits,
) -> tuple[PackageDigitalSignatureInventory, frozenset[str]]:
    """Inventory OPC signature material without making a trust decision.

    OPC package signatures can carry signer, certificate, signing-time, comment,
    and provider material. This inventory deliberately reports only structural
    counts and retains raw part digests privately. It validates recognized
    package topology and a small XMLDSIG shape, but it does not verify a
    cryptographic signature, certificate chain, timestamp, signing policy, or
    signed-content coverage.
    """

    records: list[tuple[str, ...]] = []
    origin_parts: set[str] = set()
    xml_signature_parts: set[str] = set()
    certificate_parts: set[str] = set()

    def add_origin_part(part_name: str) -> None:
        origin_parts.add(part_name)
        if len(origin_parts) > 1:
            raise DocumentFormatError("package digital signature origins are invalid")

    for part_name in members:
        if (
            part_name.casefold()
            in _PACKAGE_DIGITAL_SIGNATURE_ORIGIN_FALLBACK_PART_NAMES
        ):
            add_origin_part(part_name)

    content_type_kinds = {
        _PACKAGE_DIGITAL_SIGNATURE_ORIGIN_CONTENT_TYPE: "origin",
        _PACKAGE_DIGITAL_SIGNATURE_XML_CONTENT_TYPE: "xml_signature",
        _PACKAGE_DIGITAL_SIGNATURE_CERTIFICATE_CONTENT_TYPE: "certificate",
    }
    for part_name, content_type in sorted(content_types.items()):
        kind = content_type_kinds.get(content_type.casefold())
        if kind is None:
            continue
        records.append(
            (
                "package_digital_signature_content_type",
                kind,
                part_name,
                content_type,
            )
        )
        if kind == "origin":
            add_origin_part(part_name)
        elif kind == "xml_signature":
            xml_signature_parts.add(part_name)
        else:
            certificate_parts.add(part_name)

    origin_relationship_count = 0
    for source_part, relationships in sorted(relationship_maps.items()):
        for relationship in sorted(
            relationships.values(), key=lambda value: value.canonical_value()
        ):
            if (
                relationship.relationship_type.casefold()
                != _PACKAGE_DIGITAL_SIGNATURE_ORIGIN_RELATIONSHIP_TYPE
            ):
                continue
            origin_relationship_count += 1
            if source_part != "/" or relationship.target_mode.casefold() != "internal":
                raise DocumentFormatError(
                    "package digital signature relationships are invalid"
                )
            target = _internal_relationship_target(source_part, relationship, members)
            if (
                target is None
                or content_types.get(target, "").casefold()
                != _PACKAGE_DIGITAL_SIGNATURE_ORIGIN_CONTENT_TYPE
            ):
                raise DocumentFormatError(
                    "package digital signature relationships are invalid"
                )
            add_origin_part(target)
            records.append(
                (
                    "package_digital_signature_origin_relationship",
                    source_part,
                    *relationship.canonical_value(),
                )
            )
    if origin_relationship_count > 1:
        raise DocumentFormatError("package digital signature origins are invalid")

    for source_part, relationships in sorted(relationship_maps.items()):
        for relationship in sorted(
            relationships.values(), key=lambda value: value.canonical_value()
        ):
            if (
                relationship.relationship_type.casefold()
                != _PACKAGE_DIGITAL_SIGNATURE_RELATIONSHIP_TYPE
            ):
                continue
            if (
                source_part not in origin_parts
                or relationship.target_mode.casefold() != "internal"
            ):
                raise DocumentFormatError(
                    "package digital signature relationships are invalid"
                )
            target = _internal_relationship_target(source_part, relationship, members)
            if (
                target is None
                or content_types.get(target, "").casefold()
                != _PACKAGE_DIGITAL_SIGNATURE_XML_CONTENT_TYPE
            ):
                raise DocumentFormatError(
                    "package digital signature relationships are invalid"
                )
            xml_signature_parts.add(target)
            records.append(
                (
                    "package_digital_signature_relationship",
                    source_part,
                    *relationship.canonical_value(),
                )
            )

    for source_part, relationships in sorted(relationship_maps.items()):
        for relationship in sorted(
            relationships.values(), key=lambda value: value.canonical_value()
        ):
            if (
                relationship.relationship_type.casefold()
                != _PACKAGE_DIGITAL_SIGNATURE_CERTIFICATE_RELATIONSHIP_TYPE
            ):
                continue
            if (
                source_part not in xml_signature_parts
                or relationship.target_mode.casefold() != "internal"
            ):
                raise DocumentFormatError(
                    "package digital signature relationships are invalid"
                )
            target = _internal_relationship_target(source_part, relationship, members)
            if (
                target is None
                or content_types.get(target, "").casefold()
                != _PACKAGE_DIGITAL_SIGNATURE_CERTIFICATE_CONTENT_TYPE
            ):
                raise DocumentFormatError(
                    "package digital signature relationships are invalid"
                )
            certificate_parts.add(target)
            records.append(
                (
                    "package_digital_signature_certificate_relationship",
                    source_part,
                    *relationship.canonical_value(),
                )
            )

    counts = {
        "signed_info_reference": 0,
        "manifest_reference": 0,
        "relationship_reference": 0,
        "inline_x509_certificate": 0,
        "signature_property": 0,
    }
    for part_name in sorted(origin_parts):
        records.append(
            (
                "package_digital_signature_origin_part",
                part_name,
                _digest_bytes(_read_member(archive, members[part_name], limits)),
            )
        )
    for part_name in sorted(xml_signature_parts):
        payload = _read_member(archive, members[part_name], limits)
        root = _parse_xml(payload, limits)
        signature_counts = _validate_package_digital_signature_root(root)
        for key, value in signature_counts.items():
            counts[key] += value
        records.append(
            (
                "package_digital_signature_xml_part",
                part_name,
                _digest_bytes(payload),
            )
        )
    for part_name in sorted(certificate_parts):
        records.append(
            (
                "package_digital_signature_certificate_part",
                part_name,
                _digest_bytes(_read_member(archive, members[part_name], limits)),
            )
        )

    signature_parts = frozenset(origin_parts | xml_signature_parts | certificate_parts)
    return (
        PackageDigitalSignatureInventory(
            signature_origin_part_count=len(origin_parts),
            xml_signature_part_count=len(xml_signature_parts),
            certificate_part_count=len(certificate_parts),
            signed_info_reference_count=counts["signed_info_reference"],
            manifest_reference_count=counts["manifest_reference"],
            relationship_reference_count=counts["relationship_reference"],
            inline_x509_certificate_count=counts["inline_x509_certificate"],
            signature_property_count=counts["signature_property"],
            signature=_digest_records(records),
        ),
        signature_parts,
    )


def _package_signature_coverage_inventory(
    archive: zipfile.ZipFile,
    members: dict[str, zipfile.ZipInfo],
    content_types: dict[str, str],
    relationship_maps: dict[str, dict[str, _Relationship]],
    package_digital_signature_parts: frozenset[str],
    limits: PackageLimits,
) -> PackageSignatureCoverageInventory:
    """Resolve declared local OPC coverage without verifying a signature.

    XMLDSIG verification, certificate trust, and Office rendering behavior are
    intentionally outside this boundary.  The inventory identifies only an
    XML signature whose ``SignedInfo`` directly references an in-package
    ``ds:Object`` containing a direct package ``ds:Manifest``.  From that
    object it resolves exact local part references and selected relationship
    IDs/types.  Any unknown syntax is reported as aggregate indeterminate
    evidence instead of being treated as coverage.
    """

    signature_part_names = sorted(
        part_name
        for part_name in package_digital_signature_parts
        if content_types.get(part_name, "").casefold()
        == _PACKAGE_DIGITAL_SIGNATURE_XML_CONTENT_TYPE.casefold()
    )
    if not signature_part_names:
        return PackageSignatureCoverageInventory(
            signature_with_declared_package_coverage_count=0,
            signature_without_declared_package_coverage_count=0,
            declared_covered_word_part_count=0,
            declared_uncovered_word_part_count=0,
            declared_covered_root_document_relationship_count=0,
            declared_uncovered_root_document_relationship_count=0,
            declared_covered_word_relationship_count=0,
            declared_uncovered_word_relationship_count=0,
            unresolved_package_manifest_reference_count=0,
            unsupported_package_manifest_reference_count=0,
            signature=_digest_records([]),
        )

    records: list[tuple[str, ...]] = []
    covered_part_names: set[str] = set()
    covered_relationship_ids: set[tuple[str, str]] = set()
    signature_with_declared_package_coverage_count = 0
    signature_without_declared_package_coverage_count = 0
    unresolved_package_manifest_reference_count = 0
    unsupported_package_manifest_reference_count = 0

    for part_name in signature_part_names:
        root = _parse_xml(_read_member(archive, members[part_name], limits), limits)
        declared_coverage = _declared_package_signature_coverage(
            root,
            members,
            content_types,
            relationship_maps,
        )
        if declared_coverage.has_declared_package_coverage:
            signature_with_declared_package_coverage_count += 1
        else:
            signature_without_declared_package_coverage_count += 1
        covered_part_names.update(declared_coverage.covered_part_names)
        covered_relationship_ids.update(declared_coverage.covered_relationship_ids)
        unresolved_package_manifest_reference_count += (
            declared_coverage.unresolved_reference_count
        )
        unsupported_package_manifest_reference_count += (
            declared_coverage.unsupported_reference_count
        )
        records.append(
            (
                "package_signature_coverage_signature_part",
                part_name,
                "declared"
                if declared_coverage.has_declared_package_coverage
                else "none",
            )
        )
        records.extend(declared_coverage.records)

    word_part_names = {
        part_name
        for part_name in members
        if part_name.startswith("word/")
        and "/_rels/" not in part_name
        and not part_name.endswith(".rels")
    }
    root_document_relationship_ids = {
        ("/", relationship_id)
        for relationship_id, relationship in relationship_maps.get("/", {}).items()
        if relationship.relationship_type in _ROOT_OFFICE_DOCUMENT_RELATIONSHIP_TYPES
    }
    word_relationship_ids = {
        (source_part, relationship_id)
        for source_part, relationships in relationship_maps.items()
        if source_part.startswith("word/")
        for relationship_id in relationships
    }
    covered_root_document_relationship_ids = (
        root_document_relationship_ids & covered_relationship_ids
    )
    covered_word_relationship_ids = word_relationship_ids & covered_relationship_ids

    for part_name in sorted(word_part_names):
        records.append(
            (
                "package_signature_coverage_word_part",
                part_name,
                "covered" if part_name in covered_part_names else "uncovered",
            )
        )
    for source_part, relationship_id in sorted(root_document_relationship_ids):
        records.append(
            (
                "package_signature_coverage_root_relationship",
                source_part,
                relationship_id,
                "covered"
                if (source_part, relationship_id)
                in covered_root_document_relationship_ids
                else "uncovered",
            )
        )
    for source_part, relationship_id in sorted(word_relationship_ids):
        records.append(
            (
                "package_signature_coverage_word_relationship",
                source_part,
                relationship_id,
                "covered"
                if (source_part, relationship_id) in covered_word_relationship_ids
                else "uncovered",
            )
        )

    return PackageSignatureCoverageInventory(
        signature_with_declared_package_coverage_count=(
            signature_with_declared_package_coverage_count
        ),
        signature_without_declared_package_coverage_count=(
            signature_without_declared_package_coverage_count
        ),
        declared_covered_word_part_count=len(word_part_names & covered_part_names),
        declared_uncovered_word_part_count=len(word_part_names - covered_part_names),
        declared_covered_root_document_relationship_count=len(
            covered_root_document_relationship_ids
        ),
        declared_uncovered_root_document_relationship_count=len(
            root_document_relationship_ids - covered_relationship_ids
        ),
        declared_covered_word_relationship_count=len(covered_word_relationship_ids),
        declared_uncovered_word_relationship_count=len(
            word_relationship_ids - covered_relationship_ids
        ),
        unresolved_package_manifest_reference_count=(
            unresolved_package_manifest_reference_count
        ),
        unsupported_package_manifest_reference_count=(
            unsupported_package_manifest_reference_count
        ),
        signature=_digest_records(records),
    )


def _declared_package_signature_coverage(
    root: ET.Element,
    members: dict[str, zipfile.ZipInfo],
    content_types: dict[str, str],
    relationship_maps: dict[str, dict[str, _Relationship]],
) -> _DeclaredPackageSignatureCoverage:
    """Resolve one signature's single bounded OPC package-specific object."""

    signed_info = next(
        child
        for child in root
        if _qualified_name(child.tag) == (_XMLDSIG_NAMESPACE, "SignedInfo")
    )
    records: list[tuple[str, ...]] = []
    for reference in signed_info:
        if _qualified_name(reference.tag) != (_XMLDSIG_NAMESPACE, "Reference"):
            continue
        records.append(
            (
                "package_signature_coverage_signed_info_reference",
                _digest_bytes(ET.tostring(reference, encoding="utf-8")),
            )
        )

    manifest = _package_specific_object_manifest(root)
    package_object_references = [
        reference
        for reference in signed_info
        if _qualified_name(reference.tag) == (_XMLDSIG_NAMESPACE, "Reference")
        and _signature_fragment_identifier(reference.attrib.get("URI"))
        == _OPC_PACKAGE_SPECIFIC_OBJECT_ID
    ]
    if manifest is None or len(package_object_references) != 1:
        return _DeclaredPackageSignatureCoverage(
            has_declared_package_coverage=False,
            covered_part_names=frozenset(),
            covered_relationship_ids=frozenset(),
            unresolved_reference_count=0,
            unsupported_reference_count=0,
            records=tuple(records),
        )

    covered_part_names: set[str] = set()
    covered_relationship_ids: set[tuple[str, str]] = set()
    unresolved_reference_count = 0
    unsupported_reference_count = 0

    # OPC limits a relationships part to one relationship transform per XML
    # signature. Count this single package manifest before resolving its
    # references, so a duplicate never contributes partial coverage.
    relationship_transform_counts_by_part: dict[str, int] = {}
    for reference in manifest:
        if _qualified_name(reference.tag) != (_XMLDSIG_NAMESPACE, "Reference"):
            continue
        parsed_reference = _package_manifest_reference_member_name(
            reference.attrib.get("URI")
        )
        if parsed_reference is None:
            continue
        part_name, _ = parsed_reference
        if not part_name.endswith(".rels"):
            continue
        relationship_transform_count = _relationship_transform_count(reference)
        if relationship_transform_count:
            relationship_transform_counts_by_part[part_name] = (
                relationship_transform_counts_by_part.get(part_name, 0)
                + relationship_transform_count
            )
    duplicate_relationship_transform_part_names = frozenset(
        part_name
        for part_name, relationship_transform_count in (
            relationship_transform_counts_by_part.items()
        )
        if relationship_transform_count > 1
    )

    records.append(
        (
            "package_signature_coverage_bound_manifest_object",
            _digest_bytes(_OPC_PACKAGE_SPECIFIC_OBJECT_ID.encode("utf-8")),
        )
    )
    for reference in manifest:
        if _qualified_name(reference.tag) != (_XMLDSIG_NAMESPACE, "Reference"):
            unsupported_reference_count += 1
            records.append(
                (
                    "package_signature_coverage_unsupported_manifest_child",
                    _digest_bytes(ET.tostring(reference, encoding="utf-8")),
                )
            )
            continue
        parsed_reference = _package_manifest_reference_member_name(
            reference.attrib.get("URI")
        )
        relationship_transform_is_duplicated = (
            parsed_reference is not None
            and parsed_reference[0] in duplicate_relationship_transform_part_names
            and _relationship_transform_count(reference) > 0
        )
        if relationship_transform_is_duplicated:
            resolution = _ManifestReferenceResolution(
                covered_part_name=None,
                covered_relationship_ids=frozenset(),
                unresolved_reference_count=0,
                unsupported_reference_count=1,
            )
        else:
            resolution = _resolve_package_manifest_reference(
                reference,
                members,
                content_types,
                relationship_maps,
            )
        covered_part_names.update(
            part_name
            for part_name in (resolution.covered_part_name,)
            if part_name is not None
        )
        covered_relationship_ids.update(resolution.covered_relationship_ids)
        unresolved_reference_count += resolution.unresolved_reference_count
        unsupported_reference_count += resolution.unsupported_reference_count
        records.append(
            (
                "package_signature_coverage_manifest_reference",
                _digest_bytes(ET.tostring(reference, encoding="utf-8")),
                "resolved"
                if (
                    resolution.covered_part_name is not None
                    or resolution.covered_relationship_ids
                )
                else "indeterminate",
            )
        )

    return _DeclaredPackageSignatureCoverage(
        has_declared_package_coverage=True,
        covered_part_names=frozenset(covered_part_names),
        covered_relationship_ids=frozenset(covered_relationship_ids),
        unresolved_reference_count=unresolved_reference_count,
        unsupported_reference_count=unsupported_reference_count,
        records=tuple(records),
    )


def _package_specific_object_manifest(root: ET.Element) -> ET.Element | None:
    """Return OPC's one exact package-specific object manifest, if present."""

    package_specific_objects = [
        child
        for child in root
        if _qualified_name(child.tag) == (_XMLDSIG_NAMESPACE, "Object")
        and child.attrib.get("Id") == _OPC_PACKAGE_SPECIFIC_OBJECT_ID
    ]
    if len(package_specific_objects) != 1:
        return None

    package_object = package_specific_objects[0]
    children = list(package_object)
    if package_object.attrib != {"Id": _OPC_PACKAGE_SPECIFIC_OBJECT_ID} or [
        _qualified_name(child.tag) for child in children
    ] != [
        (_XMLDSIG_NAMESPACE, "Manifest"),
        (_XMLDSIG_NAMESPACE, "SignatureProperties"),
    ]:
        return None
    manifest, signature_properties = children
    if not _has_valid_opc_signature_time_property(
        signature_properties,
        root.attrib.get("Id"),
    ):
        return None
    return manifest


def _has_valid_opc_signature_time_property(
    signature_properties: ET.Element,
    signature_id: str | None,
) -> bool:
    """Require OPC's fixed package-object signature-time declaration."""

    properties = list(signature_properties)
    if (
        not _has_only_whitespace_interstitial_text(signature_properties)
        or len(properties) != 1
        or _qualified_name(properties[0].tag)
        != (_XMLDSIG_NAMESPACE, "SignatureProperty")
    ):
        return False

    signature_property = properties[0]
    signature_time_children = list(signature_property)
    if (
        signature_property.attrib.get("Id") != _OPC_SIGNATURE_TIME_PROPERTY_ID
        or set(signature_property.attrib) != {"Id", "Target"}
        or not _has_only_whitespace_interstitial_text(signature_property)
        or len(signature_time_children) != 1
        or _qualified_name(signature_time_children[0].tag)
        != (_OPC_DIGITAL_SIGNATURE_NAMESPACE, "SignatureTime")
    ):
        return False

    signature_target = signature_property.attrib["Target"]
    if signature_target and (
        not signature_id or signature_target != f"#{signature_id}"
    ):
        return False

    signature_time = signature_time_children[0]
    children = list(signature_time)
    if (
        signature_time.attrib
        or not _has_only_whitespace_interstitial_text(signature_time)
        or [_qualified_name(child.tag) for child in children]
        != [
            (_OPC_DIGITAL_SIGNATURE_NAMESPACE, "Format"),
            (_OPC_DIGITAL_SIGNATURE_NAMESPACE, "Value"),
        ]
    ):
        return False

    time_format, time_value = children
    format_value = time_format.text
    if (
        time_format.attrib
        or list(time_format)
        or format_value not in _OPC_SIGNATURE_TIME_VALUE_PATTERNS
        or time_value.attrib
        or list(time_value)
        or time_value.text is None
    ):
        return False
    return (
        _OPC_SIGNATURE_TIME_VALUE_PATTERNS[format_value].fullmatch(time_value.text)
        is not None
    )


def _has_only_whitespace_interstitial_text(element: ET.Element) -> bool:
    """Return whether an element's own text and child tails are whitespace."""

    return not (element.text or "").strip() and not any(
        (child.tail or "").strip() for child in element
    )


def _signature_fragment_identifier(uri: str | None) -> str | None:
    """Return a direct local XMLDSIG fragment identifier, if present."""

    if (
        uri is None
        or not uri.startswith("#")
        or len(uri) == 1
        or any(marker in uri[1:] for marker in ("#", "?", "/"))
    ):
        return None
    return uri[1:]


def _resolve_package_manifest_reference(
    reference: ET.Element,
    members: dict[str, zipfile.ZipInfo],
    content_types: dict[str, str],
    relationship_maps: dict[str, dict[str, _Relationship]],
) -> _ManifestReferenceResolution:
    """Resolve one bounded package-object manifest reference locally."""

    if not _manifest_reference_has_expected_xml_dsig_shape(reference):
        return _ManifestReferenceResolution(
            covered_part_name=None,
            covered_relationship_ids=frozenset(),
            unresolved_reference_count=0,
            unsupported_reference_count=1,
        )
    parsed_reference = _package_manifest_reference_member_name(
        reference.attrib.get("URI")
    )
    if parsed_reference is None:
        return _ManifestReferenceResolution(
            covered_part_name=None,
            covered_relationship_ids=frozenset(),
            unresolved_reference_count=0,
            unsupported_reference_count=1,
        )
    part_name, expected_content_type = parsed_reference
    if part_name not in members:
        return _ManifestReferenceResolution(
            covered_part_name=None,
            covered_relationship_ids=frozenset(),
            unresolved_reference_count=1,
            unsupported_reference_count=0,
        )
    if not part_name.endswith(".rels"):
        if not _part_manifest_reference_has_supported_transforms(reference):
            return _ManifestReferenceResolution(
                covered_part_name=None,
                covered_relationship_ids=frozenset(),
                unresolved_reference_count=0,
                unsupported_reference_count=1,
            )
        if content_types.get(part_name, "") != expected_content_type:
            return _ManifestReferenceResolution(
                covered_part_name=None,
                covered_relationship_ids=frozenset(),
                unresolved_reference_count=1,
                unsupported_reference_count=0,
            )
        return _ManifestReferenceResolution(
            covered_part_name=part_name,
            covered_relationship_ids=frozenset(),
            unresolved_reference_count=0,
            unsupported_reference_count=0,
        )
    if expected_content_type != _PACKAGE_RELATIONSHIP_CONTENT_TYPE:
        return _ManifestReferenceResolution(
            covered_part_name=None,
            covered_relationship_ids=frozenset(),
            unresolved_reference_count=0,
            unsupported_reference_count=1,
        )
    actual_content_type = content_types.get(part_name)
    if actual_content_type is not None and actual_content_type != expected_content_type:
        return _ManifestReferenceResolution(
            covered_part_name=None,
            covered_relationship_ids=frozenset(),
            unresolved_reference_count=1,
            unsupported_reference_count=0,
        )

    source_part = _relationship_source_part(part_name)
    relationships = relationship_maps.get(source_part)
    if relationships is None:
        return _ManifestReferenceResolution(
            covered_part_name=None,
            covered_relationship_ids=frozenset(),
            unresolved_reference_count=1,
            unsupported_reference_count=0,
        )
    return _relationship_manifest_reference_coverage(
        reference, source_part, relationships
    )


def _package_manifest_reference_member_name(uri: str | None) -> tuple[str, str] | None:
    """Parse the exact local part URI form used by OPC manifest references."""

    if uri is None or not uri.startswith("/") or "#" in uri:
        return None
    path, separator, query = uri.partition("?")
    if not separator or "?" in query:
        return None
    part_name = path.removeprefix("/")
    if (
        not part_name
        or part_name.startswith("/")
        or part_name.endswith("/")
        or "//" in part_name
        or any(segment in {".", ".."} for segment in part_name.split("/"))
    ):
        return None
    name, equals, value = query.partition("=")
    if name != "ContentType" or not equals or not value or "&" in value:
        return None
    return part_name, value


def _manifest_reference_has_expected_xml_dsig_shape(reference: ET.Element) -> bool:
    """Require the XMLDSIG ``Reference`` child order without verifying a digest."""

    children = list(reference)
    if (reference.text or "").strip() or any(
        (child.tail or "").strip() for child in children
    ):
        return False
    if len(children) == 2:
        digest_method, digest_value = children
    elif len(children) == 3 and _qualified_name(children[0].tag) == (
        _XMLDSIG_NAMESPACE,
        "Transforms",
    ):
        _, digest_method, digest_value = children
    else:
        return False

    if (
        _qualified_name(digest_method.tag) != (_XMLDSIG_NAMESPACE, "DigestMethod")
        or not (digest_method.attrib.get("Algorithm") or "").strip()
        or _qualified_name(digest_value.tag) != (_XMLDSIG_NAMESPACE, "DigestValue")
        or digest_value.attrib
        or list(digest_value)
        or not _element_has_text_value(digest_value)
    ):
        return False
    return True


def _relationship_manifest_reference_coverage(
    reference: ET.Element,
    source_part: str,
    relationships: dict[str, _Relationship],
) -> _ManifestReferenceResolution:
    """Resolve a bounded, standard OPC relationship-transform declaration.

    This remains a static selector audit rather than XMLDSIG verification.  It
    nevertheless requires the OPC transform sequence that gives those
    selectors their defined input: one direct relationship transform, followed
    immediately by one of OPC's two supported XML canonicalization transforms.
    """

    transforms_elements = [
        child
        for child in reference
        if _qualified_name(child.tag) == (_XMLDSIG_NAMESPACE, "Transforms")
    ]
    if len(transforms_elements) != 1:
        return _ManifestReferenceResolution(
            covered_part_name=None,
            covered_relationship_ids=frozenset(),
            unresolved_reference_count=0,
            unsupported_reference_count=1,
        )

    transforms = list(transforms_elements[0])
    if not transforms or any(
        _qualified_name(transform.tag) != (_XMLDSIG_NAMESPACE, "Transform")
        or transform.attrib.get("Algorithm")
        not in _OPC_SUPPORTED_MANIFEST_TRANSFORM_ALGORITHMS
        for transform in transforms
    ):
        return _ManifestReferenceResolution(
            covered_part_name=None,
            covered_relationship_ids=frozenset(),
            unresolved_reference_count=0,
            unsupported_reference_count=1,
        )

    relationship_transform_indexes = [
        index
        for index, transform in enumerate(transforms)
        if transform.attrib.get("Algorithm") == _OPC_RELATIONSHIP_TRANSFORM_ALGORITHM
    ]
    if len(relationship_transform_indexes) != 1:
        return _ManifestReferenceResolution(
            covered_part_name=None,
            covered_relationship_ids=frozenset(),
            unresolved_reference_count=0,
            unsupported_reference_count=1,
        )

    relationship_transform_index = relationship_transform_indexes[0]
    if (
        relationship_transform_index + 1 >= len(transforms)
        or transforms[relationship_transform_index + 1].attrib.get("Algorithm")
        not in _XML_CANONICALIZATION_TRANSFORM_ALGORITHMS
    ):
        return _ManifestReferenceResolution(
            covered_part_name=None,
            covered_relationship_ids=frozenset(),
            unresolved_reference_count=0,
            unsupported_reference_count=1,
        )

    selected_relationship_ids: set[tuple[str, str]] = set()
    unresolved_reference_count = 0
    unsupported_reference_count = 0
    selectors = list(transforms[relationship_transform_index])
    if not selectors:
        unsupported_reference_count += 1
    for selector in selectors:
        if list(selector) or (selector.text or "").strip():
            unsupported_reference_count += 1
            continue

        qualified_selector = _qualified_name(selector.tag)
        if qualified_selector == (
            _OPC_DIGITAL_SIGNATURE_NAMESPACE,
            "RelationshipReference",
        ):
            source_id = selector.attrib.get("SourceId")
            if set(selector.attrib) != {"SourceId"} or not source_id:
                unsupported_reference_count += 1
                continue
            if source_id in relationships:
                selected_relationship_ids.add((source_part, source_id))
            else:
                unresolved_reference_count += 1
            continue

        if qualified_selector == (
            _OPC_DIGITAL_SIGNATURE_NAMESPACE,
            "RelationshipsGroupReference",
        ):
            source_type = selector.attrib.get("SourceType")
            if set(selector.attrib) != {"SourceType"} or not source_type:
                unsupported_reference_count += 1
                continue
            matching_ids = [
                relationship_id
                for relationship_id, relationship in relationships.items()
                if relationship.relationship_type == source_type
            ]
            if not matching_ids:
                unresolved_reference_count += 1
            selected_relationship_ids.update(
                (source_part, relationship_id) for relationship_id in matching_ids
            )
            continue

        unsupported_reference_count += 1

    return _ManifestReferenceResolution(
        covered_part_name=None,
        covered_relationship_ids=frozenset(selected_relationship_ids),
        unresolved_reference_count=unresolved_reference_count,
        unsupported_reference_count=unsupported_reference_count,
    )


def _part_manifest_reference_has_supported_transforms(reference: ET.Element) -> bool:
    """Accept no transform or a direct C14N-only list for a package part.

    A relationships transform has defined input only for a Relationships part.
    This bounded declaration audit therefore permits ordinary part references to
    omit transforms or use OPC's two XML canonicalization algorithms, while
    refusing to credit unknown or relationship-transform sequences.
    """

    transforms_elements = [
        child
        for child in reference
        if _qualified_name(child.tag) == (_XMLDSIG_NAMESPACE, "Transforms")
    ]
    if not transforms_elements:
        return True
    if len(transforms_elements) != 1:
        return False
    transforms = list(transforms_elements[0])
    return bool(transforms) and all(
        _qualified_name(transform.tag) == (_XMLDSIG_NAMESPACE, "Transform")
        and transform.attrib.get("Algorithm")
        in _XML_CANONICALIZATION_TRANSFORM_ALGORITHMS
        for transform in transforms
    )


def _relationship_transform_count(reference: ET.Element) -> int:
    """Count direct OPC relationship transforms for one manifest reference."""

    return sum(
        _qualified_name(transform.tag) == (_XMLDSIG_NAMESPACE, "Transform")
        and transform.attrib.get("Algorithm") == _OPC_RELATIONSHIP_TRANSFORM_ALGORITHM
        for transforms in reference
        if _qualified_name(transforms.tag) == (_XMLDSIG_NAMESPACE, "Transforms")
        for transform in transforms
    )


def _validate_package_digital_signature_root(root: ET.Element) -> dict[str, int]:
    """Check a bounded XMLDSIG shape without verifying cryptographic validity."""

    if _qualified_name(root.tag) != (_XMLDSIG_NAMESPACE, "Signature"):
        raise DocumentFormatError("package digital signature parts are invalid")

    signed_infos = [
        child
        for child in root
        if _qualified_name(child.tag) == (_XMLDSIG_NAMESPACE, "SignedInfo")
    ]
    signature_values = [
        child
        for child in root
        if _qualified_name(child.tag) == (_XMLDSIG_NAMESPACE, "SignatureValue")
    ]
    if len(signed_infos) != 1 or len(signature_values) != 1:
        raise DocumentFormatError("package digital signature parts are invalid")
    if not _element_has_text_value(signature_values[0]):
        raise DocumentFormatError("package digital signature parts are invalid")

    signed_info = signed_infos[0]
    canonicalization_methods = [
        child
        for child in signed_info
        if _qualified_name(child.tag) == (_XMLDSIG_NAMESPACE, "CanonicalizationMethod")
    ]
    signature_methods = [
        child
        for child in signed_info
        if _qualified_name(child.tag) == (_XMLDSIG_NAMESPACE, "SignatureMethod")
    ]
    signed_info_references = [
        child
        for child in signed_info
        if _qualified_name(child.tag) == (_XMLDSIG_NAMESPACE, "Reference")
    ]
    if (
        len(canonicalization_methods) != 1
        or len(signature_methods) != 1
        or not signed_info_references
    ):
        raise DocumentFormatError("package digital signature parts are invalid")

    manifest_reference_count = 0
    relationship_reference_count = 0
    inline_x509_certificate_count = 0
    signature_property_count = 0
    for element in root.iter():
        qualified_name = _qualified_name(element.tag)
        if qualified_name == (_XMLDSIG_NAMESPACE, "Manifest"):
            manifest_reference_count += sum(
                _qualified_name(child.tag) == (_XMLDSIG_NAMESPACE, "Reference")
                for child in element
            )
        elif qualified_name in {
            (_OPC_DIGITAL_SIGNATURE_NAMESPACE, "RelationshipReference"),
            (_OPC_DIGITAL_SIGNATURE_NAMESPACE, "RelationshipsGroupReference"),
        }:
            relationship_reference_count += 1
        elif qualified_name == (_XMLDSIG_NAMESPACE, "X509Certificate"):
            inline_x509_certificate_count += 1
        elif qualified_name == (_XMLDSIG_NAMESPACE, "SignatureProperty"):
            signature_property_count += 1

    return {
        "signed_info_reference": len(signed_info_references),
        "manifest_reference": manifest_reference_count,
        "relationship_reference": relationship_reference_count,
        "inline_x509_certificate": inline_x509_certificate_count,
        "signature_property": signature_property_count,
    }


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


def _save_through_xslt_inventory(
    archive: zipfile.ZipFile,
    members: dict[str, zipfile.ZipInfo],
    relationship_maps: dict[str, dict[str, _Relationship]],
    settings_part_names: frozenset[str],
    limits: PackageLimits,
) -> SaveThroughXsltInventory:
    """Inventory XSLT-on-save declarations without resolving transforms.

    OOXML stores an optional custom XSLT used only when an application saves a
    document as a single XML file. A relationship-backed transform must use the
    standard ``transform`` relationship type and remain external to the
    package. A ``w:saveThroughXslt`` anchor may instead contain only its
    application-defined ``w:solutionID``, so it is retained as private
    configuration evidence without attempting to resolve it.
    """

    records: list[tuple[str, ...]] = []
    enabled_setting_count = 0
    disabled_setting_count = 0
    transform_anchor_count = 0
    transform_relationship_count = 0
    solution_identifier_count = 0

    for settings_part in sorted(settings_part_names):
        relationships = relationship_maps.get(settings_part, {})
        for relationship in _relationships_with_types(
            relationships, _SAVE_THROUGH_XSLT_RELATIONSHIP_TYPES
        ):
            _validate_save_through_xslt_relationship(relationship)
            transform_relationship_count += 1
            records.append(
                (
                    "save_through_xslt_relationship",
                    settings_part,
                    *relationship.canonical_value(),
                )
            )

        root = _read_xml(archive, members[settings_part], limits)
        if not _is_word_element(root, "settings"):
            raise DocumentFormatError("document settings are invalid")
        enabled_settings = [
            element
            for element in root
            if _is_word_element(element, "useXSLTWhenSaving")
        ]
        transforms = [
            element for element in root if _is_word_element(element, "saveThroughXslt")
        ]
        if len(enabled_settings) > 1 or len(transforms) > 1:
            raise DocumentFormatError("save-through-XSLT state is invalid")

        for enabled_setting in enabled_settings:
            _validate_save_through_xslt_leaf(enabled_setting)
            if _is_enabled(enabled_setting):
                enabled_setting_count += 1
            else:
                disabled_setting_count += 1
            records.append(
                (
                    "use_xslt_when_saving",
                    settings_part,
                    _fingerprint_element(enabled_setting, relationships),
                )
            )

        for transform in transforms:
            _validate_save_through_xslt_leaf(transform)
            _validate_save_through_xslt_anchor(transform, relationships)
            transform_anchor_count += 1
            if _word_attribute_value(transform, "solutionID") is not None:
                solution_identifier_count += 1
            records.append(
                (
                    "save_through_xslt_anchor",
                    settings_part,
                    _fingerprint_element(transform, relationships),
                )
            )

    return SaveThroughXsltInventory(
        enabled_setting_count=enabled_setting_count,
        disabled_setting_count=disabled_setting_count,
        transform_anchor_count=transform_anchor_count,
        transform_relationship_count=transform_relationship_count,
        solution_identifier_count=solution_identifier_count,
        signature=_digest_records(records),
    )


def _validate_save_through_xslt_leaf(element: ET.Element) -> None:
    if list(element) or (element.text or "").strip():
        raise DocumentFormatError("save-through-XSLT state is invalid")


def _validate_save_through_xslt_relationship(relationship: _Relationship) -> None:
    if relationship.target_mode.casefold() != "external":
        raise DocumentFormatError("save-through-XSLT relationships are invalid")


def _validate_save_through_xslt_anchor(
    element: ET.Element, relationships: dict[str, _Relationship]
) -> None:
    relationship_id = _relationship_id_value(element)
    if relationship_id is None:
        return
    relationship = relationships.get(relationship_id)
    if (
        relationship is None
        or relationship.relationship_type.casefold()
        not in _SAVE_THROUGH_XSLT_RELATIONSHIP_TYPES
        or relationship.target_mode.casefold() != "external"
    ):
        raise DocumentFormatError("save-through-XSLT markup is invalid")


def _attached_custom_xml_schema_inventory(
    archive: zipfile.ZipFile,
    members: dict[str, zipfile.ZipInfo],
    relationship_maps: dict[str, dict[str, _Relationship]],
    settings_part_names: frozenset[str],
    limits: PackageLimits,
) -> AttachedCustomXmlSchemaInventory:
    """Inventory attached custom XML schemas without exposing identifiers.

    ``w:attachedSchema`` declares a custom XML schema by target namespace for
    association when a host loads the document, provided the host has that
    schema available.  The declaration does not name a package payload or a
    fetchable target, and DocFence never resolves, retrieves, loads, or
    validates against the declared schema.
    """

    records: list[tuple[str, ...]] = []
    attached_schema_count = 0

    for settings_part in sorted(settings_part_names):
        root = _read_xml(archive, members[settings_part], limits)
        if not _is_word_element(root, "settings"):
            raise DocumentFormatError("document settings are invalid")
        relationships = relationship_maps.get(settings_part, {})
        for element in root:
            if not _is_word_element(element, "attachedSchema"):
                continue
            _validate_attached_custom_xml_schema(element)
            attached_schema_count += 1
            records.append(
                (
                    "attached_custom_xml_schema",
                    settings_part,
                    _fingerprint_element(element, relationships),
                )
            )

    return AttachedCustomXmlSchemaInventory(
        attached_schema_count=attached_schema_count,
        signature=_digest_records(records),
    )


def _validate_attached_custom_xml_schema(element: ET.Element) -> None:
    """Validate the direct ``CT_String`` attached-schema leaf."""

    if (
        not _is_word_element(element, "attachedSchema")
        or list(element)
        or (element.text or "").strip()
    ):
        raise DocumentFormatError("attached custom XML schema state is invalid")

    attributes: set[str] = set()
    for attribute in element.attrib:
        namespace, local_name = _qualified_name(attribute)
        if (
            namespace not in _WORD_NAMESPACES
            or local_name != "val"
            or local_name in attributes
        ):
            raise DocumentFormatError("attached custom XML schema state is invalid")
        attributes.add(local_name)
    if attributes != {"val"}:
        raise DocumentFormatError("attached custom XML schema state is invalid")


def _field_update_on_open_inventory(
    archive: zipfile.ZipFile,
    members: dict[str, zipfile.ZipInfo],
    settings_part_names: frozenset[str],
    limits: PackageLimits,
) -> FieldUpdateOnOpenInventory:
    """Inventory ``w:updateFields`` without evaluating document fields.

    The direct ``CT_OnOff`` Settings leaf requests automatic field-result
    recalculation when a supporting host opens the document. Its value is a
    configuration switch, not an instruction to DocFence to parse, evaluate,
    resolve, or update a field. Canonical enabled/disabled state keeps lexical
    equivalents such as ``on`` and ``true`` from creating review noise.
    """

    records: list[tuple[str, str, str]] = []
    enabled_setting_count = 0
    disabled_setting_count = 0

    for settings_part in sorted(settings_part_names):
        root = _read_xml(archive, members[settings_part], limits)
        if not _is_word_element(root, "settings"):
            raise DocumentFormatError("document settings are invalid")
        update_settings = [
            element for element in root if _is_word_element(element, "updateFields")
        ]
        if len(update_settings) > 1:
            raise DocumentFormatError("field-update-on-open state is invalid")

        for setting in update_settings:
            state = _field_update_on_open_state(setting)
            if state == "enabled":
                enabled_setting_count += 1
            else:
                disabled_setting_count += 1
            records.append(("field_update_on_open", settings_part, state))

    return FieldUpdateOnOpenInventory(
        enabled_setting_count=enabled_setting_count,
        disabled_setting_count=disabled_setting_count,
        signature=_digest_records(records),
    )


def _field_update_on_open_state(element: ET.Element) -> str:
    """Validate one direct ``w:updateFields`` ``CT_OnOff`` leaf."""

    return _word_on_off_state(
        element,
        local_name="updateFields",
        error_message="field-update-on-open state is invalid",
    )


def _template_style_update_on_open_inventory(
    archive: zipfile.ZipFile,
    members: dict[str, zipfile.ZipInfo],
    settings_part_names: frozenset[str],
    limits: PackageLimits,
) -> TemplateStyleUpdateOnOpenInventory:
    """Inventory ``w:linkStyles`` without loading a document template.

    The direct ``CT_OnOff`` Settings leaf requests automatic style updates from
    an attached document template when a supporting host opens the document.
    It is configuration evidence, not an instruction to DocFence to resolve an
    attached template, open Word, or propagate styles. Canonical enabled and
    disabled state keeps equivalent on/off spellings from creating review noise.
    """

    records: list[tuple[str, str, str]] = []
    enabled_setting_count = 0
    disabled_setting_count = 0

    for settings_part in sorted(settings_part_names):
        root = _read_xml(archive, members[settings_part], limits)
        if not _is_word_element(root, "settings"):
            raise DocumentFormatError("document settings are invalid")
        style_update_settings = [
            element for element in root if _is_word_element(element, "linkStyles")
        ]
        if len(style_update_settings) > 1:
            raise DocumentFormatError("template-style-update-on-open state is invalid")

        for setting in style_update_settings:
            state = _template_style_update_on_open_state(setting)
            if state == "enabled":
                enabled_setting_count += 1
            else:
                disabled_setting_count += 1
            records.append(("template_style_update_on_open", settings_part, state))

    return TemplateStyleUpdateOnOpenInventory(
        enabled_setting_count=enabled_setting_count,
        disabled_setting_count=disabled_setting_count,
        signature=_digest_records(records),
    )


def _template_style_update_on_open_state(element: ET.Element) -> str:
    """Validate one direct ``w:linkStyles`` ``CT_OnOff`` leaf."""

    return _word_on_off_state(
        element,
        local_name="linkStyles",
        error_message="template-style-update-on-open state is invalid",
    )


def _personal_information_removal_on_save_inventory(
    archive: zipfile.ZipFile,
    members: dict[str, zipfile.ZipInfo],
    settings_part_names: frozenset[str],
    limits: PackageLimits,
) -> PersonalInformationRemovalOnSaveInventory:
    """Inventory ``w:removePersonalInformation`` without removing data.

    The direct ``CT_OnOff`` Settings leaf records a request for a capable host
    to remove personal information when saving a document. Its value is stored
    configuration evidence, not an instruction to DocFence to identify,
    redact, rewrite, or otherwise interpret current document properties.
    Canonical enabled/disabled state keeps equivalent on/off spellings from
    creating review noise.
    """

    records: list[tuple[str, str, str]] = []
    enabled_setting_count = 0
    disabled_setting_count = 0

    for settings_part in sorted(settings_part_names):
        root = _read_xml(archive, members[settings_part], limits)
        if not _is_word_element(root, "settings"):
            raise DocumentFormatError("document settings are invalid")
        removal_settings = [
            element
            for element in root
            if _is_word_element(element, "removePersonalInformation")
        ]
        if len(removal_settings) > 1:
            raise DocumentFormatError(
                "personal-information-removal-on-save state is invalid"
            )

        for setting in removal_settings:
            state = _personal_information_removal_on_save_state(setting)
            if state == "enabled":
                enabled_setting_count += 1
            else:
                disabled_setting_count += 1
            records.append(
                ("personal_information_removal_on_save", settings_part, state)
            )

    return PersonalInformationRemovalOnSaveInventory(
        enabled_setting_count=enabled_setting_count,
        disabled_setting_count=disabled_setting_count,
        signature=_digest_records(records),
    )


def _personal_information_removal_on_save_state(element: ET.Element) -> str:
    """Validate one direct ``w:removePersonalInformation`` ``CT_OnOff`` leaf."""

    return _word_on_off_state(
        element,
        local_name="removePersonalInformation",
        error_message="personal-information-removal-on-save state is invalid",
    )


def _save_forms_data_inventory(
    archive: zipfile.ZipFile,
    members: dict[str, zipfile.ZipInfo],
    settings_part_names: frozenset[str],
    limits: PackageLimits,
) -> SaveFormsDataInventory:
    """Inventory ``w:saveFormsData`` without saving or exporting form data.

    The direct ``CT_OnOff`` Settings leaf records a request for a capable host
    to save a document as form-data-only delimited text. Its value is stored
    configuration evidence, not an instruction to DocFence to inspect form
    fields, open Word, save a document, or create an export. Canonical
    enabled/disabled state keeps equivalent on/off spellings from creating
    review noise.
    """

    records: list[tuple[str, str, str]] = []
    enabled_setting_count = 0
    disabled_setting_count = 0

    for settings_part in sorted(settings_part_names):
        root = _read_xml(archive, members[settings_part], limits)
        if not _is_word_element(root, "settings"):
            raise DocumentFormatError("document settings are invalid")
        save_forms_data_settings = [
            element for element in root if _is_word_element(element, "saveFormsData")
        ]
        if len(save_forms_data_settings) > 1:
            raise DocumentFormatError("save-forms-data state is invalid")

        for setting in save_forms_data_settings:
            state = _save_forms_data_state(setting)
            if state == "enabled":
                enabled_setting_count += 1
            else:
                disabled_setting_count += 1
            records.append(("save_forms_data", settings_part, state))

    return SaveFormsDataInventory(
        enabled_setting_count=enabled_setting_count,
        disabled_setting_count=disabled_setting_count,
        signature=_digest_records(records),
    )


def _save_forms_data_state(element: ET.Element) -> str:
    """Validate one direct ``w:saveFormsData`` ``CT_OnOff`` leaf."""

    return _word_on_off_state(
        element,
        local_name="saveFormsData",
        error_message="save-forms-data state is invalid",
    )


def _save_preview_picture_inventory(
    archive: zipfile.ZipFile,
    members: dict[str, zipfile.ZipInfo],
    settings_part_names: frozenset[str],
    limits: PackageLimits,
) -> SavePreviewPictureInventory:
    """Inventory ``w:savePreviewPicture`` without generating a thumbnail.

    The direct ``CT_OnOff`` Settings leaf records a request for a capable host
    to generate a first-page document thumbnail when saving. Its value is
    stored configuration evidence, not an instruction to DocFence to decode an
    image, render a page, open Word, save a document, or claim that a package
    thumbnail exists. Canonical enabled/disabled state keeps equivalent on/off
    spellings from creating review noise.
    """

    records: list[tuple[str, str, str]] = []
    enabled_setting_count = 0
    disabled_setting_count = 0

    for settings_part in sorted(settings_part_names):
        root = _read_xml(archive, members[settings_part], limits)
        if not _is_word_element(root, "settings"):
            raise DocumentFormatError("document settings are invalid")
        save_preview_picture_settings = [
            element
            for element in root
            if _is_word_element(element, "savePreviewPicture")
        ]
        if len(save_preview_picture_settings) > 1:
            raise DocumentFormatError("save-preview-picture state is invalid")

        for setting in save_preview_picture_settings:
            state = _save_preview_picture_state(setting)
            if state == "enabled":
                enabled_setting_count += 1
            else:
                disabled_setting_count += 1
            records.append(("save_preview_picture", settings_part, state))

    return SavePreviewPictureInventory(
        enabled_setting_count=enabled_setting_count,
        disabled_setting_count=disabled_setting_count,
        signature=_digest_records(records),
    )


def _save_preview_picture_state(element: ET.Element) -> str:
    """Validate one direct ``w:savePreviewPicture`` ``CT_OnOff`` leaf."""

    return _word_on_off_state(
        element,
        local_name="savePreviewPicture",
        error_message="save-preview-picture state is invalid",
    )


def _word_on_off_state(
    element: ET.Element, *, local_name: str, error_message: str
) -> str:
    """Return canonical state for a direct Word ``CT_OnOff`` leaf."""

    if (
        not _is_word_element(element, local_name)
        or list(element)
        or (element.text or "").strip()
    ):
        raise DocumentFormatError(error_message)

    value: str | None = None
    for attribute, raw_value in element.attrib.items():
        namespace, attribute_local_name = _qualified_name(attribute)
        if (
            namespace not in _WORD_NAMESPACES
            or attribute_local_name != "val"
            or value is not None
        ):
            raise DocumentFormatError(error_message)
        value = raw_value.strip().casefold()
        if value not in _ON_OFF_VALUES:
            raise DocumentFormatError(error_message)

    if value is None or value in _ON_OFF_TRUE_VALUES:
        return "enabled"
    return "disabled"


def _external_document_dependency_inventory(
    archive: zipfile.ZipFile,
    members: dict[str, zipfile.ZipInfo],
    relationship_maps: dict[str, dict[str, _Relationship]],
    limits: PackageLimits,
) -> tuple[ExternalDocumentDependencyInventory, frozenset[str]]:
    """Inventory standard external Word document dependencies without targets.

    Word uses three explicit relationship families to incorporate or consult a
    different document package: an attached template, master-document
    subdocuments, and frameset source files.  Every recognized relationship is
    required to remain external; DocFence only fingerprints its semantics and
    never resolves or retrieves the target.
    """

    records: list[tuple[str, ...]] = []
    attached_template_anchor_count = 0
    attached_template_relationship_count = 0
    subdocument_anchor_count = 0
    subdocument_relationship_count = 0
    frame_source_anchor_count = 0
    frame_relationship_count = 0
    dependency_part_names: set[str] = set()

    for settings_part in _linked_document_settings_parts(members, relationship_maps):
        settings_relationships = relationship_maps.get(settings_part, {})
        has_attached_template_dependency = False
        for relationship in _relationships_with_types(
            settings_relationships, _ATTACHED_TEMPLATE_RELATIONSHIP_TYPES
        ):
            _validate_external_document_dependency_relationship(relationship)
            attached_template_relationship_count += 1
            has_attached_template_dependency = True
            records.append(
                (
                    "attached_template_relationship",
                    settings_part,
                    *relationship.canonical_value(),
                )
            )

        settings_root = _read_xml(archive, members[settings_part], limits)
        if not _is_word_element(settings_root, "settings"):
            raise DocumentFormatError("document settings are invalid")
        for element in settings_root:
            if not _is_word_element(element, "attachedTemplate"):
                continue
            _validate_external_document_dependency_anchor(
                element,
                settings_relationships,
                _ATTACHED_TEMPLATE_RELATIONSHIP_TYPES,
            )
            attached_template_anchor_count += 1
            has_attached_template_dependency = True
            records.append(
                (
                    "attached_template_anchor",
                    settings_part,
                    _fingerprint_element(element, settings_relationships),
                )
            )

        if has_attached_template_dependency and settings_part != "word/settings.xml":
            dependency_part_names.add(settings_part)
            records.append(
                (
                    "attached_template_settings",
                    settings_part,
                    _fingerprint_element(settings_root, settings_relationships),
                )
            )

    main_document_part = "word/document.xml"
    main_document_relationships = relationship_maps.get(main_document_part, {})
    for relationship in _relationships_with_types(
        main_document_relationships, _SUBDOCUMENT_RELATIONSHIP_TYPES
    ):
        _validate_external_document_dependency_relationship(relationship)
        subdocument_relationship_count += 1
        records.append(
            (
                "subdocument_relationship",
                main_document_part,
                *relationship.canonical_value(),
            )
        )

    main_document_root = _read_xml(archive, members[main_document_part], limits)
    if not _is_word_element(main_document_root, "document"):
        raise DocumentFormatError("document story is invalid")
    for element in main_document_root.iter():
        if not _is_word_element(element, "subDoc"):
            continue
        _validate_external_document_dependency_anchor(
            element,
            main_document_relationships,
            _SUBDOCUMENT_RELATIONSHIP_TYPES,
        )
        subdocument_anchor_count += 1
        records.append(
            (
                "subdocument_anchor",
                main_document_part,
                _fingerprint_element(element, main_document_relationships),
            )
        )

    for web_settings_part in _linked_web_settings_parts(members, relationship_maps):
        web_settings_relationships = relationship_maps.get(web_settings_part, {})
        has_frame_dependency = False
        for relationship in _relationships_with_types(
            web_settings_relationships, _FRAME_RELATIONSHIP_TYPES
        ):
            _validate_external_document_dependency_relationship(relationship)
            frame_relationship_count += 1
            has_frame_dependency = True
            records.append(
                (
                    "frame_relationship",
                    web_settings_part,
                    *relationship.canonical_value(),
                )
            )

        web_settings_root = _read_xml(archive, members[web_settings_part], limits)
        if not _is_word_element(web_settings_root, "webSettings"):
            raise DocumentFormatError("web settings are invalid")
        for frame in web_settings_root.iter():
            if not _is_word_element(frame, "frame"):
                continue
            for element in frame:
                if not _is_word_element(element, "sourceFileName"):
                    continue
                _validate_external_document_dependency_anchor(
                    element,
                    web_settings_relationships,
                    _FRAME_RELATIONSHIP_TYPES,
                )
                frame_source_anchor_count += 1
                has_frame_dependency = True
                records.append(
                    (
                        "frame_source_anchor",
                        web_settings_part,
                        _fingerprint_element(element, web_settings_relationships),
                    )
                )

        if has_frame_dependency:
            dependency_part_names.add(web_settings_part)
            records.append(
                (
                    "frame_web_settings",
                    web_settings_part,
                    _fingerprint_element(web_settings_root, web_settings_relationships),
                )
            )

    return (
        ExternalDocumentDependencyInventory(
            attached_template_anchor_count=attached_template_anchor_count,
            attached_template_relationship_count=attached_template_relationship_count,
            subdocument_anchor_count=subdocument_anchor_count,
            subdocument_relationship_count=subdocument_relationship_count,
            frame_source_anchor_count=frame_source_anchor_count,
            frame_relationship_count=frame_relationship_count,
            signature=_digest_records(records),
        ),
        frozenset(dependency_part_names),
    )


def _relationships_with_types(
    relationships: dict[str, _Relationship], expected_types: frozenset[str]
) -> list[_Relationship]:
    return sorted(
        (
            relationship
            for relationship in relationships.values()
            if relationship.relationship_type.casefold() in expected_types
        ),
        key=lambda relationship: relationship.canonical_value(),
    )


def _linked_document_settings_parts(
    members: dict[str, zipfile.ZipInfo],
    relationship_maps: dict[str, dict[str, _Relationship]],
) -> tuple[str, ...]:
    """Return discovered Settings parts, retaining Word's canonical fallback."""

    return tuple(sorted(_document_settings_part_scopes(members, relationship_maps)))


def _document_settings_part_scopes(
    members: dict[str, zipfile.ZipInfo],
    relationship_maps: dict[str, dict[str, _Relationship]],
) -> dict[str, frozenset[str]]:
    """Map Settings parts to the main or glossary document that links them."""

    source_parts_by_settings_part: dict[str, set[str]] = {}

    def add(settings_part: str, source_part: str) -> None:
        source_parts_by_settings_part.setdefault(settings_part, set()).add(source_part)

    if "word/settings.xml" in members:
        add("word/settings.xml", "word/document.xml")
    for source_part in _SETTINGS_SOURCE_PARTS:
        if source_part not in members:
            continue
        for relationship in _relationships_with_types(
            relationship_maps.get(source_part, {}),
            _DOCUMENT_SETTINGS_RELATIONSHIP_TYPES,
        ):
            target = _internal_relationship_target(source_part, relationship, members)
            if target is None:
                raise DocumentFormatError("document settings relationships are invalid")
            add(target, source_part)
    return {
        settings_part: frozenset(source_parts)
        for settings_part, source_parts in source_parts_by_settings_part.items()
    }


def _linked_web_settings_parts(
    members: dict[str, zipfile.ZipInfo],
    relationship_maps: dict[str, dict[str, _Relationship]],
) -> tuple[str, ...]:
    """Return Web Settings parts explicitly linked from supported sources."""

    part_names: set[str] = set()
    for source_part in _SETTINGS_SOURCE_PARTS:
        if source_part not in members:
            continue
        for relationship in _relationships_with_types(
            relationship_maps.get(source_part, {}),
            _WEB_SETTINGS_RELATIONSHIP_TYPES,
        ):
            target = _internal_relationship_target(source_part, relationship, members)
            if target is None:
                raise DocumentFormatError("web settings relationships are invalid")
            part_names.add(target)
    return tuple(sorted(part_names))


def _validate_external_document_dependency_relationship(
    relationship: _Relationship,
) -> None:
    if relationship.target_mode.casefold() != "external":
        raise DocumentFormatError(
            "external document dependency relationships are invalid"
        )


def _validate_external_document_dependency_anchor(
    element: ET.Element,
    relationships: dict[str, _Relationship],
    expected_types: frozenset[str],
) -> None:
    relationship_id = _relationship_id_value(element)
    relationship = relationships.get(relationship_id) if relationship_id else None
    if (
        relationship is None
        or relationship.relationship_type.casefold() not in expected_types
        or relationship.target_mode.casefold() != "external"
    ):
        raise DocumentFormatError("external document dependency markup is invalid")


def _content_control_lock_inventory(
    references: tuple[_ContentControlLockReference, ...],
) -> ContentControlLockInventory:
    """Inventory direct content-control lock declarations without labels or text.

    Every discovered ``w:sdt`` contributes exactly one state.  A missing direct
    ``w:lock`` leaf deliberately remains distinct from an explicit
    ``w:val=\"unlocked\"`` declaration: document type can affect what a host
    treats as the implicit behavior, and DocFence does not evaluate that
    behavior.  Story-local ordinals remain only inside the private digest so a
    same-count reassignment of locks stays review-visible without exposing a
    content-control ID, tag, title, placeholder, value, or part path.
    """

    counts = {
        "no_lock_declaration": 0,
        "unlocked": 0,
        "sdt_locked": 0,
        "content_locked": 0,
        "sdt_content_locked": 0,
    }
    records: list[tuple[str, ...]] = []
    for reference in references:
        counts[reference.state] += 1
        records.append(
            (
                "content_control_lock",
                reference.story_part,
                str(reference.ordinal),
                reference.state,
            )
        )

    return ContentControlLockInventory(
        no_lock_declaration_count=counts["no_lock_declaration"],
        unlocked_count=counts["unlocked"],
        sdt_locked_count=counts["sdt_locked"],
        content_locked_count=counts["content_locked"],
        sdt_content_locked_count=counts["sdt_content_locked"],
        signature=_digest_records(records),
    )


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


def _external_field_inventory(
    references: tuple[_ExternalFieldReference, ...],
) -> ExternalFieldInventory:
    """Inventory field families that can consult material outside a package.

    The field instruction, including a potential path, connection, query, or
    object name, is hashed before it leaves the story traversal. The public
    model carries category counts only.
    """

    counts = {
        "database": 0,
        "legacy_data": 0,
        "dde": 0,
        "dde_auto": 0,
        "include_text": 0,
        "include_picture": 0,
        "link": 0,
        "referenced_document": 0,
    }
    records: list[tuple[str, ...]] = []
    for reference in references:
        counts[reference.category] += 1
        records.append(
            (
                "external_field_instruction",
                reference.story_part,
                reference.category,
                reference.instruction_signature,
            )
        )
    return ExternalFieldInventory(
        database_field_count=counts["database"],
        legacy_data_field_count=counts["legacy_data"],
        dde_field_count=counts["dde"],
        dde_auto_field_count=counts["dde_auto"],
        include_text_field_count=counts["include_text"],
        include_picture_field_count=counts["include_picture"],
        link_field_count=counts["link"],
        referenced_document_field_count=counts["referenced_document"],
        signature=_digest_records(records),
    )


def _modern_comment_metadata_inventory(
    archive: zipfile.ZipFile,
    members: dict[str, zipfile.ZipInfo],
    content_types: dict[str, str],
    relationship_maps: dict[str, dict[str, _Relationship]],
    limits: PackageLimits,
) -> tuple[ModernCommentMetadataInventory, frozenset[str]]:
    """Inventory modern Word comment metadata without exposing identities.

    Microsoft Word stores contact records, comment threading/resolution state,
    durable identifiers, and reaction metadata outside the ordinary comments
    story.  Every recognized part is root-validated and privately fingerprinted
    with its identifiers intact; public callers receive aggregate counts only.
    """

    part_kinds: dict[str, str] = {}
    records: list[tuple[str, ...]] = []

    def add_part(kind: str, part_name: str) -> None:
        existing_kind = part_kinds.get(part_name)
        if existing_kind is not None and existing_kind != kind:
            raise DocumentFormatError("modern comment metadata parts are inconsistent")
        part_kinds[part_name] = kind

    for part_name, kind in _MODERN_COMMENT_FALLBACK_PARTS.items():
        if part_name in members:
            add_part(kind, part_name)

    for part_name, content_type in content_types.items():
        kind = _MODERN_COMMENT_PART_CONTENT_TYPES.get(content_type.casefold())
        if kind is not None:
            add_part(kind, part_name)

    for source_part, relationships in sorted(relationship_maps.items()):
        for relationship in sorted(
            relationships.values(), key=lambda value: value.canonical_value()
        ):
            kind = _MODERN_COMMENT_PART_RELATIONSHIP_TYPES.get(
                relationship.relationship_type.casefold()
            )
            if kind is None:
                continue
            if relationship.target_mode.casefold() != "internal":
                raise DocumentFormatError(
                    "modern comment metadata relationships are invalid"
                )
            target = _internal_relationship_target(source_part, relationship, members)
            if target is None:
                raise DocumentFormatError(
                    "modern comment metadata relationships are invalid"
                )
            add_part(kind, target)
            records.append(
                (
                    "modern_comment_metadata_relationship",
                    kind,
                    source_part,
                    *relationship.canonical_value(),
                )
            )

    counts = {
        "people_part": 0,
        "person": 0,
        "presence_info": 0,
        "comments_extended_part": 0,
        "comment_extension": 0,
        "threaded_comment": 0,
        "resolved_comment": 0,
        "comments_id_part": 0,
        "comment_id": 0,
        "comments_extensible_part": 0,
        "comment_extensible": 0,
        "reaction": 0,
        "reaction_user": 0,
    }

    for part_name, kind in sorted(part_kinds.items()):
        root = _read_xml(archive, members[part_name], limits)
        _validate_modern_comment_metadata_root(root, kind)
        records.append(
            (
                "modern_comment_metadata_part",
                kind,
                part_name,
                _fingerprint_element(
                    root,
                    relationship_maps.get(part_name, {}),
                    preserve_volatile_attributes=True,
                ),
            )
        )

        namespace, _ = _qualified_name(root.tag)
        if kind == "people":
            counts["people_part"] += 1
            people = [
                child
                for child in root
                if _is_element_in_namespace(child, namespace, "person")
            ]
            counts["person"] += len(people)
            counts["presence_info"] += sum(
                1
                for person in people
                for child in person
                if _is_element_in_namespace(child, namespace, "presenceInfo")
            )
        elif kind == "comments_extended":
            counts["comments_extended_part"] += 1
            comment_extensions = [
                child
                for child in root
                if _is_element_in_namespace(child, namespace, "commentEx")
            ]
            counts["comment_extension"] += len(comment_extensions)
            for comment_extension in comment_extensions:
                if _has_nonempty_namespaced_attribute(
                    comment_extension, namespace, "paraIdParent"
                ):
                    counts["threaded_comment"] += 1
                if _namespaced_attribute_is_enabled(
                    comment_extension, namespace, "done"
                ):
                    counts["resolved_comment"] += 1
        elif kind == "comments_ids":
            counts["comments_id_part"] += 1
            counts["comment_id"] += sum(
                1
                for child in root
                if _is_element_in_namespace(child, namespace, "commentId")
            )
        else:
            counts["comments_extensible_part"] += 1
            counts["comment_extensible"] += sum(
                1
                for child in root
                if _is_element_in_namespace(child, namespace, "commentExtensible")
            )
            for element in root.iter():
                if _is_element_in_namespace(
                    element, _MODERN_COMMENT_REACTIONS_NAMESPACE, "reaction"
                ):
                    counts["reaction"] += 1
                elif _is_element_in_namespace(
                    element, _MODERN_COMMENT_REACTIONS_NAMESPACE, "user"
                ):
                    counts["reaction_user"] += 1

    return (
        ModernCommentMetadataInventory(
            people_part_count=counts["people_part"],
            person_count=counts["person"],
            presence_info_count=counts["presence_info"],
            comments_extended_part_count=counts["comments_extended_part"],
            comment_extension_count=counts["comment_extension"],
            threaded_comment_count=counts["threaded_comment"],
            resolved_comment_count=counts["resolved_comment"],
            comments_id_part_count=counts["comments_id_part"],
            comment_id_count=counts["comment_id"],
            comments_extensible_part_count=counts["comments_extensible_part"],
            comment_extensible_count=counts["comment_extensible"],
            reaction_count=counts["reaction"],
            reaction_user_count=counts["reaction_user"],
            signature=_digest_records(records),
        ),
        frozenset(part_kinds),
    )


def _validate_modern_comment_metadata_root(root: ET.Element, kind: str) -> None:
    expected_namespaces, expected_local_name = _MODERN_COMMENT_ROOTS[kind]
    namespace, local_name = _qualified_name(root.tag)
    if namespace not in expected_namespaces or local_name != expected_local_name:
        raise DocumentFormatError("modern comment metadata part is invalid")


def _document_task_inventory(
    archive: zipfile.ZipFile,
    members: dict[str, zipfile.ZipInfo],
    content_types: dict[str, str],
    relationship_maps: dict[str, dict[str, _Relationship]],
    limits: PackageLimits,
) -> tuple[DocumentTaskInventory, frozenset[str]]:
    """Inventory Word document tasks without releasing task or user material."""

    part_names: set[str] = {
        part_name for part_name in _DOCUMENT_TASK_FALLBACK_PARTS if part_name in members
    }
    records: list[tuple[str, ...]] = []
    for part_name, content_type in content_types.items():
        if content_type.casefold() == _DOCUMENT_TASK_CONTENT_TYPE:
            part_names.add(part_name)

    for source_part, relationships in sorted(relationship_maps.items()):
        for relationship in sorted(
            relationships.values(), key=lambda value: value.canonical_value()
        ):
            if (
                relationship.relationship_type.casefold()
                not in _DOCUMENT_TASK_RELATIONSHIP_TYPES
            ):
                continue
            if relationship.target_mode.casefold() != "internal":
                raise DocumentFormatError("document task relationships are invalid")
            target = _internal_relationship_target(source_part, relationship, members)
            if target is None:
                raise DocumentFormatError("document task relationships are invalid")
            part_names.add(target)
            records.append(
                (
                    "document_task_relationship",
                    source_part,
                    *relationship.canonical_value(),
                )
            )

    counts = {
        "document_task_part": 0,
        "task": 0,
        "task_history_event": 0,
        "task_user_reference": 0,
        "task_comment_anchor": 0,
        "assignment": 0,
        "unassignment": 0,
        "creation": 0,
        "title_change": 0,
        "schedule_change": 0,
        "progress_change": 0,
        "priority_change": 0,
        "deletion": 0,
        "restoration": 0,
        "unassign_all": 0,
        "undo": 0,
    }
    for part_name in sorted(part_names):
        root = _read_xml(archive, members[part_name], limits)
        _validate_document_task_root(root)
        relationships = relationship_maps.get(part_name, {})
        records.append(
            (
                "document_task_part",
                part_name,
                _fingerprint_element(
                    root,
                    relationships,
                    preserve_volatile_attributes=True,
                ),
            )
        )
        counts["document_task_part"] += 1
        tasks = [
            child
            for child in root
            if _is_element_in_namespace(child, _DOCUMENT_TASK_NAMESPACE, "Task")
        ]
        counts["task"] += len(tasks)
        counts["task_comment_anchor"] += sum(
            1
            for element in root.iter()
            if _is_element_in_namespace(element, _DOCUMENT_TASK_NAMESPACE, "Comment")
        )
        for task in tasks:
            for history in task:
                if not _is_element_in_namespace(
                    history, _DOCUMENT_TASK_NAMESPACE, "History"
                ):
                    continue
                for event in history:
                    if not _is_element_in_namespace(
                        event, _DOCUMENT_TASK_NAMESPACE, "Event"
                    ):
                        continue
                    counts["task_history_event"] += 1
                    for event_child in event:
                        namespace, local_name = _qualified_name(event_child.tag)
                        if namespace != _DOCUMENT_TASK_NAMESPACE:
                            continue
                        if local_name in _DOCUMENT_TASK_USER_REFERENCE_NAMES:
                            counts["task_user_reference"] += 1
                        event_count_key = _DOCUMENT_TASK_EVENT_COUNT_KEYS.get(
                            local_name
                        )
                        if event_count_key is not None:
                            counts[event_count_key] += 1

    return (
        DocumentTaskInventory(
            document_task_part_count=counts["document_task_part"],
            task_count=counts["task"],
            task_history_event_count=counts["task_history_event"],
            task_user_reference_count=counts["task_user_reference"],
            task_comment_anchor_count=counts["task_comment_anchor"],
            assignment_event_count=counts["assignment"],
            unassignment_event_count=counts["unassignment"],
            creation_event_count=counts["creation"],
            title_change_event_count=counts["title_change"],
            schedule_change_event_count=counts["schedule_change"],
            progress_change_event_count=counts["progress_change"],
            priority_change_event_count=counts["priority_change"],
            deletion_event_count=counts["deletion"],
            restoration_event_count=counts["restoration"],
            unassign_all_event_count=counts["unassign_all"],
            undo_event_count=counts["undo"],
            signature=_digest_records(records),
        ),
        frozenset(part_names),
    )


def _validate_document_task_root(root: ET.Element) -> None:
    if _qualified_name(root.tag) != (_DOCUMENT_TASK_NAMESPACE, "Tasks"):
        raise DocumentFormatError("document task part is invalid")


def _taskpane_web_extension_inventory(
    archive: zipfile.ZipFile,
    members: dict[str, zipfile.ZipInfo],
    content_types: dict[str, str],
    relationship_maps: dict[str, dict[str, _Relationship]],
    control_references: tuple[_WebExtensionControlReference, ...],
    limits: PackageLimits,
) -> tuple[TaskpaneWebExtensionInventory, frozenset[str]]:
    """Inventory document-borne task-pane Office add-in state safely.

    The stored references, property values, bindings, and part paths are kept
    exclusively in private fingerprints.  Public output is limited to counts
    that let a reviewer decide whether this document carries add-in state.
    """

    part_kinds: dict[str, str] = {}
    records: list[tuple[str, ...]] = []

    def add_part(kind: str, part_name: str) -> None:
        existing_kind = part_kinds.get(part_name)
        if existing_kind is not None and existing_kind != kind:
            raise DocumentFormatError("task-pane web extension parts are inconsistent")
        part_kinds[part_name] = kind

    for part_name, kind in _TASKPANE_WEB_EXTENSION_FALLBACK_PARTS.items():
        if part_name in members:
            add_part(kind, part_name)
    for part_name in members:
        if _is_conventional_web_extension_part_name(part_name):
            add_part("web_extension", part_name)
    for part_name, content_type in content_types.items():
        kind = _TASKPANE_WEB_EXTENSION_PART_CONTENT_TYPES.get(content_type.casefold())
        if kind is not None:
            add_part(kind, part_name)

    for source_part, relationships in sorted(relationship_maps.items()):
        for relationship in sorted(
            relationships.values(), key=lambda value: value.canonical_value()
        ):
            kind = _TASKPANE_WEB_EXTENSION_PART_RELATIONSHIP_TYPES.get(
                relationship.relationship_type.casefold()
            )
            if kind is None:
                continue
            if relationship.target_mode.casefold() != "internal":
                raise DocumentFormatError(
                    "task-pane web extension relationships are invalid"
                )
            target = _internal_relationship_target(source_part, relationship, members)
            if target is None:
                raise DocumentFormatError(
                    "task-pane web extension relationships are invalid"
                )
            add_part(kind, target)
            records.append(
                (
                    "taskpane_web_extension_relationship",
                    kind,
                    source_part,
                    *relationship.canonical_value(),
                )
            )

    counts = {
        "taskpane_part": 0,
        "taskpane": 0,
        "visible_taskpane": 0,
        "locked_taskpane": 0,
        "web_extension_part": 0,
        "web_extension_reference": 0,
        "web_extension_property": 0,
        "web_extension_binding": 0,
        "auto_show_taskpane_setting": 0,
    }
    taskpane_part_names = sorted(
        part_name for part_name, kind in part_kinds.items() if kind == "taskpanes"
    )
    for part_name in taskpane_part_names:
        root = _read_xml(archive, members[part_name], limits)
        _validate_taskpane_web_extension_root(root, "taskpanes")
        relationships = relationship_maps.get(part_name, {})
        records.append(
            (
                "taskpane_web_extension_part",
                "taskpanes",
                part_name,
                _fingerprint_element(
                    root,
                    relationships,
                    preserve_volatile_attributes=True,
                ),
            )
        )
        counts["taskpane_part"] += 1
        taskpanes = [
            child
            for child in root
            if _is_element_in_namespace(
                child, _TASKPANE_WEB_EXTENSION_TASKPANES_NAMESPACE, "taskpane"
            )
        ]
        counts["taskpane"] += len(taskpanes)
        for taskpane in taskpanes:
            if _unqualified_attribute_is_enabled(taskpane, "visibility"):
                counts["visible_taskpane"] += 1
            if _unqualified_attribute_is_enabled(taskpane, "locked"):
                counts["locked_taskpane"] += 1
            references = [
                child
                for child in taskpane
                if _qualified_name(child.tag) in _TASKPANE_REFERENCE_TAGS
            ]
            if len(references) != 1:
                raise DocumentFormatError("task-pane web extension markup is invalid")
            relationship_id = _relationship_id_value(references[0])
            relationship = (
                relationships.get(relationship_id)
                if relationship_id is not None
                else None
            )
            if (
                relationship is None
                or relationship.relationship_type.casefold()
                != "http://schemas.microsoft.com/office/2011/relationships/webextension"
                or relationship.target_mode.casefold() != "internal"
            ):
                raise DocumentFormatError("task-pane web extension markup is invalid")
            target = _internal_relationship_target(part_name, relationship, members)
            if target is None:
                raise DocumentFormatError("task-pane web extension markup is invalid")
            add_part("web_extension", target)
            records.append(
                (
                    "taskpane_web_extension_reference",
                    part_name,
                    *relationship.canonical_value(),
                )
            )

    web_extension_part_names = sorted(
        part_name for part_name, kind in part_kinds.items() if kind == "web_extension"
    )
    for part_name in web_extension_part_names:
        root = _read_xml(archive, members[part_name], limits)
        _validate_taskpane_web_extension_root(root, "web_extension")
        records.append(
            (
                "taskpane_web_extension_part",
                "web_extension",
                part_name,
                _fingerprint_element(
                    root,
                    relationship_maps.get(part_name, {}),
                    preserve_volatile_attributes=True,
                ),
            )
        )
        counts["web_extension_part"] += 1
        for child in root:
            namespace, local_name = _qualified_name(child.tag)
            if namespace != _TASKPANE_WEB_EXTENSION_NAMESPACE:
                continue
            if local_name == "reference":
                counts["web_extension_reference"] += 1
            elif local_name == "alternateReferences":
                counts["web_extension_reference"] += sum(
                    1
                    for reference in child
                    if _is_element_in_namespace(
                        reference, _TASKPANE_WEB_EXTENSION_NAMESPACE, "reference"
                    )
                )
            elif local_name == "properties":
                properties = [
                    property_element
                    for property_element in child
                    if _is_element_in_namespace(
                        property_element, _TASKPANE_WEB_EXTENSION_NAMESPACE, "property"
                    )
                ]
                counts["web_extension_property"] += len(properties)
                counts["auto_show_taskpane_setting"] += sum(
                    1
                    for property_element in properties
                    if _unqualified_attribute_value(property_element, "name")
                    == _AUTO_SHOW_TASKPANE_PROPERTY_NAME
                    and _unqualified_attribute_is_enabled(property_element, "value")
                )
            elif local_name == "bindings":
                counts["web_extension_binding"] += sum(
                    1
                    for binding in child
                    if _is_element_in_namespace(
                        binding, _TASKPANE_WEB_EXTENSION_NAMESPACE, "binding"
                    )
                )

    records.extend(
        (
            "web_extension_bound_content_control",
            reference.story_part,
            reference.marker_kind,
            reference.signature,
        )
        for reference in control_references
    )
    return (
        TaskpaneWebExtensionInventory(
            taskpane_part_count=counts["taskpane_part"],
            taskpane_count=counts["taskpane"],
            visible_taskpane_count=counts["visible_taskpane"],
            locked_taskpane_count=counts["locked_taskpane"],
            web_extension_part_count=counts["web_extension_part"],
            web_extension_reference_count=counts["web_extension_reference"],
            web_extension_property_count=counts["web_extension_property"],
            web_extension_binding_count=counts["web_extension_binding"],
            auto_show_taskpane_setting_count=counts["auto_show_taskpane_setting"],
            web_extension_bound_content_control_count=len(control_references),
            signature=_digest_records(records),
        ),
        frozenset(part_kinds),
    )


def _is_conventional_web_extension_part_name(part_name: str) -> bool:
    return (
        part_name
        in {"word/webextensions/webextension", "word/webextensions/webextension.xml"}
        or (
            part_name.startswith("word/webextensions/webextension")
            and part_name.endswith(".xml")
        )
    ) and "/_rels/" not in part_name


def _validate_taskpane_web_extension_root(root: ET.Element, kind: str) -> None:
    expected = (
        (_TASKPANE_WEB_EXTENSION_TASKPANES_NAMESPACE, "taskpanes")
        if kind == "taskpanes"
        else (_TASKPANE_WEB_EXTENSION_NAMESPACE, "webextension")
    )
    if _qualified_name(root.tag) != expected:
        raise DocumentFormatError("task-pane web extension part is invalid")


def _is_element_in_namespace(
    element: ET.Element, namespace: str, local_name: str
) -> bool:
    return _qualified_name(element.tag) == (namespace, local_name)


def _has_nonempty_namespaced_attribute(
    element: ET.Element, namespace: str, local_name: str
) -> bool:
    value = _namespaced_attribute_value(element, namespace, local_name)
    return bool(value and value.strip())


def _namespaced_attribute_is_enabled(
    element: ET.Element, namespace: str, local_name: str
) -> bool:
    value = _namespaced_attribute_value(element, namespace, local_name)
    return value is not None and value.strip().casefold() not in _FALSE_VALUES


def _namespaced_attribute_value(
    element: ET.Element, namespace: str, local_name: str
) -> str | None:
    for attribute, value in element.attrib.items():
        if _qualified_name(attribute) == (namespace, local_name):
            return value
    return None


def _unqualified_attribute_value(element: ET.Element, local_name: str) -> str | None:
    for attribute, value in element.attrib.items():
        attribute_namespace, attribute_local_name = _qualified_name(attribute)
        if attribute_namespace == "" and attribute_local_name == local_name:
            return value
    return None


def _unqualified_attribute_is_enabled(element: ET.Element, local_name: str) -> bool:
    value = _unqualified_attribute_value(element, local_name)
    return value is not None and value.strip().casefold() not in _FALSE_VALUES


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
) -> tuple[
    tuple[StorySnapshot, ...],
    int,
    tuple[_ContentControlLockReference, ...],
    tuple[_DataBindingReference, ...],
    tuple[_ExternalFieldReference, ...],
    tuple[_DocumentVariableFieldReference, ...],
    tuple[_HyperlinkFieldReference, ...],
    tuple[_HyperlinkMarkupReference, ...],
    tuple[_DrawingHyperlinkReference, ...],
    tuple[_DrawingVisibilityReference, ...],
    tuple[_DrawingLinkedPictureReference, ...],
    tuple[_VmlHyperlinkReference, ...],
    tuple[_VmlExternalImageReference, ...],
    tuple[_VmlImageHyperlinkReference, ...],
    tuple[_VmlLinkedOleObjectReference, ...],
    tuple[_VmlEmbeddedOleObjectReference, ...],
    tuple[_WordObjectLinkReference, ...],
    tuple[_WordEmbeddedControlReference, ...],
    tuple[_WebExtensionControlReference, ...],
]:
    stories: list[StorySnapshot] = []
    content_control_lock_references: list[_ContentControlLockReference] = []
    data_binding_references: list[_DataBindingReference] = []
    external_field_references: list[_ExternalFieldReference] = []
    document_variable_field_references: list[_DocumentVariableFieldReference] = []
    hyperlink_field_references: list[_HyperlinkFieldReference] = []
    hyperlink_markup_references: list[_HyperlinkMarkupReference] = []
    drawing_hyperlink_references: list[_DrawingHyperlinkReference] = []
    drawing_visibility_references: list[_DrawingVisibilityReference] = []
    drawing_linked_picture_references: list[_DrawingLinkedPictureReference] = []
    vml_hyperlink_references: list[_VmlHyperlinkReference] = []
    vml_external_image_references: list[_VmlExternalImageReference] = []
    vml_image_hyperlink_references: list[_VmlImageHyperlinkReference] = []
    vml_linked_ole_object_references: list[_VmlLinkedOleObjectReference] = []
    vml_embedded_ole_object_references: list[_VmlEmbeddedOleObjectReference] = []
    word_object_link_references: list[_WordObjectLinkReference] = []
    word_embedded_control_references: list[_WordEmbeddedControlReference] = []
    web_extension_control_references: list[_WebExtensionControlReference] = []
    comment_count = 0
    for part_key, kind in _discover_story_parts(members, content_types):
        root = _read_xml(archive, members[part_key], limits)
        (
            story,
            story_content_control_lock_references,
            story_data_binding_references,
            story_external_field_references,
            story_document_variable_field_references,
            story_hyperlink_field_references,
            story_hyperlink_markup_references,
            story_drawing_hyperlink_references,
            story_drawing_visibility_references,
            story_drawing_linked_picture_references,
            story_vml_hyperlink_references,
            story_vml_external_image_references,
            story_vml_image_hyperlink_references,
            story_vml_linked_ole_object_references,
            story_vml_embedded_ole_object_references,
            story_word_object_link_references,
            story_word_embedded_control_references,
        ) = _snapshot_story(root, part_key, kind, relationship_maps.get(part_key, {}))
        stories.append(story)
        content_control_lock_references.extend(story_content_control_lock_references)
        data_binding_references.extend(story_data_binding_references)
        external_field_references.extend(story_external_field_references)
        document_variable_field_references.extend(
            story_document_variable_field_references
        )
        hyperlink_field_references.extend(story_hyperlink_field_references)
        hyperlink_markup_references.extend(story_hyperlink_markup_references)
        drawing_hyperlink_references.extend(story_drawing_hyperlink_references)
        drawing_visibility_references.extend(story_drawing_visibility_references)
        drawing_linked_picture_references.extend(
            story_drawing_linked_picture_references
        )
        vml_hyperlink_references.extend(story_vml_hyperlink_references)
        vml_external_image_references.extend(story_vml_external_image_references)
        vml_image_hyperlink_references.extend(story_vml_image_hyperlink_references)
        vml_linked_ole_object_references.extend(story_vml_linked_ole_object_references)
        vml_embedded_ole_object_references.extend(
            story_vml_embedded_ole_object_references
        )
        word_object_link_references.extend(story_word_object_link_references)
        word_embedded_control_references.extend(story_word_embedded_control_references)
        web_extension_control_references.extend(
            _web_extension_control_references(
                root, part_key, relationship_maps.get(part_key, {})
            )
        )
        if kind == "comment":
            comment_count += _count_word_elements(root, "comment")
    if not any(story.kind == "body" for story in stories):
        raise DocumentFormatError("package does not contain a document story")
    return (
        tuple(stories),
        comment_count,
        tuple(content_control_lock_references),
        tuple(data_binding_references),
        tuple(external_field_references),
        tuple(document_variable_field_references),
        tuple(hyperlink_field_references),
        tuple(hyperlink_markup_references),
        tuple(drawing_hyperlink_references),
        tuple(drawing_visibility_references),
        tuple(drawing_linked_picture_references),
        tuple(vml_hyperlink_references),
        tuple(vml_external_image_references),
        tuple(vml_image_hyperlink_references),
        tuple(vml_linked_ole_object_references),
        tuple(vml_embedded_ole_object_references),
        tuple(word_object_link_references),
        tuple(word_embedded_control_references),
        tuple(web_extension_control_references),
    )


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
) -> tuple[
    StorySnapshot,
    tuple[_ContentControlLockReference, ...],
    tuple[_DataBindingReference, ...],
    tuple[_ExternalFieldReference, ...],
    tuple[_DocumentVariableFieldReference, ...],
    tuple[_HyperlinkFieldReference, ...],
    tuple[_HyperlinkMarkupReference, ...],
    tuple[_DrawingHyperlinkReference, ...],
    tuple[_DrawingVisibilityReference, ...],
    tuple[_DrawingLinkedPictureReference, ...],
    tuple[_VmlHyperlinkReference, ...],
    tuple[_VmlExternalImageReference, ...],
    tuple[_VmlImageHyperlinkReference, ...],
    tuple[_VmlLinkedOleObjectReference, ...],
    tuple[_VmlEmbeddedOleObjectReference, ...],
    tuple[_WordObjectLinkReference, ...],
    tuple[_WordEmbeddedControlReference, ...],
]:
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
    field_instruction_references = _field_instruction_references(root, part_key)
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
        _content_control_lock_references(root, part_key),
        _data_binding_references(root, part_key, relationships),
        _external_field_references(field_instruction_references),
        _document_variable_field_references(field_instruction_references),
        _hyperlink_field_references(field_instruction_references),
        _hyperlink_markup_references(root, part_key, relationships),
        _drawing_hyperlink_references(root, part_key, relationships),
        _drawing_visibility_references(root, part_key),
        _drawing_linked_picture_references(root, part_key, relationships),
        _vml_hyperlink_references(root, part_key, relationships),
        _vml_external_image_references(root, part_key, relationships),
        _vml_image_hyperlink_references(root, part_key, relationships),
        _vml_linked_ole_object_references(root, part_key, relationships),
        _vml_embedded_ole_object_references(root, part_key, relationships),
        _word_object_link_references(root, part_key, relationships),
        _word_embedded_control_references(root, part_key, relationships),
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


def _content_control_lock_references(
    root: ET.Element, story_part: str
) -> tuple[_ContentControlLockReference, ...]:
    """Find one direct lock state for every standard Word content control.

    A lock is meaningful here only as a direct ``w:sdtPr/w:lock`` leaf of the
    particular ``w:sdt`` being counted.  Markup that merely has the same local
    names in another nesting context is intentionally outside this inventory.
    """

    references: list[_ContentControlLockReference] = []
    controls = (element for element in root.iter() if _is_word_element(element, "sdt"))
    for ordinal, control in enumerate(controls):
        properties = [child for child in control if _is_word_element(child, "sdtPr")]
        if len(properties) > 1:
            raise DocumentFormatError("content-control lock state is invalid")
        locks = (
            [child for child in properties[0] if _is_word_element(child, "lock")]
            if properties
            else []
        )
        if len(locks) > 1:
            raise DocumentFormatError("content-control lock state is invalid")
        state = (
            _content_control_lock_state(locks[0]) if locks else "no_lock_declaration"
        )
        references.append(
            _ContentControlLockReference(
                story_part=story_part,
                ordinal=ordinal,
                state=state,
            )
        )
    return tuple(references)


def _content_control_lock_state(element: ET.Element) -> str:
    """Validate one direct ``w:sdtPr/w:lock`` leaf and normalize its state."""

    if (
        not _is_word_element(element, "lock")
        or list(element)
        or (element.text or "").strip()
        or len(element.attrib) != 1
    ):
        raise DocumentFormatError("content-control lock state is invalid")
    value = _word_attribute_value(element, "val")
    state = _CONTENT_CONTROL_LOCK_STATE_BY_VALUE.get(value)
    if state is None:
        raise DocumentFormatError("content-control lock state is invalid")
    return state


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


def _web_extension_control_references(
    root: ET.Element,
    story_part: str,
    relationships: dict[str, _Relationship],
) -> tuple[_WebExtensionControlReference, ...]:
    """Find enabled Word content-control markers bound to web extensions.

    Word gives ``webExtensionCreated`` precedence over ``webExtensionLinked``
    when both appear in the same ``w:sdtPr``.  The marker values and every
    other property remain private inside a digest; this inventory publishes a
    count of enabled, document-borne bindings only.
    """

    references: list[_WebExtensionControlReference] = []
    for properties in root.iter():
        if not _is_word_element(properties, "sdtPr"):
            continue
        created_markers = [
            child
            for child in properties
            if _is_element_in_namespace(
                child, _WORD_2012_NAMESPACE, "webExtensionCreated"
            )
        ]
        linked_markers = [
            child
            for child in properties
            if _is_element_in_namespace(
                child, _WORD_2012_NAMESPACE, "webExtensionLinked"
            )
        ]
        if created_markers:
            marker_kind = "created"
            selected_markers = created_markers
        elif linked_markers:
            marker_kind = "linked"
            selected_markers = linked_markers
        else:
            continue
        if not any(_is_enabled(marker) for marker in selected_markers):
            continue
        references.append(
            _WebExtensionControlReference(
                story_part=story_part,
                marker_kind=marker_kind,
                signature=_digest_records(
                    [
                        (
                            "web_extension_control_marker",
                            _fingerprint_element(
                                marker,
                                relationships,
                                preserve_volatile_attributes=True,
                            ),
                        )
                        for marker in selected_markers
                    ]
                ),
            )
        )
    return tuple(references)


def _field_instruction_references(
    root: ET.Element, story_part: str
) -> tuple[_FieldInstructionReference, ...]:
    """Collect complete Word field instructions in both OOXML encodings.

    A ``w:fldSimple`` stores its instruction directly in ``w:instr``. Complex
    fields use begin/separate/end characters and may split their instruction
    across multiple ``w:instrText`` runs. Revision markup can preserve both a
    current instruction and a deleted instruction in one complex field, so the
    two variants are assembled independently. Instruction text outside a
    complete complex field's instruction portion is ordinary text under
    ISO/IEC 29500, so it is intentionally ignored here.

    Nested fields occurring before a parent complex field's separator make the
    parent's assembled argument incomplete. The parent remains a complete
    field instruction for family-level inventories, but the nested marker lets
    literal argument consumers remain deliberately conservative.
    """

    references: list[_FieldInstructionReference] = []
    complex_fields: list[_ComplexFieldState] = []

    def add_instruction(
        instruction: str | None, *, contains_nested_instruction_field: bool
    ) -> None:
        if instruction is None:
            return
        references.append(
            _FieldInstructionReference(
                story_part=story_part,
                instruction=instruction,
                contains_nested_instruction_field=contains_nested_instruction_field,
            )
        )

    def add_complex_field_instructions(field: _ComplexFieldState) -> None:
        seen_instructions: set[str] = set()
        for variant in _FIELD_INSTRUCTION_VARIANTS:
            instruction = "".join(field.instruction_chunks[variant])
            if not instruction or instruction in seen_instructions:
                continue
            seen_instructions.add(instruction)
            add_instruction(
                instruction,
                contains_nested_instruction_field=(
                    variant in field.nested_instruction_field_variants
                ),
            )

    def append_instruction_text(
        field: _ComplexFieldState, text: str, variants: frozenset[str]
    ) -> None:
        for variant in variants:
            field.instruction_chunks[variant].append(text)

    def mark_nested_instruction_field(variants: frozenset[str]) -> None:
        for field in complex_fields:
            if field.accepts_instructions:
                field.nested_instruction_field_variants.update(variants)

    def visit(
        element: ET.Element,
        simple_field_depth: int,
        instruction_variants: frozenset[str],
        in_deleted_revision: bool,
    ) -> None:
        namespace, local_name = _qualified_name(element.tag)
        child_simple_field_depth = simple_field_depth
        child_instruction_variants = instruction_variants
        child_in_deleted_revision = in_deleted_revision
        if namespace in _WORD_NAMESPACES:
            if local_name in _CURRENT_FIELD_REVISION_TAGS:
                child_instruction_variants = (
                    instruction_variants & _CURRENT_FIELD_INSTRUCTION_VARIANTS
                )
                child_in_deleted_revision = False
            elif local_name in _DELETED_FIELD_REVISION_TAGS:
                child_instruction_variants = (
                    instruction_variants & _DELETED_FIELD_INSTRUCTION_VARIANTS
                )
                child_in_deleted_revision = True

            if local_name == "fldSimple":
                if instruction_variants:
                    mark_nested_instruction_field(instruction_variants)
                    add_instruction(
                        _word_attribute_value(element, "instr"),
                        contains_nested_instruction_field=False,
                    )
                child_simple_field_depth += 1
            elif local_name == "fldChar":
                field_char_type = _field_char_type(element)
                if field_char_type == "begin":
                    mark_nested_instruction_field(instruction_variants)
                    complex_fields.append(
                        _ComplexFieldState(
                            simple_field_depth=simple_field_depth,
                            instruction_chunks={
                                variant: [] for variant in _FIELD_INSTRUCTION_VARIANTS
                            },
                            nested_instruction_field_variants=set(),
                        )
                    )
                elif field_char_type == "separate" and complex_fields:
                    complex_fields[-1].accepts_instructions = False
                elif field_char_type == "end" and complex_fields:
                    field = complex_fields.pop()
                    add_complex_field_instructions(field)
            elif (
                local_name == "instrText"
                and complex_fields
                and complex_fields[-1].accepts_instructions
                and complex_fields[-1].simple_field_depth == simple_field_depth
                and not in_deleted_revision
            ):
                append_instruction_text(
                    complex_fields[-1], element.text or "", instruction_variants
                )
            elif (
                local_name == "delInstrText"
                and complex_fields
                and complex_fields[-1].accepts_instructions
                and complex_fields[-1].simple_field_depth == simple_field_depth
                and in_deleted_revision
            ):
                append_instruction_text(
                    complex_fields[-1], element.text or "", instruction_variants
                )

        for child in element:
            visit(
                child,
                child_simple_field_depth,
                child_instruction_variants,
                child_in_deleted_revision,
            )

    visit(root, 0, frozenset(_FIELD_INSTRUCTION_VARIANTS), False)
    return tuple(references)


def _external_field_references(
    field_instruction_references: tuple[_FieldInstructionReference, ...],
) -> tuple[_ExternalFieldReference, ...]:
    """Select external-source field instructions from complete field codes."""

    references: list[_ExternalFieldReference] = []
    for reference in field_instruction_references:
        category = _external_field_category(reference.instruction)
        if category is None:
            continue
        references.append(
            _ExternalFieldReference(
                story_part=reference.story_part,
                category=category,
                instruction_signature=_digest_bytes(
                    reference.instruction.encode("utf-8")
                ),
            )
        )
    return tuple(references)


def _document_variable_field_references(
    field_instruction_references: tuple[_FieldInstructionReference, ...],
) -> tuple[_DocumentVariableFieldReference, ...]:
    """Select ``DOCVARIABLE`` field codes without retaining names publicly."""

    references: list[_DocumentVariableFieldReference] = []
    for reference in field_instruction_references:
        if _field_instruction_keyword(reference.instruction) != "docvariable":
            continue
        references.append(
            _DocumentVariableFieldReference(
                story_part=reference.story_part,
                document_scope=_document_scope_for_story(reference.story_part),
                instruction_signature=_digest_bytes(
                    reference.instruction.encode("utf-8")
                ),
                literal_name=_document_variable_field_literal_name(
                    reference.instruction,
                    contains_nested_instruction_field=(
                        reference.contains_nested_instruction_field
                    ),
                ),
            )
        )
    return tuple(references)


def _hyperlink_field_references(
    field_instruction_references: tuple[_FieldInstructionReference, ...],
) -> tuple[_HyperlinkFieldReference, ...]:
    """Select ``HYPERLINK`` fields without retaining any destination publicly.

    A Hyperlink field can point to a web address, a file, or a bookmark with no
    separate OOXML relationship. This is a stored-field inventory only: it
    neither resolves nor follows any target, and it does not label a literal
    destination as external because Word also accepts bookmarks there.
    """

    references: list[_HyperlinkFieldReference] = []
    for reference in field_instruction_references:
        if _field_instruction_keyword(reference.instruction) != "hyperlink":
            continue
        references.append(
            _HyperlinkFieldReference(
                story_part=reference.story_part,
                instruction_signature=_digest_bytes(
                    reference.instruction.encode("utf-8")
                ),
                classification=_hyperlink_field_classification(
                    reference.instruction,
                    contains_nested_instruction_field=(
                        reference.contains_nested_instruction_field
                    ),
                ),
            )
        )
    return tuple(references)


def _document_variable_field_literal_name(
    instruction: str, *, contains_nested_instruction_field: bool
) -> str | None:
    """Return one conservative, complete literal ``DOCVARIABLE`` argument.

    A plain argument with no whitespace, or one wholly enclosed by one pair of
    double quotes, is accepted. Trailing Word field-switch material beginning
    with ``\\`` does not change that leading argument. Escaped/nested or
    otherwise compound field expressions deliberately remain nonliteral:
    DocFence reports their stored presence without implementing Word's field
    evaluator.
    """

    if contains_nested_instruction_field:
        return None
    return _leading_literal_field_argument(_field_instruction_argument(instruction))


def _hyperlink_field_classification(
    instruction: str, *, contains_nested_instruction_field: bool
) -> str:
    """Classify a complete ``HYPERLINK`` field without exposing its target.

    ``HYPERLINK field-argument [ switches ]`` can use a URL, file name, or
    bookmark as its first argument. ``HYPERLINK \\l field-argument`` is the
    standard internal-location-only form. We recognize only a leading plain or
    wholly quoted literal followed by optional switch material. Any nested,
    compound, missing, or malformed argument stays dynamic or unparseable;
    DocFence intentionally does not implement Word's field evaluator.
    """

    if contains_nested_instruction_field:
        return "dynamic_or_unparseable"
    argument = _field_instruction_argument(instruction)
    if _leading_literal_field_argument(argument) is not None:
        return "literal_destination"
    if _hyperlink_internal_location_only_literal(argument) is not None:
        return "literal_internal_location_only"
    return "dynamic_or_unparseable"


def _field_instruction_argument(instruction: str) -> str | None:
    """Return the trimmed material following one field keyword, if present."""

    tokens = instruction.lstrip().split(maxsplit=1)
    if len(tokens) != 2:
        return None
    argument = tokens[1].strip()
    return argument or None


def _leading_literal_field_argument(argument: str | None) -> str | None:
    """Return one conservative leading literal field argument.

    A literal is a plain, unquoted token or one wholly enclosed by one pair of
    double quotes. Any remaining material must start with a Word field switch.
    This parser deliberately declines escaped quotes and compound expressions
    rather than trying to reproduce Word field evaluation semantics.
    """

    if not argument:
        return None
    if not argument.startswith('"'):
        argument_tokens = argument.split(maxsplit=1)
        literal_argument = argument_tokens[0]
        suffix = argument_tokens[1] if len(argument_tokens) == 2 else ""
        if (
            literal_argument.startswith("\\")
            or '"' in literal_argument
            or (suffix and not suffix.startswith("\\"))
        ):
            return None
        return literal_argument

    closing_quote = argument.find('"', 1)
    suffix = argument[closing_quote + 1 :].strip() if closing_quote >= 0 else ""
    if closing_quote < 0 or (suffix and not suffix.startswith("\\")):
        return None
    return argument[1:closing_quote]


def _hyperlink_internal_location_only_literal(argument: str | None) -> str | None:
    """Return a literal ``\\l`` location for switch-only Hyperlink fields."""

    if not argument:
        return None
    switch_tokens = argument.split(maxsplit=1)
    if len(switch_tokens) != 2 or switch_tokens[0].casefold() != "\\l":
        return None
    return _leading_literal_field_argument(switch_tokens[1])


def _document_scope_for_story(story_part: str) -> str:
    """Map a supported story to its main or glossary document scope."""

    if story_part == "word/glossary/document.xml":
        return story_part
    return "word/document.xml"


def _external_field_category(instruction: str) -> str | None:
    """Return the review category for a complete Word field instruction."""

    keyword = _field_instruction_keyword(instruction)
    if keyword is None:
        return None
    return _EXTERNAL_FIELD_CATEGORY_BY_KEYWORD.get(keyword)


def _field_instruction_keyword(instruction: str) -> str | None:
    tokens = instruction.lstrip().split(maxsplit=1)
    if not tokens:
        return None
    return tokens[0].casefold()


def _word_attribute_value(element: ET.Element, local_name: str) -> str | None:
    for attribute, value in element.attrib.items():
        namespace, attribute_local_name = _qualified_name(attribute)
        if namespace in _WORD_NAMESPACES and attribute_local_name == local_name:
            return value
    return None


def _unqualified_attribute_value(element: ET.Element, local_name: str) -> str | None:
    """Return a direct unqualified attribute without treating empty as absent."""

    for attribute, value in element.attrib.items():
        namespace, attribute_local_name = _qualified_name(attribute)
        if not namespace and attribute_local_name == local_name:
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
    return _relationship_attribute_value(element, "id")


def _relationship_attribute_value(element: ET.Element, local_name: str) -> str | None:
    """Return one direct relationship-namespace attribute, including empty."""

    for attribute, value in element.attrib.items():
        namespace, attribute_local_name = _qualified_name(attribute)
        if (
            namespace in _REL_ATTRIBUTE_NAMESPACES
            and attribute_local_name == local_name
        ):
            return value
    return None


def _field_char_is_begin(element: ET.Element) -> bool:
    return _field_char_type(element) == "begin"


def _field_char_type(element: ET.Element) -> str | None:
    for attribute, value in element.attrib.items():
        if _local_name(attribute) == "fldCharType":
            return value.strip().casefold()
    return None


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
) -> tuple[bool, str, frozenset[str], dict[str, frozenset[str]]]:
    """Fingerprint every discovered document Settings part."""

    settings_part_scopes = _document_settings_part_scopes(members, relationship_maps)
    settings_part_names = tuple(sorted(settings_part_scopes))
    if not settings_part_names:
        return False, _digest_bytes(b"settings-absent"), frozenset(), {}

    enabled = False
    records: list[tuple[str, str, str]] = []
    for settings_part in settings_part_names:
        root = _read_xml(archive, members[settings_part], limits)
        if not _is_word_element(root, "settings"):
            raise DocumentFormatError("document settings are invalid")
        if any(
            _is_word_element(element, "trackRevisions") and _is_enabled(element)
            for element in root.iter()
        ):
            enabled = True
        records.append(
            (
                "settings",
                settings_part,
                _fingerprint_element(
                    root,
                    relationship_maps.get(settings_part, {}),
                    ignore_rsids=True,
                ),
            )
        )
    return (
        enabled,
        _digest_records(records),
        frozenset(settings_part_names),
        settings_part_scopes,
    )


def _word_protection_inventory(
    archive: zipfile.ZipFile,
    members: dict[str, zipfile.ZipInfo],
    relationship_maps: dict[str, dict[str, _Relationship]],
    settings_part_names: frozenset[str],
    limits: PackageLimits,
) -> WordProtectionInventory:
    """Inventory stored Word editing/write protection without exposing verifiers.

    The protection elements are direct children of a document Settings part.
    Their opaque password-verifier and cryptographic fields are retained only
    inside the private element fingerprint. This inventory deliberately does
    not validate password construction, strength, or effective application
    enforcement.
    """

    records: list[tuple[str, str, str]] = []
    document_protection_count = 0
    document_protection_enforcement_enabled_count = 0
    document_protection_formatting_restricted_count = 0
    document_protection_read_only_count = 0
    document_protection_comments_count = 0
    document_protection_tracked_changes_count = 0
    document_protection_forms_count = 0
    document_protection_password_material_count = 0
    write_protection_count = 0
    write_protection_recommended_count = 0
    write_protection_password_material_count = 0

    for settings_part in sorted(settings_part_names):
        root = _read_xml(archive, members[settings_part], limits)
        if not _is_word_element(root, "settings"):
            raise DocumentFormatError("document settings are invalid")
        document_protections = [
            element
            for element in root
            if _is_word_element(element, "documentProtection")
        ]
        write_protections = [
            element for element in root if _is_word_element(element, "writeProtection")
        ]
        if len(document_protections) > 1 or len(write_protections) > 1:
            raise DocumentFormatError("Word protection state is invalid")

        relationships = relationship_maps.get(settings_part, {})
        for element in document_protections:
            attributes = _validate_word_protection_element(
                element,
                "documentProtection",
                _DOCUMENT_PROTECTION_ATTRIBUTE_NAMES,
            )
            document_protection_count += 1
            if _word_protection_attribute_enabled(attributes, "enforcement"):
                document_protection_enforcement_enabled_count += 1
            if _word_protection_attribute_enabled(attributes, "formatting"):
                document_protection_formatting_restricted_count += 1
            edit_mode = attributes.get("edit")
            if edit_mode == "readOnly":
                document_protection_read_only_count += 1
            elif edit_mode == "comments":
                document_protection_comments_count += 1
            elif edit_mode == "trackedChanges":
                document_protection_tracked_changes_count += 1
            elif edit_mode == "forms":
                document_protection_forms_count += 1
            if _WORD_PROTECTION_PASSWORD_MATERIAL_ATTRIBUTE_NAMES & attributes.keys():
                document_protection_password_material_count += 1
            records.append(
                (
                    "document_protection",
                    settings_part,
                    _fingerprint_element(element, relationships),
                )
            )

        for element in write_protections:
            attributes = _validate_word_protection_element(
                element,
                "writeProtection",
                _WRITE_PROTECTION_ATTRIBUTE_NAMES,
            )
            write_protection_count += 1
            if _word_protection_attribute_enabled(attributes, "recommended"):
                write_protection_recommended_count += 1
            if _WORD_PROTECTION_PASSWORD_MATERIAL_ATTRIBUTE_NAMES & attributes.keys():
                write_protection_password_material_count += 1
            records.append(
                (
                    "write_protection",
                    settings_part,
                    _fingerprint_element(element, relationships),
                )
            )

    return WordProtectionInventory(
        document_protection_count=document_protection_count,
        document_protection_enforcement_enabled_count=(
            document_protection_enforcement_enabled_count
        ),
        document_protection_formatting_restricted_count=(
            document_protection_formatting_restricted_count
        ),
        document_protection_read_only_count=document_protection_read_only_count,
        document_protection_comments_count=document_protection_comments_count,
        document_protection_tracked_changes_count=(
            document_protection_tracked_changes_count
        ),
        document_protection_forms_count=document_protection_forms_count,
        document_protection_password_material_count=(
            document_protection_password_material_count
        ),
        write_protection_count=write_protection_count,
        write_protection_recommended_count=write_protection_recommended_count,
        write_protection_password_material_count=(
            write_protection_password_material_count
        ),
        signature=_digest_records(records),
    )


def _validate_word_protection_element(
    element: ET.Element,
    expected_local_name: str,
    allowed_attribute_names: frozenset[str],
) -> dict[str, str]:
    if (
        not _is_word_element(element, expected_local_name)
        or list(element)
        or (element.text or "").strip()
    ):
        raise DocumentFormatError("Word protection state is invalid")

    attributes: dict[str, str] = {}
    for attribute, value in element.attrib.items():
        namespace, local_name = _qualified_name(attribute)
        if (
            namespace not in _WORD_NAMESPACES
            or local_name not in allowed_attribute_names
            or local_name in attributes
        ):
            raise DocumentFormatError("Word protection state is invalid")
        attributes[local_name] = value

    if expected_local_name == "documentProtection":
        edit_mode = attributes.get("edit")
        if edit_mode is not None and edit_mode not in _DOCUMENT_PROTECTION_EDIT_VALUES:
            raise DocumentFormatError("Word protection state is invalid")
        boolean_attributes = ("formatting", "enforcement")
    else:
        boolean_attributes = ("recommended",)
    for attribute_name in boolean_attributes:
        value = attributes.get(attribute_name)
        if value is not None and value.strip().casefold() not in (
            _WORD_PROTECTION_TRUE_VALUES | _WORD_PROTECTION_FALSE_VALUES
        ):
            raise DocumentFormatError("Word protection state is invalid")
    return attributes


def _word_protection_attribute_enabled(
    attributes: dict[str, str], attribute_name: str
) -> bool:
    value = attributes.get(attribute_name)
    return (
        value is not None and value.strip().casefold() in _WORD_PROTECTION_TRUE_VALUES
    )


def _word_document_variable_inventory(
    archive: zipfile.ZipFile,
    members: dict[str, zipfile.ZipInfo],
    relationship_maps: dict[str, dict[str, _Relationship]],
    settings_part_names: frozenset[str],
    settings_part_scopes: dict[str, frozenset[str]],
    limits: PackageLimits,
) -> tuple[WordDocumentVariableInventory, tuple[_StoredDocumentVariable, ...]]:
    """Inventory persisted Word document variables without exposing contents.

    ``w:docVars`` is a direct child of a Word Settings part and stores arbitrary
    name/value pairs in ``w:docVar`` leaves. Names and values may hold template
    or automation state, so every recognized container is privately
    fingerprinted while only aggregate counts are emitted.
    """

    records: list[tuple[str, str, str]] = []
    document_variable_container_count = 0
    document_variable_count = 0
    empty_document_variable_value_count = 0
    stored_variables: list[_StoredDocumentVariable] = []

    for settings_part in sorted(settings_part_names):
        root = _read_xml(archive, members[settings_part], limits)
        if not _is_word_element(root, "settings"):
            raise DocumentFormatError("document settings are invalid")
        containers = [
            element for element in root if _is_word_element(element, "docVars")
        ]
        if len(containers) > 1:
            raise DocumentFormatError("Word document variables are invalid")

        relationships = relationship_maps.get(settings_part, {})
        for container in containers:
            variable_count, empty_value_count, variable_names = (
                _validate_word_document_variables_element(container)
            )
            document_variable_container_count += 1
            document_variable_count += variable_count
            empty_document_variable_value_count += empty_value_count
            stored_variables.extend(
                _StoredDocumentVariable(
                    document_scopes=settings_part_scopes[settings_part],
                    name=name,
                )
                for name in variable_names
            )
            records.append(
                (
                    "word_document_variables",
                    settings_part,
                    _fingerprint_element(container, relationships),
                )
            )

    return (
        WordDocumentVariableInventory(
            document_variable_container_count=document_variable_container_count,
            document_variable_count=document_variable_count,
            empty_document_variable_value_count=empty_document_variable_value_count,
            signature=_digest_records(records),
        ),
        tuple(stored_variables),
    )


def _word_document_variable_field_inventory(
    references: tuple[_DocumentVariableFieldReference, ...],
    stored_variables: tuple[_StoredDocumentVariable, ...],
) -> WordDocumentVariableFieldInventory:
    """Inventory stored ``DOCVARIABLE`` fields without disclosing arguments.

    The only association attempted is an exact, conservatively parsed literal
    name against a validated ``w:docVar`` in the same main/glossary package
    document scope. It is evidence about stored package state, not a Word
    field evaluation or a statement about an attached template.
    """

    stored_names_by_scope: dict[str, set[str]] = {}
    for variable in stored_variables:
        for document_scope in variable.document_scopes:
            stored_names_by_scope.setdefault(document_scope, set()).add(variable.name)

    literal_reference_count = 0
    matching_stored_variable_count = 0
    not_matching_stored_variable_count = 0
    records: list[tuple[str, ...]] = []
    for reference in references:
        if reference.literal_name is None:
            classification = "nonliteral"
        else:
            literal_reference_count += 1
            if reference.literal_name in stored_names_by_scope.get(
                reference.document_scope, set()
            ):
                matching_stored_variable_count += 1
                classification = "literal_matching_stored_variable"
            else:
                not_matching_stored_variable_count += 1
                classification = "literal_not_matching_stored_variable"
        records.append(
            (
                "word_document_variable_field",
                reference.story_part,
                reference.instruction_signature,
                classification,
            )
        )

    return WordDocumentVariableFieldInventory(
        document_variable_field_reference_count=len(references),
        document_variable_field_story_count=len(
            {reference.story_part for reference in references}
        ),
        literal_document_variable_field_reference_count=literal_reference_count,
        nonliteral_document_variable_field_reference_count=(
            len(references) - literal_reference_count
        ),
        literal_document_variable_field_reference_matching_stored_variable_count=(
            matching_stored_variable_count
        ),
        literal_document_variable_field_reference_not_matching_stored_variable_count=(
            not_matching_stored_variable_count
        ),
        signature=_digest_records(records),
    )


def _word_hyperlink_field_inventory(
    references: tuple[_HyperlinkFieldReference, ...],
) -> WordHyperlinkFieldInventory:
    """Inventory stored Hyperlink field codes without disclosing targets.

    The classification is lexical evidence about a complete stored field code,
    not a claim that Word will render a link, that its destination is reachable,
    or that a literal destination is external. The raw instruction remains
    solely in a private digest.
    """

    counts = {
        "literal_destination": 0,
        "literal_internal_location_only": 0,
        "dynamic_or_unparseable": 0,
    }
    records: list[tuple[str, ...]] = []
    for reference in references:
        counts[reference.classification] += 1
        records.append(
            (
                "word_hyperlink_field",
                reference.story_part,
                reference.instruction_signature,
                reference.classification,
            )
        )

    return WordHyperlinkFieldInventory(
        hyperlink_field_reference_count=len(references),
        hyperlink_field_story_count=len(
            {reference.story_part for reference in references}
        ),
        literal_destination_hyperlink_field_count=counts["literal_destination"],
        literal_internal_location_only_hyperlink_field_count=counts[
            "literal_internal_location_only"
        ],
        dynamic_or_unparseable_hyperlink_field_count=counts["dynamic_or_unparseable"],
        signature=_digest_records(records),
    )


def _hyperlink_markup_references(
    root: ET.Element,
    story_part: str,
    relationships: dict[str, _Relationship],
) -> tuple[_HyperlinkMarkupReference, ...]:
    """Retain direct ``w:hyperlink`` semantics without exposing link material.

    A relationship ID takes precedence over ``w:anchor`` in WordprocessingML.
    The standard permits a hyperlink relationship target to be either internal
    or external, so neither relationship-backed class is inferred from target
    text. Attributes such as anchors, locations, tooltips, target frames, and
    history remain in the element's private fingerprint only.
    """

    references: list[_HyperlinkMarkupReference] = []
    for element in root.iter():
        if not _is_word_element(element, "hyperlink"):
            continue
        relationship_id = _relationship_id_value(element)
        anchor_present = _word_attribute_value(element, "anchor") is not None
        if relationship_id is None:
            classification = (
                "anchor_only" if anchor_present else "default_document_start"
            )
            relationship_backed_anchor_attribute = False
        else:
            classification = _hyperlink_markup_relationship_classification(
                relationships.get(relationship_id)
            )
            relationship_backed_anchor_attribute = anchor_present
        references.append(
            _HyperlinkMarkupReference(
                story_part=story_part,
                markup_signature=_fingerprint_element(element, relationships),
                classification=classification,
                relationship_backed_anchor_attribute=(
                    relationship_backed_anchor_attribute
                ),
            )
        )
    return tuple(references)


def _hyperlink_markup_relationship_classification(
    relationship: _Relationship | None,
) -> str:
    """Classify only a resolved standard relationship's stored mode."""

    if (
        relationship is None
        or relationship.relationship_type.casefold()
        not in _HYPERLINK_RELATIONSHIP_TYPES
    ):
        return "unsupported_relationship"
    target_mode = relationship.target_mode.casefold()
    if target_mode == "external":
        return "external_relationship"
    if target_mode == "internal":
        return "internal_relationship"
    return "unsupported_relationship"


def _drawing_linked_picture_relationship_classification(
    relationship: _Relationship | None,
) -> str:
    """Classify the stored image relationship behind an ``a:blip/@r:link``."""

    if (
        relationship is None
        or relationship.relationship_type.casefold() not in _IMAGE_RELATIONSHIP_TYPES
    ):
        return "unsupported_relationship"
    target_mode = relationship.target_mode.casefold()
    if target_mode == "external":
        return "external_image_relationship"
    if target_mode == "internal":
        return "internal_image_relationship"
    return "unsupported_relationship"


def _word_hyperlink_markup_inventory(
    references: tuple[_HyperlinkMarkupReference, ...],
) -> WordHyperlinkMarkupInventory:
    """Aggregate direct hyperlink markup without emitting targets or labels."""

    counts = {
        "external_relationship": 0,
        "internal_relationship": 0,
        "unsupported_relationship": 0,
        "anchor_only": 0,
        "default_document_start": 0,
    }
    records: list[tuple[str, ...]] = []
    relationship_backed_anchor_attribute_count = 0
    for reference in references:
        counts[reference.classification] += 1
        relationship_backed_anchor_attribute_count += int(
            reference.relationship_backed_anchor_attribute
        )
        records.append(
            (
                "word_hyperlink_markup",
                reference.story_part,
                reference.markup_signature,
                reference.classification,
                str(reference.relationship_backed_anchor_attribute),
            )
        )

    relationship_backed_count = sum(
        counts[classification]
        for classification in (
            "external_relationship",
            "internal_relationship",
            "unsupported_relationship",
        )
    )
    return WordHyperlinkMarkupInventory(
        hyperlink_element_count=len(references),
        hyperlink_story_count=len({reference.story_part for reference in references}),
        relationship_backed_hyperlink_count=relationship_backed_count,
        external_relationship_hyperlink_count=counts["external_relationship"],
        internal_relationship_hyperlink_count=counts["internal_relationship"],
        unsupported_relationship_hyperlink_count=counts["unsupported_relationship"],
        anchor_only_hyperlink_count=counts["anchor_only"],
        default_document_start_hyperlink_count=counts["default_document_start"],
        relationship_backed_anchor_attribute_count=(
            relationship_backed_anchor_attribute_count
        ),
        signature=_digest_records(records),
    )


def _drawing_hyperlink_references(
    root: ET.Element,
    story_part: str,
    relationships: dict[str, _Relationship],
) -> tuple[_DrawingHyperlinkReference, ...]:
    """Retain direct DrawingML hyperlink actions without exposing link material.

    This deliberately inventories each stored marker in supported Word stories,
    including duplicate markers that point at the same relationship and markers
    inside markup-compatibility branches. It does not select a rendering branch,
    deduplicate visual objects, retrieve a target, or claim an action executes.
    ``r:id`` is schema-required for these elements; an omitted attribute remains
    reviewable malformed stored evidence instead of being resolved as a target.
    """

    references: list[_DrawingHyperlinkReference] = []
    for element in root.iter():
        namespace, local_name = _qualified_name(element.tag)
        if namespace not in _DRAWING_NAMESPACES:
            continue
        kind = _DRAWING_HYPERLINK_REFERENCE_KINDS.get(local_name)
        if kind is None:
            continue
        relationship_id = _relationship_id_value(element)
        classification = (
            "missing_relationship_id"
            if relationship_id is None
            else _hyperlink_markup_relationship_classification(
                relationships.get(relationship_id)
            )
        )
        references.append(
            _DrawingHyperlinkReference(
                story_part=story_part,
                markup_signature=_fingerprint_element(element, relationships),
                kind=kind,
                classification=classification,
                action_attribute_present=(
                    _unqualified_attribute_value(element, "action") is not None
                ),
                invalid_url_attribute_present=(
                    _unqualified_attribute_value(element, "invalidUrl") is not None
                ),
            )
        )
    return tuple(references)


def _drawing_visibility_references(
    root: ET.Element,
    story_part: str,
) -> tuple[_DrawingVisibilityReference, ...]:
    """Retain direct supported nonvisual hidden declarations in Word stories.

    This reads every stored supported marker, including duplicates and markers
    in markup-compatibility branches. It does not infer effective visibility,
    select a markup-compatibility branch, resolve object identity, inspect a
    drawing's contents, or predict any application's rendering behavior.
    """

    references: list[_DrawingVisibilityReference] = []
    for element in root.iter():
        namespace, local_name = _qualified_name(element.tag)
        if (namespace, local_name) not in _WORD_DRAWING_VISIBILITY_ELEMENTS:
            continue
        value = _unqualified_attribute_value(element, "hidden")
        if value is None:
            continue
        references.append(
            _DrawingVisibilityReference(
                story_part=story_part,
                namespace=namespace,
                local_name=local_name,
                state=_drawing_visibility_state(value),
            )
        )
    return tuple(references)


def _drawing_visibility_state(value: str) -> str:
    """Canonicalize valid XML Boolean spellings without losing invalid values."""

    normalized = value.strip()
    if normalized in _XML_BOOLEAN_TRUE_VALUES:
        return "hidden"
    if normalized in _XML_BOOLEAN_FALSE_VALUES:
        return "explicitly_shown"
    return f"invalid:{value}"


def _word_drawing_visibility_inventory(
    references: tuple[_DrawingVisibilityReference, ...],
) -> WordDrawingVisibilityInventory:
    """Aggregate direct nonvisual visibility declarations without object data."""

    state_counts = {
        "hidden": 0,
        "explicitly_shown": 0,
        "invalid": 0,
    }
    records: list[tuple[str, ...]] = []
    for reference in references:
        state = reference.state
        if state.startswith("invalid:"):
            state_counts["invalid"] += 1
        else:
            state_counts[state] += 1
        records.append(
            (
                "word_drawing_visibility",
                reference.story_part,
                reference.namespace,
                reference.local_name,
                state,
            )
        )

    return WordDrawingVisibilityInventory(
        visibility_declaration_count=len(references),
        visibility_declaration_story_count=len(
            {reference.story_part for reference in references}
        ),
        hidden_drawing_object_count=state_counts["hidden"],
        explicitly_shown_drawing_object_count=state_counts["explicitly_shown"],
        invalid_hidden_attribute_count=state_counts["invalid"],
        signature=_digest_records(records),
    )


def _word_drawing_hyperlink_inventory(
    references: tuple[_DrawingHyperlinkReference, ...],
) -> WordDrawingHyperlinkInventory:
    """Aggregate DrawingML hyperlink-action markers without emitting targets."""

    kind_counts = {"click": 0, "hover": 0, "mouse_over": 0}
    classification_counts = {
        "external_relationship": 0,
        "internal_relationship": 0,
        "unsupported_relationship": 0,
        "missing_relationship_id": 0,
    }
    records: list[tuple[str, ...]] = []
    action_attribute_count = 0
    invalid_url_attribute_count = 0
    for reference in references:
        kind_counts[reference.kind] += 1
        classification_counts[reference.classification] += 1
        action_attribute_count += int(reference.action_attribute_present)
        invalid_url_attribute_count += int(reference.invalid_url_attribute_present)
        records.append(
            (
                "word_drawing_hyperlink",
                reference.story_part,
                reference.markup_signature,
                reference.kind,
                reference.classification,
                str(reference.action_attribute_present),
                str(reference.invalid_url_attribute_present),
            )
        )

    return WordDrawingHyperlinkInventory(
        drawing_hyperlink_reference_count=len(references),
        drawing_hyperlink_story_count=len(
            {reference.story_part for reference in references}
        ),
        click_drawing_hyperlink_count=kind_counts["click"],
        hover_drawing_hyperlink_count=kind_counts["hover"],
        mouse_over_drawing_hyperlink_count=kind_counts["mouse_over"],
        external_relationship_drawing_hyperlink_count=classification_counts[
            "external_relationship"
        ],
        internal_relationship_drawing_hyperlink_count=classification_counts[
            "internal_relationship"
        ],
        unsupported_relationship_drawing_hyperlink_count=classification_counts[
            "unsupported_relationship"
        ],
        missing_relationship_id_drawing_hyperlink_count=classification_counts[
            "missing_relationship_id"
        ],
        action_attribute_drawing_hyperlink_count=action_attribute_count,
        invalid_url_attribute_drawing_hyperlink_count=invalid_url_attribute_count,
        signature=_digest_records(records),
    )


def _drawing_linked_picture_references(
    root: ET.Element,
    story_part: str,
    relationships: dict[str, _Relationship],
) -> tuple[_DrawingLinkedPictureReference, ...]:
    """Retain direct ``a:blip/@r:link`` markers without exposing image data.

    ``r:link`` is the DrawingML linked-picture reference. This inventories every
    stored marker in supported Word stories, including duplicate markers that
    share a relationship and markers in markup-compatibility branches. It does
    not select a rendering branch, deduplicate visual objects, retrieve an
    image, or claim a Word client will load one. An ``r:link`` attribute is
    always retained as stored evidence; its resolved relationship is classified
    only by standard image type and stored target mode.
    """

    references: list[_DrawingLinkedPictureReference] = []
    for element in root.iter():
        namespace, local_name = _qualified_name(element.tag)
        if namespace not in _DRAWING_NAMESPACES or local_name != "blip":
            continue
        relationship_id = _relationship_attribute_value(element, "link")
        if relationship_id is None:
            continue
        references.append(
            _DrawingLinkedPictureReference(
                story_part=story_part,
                markup_signature=_fingerprint_element(element, relationships),
                classification=_drawing_linked_picture_relationship_classification(
                    relationships.get(relationship_id)
                ),
            )
        )
    return tuple(references)


def _word_drawing_linked_picture_inventory(
    references: tuple[_DrawingLinkedPictureReference, ...],
) -> WordDrawingLinkedPictureInventory:
    """Aggregate linked-picture markers without emitting their target material."""

    classification_counts = {
        "external_image_relationship": 0,
        "internal_image_relationship": 0,
        "unsupported_relationship": 0,
    }
    records: list[tuple[str, ...]] = []
    for reference in references:
        classification_counts[reference.classification] += 1
        records.append(
            (
                "word_drawing_linked_picture",
                reference.story_part,
                reference.markup_signature,
                reference.classification,
            )
        )

    return WordDrawingLinkedPictureInventory(
        drawing_linked_picture_reference_count=len(references),
        drawing_linked_picture_story_count=len(
            {reference.story_part for reference in references}
        ),
        external_image_relationship_drawing_linked_picture_count=(
            classification_counts["external_image_relationship"]
        ),
        internal_image_relationship_drawing_linked_picture_count=(
            classification_counts["internal_image_relationship"]
        ),
        unsupported_relationship_drawing_linked_picture_count=(
            classification_counts["unsupported_relationship"]
        ),
        signature=_digest_records(records),
    )


def _vml_hyperlink_references(
    root: ET.Element,
    story_part: str,
    relationships: dict[str, _Relationship],
) -> tuple[_VmlHyperlinkReference, ...]:
    """Retain direct legacy VML shape ``href`` markup without link material.

    This is intentionally a narrow stored-markup inventory.  It examines only
    the VML shape, group, and shape-template element kinds that define an
    ``href`` surface, and counts a marker whenever its direct unqualified
    attribute is present, including an empty value.  It neither computes an
    effective inherited link nor selects a rendering branch, resolves a URL,
    follows a target, or claims that a click action will execute.

    Nested VML data and formatting elements are excluded even if a producer
    happens to add an ``href``-named attribute there: this inventory is about
    the documented direct shape-link surface, not arbitrary legacy markup.
    """

    references: list[_VmlHyperlinkReference] = []
    for element in root.iter():
        namespace, local_name = _qualified_name(element.tag)
        if namespace != _VML_NAMESPACE:
            continue
        kind = _VML_HYPERLINK_REFERENCE_KINDS.get(local_name)
        if kind is None or _unqualified_attribute_value(element, "href") is None:
            continue
        references.append(
            _VmlHyperlinkReference(
                story_part=story_part,
                markup_signature=_fingerprint_element(element, relationships),
                kind=kind,
                target_attribute_present=(
                    _unqualified_attribute_value(element, "target") is not None
                ),
            )
        )
    return tuple(references)


def _word_vml_hyperlink_inventory(
    references: tuple[_VmlHyperlinkReference, ...],
) -> WordVmlHyperlinkInventory:
    """Aggregate direct VML shape-link markers without emitting raw markup."""

    kind_counts = {"concrete_shape": 0, "group": 0, "shape_type": 0}
    records: list[tuple[str, ...]] = []
    target_attribute_count = 0
    for reference in references:
        kind_counts[reference.kind] += 1
        target_attribute_count += int(reference.target_attribute_present)
        records.append(
            (
                "word_vml_hyperlink",
                reference.story_part,
                reference.markup_signature,
                reference.kind,
                str(reference.target_attribute_present),
            )
        )

    return WordVmlHyperlinkInventory(
        vml_hyperlink_element_count=len(references),
        vml_hyperlink_story_count=len(
            {reference.story_part for reference in references}
        ),
        concrete_shape_vml_hyperlink_count=kind_counts["concrete_shape"],
        group_vml_hyperlink_count=kind_counts["group"],
        shape_type_vml_hyperlink_count=kind_counts["shape_type"],
        target_attribute_vml_hyperlink_count=target_attribute_count,
        signature=_digest_records(records),
    )


def _vml_external_image_relationship_classification(
    relationship: _Relationship | None,
) -> str | None:
    """Classify only an externally stored relation behind VML image data.

    VML ``imagedata/@r:id`` also represents ordinary embedded images. Those
    internal relationships are intentionally outside this external-image review
    boundary, as are absent or malformed relationship records. An external
    relation with a non-image type stays reviewable as unsupported stored
    evidence rather than being treated as a linked image.
    """

    if relationship is None or relationship.target_mode.casefold() != "external":
        return None
    if relationship.relationship_type.casefold() in _IMAGE_RELATIONSHIP_TYPES:
        return "external_image_relationship"
    return "unsupported_relationship"


def _vml_external_image_marker_signature(relationship: _Relationship) -> str:
    """Fingerprint only the reviewed relationship surface of a VML marker.

    In particular, raw VML ``src`` and the other legacy relationship attributes
    are not folded into this inventory's signature. They remain separate
    surfaces rather than making this narrow ``r:id`` review boundary wider.
    """

    return _digest_records(
        [("word_vml_external_image_marker", *relationship.canonical_value())]
    )


def _vml_external_image_references(
    root: ET.Element,
    story_part: str,
    relationships: dict[str, _Relationship],
) -> tuple[_VmlExternalImageReference, ...]:
    """Retain direct VML ``imagedata/@r:id`` external-image markers.

    This deliberately inventories each stored ``r:id`` marker only when its
    resolved relationship has stored ``TargetMode=External``. It includes
    duplicate markers that share a relationship and markers in
    markup-compatibility branches, but does not select a rendering branch,
    deduplicate visual objects, retrieve an image, or claim a Word client will
    load one. Internal image relationships, raw VML ``src``, ``r:pict``,
    ``r:href``, and ``o:relid`` are not part of this boundary.
    """

    references: list[_VmlExternalImageReference] = []
    for element in root.iter():
        namespace, local_name = _qualified_name(element.tag)
        if namespace != _VML_NAMESPACE or local_name != "imagedata":
            continue
        relationship_id = _relationship_id_value(element)
        if relationship_id is None:
            continue
        relationship = relationships.get(relationship_id)
        classification = _vml_external_image_relationship_classification(relationship)
        if classification is None or relationship is None:
            continue
        references.append(
            _VmlExternalImageReference(
                story_part=story_part,
                marker_signature=_vml_external_image_marker_signature(relationship),
                classification=classification,
            )
        )
    return tuple(references)


def _word_vml_external_image_inventory(
    references: tuple[_VmlExternalImageReference, ...],
) -> WordVmlExternalImageInventory:
    """Aggregate VML external-image markers without emitting image data."""

    classification_counts = {
        "external_image_relationship": 0,
        "unsupported_relationship": 0,
    }
    records: list[tuple[str, ...]] = []
    for reference in references:
        classification_counts[reference.classification] += 1
        records.append(
            (
                "word_vml_external_image",
                reference.story_part,
                reference.marker_signature,
                reference.classification,
            )
        )

    return WordVmlExternalImageInventory(
        vml_external_image_reference_count=len(references),
        vml_external_image_story_count=len(
            {reference.story_part for reference in references}
        ),
        external_image_relationship_vml_external_image_count=(
            classification_counts["external_image_relationship"]
        ),
        unsupported_relationship_vml_external_image_count=(
            classification_counts["unsupported_relationship"]
        ),
        signature=_digest_records(records),
    )


def _vml_image_hyperlink_marker_signature(relationship: _Relationship) -> str:
    """Fingerprint only the reviewed VML image-data hyperlink relationship."""

    return _digest_records(
        [("word_vml_image_hyperlink_marker", *relationship.canonical_value())]
    )


def _vml_image_hyperlink_references(
    root: ET.Element,
    story_part: str,
    relationships: dict[str, _Relationship],
) -> tuple[_VmlImageHyperlinkReference, ...]:
    """Retain direct VML ``imagedata/@r:href`` markers without link material.

    ``r:href`` is the VML image-data relationship attribute documented for a
    hyperlink target. This inventories each stored direct marker in supported
    Word stories, including duplicate markers that share a relationship and
    markers in markup-compatibility branches. It does not select a rendering
    branch, associate a marker with a visual image, resolve a target, follow a
    link, or claim that a client will honor one. Other VML image-data surfaces
    (including ``r:id``, ``r:pict``, ``src``, and ``o:relid``) stay outside this
    narrow marker and signature boundary.
    """

    references: list[_VmlImageHyperlinkReference] = []
    for element in root.iter():
        namespace, local_name = _qualified_name(element.tag)
        if namespace != _VML_NAMESPACE or local_name != "imagedata":
            continue
        relationship_id = _relationship_attribute_value(element, "href")
        if relationship_id is None:
            continue
        relationship = relationships.get(relationship_id)
        if relationship is None:
            continue
        references.append(
            _VmlImageHyperlinkReference(
                story_part=story_part,
                marker_signature=_vml_image_hyperlink_marker_signature(relationship),
                classification=_hyperlink_markup_relationship_classification(
                    relationship
                ),
            )
        )
    return tuple(references)


def _word_vml_image_hyperlink_inventory(
    references: tuple[_VmlImageHyperlinkReference, ...],
) -> WordVmlImageHyperlinkInventory:
    """Aggregate VML image-data hyperlink markers without emitting targets."""

    classification_counts = {
        "external_relationship": 0,
        "internal_relationship": 0,
        "unsupported_relationship": 0,
    }
    records: list[tuple[str, ...]] = []
    for reference in references:
        classification_counts[reference.classification] += 1
        records.append(
            (
                "word_vml_image_hyperlink",
                reference.story_part,
                reference.marker_signature,
                reference.classification,
            )
        )

    return WordVmlImageHyperlinkInventory(
        vml_image_hyperlink_reference_count=len(references),
        vml_image_hyperlink_story_count=len(
            {reference.story_part for reference in references}
        ),
        external_relationship_vml_image_hyperlink_count=(
            classification_counts["external_relationship"]
        ),
        internal_relationship_vml_image_hyperlink_count=(
            classification_counts["internal_relationship"]
        ),
        unsupported_relationship_vml_image_hyperlink_count=(
            classification_counts["unsupported_relationship"]
        ),
        signature=_digest_records(records),
    )


def _ole_object_relationship_classification(
    relationship: _Relationship | None,
) -> str:
    """Classify a stored OLE-object relationship without resolving its target."""

    if (
        relationship is None
        or relationship.relationship_type.casefold()
        not in _OLE_OBJECT_RELATIONSHIP_TYPES
    ):
        return "unsupported_relationship"
    target_mode = relationship.target_mode.casefold()
    if target_mode == "external":
        return "external_standard_ole_object_relationship"
    if target_mode == "internal":
        return "internal_standard_ole_object_relationship"
    return "unsupported_relationship"


def _vml_linked_ole_object_references(
    root: ET.Element,
    story_part: str,
    relationships: dict[str, _Relationship],
) -> tuple[_VmlLinkedOleObjectReference, ...]:
    """Retain direct legacy VML Office ``OLEObject Type=Link`` markers.

    This is a narrow stored-markup inventory. It records each direct Office VML
    OLE-object marker whose unqualified ``Type`` is ``Link`` in a supported Word
    story, including duplicate markers and markers in markup-compatibility
    branches. The direct marker remains privately fingerprinted so source and
    update metadata never reach output while same-count rewrites stay
    review-visible. It does not select a rendering branch, associate a marker
    with a shape, resolve or retrieve a source, update or activate an OLE
    object, or claim that a client will honor one.

    An OLE relationship is optional in the legacy marker, so a direct marker
    without ``r:id`` remains stored evidence in its own public class. VML image
    data, VML shape links, WordprocessingML ``w:objectLink`` markup, and broad
    embedded-object relationship totals remain separate surfaces.
    """

    references: list[_VmlLinkedOleObjectReference] = []
    for element in root.iter():
        namespace, local_name = _qualified_name(element.tag)
        if namespace != _OFFICE_VML_NAMESPACE or local_name != "OLEObject":
            continue
        marker_type = _unqualified_attribute_value(element, "Type")
        if marker_type is None or marker_type.strip().casefold() != "link":
            continue
        relationship_id = _relationship_id_value(element)
        if relationship_id is None:
            relationship_classification = "without_relationship_id"
        else:
            relationship_classification = _ole_object_relationship_classification(
                relationships.get(relationship_id)
            )
        update_mode = _unqualified_attribute_value(element, "UpdateMode")
        update_mode_classification = (
            "automatic_update"
            if update_mode is not None and update_mode.strip().casefold() == "always"
            else "nonautomatic_or_unspecified_update"
        )
        references.append(
            _VmlLinkedOleObjectReference(
                story_part=story_part,
                markup_signature=_fingerprint_element(element, relationships),
                update_mode_classification=update_mode_classification,
                relationship_classification=relationship_classification,
            )
        )
    return tuple(references)


def _word_vml_linked_ole_object_inventory(
    references: tuple[_VmlLinkedOleObjectReference, ...],
) -> WordVmlLinkedOleObjectInventory:
    """Aggregate VML linked-OLE-object markers without emitting source data."""

    update_mode_counts = {
        "automatic_update": 0,
        "nonautomatic_or_unspecified_update": 0,
    }
    relationship_counts = {
        "external_standard_ole_object_relationship": 0,
        "internal_standard_ole_object_relationship": 0,
        "unsupported_relationship": 0,
        "without_relationship_id": 0,
    }
    records: list[tuple[str, ...]] = []
    for reference in references:
        update_mode_counts[reference.update_mode_classification] += 1
        relationship_counts[reference.relationship_classification] += 1
        records.append(
            (
                "word_vml_linked_ole_object",
                reference.story_part,
                reference.markup_signature,
                reference.update_mode_classification,
                reference.relationship_classification,
            )
        )

    return WordVmlLinkedOleObjectInventory(
        vml_linked_ole_object_count=len(references),
        vml_linked_ole_object_story_count=len(
            {reference.story_part for reference in references}
        ),
        automatic_update_vml_linked_ole_object_count=(
            update_mode_counts["automatic_update"]
        ),
        nonautomatic_or_unspecified_update_vml_linked_ole_object_count=(
            update_mode_counts["nonautomatic_or_unspecified_update"]
        ),
        external_standard_ole_object_relationship_vml_linked_ole_object_count=(
            relationship_counts["external_standard_ole_object_relationship"]
        ),
        internal_standard_ole_object_relationship_vml_linked_ole_object_count=(
            relationship_counts["internal_standard_ole_object_relationship"]
        ),
        unsupported_relationship_vml_linked_ole_object_count=(
            relationship_counts["unsupported_relationship"]
        ),
        without_relationship_id_vml_linked_ole_object_count=(
            relationship_counts["without_relationship_id"]
        ),
        signature=_digest_records(records),
    )


def _vml_embedded_ole_object_references(
    root: ET.Element,
    story_part: str,
    relationships: dict[str, _Relationship],
) -> tuple[_VmlEmbeddedOleObjectReference, ...]:
    """Retain direct legacy VML Office ``OLEObject Type=Embed`` markers.

    This narrow stored-markup inventory records each Office VML OLE-object
    marker whose unqualified ``Type`` is ``Embed`` in a supported Word story,
    including duplicate markers and markers in markup-compatibility branches.
    The Office VML contract permits the element in several parent forms, so the
    boundary follows the direct marker rather than assigning it a rendering
    position. The full marker remains privately fingerprinted so program,
    shape, object, field-code, and update metadata never reach output while
    same-count rewrites stay review-visible. It does not select a rendering
    branch, associate a marker with a shape, inspect an object payload, load or
    activate an OLE object, or claim that a client will honor one.

    An OLE relationship is optional in the legacy marker, so a direct marker
    without ``r:id`` remains stored evidence in its own public class. The
    standard permits an OLE-object relationship with an internal or external
    target mode; those stored modes remain separate. ``UpdateMode`` applies to
    the Link type, so this embedded-object inventory retains it only inside the
    private marker signature rather than presenting it as an embedded-object
    behavior. VML linked-OLE markup, WordprocessingML ``w:objectEmbed`` and
    ``w:objectLink`` markup, VML image data and shape links, fields, and broad
    embedded-object relationship/payload totals remain separate surfaces.
    """

    references: list[_VmlEmbeddedOleObjectReference] = []
    for element in root.iter():
        namespace, local_name = _qualified_name(element.tag)
        if namespace != _OFFICE_VML_NAMESPACE or local_name != "OLEObject":
            continue
        marker_type = _unqualified_attribute_value(element, "Type")
        if marker_type is None or marker_type.strip().casefold() != "embed":
            continue
        relationship_id = _relationship_id_value(element)
        relationship_classification = (
            "without_relationship_id"
            if relationship_id is None
            else _ole_object_relationship_classification(
                relationships.get(relationship_id)
            )
        )
        references.append(
            _VmlEmbeddedOleObjectReference(
                story_part=story_part,
                markup_signature=_fingerprint_element(element, relationships),
                relationship_classification=relationship_classification,
            )
        )
    return tuple(references)


def _word_vml_embedded_ole_object_inventory(
    references: tuple[_VmlEmbeddedOleObjectReference, ...],
) -> WordVmlEmbeddedOleObjectInventory:
    """Aggregate VML embedded-OLE-object markers without emitting metadata."""

    relationship_counts = {
        "external_standard_ole_object_relationship": 0,
        "internal_standard_ole_object_relationship": 0,
        "unsupported_relationship": 0,
        "without_relationship_id": 0,
    }
    records: list[tuple[str, ...]] = []
    for reference in references:
        relationship_counts[reference.relationship_classification] += 1
        records.append(
            (
                "word_vml_embedded_ole_object",
                reference.story_part,
                reference.markup_signature,
                reference.relationship_classification,
            )
        )

    return WordVmlEmbeddedOleObjectInventory(
        vml_embedded_ole_object_count=len(references),
        vml_embedded_ole_object_story_count=len(
            {reference.story_part for reference in references}
        ),
        external_standard_ole_object_relationship_vml_embedded_ole_object_count=(
            relationship_counts["external_standard_ole_object_relationship"]
        ),
        internal_standard_ole_object_relationship_vml_embedded_ole_object_count=(
            relationship_counts["internal_standard_ole_object_relationship"]
        ),
        unsupported_relationship_vml_embedded_ole_object_count=(
            relationship_counts["unsupported_relationship"]
        ),
        without_relationship_id_vml_embedded_ole_object_count=(
            relationship_counts["without_relationship_id"]
        ),
        signature=_digest_records(records),
    )


def _word_object_link_references(
    root: ET.Element,
    story_part: str,
    relationships: dict[str, _Relationship],
) -> tuple[_WordObjectLinkReference, ...]:
    """Retain direct WordprocessingML linked-object-property markers.

    This is a narrow stored-markup inventory. It records every direct
    WordprocessingML objectLink child of a Word object in a supported Word
    story, including duplicate markers and markers in markup-compatibility
    branches. The direct marker remains privately fingerprinted so program,
    field, locking, and relationship metadata never reach output while
    same-count rewrites stay review-visible. It does not select a rendering
    branch, associate a marker with a visual object, resolve or retrieve a
    source, update or activate an OLE object, or claim that a client will honor
    one.

    The standard requires both r:id and w:updateMode, but malformed stored
    direct markers remain reviewable evidence: a missing r:id has its own
    public class, while an absent or nonstandard update token is counted as
    unsupported-or-missing. Only the schema tokens always and onCall receive
    the two named stored-mode classes. WordprocessingML objectEmbed markup,
    legacy Office VML OLEObject markup, VML image data and shape links, fields,
    and broad embedded-object relationship totals remain separate surfaces.
    """

    references: list[_WordObjectLinkReference] = []
    for parent in root.iter():
        if not _is_word_element(parent, "object"):
            continue
        for element in parent:
            if not _is_word_element(element, "objectLink"):
                continue
            relationship_id = _relationship_id_value(element)
            if relationship_id is None:
                relationship_classification = "without_relationship_id"
            else:
                relationship_classification = _ole_object_relationship_classification(
                    relationships.get(relationship_id)
                )
            update_mode = _word_attribute_value(element, "updateMode")
            if update_mode == "always":
                update_mode_classification = "automatic_update"
            elif update_mode == "onCall":
                update_mode_classification = "on_call_update"
            else:
                update_mode_classification = "unsupported_or_missing_update_mode"
            references.append(
                _WordObjectLinkReference(
                    story_part=story_part,
                    markup_signature=_fingerprint_element(element, relationships),
                    update_mode_classification=update_mode_classification,
                    relationship_classification=relationship_classification,
                )
            )
    return tuple(references)


def _word_object_link_inventory(
    references: tuple[_WordObjectLinkReference, ...],
) -> WordObjectLinkInventory:
    """Aggregate WordprocessingML objectLink markers without emitting metadata."""

    update_mode_counts = {
        "automatic_update": 0,
        "on_call_update": 0,
        "unsupported_or_missing_update_mode": 0,
    }
    relationship_counts = {
        "external_standard_ole_object_relationship": 0,
        "internal_standard_ole_object_relationship": 0,
        "unsupported_relationship": 0,
        "without_relationship_id": 0,
    }
    records: list[tuple[str, ...]] = []
    for reference in references:
        update_mode_counts[reference.update_mode_classification] += 1
        relationship_counts[reference.relationship_classification] += 1
        records.append(
            (
                "word_object_link",
                reference.story_part,
                reference.markup_signature,
                reference.update_mode_classification,
                reference.relationship_classification,
            )
        )

    return WordObjectLinkInventory(
        object_link_count=len(references),
        object_link_story_count=len({reference.story_part for reference in references}),
        automatic_update_object_link_count=update_mode_counts["automatic_update"],
        on_call_update_object_link_count=update_mode_counts["on_call_update"],
        unsupported_or_missing_update_mode_object_link_count=(
            update_mode_counts["unsupported_or_missing_update_mode"]
        ),
        external_standard_ole_object_relationship_object_link_count=(
            relationship_counts["external_standard_ole_object_relationship"]
        ),
        internal_standard_ole_object_relationship_object_link_count=(
            relationship_counts["internal_standard_ole_object_relationship"]
        ),
        unsupported_relationship_object_link_count=(
            relationship_counts["unsupported_relationship"]
        ),
        without_relationship_id_object_link_count=(
            relationship_counts["without_relationship_id"]
        ),
        signature=_digest_records(records),
    )


def _control_relationship_classification(
    relationship: _Relationship | None,
) -> str:
    """Classify a Word embedded-control property relationship as stored.

    A control-persistence relationship must be internal under the OOXML part
    model. The external class therefore records a standard control relationship
    type with a nonconforming stored target mode rather than treating it as a
    usable control payload.
    """

    if (
        relationship is None
        or relationship.relationship_type.casefold() not in _CONTROL_RELATIONSHIP_TYPES
    ):
        return "unsupported_relationship"
    target_mode = relationship.target_mode.casefold()
    if target_mode == "internal":
        return "internal_standard_control_relationship"
    if target_mode == "external":
        return "external_standard_control_relationship"
    return "unsupported_relationship"


def _word_embedded_control_references(
    root: ET.Element,
    story_part: str,
    relationships: dict[str, _Relationship],
) -> tuple[_WordEmbeddedControlReference, ...]:
    """Retain direct WordprocessingML embedded-control anchors.

    This narrow stored-markup inventory records each direct ``w:control`` child
    of ``w:object`` or ``w:pict`` in a supported Word story, including duplicate
    markers and markers in markup-compatibility branches. The standard assigns
    those two parent positions distinct embedded-control representations. The
    full direct marker remains privately fingerprinted so control names, shape
    identifiers, and relationship metadata never reach output while same-count
    rewrites stay review-visible. It does not select a rendering branch,
    associate a marker with a visual shape, instantiate or load a control,
    inspect a persistence payload, or claim that a client will honor one.

    The ``r:id`` relationship is optional: an omitted ID indicates that the
    control has no property bag when instantiated, so a direct marker without
    it remains stored evidence in its own public class. A resolved standard
    control relationship is classified by its stored target mode. The standard
    requires that persistence-part relationship to be internal; an external
    mode remains separately reviewable nonconforming evidence. Generic embedded
    control relationship/payload totals, ActiveX-binary relationships, VML
    image data and shapes, ``w:objectLink`` and ``w:objectEmbed`` markup, fields,
    and arbitrary ``w:control`` elements outside these direct parent positions
    stay outside this boundary.
    """

    references: list[_WordEmbeddedControlReference] = []
    for parent in root.iter():
        if _is_word_element(parent, "object"):
            parent_kind = "object"
        elif _is_word_element(parent, "pict"):
            parent_kind = "pict"
        else:
            continue
        for element in parent:
            if not _is_word_element(element, "control"):
                continue
            relationship_id = _relationship_id_value(element)
            relationship_classification = (
                "without_relationship_id"
                if relationship_id is None
                else _control_relationship_classification(
                    relationships.get(relationship_id)
                )
            )
            references.append(
                _WordEmbeddedControlReference(
                    story_part=story_part,
                    markup_signature=_fingerprint_element(element, relationships),
                    parent_kind=parent_kind,
                    relationship_classification=relationship_classification,
                )
            )
    return tuple(references)


def _word_embedded_control_inventory(
    references: tuple[_WordEmbeddedControlReference, ...],
) -> WordEmbeddedControlInventory:
    """Aggregate Word embedded-control anchors without emitting metadata."""

    parent_counts = {"object": 0, "pict": 0}
    relationship_counts = {
        "internal_standard_control_relationship": 0,
        "external_standard_control_relationship": 0,
        "unsupported_relationship": 0,
        "without_relationship_id": 0,
    }
    records: list[tuple[str, ...]] = []
    for reference in references:
        parent_counts[reference.parent_kind] += 1
        relationship_counts[reference.relationship_classification] += 1
        records.append(
            (
                "word_embedded_control",
                reference.story_part,
                reference.markup_signature,
                reference.parent_kind,
                reference.relationship_classification,
            )
        )

    return WordEmbeddedControlInventory(
        embedded_control_count=len(references),
        embedded_control_story_count=len(
            {reference.story_part for reference in references}
        ),
        object_parent_embedded_control_count=parent_counts["object"],
        pict_parent_embedded_control_count=parent_counts["pict"],
        internal_standard_control_relationship_embedded_control_count=(
            relationship_counts["internal_standard_control_relationship"]
        ),
        external_standard_control_relationship_embedded_control_count=(
            relationship_counts["external_standard_control_relationship"]
        ),
        unsupported_relationship_embedded_control_count=(
            relationship_counts["unsupported_relationship"]
        ),
        without_relationship_id_embedded_control_count=(
            relationship_counts["without_relationship_id"]
        ),
        signature=_digest_records(records),
    )


def _validate_word_document_variables_element(
    element: ET.Element,
) -> tuple[int, int, tuple[str, ...]]:
    """Validate Word's direct document-variable container and leaves."""

    if (
        not _is_word_element(element, "docVars")
        or element.attrib
        or (element.text or "").strip()
    ):
        raise DocumentFormatError("Word document variables are invalid")

    variable_count = 0
    empty_value_count = 0
    variable_names: list[str] = []
    for child in element:
        attributes = _validate_word_document_variable_element(child)
        if (child.tail or "").strip():
            raise DocumentFormatError("Word document variables are invalid")
        variable_count += 1
        variable_names.append(attributes["name"])
        if not attributes["val"]:
            empty_value_count += 1
    return variable_count, empty_value_count, tuple(variable_names)


def _validate_word_document_variable_element(element: ET.Element) -> dict[str, str]:
    if (
        not _is_word_element(element, "docVar")
        or list(element)
        or (element.text or "").strip()
    ):
        raise DocumentFormatError("Word document variables are invalid")

    attributes: dict[str, str] = {}
    for attribute, value in element.attrib.items():
        namespace, local_name = _qualified_name(attribute)
        if (
            namespace not in _WORD_NAMESPACES
            or local_name not in _WORD_DOCUMENT_VARIABLE_ATTRIBUTE_NAMES
            or local_name in attributes
        ):
            raise DocumentFormatError("Word document variables are invalid")
        attributes[local_name] = value

    if set(attributes) != _WORD_DOCUMENT_VARIABLE_ATTRIBUTE_NAMES:
        raise DocumentFormatError("Word document variables are invalid")
    if (
        not 1
        <= _utf16_code_unit_length(attributes["name"])
        <= (_WORD_DOCUMENT_VARIABLE_NAME_MAX_UTF16_CODE_UNITS)
    ):
        raise DocumentFormatError("Word document variables are invalid")
    if _utf16_code_unit_length(attributes["val"]) > (
        _WORD_DOCUMENT_VARIABLE_VALUE_MAX_UTF16_CODE_UNITS
    ):
        raise DocumentFormatError("Word document variables are invalid")
    return attributes


def _utf16_code_unit_length(value: str) -> int:
    """Match the Open XML SDK StringValue length convention."""

    return len(value.encode("utf-16-le")) // 2


def _word_permission_range_inventory(
    archive: zipfile.ZipFile,
    members: dict[str, zipfile.ZipInfo],
    content_types: dict[str, str],
    relationship_maps: dict[str, dict[str, _Relationship]],
    limits: PackageLimits,
) -> WordPermissionRangeInventory:
    """Inventory editable-range markup without exposing editor identities.

    Word stores individual editors in ``w:ed`` and predefined application
    groups in ``w:edGrp`` on ``w:permStart``.  The former can contain an email
    address, alias, or domain identity.  Marker IDs, editor values, exact table
    columns, and placement stay solely in the private element fingerprints.

    An unmatched marker is still stored review state, so it is counted rather
    than treated as an effective permission or discarded.  Exact duplicate
    boundary IDs in one story are rejected because they make a pairing
    ambiguous.
    """

    records: list[tuple[str, ...]] = []
    permission_range_story_count = 0
    permission_start_count = 0
    permission_end_count = 0
    paired_permission_range_count = 0
    unpaired_permission_start_count = 0
    unpaired_permission_end_count = 0
    individual_editor_assignment_count = 0
    editor_group_assignment_count = 0
    editor_group_counts = {group: 0 for group in _WORD_PERMISSION_EDITOR_GROUPS}
    table_column_permission_range_start_count = 0
    custom_xml_displaced_permission_marker_count = 0

    for part_key, kind in _discover_story_parts(members, content_types):
        root = _read_xml(archive, members[part_key], limits)
        if not _is_word_element(root, _STORY_ROOT_NAMES[kind]):
            raise DocumentFormatError("document story is invalid")

        relationships = relationship_maps.get(part_key, {})
        starts: dict[str, _WordPermissionMarker] = {}
        ends: dict[str, _WordPermissionMarker] = {}
        marker_ordinal = 0

        for element in root.iter():
            if _is_word_element(element, "permStart"):
                attributes = _validate_word_permission_range_element(
                    element, "permStart"
                )
                identifier = attributes["id"]
                if identifier in starts:
                    raise DocumentFormatError("Word permission range is invalid")
                marker = _WordPermissionMarker(marker_ordinal, identifier)
                starts[identifier] = marker
                permission_start_count += 1
                if "ed" in attributes:
                    individual_editor_assignment_count += 1
                editor_group = attributes.get("edGrp")
                if editor_group is not None:
                    editor_group_assignment_count += 1
                    editor_group_counts[editor_group] += 1
                if "colFirst" in attributes or "colLast" in attributes:
                    table_column_permission_range_start_count += 1
            elif _is_word_element(element, "permEnd"):
                attributes = _validate_word_permission_range_element(element, "permEnd")
                identifier = attributes["id"]
                if identifier in ends:
                    raise DocumentFormatError("Word permission range is invalid")
                marker = _WordPermissionMarker(marker_ordinal, identifier)
                ends[identifier] = marker
                permission_end_count += 1
            else:
                continue

            if "displacedByCustomXml" in attributes:
                custom_xml_displaced_permission_marker_count += 1
            records.append(
                (
                    "word_permission_marker",
                    part_key,
                    str(marker.ordinal),
                    _fingerprint_element(element, relationships),
                )
            )
            marker_ordinal += 1

        if marker_ordinal:
            permission_range_story_count += 1
        for marker in starts.values():
            end = ends.get(marker.identifier)
            if end is not None and end.ordinal > marker.ordinal:
                paired_permission_range_count += 1
            else:
                unpaired_permission_start_count += 1
        for marker in ends.values():
            start = starts.get(marker.identifier)
            if start is None or start.ordinal >= marker.ordinal:
                unpaired_permission_end_count += 1

    return WordPermissionRangeInventory(
        permission_range_story_count=permission_range_story_count,
        permission_start_count=permission_start_count,
        permission_end_count=permission_end_count,
        paired_permission_range_count=paired_permission_range_count,
        unpaired_permission_start_count=unpaired_permission_start_count,
        unpaired_permission_end_count=unpaired_permission_end_count,
        individual_editor_assignment_count=individual_editor_assignment_count,
        editor_group_assignment_count=editor_group_assignment_count,
        editor_group_none_count=editor_group_counts["none"],
        editor_group_everyone_count=editor_group_counts["everyone"],
        editor_group_administrators_count=editor_group_counts["administrators"],
        editor_group_contributors_count=editor_group_counts["contributors"],
        editor_group_editors_count=editor_group_counts["editors"],
        editor_group_owners_count=editor_group_counts["owners"],
        editor_group_current_count=editor_group_counts["current"],
        table_column_permission_range_start_count=(
            table_column_permission_range_start_count
        ),
        custom_xml_displaced_permission_marker_count=(
            custom_xml_displaced_permission_marker_count
        ),
        signature=_digest_records(records),
    )


def _validate_word_permission_range_element(
    element: ET.Element, expected_local_name: str
) -> dict[str, str]:
    """Validate the narrow, standard range-permission leaf shape."""

    allowed_attribute_names = (
        _WORD_PERMISSION_START_ATTRIBUTE_NAMES
        if expected_local_name == "permStart"
        else _WORD_PERMISSION_END_ATTRIBUTE_NAMES
    )
    if (
        not _is_word_element(element, expected_local_name)
        or list(element)
        or (element.text or "").strip()
    ):
        raise DocumentFormatError("Word permission range is invalid")

    attributes: dict[str, str] = {}
    for attribute, value in element.attrib.items():
        namespace, local_name = _qualified_name(attribute)
        if (
            namespace not in _WORD_NAMESPACES
            or local_name not in allowed_attribute_names
            or local_name in attributes
        ):
            raise DocumentFormatError("Word permission range is invalid")
        attributes[local_name] = value

    if "id" not in attributes:
        raise DocumentFormatError("Word permission range is invalid")
    if expected_local_name == "permStart":
        editor_group = attributes.get("edGrp")
        if (
            editor_group is not None
            and editor_group not in _WORD_PERMISSION_EDITOR_GROUPS
        ):
            raise DocumentFormatError("Word permission range is invalid")
        for attribute_name in ("colFirst", "colLast"):
            value = attributes.get(attribute_name)
            if value is not None and not _WORD_PERMISSION_COLUMN_VALUE.fullmatch(value):
                raise DocumentFormatError("Word permission range is invalid")

    displacement = attributes.get("displacedByCustomXml")
    if (
        displacement is not None
        and displacement not in _WORD_PERMISSION_DISPLACED_BY_CUSTOM_XML_VALUES
    ):
        raise DocumentFormatError("Word permission range is invalid")
    return attributes


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
    preserve_volatile_attributes: bool = False,
) -> str:
    canonical = _canonical_element(
        element,
        relationships,
        ignore_rsids=ignore_rsids,
        preserve_volatile_attributes=preserve_volatile_attributes,
    )
    encoded = json.dumps(canonical, ensure_ascii=False, separators=(",", ":")).encode(
        "utf-8"
    )
    return _digest_bytes(encoded)


def _canonical_element(
    element: ET.Element,
    relationships: dict[str, _Relationship],
    *,
    ignore_rsids: bool,
    preserve_volatile_attributes: bool,
) -> list[object]:
    namespace, local_name = _qualified_name(element.tag)
    attributes: list[tuple[str, object]] = []
    for attribute, value in sorted(element.attrib.items()):
        if _ignore_attribute(
            namespace,
            local_name,
            attribute,
            preserve_volatile_attributes=preserve_volatile_attributes,
        ):
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
            _canonical_element(
                child,
                relationships,
                ignore_rsids=ignore_rsids,
                preserve_volatile_attributes=preserve_volatile_attributes,
            )
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
    *,
    preserve_volatile_attributes: bool,
) -> bool:
    attribute_namespace, attribute_local_name = _qualified_name(attribute)
    if (
        attribute_namespace in _WORD_NAMESPACES
        and attribute_local_name.casefold().startswith("rsid")
    ):
        return True
    if (
        not preserve_volatile_attributes
        and attribute_local_name in _VOLATILE_ATTRIBUTE_NAMES
    ):
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
