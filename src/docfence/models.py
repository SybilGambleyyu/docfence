"""Private snapshot state and public-safe report models."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Literal

Severity = Literal["critical", "high", "medium", "low"]


@dataclass(frozen=True)
class RevisionInventory:
    """Counts of stored WordprocessingML revision markup."""

    insertions: int = 0
    deletions: int = 0
    move_from: int = 0
    move_to: int = 0
    property_changes: int = 0

    @property
    def unresolved_count(self) -> int:
        return (
            self.insertions
            + self.deletions
            + self.move_from
            + self.move_to
            + self.property_changes
        )

    def public_dict(self) -> dict[str, int]:
        return {
            "insertions": self.insertions,
            "deletions": self.deletions,
            "move_from": self.move_from,
            "move_to": self.move_to,
            "property_changes": self.property_changes,
            "unresolved_count": self.unresolved_count,
        }


@dataclass(frozen=True)
class StyleInventory:
    """Stored style/default declarations relevant to hidden-text review coverage."""

    style_definition_count: int
    hidden_text_style_definition_count: int
    document_default_hidden_text_enabled: bool
    signature: str

    def public_dict(self) -> dict[str, object]:
        return {
            "style_definition_count": self.style_definition_count,
            "hidden_text_style_definition_count": (
                self.hidden_text_style_definition_count
            ),
            "document_default_hidden_text_enabled": (
                self.document_default_hidden_text_enabled
            ),
        }


@dataclass(frozen=True)
class EmbeddedObjectInventory:
    """Stored OLE/package object and embedded-control evidence."""

    object_relationship_count: int
    object_part_count: int
    control_relationship_count: int
    control_part_count: int
    signature: str

    def public_dict(self) -> dict[str, int]:
        return {
            "embedded_object_relationship_count": self.object_relationship_count,
            "embedded_object_part_count": self.object_part_count,
            "embedded_control_relationship_count": self.control_relationship_count,
            "embedded_control_part_count": self.control_part_count,
        }


@dataclass(frozen=True)
class AlternativeFormatImportInventory:
    """Stored OOXML alternative-format import relationships and payloads."""

    relationship_count: int
    payload_part_count: int
    signature: str

    def public_dict(self) -> dict[str, int]:
        return {
            "alternative_format_import_relationship_count": self.relationship_count,
            "alternative_format_import_payload_part_count": self.payload_part_count,
        }


@dataclass(frozen=True)
class DocumentPropertyInventory:
    """Stored core, extended, and custom document-property evidence."""

    core_property_part_count: int
    core_property_value_count: int
    extended_property_part_count: int
    extended_property_value_count: int
    custom_property_part_count: int
    custom_property_count: int
    signature: str

    def public_dict(self) -> dict[str, int]:
        return {
            "core_property_part_count": self.core_property_part_count,
            "core_property_value_count": self.core_property_value_count,
            "extended_property_part_count": self.extended_property_part_count,
            "extended_property_value_count": self.extended_property_value_count,
            "custom_property_part_count": self.custom_property_part_count,
            "custom_property_count": self.custom_property_count,
        }


@dataclass(frozen=True)
class SensitivityLabelInventory:
    """Stored Microsoft Purview sensitivity-label evidence, aggregate only."""

    label_info_part_count: int
    label_info_label_count: int
    label_info_enabled_label_count: int
    label_info_removed_label_count: int
    label_info_extension_count: int
    legacy_mip_label_count: int
    legacy_mip_property_count: int
    legacy_sensitivity_property_count: int
    word_content_marking_property_count: int
    signature: str

    def public_dict(self) -> dict[str, int]:
        return {
            "label_info_part_count": self.label_info_part_count,
            "label_info_label_count": self.label_info_label_count,
            "label_info_enabled_label_count": self.label_info_enabled_label_count,
            "label_info_removed_label_count": self.label_info_removed_label_count,
            "label_info_extension_count": self.label_info_extension_count,
            "legacy_mip_label_count": self.legacy_mip_label_count,
            "legacy_mip_property_count": self.legacy_mip_property_count,
            "legacy_sensitivity_property_count": (
                self.legacy_sensitivity_property_count
            ),
            "word_content_marking_property_count": (
                self.word_content_marking_property_count
            ),
        }


@dataclass(frozen=True)
class PackageDigitalSignatureInventory:
    """Stored OPC package digital-signature evidence, aggregate only."""

    signature_origin_part_count: int
    xml_signature_part_count: int
    certificate_part_count: int
    signed_info_reference_count: int
    manifest_reference_count: int
    relationship_reference_count: int
    inline_x509_certificate_count: int
    signature_property_count: int
    signature: str

    def public_dict(self) -> dict[str, int]:
        return {
            "signature_origin_part_count": self.signature_origin_part_count,
            "xml_signature_part_count": self.xml_signature_part_count,
            "certificate_part_count": self.certificate_part_count,
            "signed_info_reference_count": self.signed_info_reference_count,
            "manifest_reference_count": self.manifest_reference_count,
            "relationship_reference_count": self.relationship_reference_count,
            "inline_x509_certificate_count": self.inline_x509_certificate_count,
            "signature_property_count": self.signature_property_count,
        }


@dataclass(frozen=True)
class WordProtectionInventory:
    """Stored Word editing and write-protection evidence, aggregate only."""

    document_protection_count: int
    document_protection_enforcement_enabled_count: int
    document_protection_formatting_restricted_count: int
    document_protection_read_only_count: int
    document_protection_comments_count: int
    document_protection_tracked_changes_count: int
    document_protection_forms_count: int
    document_protection_password_material_count: int
    write_protection_count: int
    write_protection_recommended_count: int
    write_protection_password_material_count: int
    signature: str

    def public_dict(self) -> dict[str, int]:
        return {
            "document_protection_count": self.document_protection_count,
            "document_protection_enforcement_enabled_count": (
                self.document_protection_enforcement_enabled_count
            ),
            "document_protection_formatting_restricted_count": (
                self.document_protection_formatting_restricted_count
            ),
            "document_protection_read_only_count": (
                self.document_protection_read_only_count
            ),
            "document_protection_comments_count": (
                self.document_protection_comments_count
            ),
            "document_protection_tracked_changes_count": (
                self.document_protection_tracked_changes_count
            ),
            "document_protection_forms_count": self.document_protection_forms_count,
            "document_protection_password_material_count": (
                self.document_protection_password_material_count
            ),
            "write_protection_count": self.write_protection_count,
            "write_protection_recommended_count": (
                self.write_protection_recommended_count
            ),
            "write_protection_password_material_count": (
                self.write_protection_password_material_count
            ),
        }


@dataclass(frozen=True)
class WordDocumentVariableInventory:
    """Stored Word document-variable state, aggregate only.

    Variable names and values can preserve template or automation state. They
    remain inside the private signature; public output reports only container
    and value-state counts.
    """

    document_variable_container_count: int
    document_variable_count: int
    empty_document_variable_value_count: int
    signature: str

    def public_dict(self) -> dict[str, int]:
        return {
            "document_variable_container_count": (
                self.document_variable_container_count
            ),
            "document_variable_count": self.document_variable_count,
            "empty_document_variable_value_count": (
                self.empty_document_variable_value_count
            ),
        }


@dataclass(frozen=True)
class WordDocumentVariableFieldInventory:
    """Stored ``DOCVARIABLE`` field references, aggregate only.

    Field instructions and variable names can reveal automation or template
    state.  They remain inside the private signature; public output reports
    only reference counts and conservative exact-literal associations with a
    discovered variable in the same package document scope.
    """

    document_variable_field_reference_count: int
    document_variable_field_story_count: int
    literal_document_variable_field_reference_count: int
    nonliteral_document_variable_field_reference_count: int
    literal_document_variable_field_reference_matching_stored_variable_count: int
    literal_document_variable_field_reference_not_matching_stored_variable_count: int
    signature: str

    def public_dict(self) -> dict[str, int]:
        return {
            "document_variable_field_reference_count": (
                self.document_variable_field_reference_count
            ),
            "document_variable_field_story_count": (
                self.document_variable_field_story_count
            ),
            "literal_document_variable_field_reference_count": (
                self.literal_document_variable_field_reference_count
            ),
            "nonliteral_document_variable_field_reference_count": (
                self.nonliteral_document_variable_field_reference_count
            ),
            (
                "literal_document_variable_field_reference_"
                "matching_stored_variable_count"
            ): (
                self.literal_document_variable_field_reference_matching_stored_variable_count
            ),
            (
                "literal_document_variable_field_reference_"
                "not_matching_stored_variable_count"
            ): (
                self.literal_document_variable_field_reference_not_matching_stored_variable_count
            ),
        }


@dataclass(frozen=True)
class WordHyperlinkFieldInventory:
    """Stored ``HYPERLINK`` field references, aggregate only.

    Field instructions can embed destinations without an OOXML relationship.
    Destinations, internal locations, ScreenTips, targets, story paths, and
    fingerprints remain in the private signature; public output reports only
    conservative lexical classifications of complete field instructions.
    """

    hyperlink_field_reference_count: int
    hyperlink_field_story_count: int
    literal_destination_hyperlink_field_count: int
    literal_internal_location_only_hyperlink_field_count: int
    dynamic_or_unparseable_hyperlink_field_count: int
    signature: str

    def public_dict(self) -> dict[str, int]:
        return {
            "hyperlink_field_reference_count": self.hyperlink_field_reference_count,
            "hyperlink_field_story_count": self.hyperlink_field_story_count,
            "literal_destination_hyperlink_field_count": (
                self.literal_destination_hyperlink_field_count
            ),
            "literal_internal_location_only_hyperlink_field_count": (
                self.literal_internal_location_only_hyperlink_field_count
            ),
            "dynamic_or_unparseable_hyperlink_field_count": (
                self.dynamic_or_unparseable_hyperlink_field_count
            ),
        }


@dataclass(frozen=True)
class WordHyperlinkMarkupInventory:
    """Direct WordprocessingML ``w:hyperlink`` elements, aggregate only.

    Relationship targets, local anchors, locations, tooltips, frame names,
    display text, story paths, and fingerprints can reveal document context.
    They remain in the private signature; public output reports only stored
    target-mechanism and relationship-mode evidence.
    """

    hyperlink_element_count: int
    hyperlink_story_count: int
    relationship_backed_hyperlink_count: int
    external_relationship_hyperlink_count: int
    internal_relationship_hyperlink_count: int
    unsupported_relationship_hyperlink_count: int
    anchor_only_hyperlink_count: int
    default_document_start_hyperlink_count: int
    relationship_backed_anchor_attribute_count: int
    signature: str

    def public_dict(self) -> dict[str, int]:
        return {
            "hyperlink_element_count": self.hyperlink_element_count,
            "hyperlink_story_count": self.hyperlink_story_count,
            "relationship_backed_hyperlink_count": (
                self.relationship_backed_hyperlink_count
            ),
            "external_relationship_hyperlink_count": (
                self.external_relationship_hyperlink_count
            ),
            "internal_relationship_hyperlink_count": (
                self.internal_relationship_hyperlink_count
            ),
            "unsupported_relationship_hyperlink_count": (
                self.unsupported_relationship_hyperlink_count
            ),
            "anchor_only_hyperlink_count": self.anchor_only_hyperlink_count,
            "default_document_start_hyperlink_count": (
                self.default_document_start_hyperlink_count
            ),
            "relationship_backed_anchor_attribute_count": (
                self.relationship_backed_anchor_attribute_count
            ),
        }


@dataclass(frozen=True)
class WordDrawingHyperlinkInventory:
    """Direct DrawingML hyperlink-action markup in Word stories, aggregate only.

    Drawing hyperlink targets, invalid URLs, actions, tooltips, frame names,
    history settings, story paths, and fingerprints can reveal document context.
    They remain in the private signature; public output reports only stored
    action kind, relationship-mode, and selected attribute-presence evidence.
    """

    drawing_hyperlink_reference_count: int
    drawing_hyperlink_story_count: int
    click_drawing_hyperlink_count: int
    hover_drawing_hyperlink_count: int
    mouse_over_drawing_hyperlink_count: int
    external_relationship_drawing_hyperlink_count: int
    internal_relationship_drawing_hyperlink_count: int
    unsupported_relationship_drawing_hyperlink_count: int
    missing_relationship_id_drawing_hyperlink_count: int
    action_attribute_drawing_hyperlink_count: int
    invalid_url_attribute_drawing_hyperlink_count: int
    signature: str

    def public_dict(self) -> dict[str, int]:
        return {
            "drawing_hyperlink_reference_count": (
                self.drawing_hyperlink_reference_count
            ),
            "drawing_hyperlink_story_count": self.drawing_hyperlink_story_count,
            "click_drawing_hyperlink_count": self.click_drawing_hyperlink_count,
            "hover_drawing_hyperlink_count": self.hover_drawing_hyperlink_count,
            "mouse_over_drawing_hyperlink_count": (
                self.mouse_over_drawing_hyperlink_count
            ),
            "external_relationship_drawing_hyperlink_count": (
                self.external_relationship_drawing_hyperlink_count
            ),
            "internal_relationship_drawing_hyperlink_count": (
                self.internal_relationship_drawing_hyperlink_count
            ),
            "unsupported_relationship_drawing_hyperlink_count": (
                self.unsupported_relationship_drawing_hyperlink_count
            ),
            "missing_relationship_id_drawing_hyperlink_count": (
                self.missing_relationship_id_drawing_hyperlink_count
            ),
            "action_attribute_drawing_hyperlink_count": (
                self.action_attribute_drawing_hyperlink_count
            ),
            "invalid_url_attribute_drawing_hyperlink_count": (
                self.invalid_url_attribute_drawing_hyperlink_count
            ),
        }


@dataclass(frozen=True)
class WordDrawingLinkedPictureInventory:
    """Direct DrawingML linked-picture markup in Word stories, aggregate only.

    Linked-picture targets, relationship IDs, surrounding drawing markup, story
    paths, and fingerprints can reveal document context. They remain in the
    private signature; public output reports only marker and relationship
    classification counts. A stored marker is not evidence that a Word client
    will load, render, or otherwise honor the referenced image.
    """

    drawing_linked_picture_reference_count: int
    drawing_linked_picture_story_count: int
    external_image_relationship_drawing_linked_picture_count: int
    internal_image_relationship_drawing_linked_picture_count: int
    unsupported_relationship_drawing_linked_picture_count: int
    signature: str

    def public_dict(self) -> dict[str, int]:
        return {
            "drawing_linked_picture_reference_count": (
                self.drawing_linked_picture_reference_count
            ),
            "drawing_linked_picture_story_count": (
                self.drawing_linked_picture_story_count
            ),
            "external_image_relationship_drawing_linked_picture_count": (
                self.external_image_relationship_drawing_linked_picture_count
            ),
            "internal_image_relationship_drawing_linked_picture_count": (
                self.internal_image_relationship_drawing_linked_picture_count
            ),
            "unsupported_relationship_drawing_linked_picture_count": (
                self.unsupported_relationship_drawing_linked_picture_count
            ),
        }


@dataclass(frozen=True)
class WordVmlHyperlinkInventory:
    """Direct legacy VML ``href`` markup in Word stories, aggregate only.

    VML href values, frame targets, titles, alternate text, shape identifiers,
    story paths, and fingerprints can reveal document context. They remain in
    the private signature; public output reports only marker category and
    target-attribute-presence evidence.
    """

    vml_hyperlink_element_count: int
    vml_hyperlink_story_count: int
    concrete_shape_vml_hyperlink_count: int
    group_vml_hyperlink_count: int
    shape_type_vml_hyperlink_count: int
    target_attribute_vml_hyperlink_count: int
    signature: str

    def public_dict(self) -> dict[str, int]:
        return {
            "vml_hyperlink_element_count": self.vml_hyperlink_element_count,
            "vml_hyperlink_story_count": self.vml_hyperlink_story_count,
            "concrete_shape_vml_hyperlink_count": (
                self.concrete_shape_vml_hyperlink_count
            ),
            "group_vml_hyperlink_count": self.group_vml_hyperlink_count,
            "shape_type_vml_hyperlink_count": (
                self.shape_type_vml_hyperlink_count
            ),
            "target_attribute_vml_hyperlink_count": (
                self.target_attribute_vml_hyperlink_count
            ),
        }


@dataclass(frozen=True)
class WordVmlExternalImageInventory:
    """Direct legacy VML external-image markup in Word stories, aggregate only.

    Image targets, relationship IDs, VML markup, story paths, and fingerprints
    can reveal document context. Public output reports only stored direct-marker
    and relationship classification counts; the private signature carries only
    the reviewed relationship semantics, not other VML image-data attributes.
    A stored marker is not evidence that a Word client will load, render,
    update, or otherwise honor the referenced image.
    """

    vml_external_image_reference_count: int
    vml_external_image_story_count: int
    external_image_relationship_vml_external_image_count: int
    unsupported_relationship_vml_external_image_count: int
    signature: str

    def public_dict(self) -> dict[str, int]:
        return {
            "vml_external_image_reference_count": (
                self.vml_external_image_reference_count
            ),
            "vml_external_image_story_count": self.vml_external_image_story_count,
            "external_image_relationship_vml_external_image_count": (
                self.external_image_relationship_vml_external_image_count
            ),
            "unsupported_relationship_vml_external_image_count": (
                self.unsupported_relationship_vml_external_image_count
            ),
        }


@dataclass(frozen=True)
class WordVmlImageHyperlinkInventory:
    """Direct legacy VML image-data hyperlink markup, aggregate only.

    Relationship targets and IDs, VML markup, story paths, and fingerprints can
    reveal document context. Public output reports only direct-marker and
    relationship-classification counts; the private signature carries only the
    reviewed relationship semantics, not other VML image-data attributes. A
    stored marker is not evidence that a Word client will select, follow, or
    otherwise honor a target.
    """

    vml_image_hyperlink_reference_count: int
    vml_image_hyperlink_story_count: int
    external_relationship_vml_image_hyperlink_count: int
    internal_relationship_vml_image_hyperlink_count: int
    unsupported_relationship_vml_image_hyperlink_count: int
    signature: str

    def public_dict(self) -> dict[str, int]:
        return {
            "vml_image_hyperlink_reference_count": (
                self.vml_image_hyperlink_reference_count
            ),
            "vml_image_hyperlink_story_count": (
                self.vml_image_hyperlink_story_count
            ),
            "external_relationship_vml_image_hyperlink_count": (
                self.external_relationship_vml_image_hyperlink_count
            ),
            "internal_relationship_vml_image_hyperlink_count": (
                self.internal_relationship_vml_image_hyperlink_count
            ),
            "unsupported_relationship_vml_image_hyperlink_count": (
                self.unsupported_relationship_vml_image_hyperlink_count
            ),
        }


@dataclass(frozen=True)
class WordVmlLinkedOleObjectInventory:
    """Direct legacy VML Office linked-OLE markup, aggregate only.

    Link sources, program identifiers, IDs, field codes, VML markup, story
    paths, and fingerprints can reveal document context. Public output reports
    only marker, update-mode, and relationship-classification counts; the
    private signature retains the direct reviewed marker. A stored marker is not
    evidence that a Word client will update, retrieve, activate, or otherwise
    honor an OLE object.
    """

    vml_linked_ole_object_count: int
    vml_linked_ole_object_story_count: int
    automatic_update_vml_linked_ole_object_count: int
    nonautomatic_or_unspecified_update_vml_linked_ole_object_count: int
    external_standard_ole_object_relationship_vml_linked_ole_object_count: int
    internal_standard_ole_object_relationship_vml_linked_ole_object_count: int
    unsupported_relationship_vml_linked_ole_object_count: int
    without_relationship_id_vml_linked_ole_object_count: int
    signature: str

    def public_dict(self) -> dict[str, int]:
        return {
            "vml_linked_ole_object_count": self.vml_linked_ole_object_count,
            "vml_linked_ole_object_story_count": (
                self.vml_linked_ole_object_story_count
            ),
            "automatic_update_vml_linked_ole_object_count": (
                self.automatic_update_vml_linked_ole_object_count
            ),
            "nonautomatic_or_unspecified_update_vml_linked_ole_object_count": (
                self.nonautomatic_or_unspecified_update_vml_linked_ole_object_count
            ),
            "external_standard_ole_object_relationship_vml_linked_ole_object_count": (
                self.external_standard_ole_object_relationship_vml_linked_ole_object_count
            ),
            "internal_standard_ole_object_relationship_vml_linked_ole_object_count": (
                self.internal_standard_ole_object_relationship_vml_linked_ole_object_count
            ),
            "unsupported_relationship_vml_linked_ole_object_count": (
                self.unsupported_relationship_vml_linked_ole_object_count
            ),
            "without_relationship_id_vml_linked_ole_object_count": (
                self.without_relationship_id_vml_linked_ole_object_count
            ),
        }


@dataclass(frozen=True)
class WordVmlEmbeddedOleObjectInventory:
    """Direct legacy VML Office embedded-OLE markup, aggregate only.

    Program, shape, and object identifiers, relationship IDs and targets,
    update metadata, field codes, VML markup, story paths, and fingerprints can
    reveal document context. Public output reports only marker and
    relationship-classification counts; the private signature retains the
    direct reviewed marker. A stored marker is not evidence that a Word client
    will load, render, activate, or otherwise honor an OLE object.
    """

    vml_embedded_ole_object_count: int
    vml_embedded_ole_object_story_count: int
    external_standard_ole_object_relationship_vml_embedded_ole_object_count: int
    internal_standard_ole_object_relationship_vml_embedded_ole_object_count: int
    unsupported_relationship_vml_embedded_ole_object_count: int
    without_relationship_id_vml_embedded_ole_object_count: int
    signature: str

    def public_dict(self) -> dict[str, int]:
        return {
            "vml_embedded_ole_object_count": self.vml_embedded_ole_object_count,
            "vml_embedded_ole_object_story_count": (
                self.vml_embedded_ole_object_story_count
            ),
            "external_standard_ole_object_relationship_vml_embedded_ole_object_count": (
                self.external_standard_ole_object_relationship_vml_embedded_ole_object_count
            ),
            "internal_standard_ole_object_relationship_vml_embedded_ole_object_count": (
                self.internal_standard_ole_object_relationship_vml_embedded_ole_object_count
            ),
            "unsupported_relationship_vml_embedded_ole_object_count": (
                self.unsupported_relationship_vml_embedded_ole_object_count
            ),
            "without_relationship_id_vml_embedded_ole_object_count": (
                self.without_relationship_id_vml_embedded_ole_object_count
            ),
        }


@dataclass(frozen=True)
class WordObjectLinkInventory:
    """Direct WordprocessingML linked-object-property markup, aggregate only.

    Relationship targets and IDs, program and shape identifiers, field codes,
    markup, story paths, and fingerprints can reveal document context. Public
    output reports only direct-marker, exact stored update-mode, and
    relationship-classification counts; the private signature retains the
    direct reviewed marker. A stored marker is not evidence that a Word client
    will update, retrieve, activate, or otherwise honor an OLE object.
    """

    object_link_count: int
    object_link_story_count: int
    automatic_update_object_link_count: int
    on_call_update_object_link_count: int
    unsupported_or_missing_update_mode_object_link_count: int
    external_standard_ole_object_relationship_object_link_count: int
    internal_standard_ole_object_relationship_object_link_count: int
    unsupported_relationship_object_link_count: int
    without_relationship_id_object_link_count: int
    signature: str

    def public_dict(self) -> dict[str, int]:
        return {
            "object_link_count": self.object_link_count,
            "object_link_story_count": self.object_link_story_count,
            "automatic_update_object_link_count": (
                self.automatic_update_object_link_count
            ),
            "on_call_update_object_link_count": self.on_call_update_object_link_count,
            "unsupported_or_missing_update_mode_object_link_count": (
                self.unsupported_or_missing_update_mode_object_link_count
            ),
            "external_standard_ole_object_relationship_object_link_count": (
                self.external_standard_ole_object_relationship_object_link_count
            ),
            "internal_standard_ole_object_relationship_object_link_count": (
                self.internal_standard_ole_object_relationship_object_link_count
            ),
            "unsupported_relationship_object_link_count": (
                self.unsupported_relationship_object_link_count
            ),
            "without_relationship_id_object_link_count": (
                self.without_relationship_id_object_link_count
            ),
        }


@dataclass(frozen=True)
class WordEmbeddedControlInventory:
    """Direct WordprocessingML embedded-control anchors, aggregate only.

    Relationship targets and IDs, control names, shape identifiers, markup,
    story paths, and fingerprints can reveal document context. Public output
    reports only direct-marker, parent-position, and relationship-classification
    counts; the private signature retains the direct reviewed marker. A stored
    marker is not evidence that a Word client will load, render, activate, or
    otherwise honor an embedded control.
    """

    embedded_control_count: int
    embedded_control_story_count: int
    object_parent_embedded_control_count: int
    pict_parent_embedded_control_count: int
    internal_standard_control_relationship_embedded_control_count: int
    external_standard_control_relationship_embedded_control_count: int
    unsupported_relationship_embedded_control_count: int
    without_relationship_id_embedded_control_count: int
    signature: str

    def public_dict(self) -> dict[str, int]:
        return {
            "embedded_control_count": self.embedded_control_count,
            "embedded_control_story_count": self.embedded_control_story_count,
            "object_parent_embedded_control_count": (
                self.object_parent_embedded_control_count
            ),
            "pict_parent_embedded_control_count": (
                self.pict_parent_embedded_control_count
            ),
            "internal_standard_control_relationship_embedded_control_count": (
                self.internal_standard_control_relationship_embedded_control_count
            ),
            "external_standard_control_relationship_embedded_control_count": (
                self.external_standard_control_relationship_embedded_control_count
            ),
            "unsupported_relationship_embedded_control_count": (
                self.unsupported_relationship_embedded_control_count
            ),
            "without_relationship_id_embedded_control_count": (
                self.without_relationship_id_embedded_control_count
            ),
        }


@dataclass(frozen=True)
class WordPermissionRangeInventory:
    """Stored Word editable-range permission markup, aggregate only.

    Individual editor values can be email addresses, aliases, or domain
    identities. They remain inside the private signature; the public contract
    reports only marker, pairing, and predefined-group categories.
    """

    permission_range_story_count: int
    permission_start_count: int
    permission_end_count: int
    paired_permission_range_count: int
    unpaired_permission_start_count: int
    unpaired_permission_end_count: int
    individual_editor_assignment_count: int
    editor_group_assignment_count: int
    editor_group_none_count: int
    editor_group_everyone_count: int
    editor_group_administrators_count: int
    editor_group_contributors_count: int
    editor_group_editors_count: int
    editor_group_owners_count: int
    editor_group_current_count: int
    table_column_permission_range_start_count: int
    custom_xml_displaced_permission_marker_count: int
    signature: str

    def public_dict(self) -> dict[str, int]:
        return {
            "permission_range_story_count": self.permission_range_story_count,
            "permission_start_count": self.permission_start_count,
            "permission_end_count": self.permission_end_count,
            "paired_permission_range_count": self.paired_permission_range_count,
            "unpaired_permission_start_count": self.unpaired_permission_start_count,
            "unpaired_permission_end_count": self.unpaired_permission_end_count,
            "individual_editor_assignment_count": (
                self.individual_editor_assignment_count
            ),
            "editor_group_assignment_count": self.editor_group_assignment_count,
            "editor_group_none_count": self.editor_group_none_count,
            "editor_group_everyone_count": self.editor_group_everyone_count,
            "editor_group_administrators_count": self.editor_group_administrators_count,
            "editor_group_contributors_count": self.editor_group_contributors_count,
            "editor_group_editors_count": self.editor_group_editors_count,
            "editor_group_owners_count": self.editor_group_owners_count,
            "editor_group_current_count": self.editor_group_current_count,
            "table_column_permission_range_start_count": (
                self.table_column_permission_range_start_count
            ),
            "custom_xml_displaced_permission_marker_count": (
                self.custom_xml_displaced_permission_marker_count
            ),
        }


@dataclass(frozen=True)
class MailMergeInventory:
    """Stored mail-merge configuration and recipient-data evidence."""

    configuration_count: int
    data_source_relationship_count: int
    header_source_relationship_count: int
    recipient_data_relationship_count: int
    recipient_data_part_count: int
    signature: str

    def public_dict(self) -> dict[str, int]:
        return {
            "mail_merge_configuration_count": self.configuration_count,
            "mail_merge_data_source_relationship_count": (
                self.data_source_relationship_count
            ),
            "mail_merge_header_source_relationship_count": (
                self.header_source_relationship_count
            ),
            "mail_merge_recipient_data_relationship_count": (
                self.recipient_data_relationship_count
            ),
            "mail_merge_recipient_data_part_count": (self.recipient_data_part_count),
        }


@dataclass(frozen=True)
class SaveThroughXsltInventory:
    """Stored XSLT-on-single-XML-save configuration, aggregate only.

    Transform targets, relationship identifiers, local solution identifiers,
    Settings-part paths, and private fingerprints can expose environment or
    workflow details. Public output therefore reports only the configuration
    shape and enabled-setting state.
    """

    enabled_setting_count: int
    disabled_setting_count: int
    transform_anchor_count: int
    transform_relationship_count: int
    solution_identifier_count: int
    signature: str

    def public_dict(self) -> dict[str, int]:
        return {
            "save_through_xslt_enabled_setting_count": self.enabled_setting_count,
            "save_through_xslt_disabled_setting_count": self.disabled_setting_count,
            "save_through_xslt_anchor_count": self.transform_anchor_count,
            "save_through_xslt_relationship_count": (
                self.transform_relationship_count
            ),
            "save_through_xslt_solution_identifier_count": (
                self.solution_identifier_count
            ),
        }


@dataclass(frozen=True)
class AttachedCustomXmlSchemaInventory:
    """Stored custom-XML schema associations, exposed only as a count.

    Attached-schema namespace identifiers can identify private vocabularies or
    workflow systems.  Their values and Settings-part paths stay inside the
    private signature; public output reports only how many declarations exist.
    """

    attached_schema_count: int
    signature: str

    def public_dict(self) -> dict[str, int]:
        return {"attached_custom_xml_schema_count": self.attached_schema_count}


@dataclass(frozen=True)
class FieldUpdateOnOpenInventory:
    """Stored automatic-field-recalculation settings, aggregate only.

    Settings-part paths and private fingerprints remain local. Public output
    records only whether explicitly stored settings request or decline automatic
    recalculation when a capable document host opens the package.
    """

    enabled_setting_count: int
    disabled_setting_count: int
    signature: str

    def public_dict(self) -> dict[str, int]:
        return {
            "field_update_on_open_enabled_setting_count": self.enabled_setting_count,
            "field_update_on_open_disabled_setting_count": self.disabled_setting_count,
        }


@dataclass(frozen=True)
class TemplateStyleUpdateOnOpenInventory:
    """Stored automatic-template-style-update settings, aggregate only.

    Settings-part paths and private fingerprints remain local. Public output
    records only whether explicitly stored settings request or decline automatic
    style updates from an attached template when a capable document host opens
    the package.
    """

    enabled_setting_count: int
    disabled_setting_count: int
    signature: str

    def public_dict(self) -> dict[str, int]:
        return {
            "template_style_update_on_open_enabled_setting_count": (
                self.enabled_setting_count
            ),
            "template_style_update_on_open_disabled_setting_count": (
                self.disabled_setting_count
            ),
        }


@dataclass(frozen=True)
class DataBindingInventory:
    """Stored content-control-to-custom-XML mapping evidence."""

    binding_count: int
    binding_with_store_item_id_count: int
    binding_without_store_item_id_count: int
    referenced_custom_xml_part_count: int
    unmatched_store_item_id_count: int
    signature: str

    def public_dict(self) -> dict[str, int]:
        return {
            "data_binding_count": self.binding_count,
            "data_binding_with_store_item_id_count": (
                self.binding_with_store_item_id_count
            ),
            "data_binding_without_store_item_id_count": (
                self.binding_without_store_item_id_count
            ),
            "data_binding_referenced_custom_xml_part_count": (
                self.referenced_custom_xml_part_count
            ),
            "data_binding_unmatched_store_item_id_count": (
                self.unmatched_store_item_id_count
            ),
        }


@dataclass(frozen=True)
class ExternalFieldInventory:
    """Stored Word field instructions that can source outside material."""

    database_field_count: int
    legacy_data_field_count: int
    dde_field_count: int
    dde_auto_field_count: int
    include_text_field_count: int
    include_picture_field_count: int
    link_field_count: int
    referenced_document_field_count: int
    signature: str

    def public_dict(self) -> dict[str, int]:
        return {
            "database_field_count": self.database_field_count,
            "legacy_data_field_count": self.legacy_data_field_count,
            "dde_field_count": self.dde_field_count,
            "dde_auto_field_count": self.dde_auto_field_count,
            "include_text_field_count": self.include_text_field_count,
            "include_picture_field_count": self.include_picture_field_count,
            "link_field_count": self.link_field_count,
            "referenced_document_field_count": self.referenced_document_field_count,
        }


@dataclass(frozen=True)
class ModernCommentMetadataInventory:
    """Stored modern Word comment metadata, exposed only as aggregate counts."""

    people_part_count: int
    person_count: int
    presence_info_count: int
    comments_extended_part_count: int
    comment_extension_count: int
    threaded_comment_count: int
    resolved_comment_count: int
    comments_id_part_count: int
    comment_id_count: int
    comments_extensible_part_count: int
    comment_extensible_count: int
    reaction_count: int
    reaction_user_count: int
    signature: str

    def public_dict(self) -> dict[str, int]:
        return {
            "people_part_count": self.people_part_count,
            "person_count": self.person_count,
            "presence_info_count": self.presence_info_count,
            "comments_extended_part_count": self.comments_extended_part_count,
            "comment_extension_count": self.comment_extension_count,
            "threaded_comment_count": self.threaded_comment_count,
            "resolved_comment_count": self.resolved_comment_count,
            "comments_id_part_count": self.comments_id_part_count,
            "comment_id_count": self.comment_id_count,
            "comments_extensible_part_count": self.comments_extensible_part_count,
            "comment_extensible_count": self.comment_extensible_count,
            "reaction_count": self.reaction_count,
            "reaction_user_count": self.reaction_user_count,
        }


@dataclass(frozen=True)
class DocumentTaskInventory:
    """Stored Word document tasks, exposed only as aggregate workflow counts."""

    document_task_part_count: int
    task_count: int
    task_history_event_count: int
    task_user_reference_count: int
    task_comment_anchor_count: int
    assignment_event_count: int
    unassignment_event_count: int
    creation_event_count: int
    title_change_event_count: int
    schedule_change_event_count: int
    progress_change_event_count: int
    priority_change_event_count: int
    deletion_event_count: int
    restoration_event_count: int
    unassign_all_event_count: int
    undo_event_count: int
    signature: str

    def public_dict(self) -> dict[str, int]:
        return {
            "document_task_part_count": self.document_task_part_count,
            "task_count": self.task_count,
            "task_history_event_count": self.task_history_event_count,
            "task_user_reference_count": self.task_user_reference_count,
            "task_comment_anchor_count": self.task_comment_anchor_count,
            "assignment_event_count": self.assignment_event_count,
            "unassignment_event_count": self.unassignment_event_count,
            "creation_event_count": self.creation_event_count,
            "title_change_event_count": self.title_change_event_count,
            "schedule_change_event_count": self.schedule_change_event_count,
            "progress_change_event_count": self.progress_change_event_count,
            "priority_change_event_count": self.priority_change_event_count,
            "deletion_event_count": self.deletion_event_count,
            "restoration_event_count": self.restoration_event_count,
            "unassign_all_event_count": self.unassign_all_event_count,
            "undo_event_count": self.undo_event_count,
        }


@dataclass(frozen=True)
class TaskpaneWebExtensionInventory:
    """Stored task-pane Office add-in configuration, kept aggregate-only."""

    taskpane_part_count: int
    taskpane_count: int
    visible_taskpane_count: int
    locked_taskpane_count: int
    web_extension_part_count: int
    web_extension_reference_count: int
    web_extension_property_count: int
    web_extension_binding_count: int
    auto_show_taskpane_setting_count: int
    web_extension_bound_content_control_count: int
    signature: str

    def public_dict(self) -> dict[str, int]:
        return {
            "taskpane_part_count": self.taskpane_part_count,
            "taskpane_count": self.taskpane_count,
            "visible_taskpane_count": self.visible_taskpane_count,
            "locked_taskpane_count": self.locked_taskpane_count,
            "web_extension_part_count": self.web_extension_part_count,
            "web_extension_reference_count": self.web_extension_reference_count,
            "web_extension_property_count": self.web_extension_property_count,
            "web_extension_binding_count": self.web_extension_binding_count,
            "auto_show_taskpane_setting_count": self.auto_show_taskpane_setting_count,
            "web_extension_bound_content_control_count": (
                self.web_extension_bound_content_control_count
            ),
        }


@dataclass(frozen=True)
class ExternalDocumentDependencyInventory:
    """Stored references to external Word document packages."""

    attached_template_anchor_count: int
    attached_template_relationship_count: int
    subdocument_anchor_count: int
    subdocument_relationship_count: int
    frame_source_anchor_count: int
    frame_relationship_count: int
    signature: str

    def public_dict(self) -> dict[str, int]:
        return {
            "attached_template_anchor_count": self.attached_template_anchor_count,
            "attached_template_relationship_count": (
                self.attached_template_relationship_count
            ),
            "subdocument_anchor_count": self.subdocument_anchor_count,
            "subdocument_relationship_count": self.subdocument_relationship_count,
            "frame_source_anchor_count": self.frame_source_anchor_count,
            "frame_relationship_count": self.frame_relationship_count,
        }


@dataclass(frozen=True)
class StorySnapshot:
    """One private document story fingerprint plus aggregate review state."""

    part_key: str
    kind: str
    block_signatures: tuple[str, ...]
    structure_signature: str
    paragraph_count: int
    table_count: int
    text_run_count: int
    hidden_text_run_count: int
    hidden_paragraph_mark_count: int
    alternative_format_import_anchor_count: int
    field_code_count: int
    content_control_count: int
    comment_anchor_count: int
    revisions: RevisionInventory

    def public_dict(self) -> dict[str, object]:
        return {
            "kind": self.kind,
            "block_count": len(self.block_signatures),
            "paragraph_count": self.paragraph_count,
            "table_count": self.table_count,
            "text_run_count": self.text_run_count,
            "hidden_text_run_count": self.hidden_text_run_count,
            "hidden_paragraph_mark_count": self.hidden_paragraph_mark_count,
            "alternative_format_import_anchor_count": (
                self.alternative_format_import_anchor_count
            ),
            "field_code_count": self.field_code_count,
            "content_control_count": self.content_control_count,
            "comment_anchor_count": self.comment_anchor_count,
            "revisions": self.revisions.public_dict(),
        }


@dataclass(frozen=True)
class RelationshipInventory:
    """Private relationship signature with safe public aggregate counts."""

    relationship_count: int
    external_count: int
    relationship_signature: str
    external_signature: str

    def public_dict(self) -> dict[str, int]:
        return {
            "relationship_count": self.relationship_count,
            "external_relationship_count": self.external_count,
        }


@dataclass(frozen=True)
class DocumentSnapshot:
    """One bounded private Word OOXML document or template snapshot."""

    format: str
    package_member_count: int
    stories: tuple[StorySnapshot, ...]
    relationships: RelationshipInventory
    styles: StyleInventory
    embedded_objects: EmbeddedObjectInventory
    alternative_format_imports: AlternativeFormatImportInventory
    document_properties: DocumentPropertyInventory
    sensitivity_labels: SensitivityLabelInventory
    package_digital_signatures: PackageDigitalSignatureInventory
    word_protection: WordProtectionInventory
    word_document_variables: WordDocumentVariableInventory
    word_document_variable_fields: WordDocumentVariableFieldInventory
    word_hyperlink_fields: WordHyperlinkFieldInventory
    word_hyperlink_markup: WordHyperlinkMarkupInventory
    word_drawing_hyperlinks: WordDrawingHyperlinkInventory
    word_drawing_linked_pictures: WordDrawingLinkedPictureInventory
    word_vml_hyperlinks: WordVmlHyperlinkInventory
    word_vml_external_images: WordVmlExternalImageInventory
    word_vml_image_hyperlinks: WordVmlImageHyperlinkInventory
    word_vml_linked_ole_objects: WordVmlLinkedOleObjectInventory
    word_vml_embedded_ole_objects: WordVmlEmbeddedOleObjectInventory
    word_object_links: WordObjectLinkInventory
    word_embedded_controls: WordEmbeddedControlInventory
    word_permission_ranges: WordPermissionRangeInventory
    mail_merge: MailMergeInventory
    save_through_xslt: SaveThroughXsltInventory
    attached_custom_xml_schemas: AttachedCustomXmlSchemaInventory
    field_updates_on_open: FieldUpdateOnOpenInventory
    template_style_updates_on_open: TemplateStyleUpdateOnOpenInventory
    data_bindings: DataBindingInventory
    external_fields: ExternalFieldInventory
    modern_comment_metadata: ModernCommentMetadataInventory
    document_tasks: DocumentTaskInventory
    taskpane_web_extensions: TaskpaneWebExtensionInventory
    external_document_dependencies: ExternalDocumentDependencyInventory
    track_revisions_enabled: bool
    comment_count: int
    custom_xml_part_count: int
    custom_xml_signature: str
    macro_present: bool
    macro_signature: str
    settings_signature: str
    unclassified_part_count: int
    unclassified_signature: str

    @property
    def revisions(self) -> RevisionInventory:
        return RevisionInventory(
            insertions=sum(story.revisions.insertions for story in self.stories),
            deletions=sum(story.revisions.deletions for story in self.stories),
            move_from=sum(story.revisions.move_from for story in self.stories),
            move_to=sum(story.revisions.move_to for story in self.stories),
            property_changes=sum(
                story.revisions.property_changes for story in self.stories
            ),
        )

    @property
    def hidden_text_run_count(self) -> int:
        return sum(story.hidden_text_run_count for story in self.stories)

    @property
    def hidden_paragraph_mark_count(self) -> int:
        return sum(story.hidden_paragraph_mark_count for story in self.stories)

    @property
    def alternative_format_import_anchor_count(self) -> int:
        return sum(
            story.alternative_format_import_anchor_count for story in self.stories
        )

    @property
    def field_code_count(self) -> int:
        return sum(story.field_code_count for story in self.stories)

    @property
    def content_control_count(self) -> int:
        return sum(story.content_control_count for story in self.stories)

    @property
    def story_kind_counts(self) -> dict[str, int]:
        counts: dict[str, int] = {}
        for story in self.stories:
            counts[story.kind] = counts.get(story.kind, 0) + 1
        return dict(sorted(counts.items()))

    def public_dict(self) -> dict[str, object]:
        return {
            "format": self.format,
            "package_member_count": self.package_member_count,
            "story_count": len(self.stories),
            "story_kind_counts": self.story_kind_counts,
            "paragraph_count": sum(story.paragraph_count for story in self.stories),
            "table_count": sum(story.table_count for story in self.stories),
            "text_run_count": sum(story.text_run_count for story in self.stories),
            "hidden_text_run_count": self.hidden_text_run_count,
            "hidden_paragraph_mark_count": self.hidden_paragraph_mark_count,
            "alternative_format_import_anchor_count": (
                self.alternative_format_import_anchor_count
            ),
            "field_code_count": self.field_code_count,
            "content_control_count": self.content_control_count,
            "comment_anchor_count": sum(
                story.comment_anchor_count for story in self.stories
            ),
            "comment_count": self.comment_count,
            "track_revisions_enabled": self.track_revisions_enabled,
            "revisions": self.revisions.public_dict(),
            "relationships": self.relationships.public_dict(),
            "styles": self.styles.public_dict(),
            "embedded_objects": self.embedded_objects.public_dict(),
            "alternative_format_imports": self.alternative_format_imports.public_dict(),
            "document_properties": self.document_properties.public_dict(),
            "sensitivity_labels": self.sensitivity_labels.public_dict(),
            "package_digital_signatures": (
                self.package_digital_signatures.public_dict()
            ),
            "word_protection": self.word_protection.public_dict(),
            "word_document_variables": (
                self.word_document_variables.public_dict()
            ),
            "word_document_variable_fields": (
                self.word_document_variable_fields.public_dict()
            ),
            "word_hyperlink_fields": self.word_hyperlink_fields.public_dict(),
            "word_hyperlink_markup": self.word_hyperlink_markup.public_dict(),
            "word_drawing_hyperlinks": self.word_drawing_hyperlinks.public_dict(),
            "word_drawing_linked_pictures": (
                self.word_drawing_linked_pictures.public_dict()
            ),
            "word_vml_hyperlinks": self.word_vml_hyperlinks.public_dict(),
            "word_vml_external_images": (
                self.word_vml_external_images.public_dict()
            ),
            "word_vml_image_hyperlinks": (
                self.word_vml_image_hyperlinks.public_dict()
            ),
            "word_vml_linked_ole_objects": (
                self.word_vml_linked_ole_objects.public_dict()
            ),
            "word_vml_embedded_ole_objects": (
                self.word_vml_embedded_ole_objects.public_dict()
            ),
            "word_object_links": self.word_object_links.public_dict(),
            "word_embedded_controls": self.word_embedded_controls.public_dict(),
            "word_permission_ranges": self.word_permission_ranges.public_dict(),
            "mail_merge": self.mail_merge.public_dict(),
            "save_through_xslt": self.save_through_xslt.public_dict(),
            "attached_custom_xml_schemas": (
                self.attached_custom_xml_schemas.public_dict()
            ),
            "field_updates_on_open": self.field_updates_on_open.public_dict(),
            "template_style_updates_on_open": (
                self.template_style_updates_on_open.public_dict()
            ),
            "data_bindings": self.data_bindings.public_dict(),
            "external_fields": self.external_fields.public_dict(),
            "modern_comment_metadata": self.modern_comment_metadata.public_dict(),
            "document_tasks": self.document_tasks.public_dict(),
            "taskpane_web_extensions": self.taskpane_web_extensions.public_dict(),
            "external_document_dependencies": (
                self.external_document_dependencies.public_dict()
            ),
            "custom_xml_part_count": self.custom_xml_part_count,
            "macro_present": self.macro_present,
            "unclassified_part_count": self.unclassified_part_count,
        }


@dataclass(frozen=True)
class Change:
    """One public-safe semantic document change."""

    kind: str
    message: str
    details: dict[str, object]

    def public_dict(self) -> dict[str, object]:
        return {
            "kind": self.kind,
            "message": self.message,
            "details": self.details,
        }


@dataclass(frozen=True)
class Finding:
    """A policy finding suitable for a CI gate."""

    rule_id: str
    severity: Severity
    message: str
    details: dict[str, object]

    def public_dict(self) -> dict[str, object]:
        return {
            "rule_id": self.rule_id,
            "severity": self.severity,
            "message": self.message,
            "details": self.details,
        }


@dataclass(frozen=True)
class DiffReport:
    """A private diff with explicitly public-safe changes and findings."""

    before: DocumentSnapshot
    after: DocumentSnapshot
    changes: tuple[Change, ...]
    findings: tuple[Finding, ...] = ()

    def public_dict(self) -> dict[str, object]:
        return {
            "schema_version": 1,
            "before": self.before.public_dict(),
            "after": self.after.public_dict(),
            "changes": [change.public_dict() for change in self.changes],
            "findings": [finding.public_dict() for finding in self.findings],
        }
