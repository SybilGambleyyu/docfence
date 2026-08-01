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
    mail_merge: MailMergeInventory
    data_bindings: DataBindingInventory
    external_fields: ExternalFieldInventory
    modern_comment_metadata: ModernCommentMetadataInventory
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
            "mail_merge": self.mail_merge.public_dict(),
            "data_bindings": self.data_bindings.public_dict(),
            "external_fields": self.external_fields.public_dict(),
            "modern_comment_metadata": self.modern_comment_metadata.public_dict(),
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
