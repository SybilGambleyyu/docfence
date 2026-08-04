"""Privacy-safe semantic comparisons for DocFence snapshots."""

from __future__ import annotations

from collections import Counter
from difflib import SequenceMatcher
from pathlib import Path

from docfence.models import Change, DiffReport, DocumentSnapshot, StorySnapshot
from docfence.snapshot import load_snapshot

_MAX_EXACT_BLOCKS = 20_000
_MAX_SEQUENCE_CANDIDATES = 1_000_000


def diff_documents(
    before: str | Path | DocumentSnapshot,
    after: str | Path | DocumentSnapshot,
) -> DiffReport:
    """Compare two document inputs without including their material in a report."""

    before_snapshot = _as_snapshot(before)
    after_snapshot = _as_snapshot(after)
    changes: list[Change] = []

    if before_snapshot.format != after_snapshot.format:
        changes.append(
            Change(
                kind="document_format_changed",
                message="The stored document format changed.",
                details={
                    "before_format": before_snapshot.format,
                    "after_format": after_snapshot.format,
                },
            )
        )

    matched, removed, added = _match_stories(before_snapshot, after_snapshot)
    for before_story, after_story in matched:
        change = _content_change(before_story, after_story)
        if change is not None:
            changes.append(change)
        elif before_story.structure_signature != after_story.structure_signature:
            changes.append(
                Change(
                    kind="story_structure_changed",
                    message="Stored non-block structure changed in a document story.",
                    details={
                        "story_kind": before_story.kind,
                        "before_block_count": len(before_story.block_signatures),
                        "after_block_count": len(after_story.block_signatures),
                    },
                )
            )

    for story in removed:
        changes.append(
            Change(
                kind="story_removed",
                message="A document story was removed.",
                details=_story_size_details(story),
            )
        )
    for story in added:
        changes.append(
            Change(
                kind="story_added",
                message="A document story was added.",
                details=_story_size_details(story),
            )
        )

    _append_review_inventory_changes(changes, before_snapshot, after_snapshot)
    _append_package_changes(changes, before_snapshot, after_snapshot)
    return DiffReport(
        before=before_snapshot, after=after_snapshot, changes=tuple(changes)
    )


def _as_snapshot(value: str | Path | DocumentSnapshot) -> DocumentSnapshot:
    if isinstance(value, DocumentSnapshot):
        return value
    return load_snapshot(value)


def _match_stories(
    before: DocumentSnapshot, after: DocumentSnapshot
) -> tuple[
    list[tuple[StorySnapshot, StorySnapshot]],
    list[StorySnapshot],
    list[StorySnapshot],
]:
    before_by_key = {story.part_key: story for story in before.stories}
    after_by_key = {story.part_key: story for story in after.stories}
    matched: list[tuple[StorySnapshot, StorySnapshot]] = []
    unmatched_before: list[StorySnapshot] = []
    unmatched_after: list[StorySnapshot] = []

    for part_key in sorted(set(before_by_key) | set(after_by_key)):
        before_story = before_by_key.get(part_key)
        after_story = after_by_key.get(part_key)
        if (
            before_story is not None
            and after_story is not None
            and before_story.kind == after_story.kind
        ):
            matched.append((before_story, after_story))
        elif before_story is not None:
            unmatched_before.append(before_story)
            if after_story is not None:
                unmatched_after.append(after_story)
        elif after_story is not None:
            unmatched_after.append(after_story)

    by_exact_after: dict[tuple[str, str], list[StorySnapshot]] = {}
    for story in unmatched_after:
        by_exact_after.setdefault((story.kind, story.structure_signature), []).append(
            story
        )
    still_unmatched_before: list[StorySnapshot] = []
    matched_after_ids: set[int] = set()
    for story in unmatched_before:
        candidates = by_exact_after.get((story.kind, story.structure_signature), [])
        if candidates:
            counterpart = candidates.pop(0)
            matched.append((story, counterpart))
            matched_after_ids.add(id(counterpart))
        else:
            still_unmatched_before.append(story)
    still_unmatched_after = [
        story for story in unmatched_after if id(story) not in matched_after_ids
    ]

    remaining_before: list[StorySnapshot] = []
    remaining_after: list[StorySnapshot] = []
    kinds = {story.kind for story in still_unmatched_before} | {
        story.kind for story in still_unmatched_after
    }
    for kind in sorted(kinds):
        before_group = sorted(
            (story for story in still_unmatched_before if story.kind == kind),
            key=lambda story: story.part_key,
        )
        after_group = sorted(
            (story for story in still_unmatched_after if story.kind == kind),
            key=lambda story: story.part_key,
        )
        paired_count = min(len(before_group), len(after_group))
        matched.extend(
            zip(
                before_group[:paired_count],
                after_group[:paired_count],
                strict=True,
            )
        )
        remaining_before.extend(before_group[paired_count:])
        remaining_after.extend(after_group[paired_count:])

    return matched, remaining_before, remaining_after


def _content_change(before: StorySnapshot, after: StorySnapshot) -> Change | None:
    if before.block_signatures == after.block_signatures:
        return None
    before_count = len(before.block_signatures)
    after_count = len(after.block_signatures)
    details: dict[str, object] = {
        "story_kind": before.kind,
        "before_block_count": before_count,
        "after_block_count": after_count,
    }
    if _requires_bounded_summary(before.block_signatures, after.block_signatures):
        details["comparison_mode"] = "bounded_summary"
        return Change(
            kind="content_changed",
            message="Stored content blocks changed in a document story.",
            details=details,
        )

    added = 0
    removed = 0
    replaced_before = 0
    replaced_after = 0
    matcher = SequenceMatcher(
        a=before.block_signatures, b=after.block_signatures, autojunk=False
    )
    for (
        opcode,
        before_start,
        before_end,
        after_start,
        after_end,
    ) in matcher.get_opcodes():
        if opcode == "insert":
            added += after_end - after_start
        elif opcode == "delete":
            removed += before_end - before_start
        elif opcode == "replace":
            replaced_before += before_end - before_start
            replaced_after += after_end - after_start
    details.update(
        {
            "added_block_count": added,
            "removed_block_count": removed,
            "replaced_before_block_count": replaced_before,
            "replaced_after_block_count": replaced_after,
            "comparison_mode": "sequence",
        }
    )
    return Change(
        kind="content_changed",
        message="Stored content blocks changed in a document story.",
        details=details,
    )


def _requires_bounded_summary(
    before_blocks: tuple[str, ...], after_blocks: tuple[str, ...]
) -> bool:
    if len(before_blocks) + len(after_blocks) > _MAX_EXACT_BLOCKS:
        return True
    before_counts = Counter(before_blocks)
    after_counts = Counter(after_blocks)
    candidates = sum(
        count * after_counts.get(signature, 0)
        for signature, count in before_counts.items()
    )
    return candidates > _MAX_SEQUENCE_CANDIDATES


def _story_size_details(story: StorySnapshot) -> dict[str, object]:
    return {
        "story_kind": story.kind,
        "block_count": len(story.block_signatures),
        "paragraph_count": story.paragraph_count,
        "table_count": story.table_count,
    }


def _append_review_inventory_changes(
    changes: list[Change], before: DocumentSnapshot, after: DocumentSnapshot
) -> None:
    if before.revisions != after.revisions:
        changes.append(
            Change(
                kind="revision_inventory_changed",
                message="Stored revision markup changed.",
                details={
                    "before": before.revisions.public_dict(),
                    "after": after.revisions.public_dict(),
                },
            )
        )
    _append_count_change(
        changes,
        "hidden_text_inventory_changed",
        "Stored hidden-text run inventory changed.",
        "hidden_text_run_count",
        before.hidden_text_run_count,
        after.hidden_text_run_count,
    )
    _append_count_change(
        changes,
        "hidden_paragraph_mark_inventory_changed",
        "Stored hidden paragraph-mark inventory changed.",
        "hidden_paragraph_mark_count",
        before.hidden_paragraph_mark_count,
        after.hidden_paragraph_mark_count,
    )
    _append_count_change(
        changes,
        "alternative_format_import_anchor_inventory_changed",
        "Stored alternative-format import anchor inventory changed.",
        "alternative_format_import_anchor_count",
        before.alternative_format_import_anchor_count,
        after.alternative_format_import_anchor_count,
    )
    _append_count_change(
        changes,
        "field_code_inventory_changed",
        "Stored field-code inventory changed.",
        "field_code_count",
        before.field_code_count,
        after.field_code_count,
    )
    _append_count_change(
        changes,
        "content_control_inventory_changed",
        "Stored content-control inventory changed.",
        "content_control_count",
        before.content_control_count,
        after.content_control_count,
    )
    _append_count_change(
        changes,
        "comment_anchor_inventory_changed",
        "Stored comment-anchor inventory changed.",
        "comment_anchor_count",
        sum(story.comment_anchor_count for story in before.stories),
        sum(story.comment_anchor_count for story in after.stories),
    )
    _append_count_change(
        changes,
        "comment_inventory_changed",
        "Stored comment inventory changed.",
        "comment_count",
        before.comment_count,
        after.comment_count,
    )
    if before.track_revisions_enabled != after.track_revisions_enabled:
        changes.append(
            Change(
                kind="track_revisions_setting_changed",
                message="The stored Track Changes setting changed.",
                details={
                    "before_enabled": before.track_revisions_enabled,
                    "after_enabled": after.track_revisions_enabled,
                },
            )
        )
    if before.styles.signature != after.styles.signature:
        changes.append(
            Change(
                kind="style_inventory_changed",
                message="Stored style inventory changed.",
                details={
                    "before": before.styles.public_dict(),
                    "after": after.styles.public_dict(),
                },
            )
        )


def _append_package_changes(
    changes: list[Change], before: DocumentSnapshot, after: DocumentSnapshot
) -> None:
    before_relationships = before.relationships
    after_relationships = after.relationships
    before_external_count = before_relationships.external_count
    after_external_count = after_relationships.external_count
    before_relationship_count = before_relationships.relationship_count
    after_relationship_count = after_relationships.relationship_count
    if (
        before_relationships.external_signature
        != after_relationships.external_signature
    ):
        changes.append(
            Change(
                kind="external_relationships_changed",
                message="Stored external relationship inventory changed.",
                details={
                    "before_external_relationship_count": before_external_count,
                    "after_external_relationship_count": after_external_count,
                },
            )
        )
    if (
        before_relationships.relationship_signature
        != after_relationships.relationship_signature
        and before_relationships.external_signature
        == after_relationships.external_signature
    ):
        changes.append(
            Change(
                kind="relationships_changed",
                message="Stored package relationship inventory changed.",
                details={
                    "before_relationship_count": before_relationship_count,
                    "after_relationship_count": after_relationship_count,
                },
            )
        )
    if before.custom_xml_signature != after.custom_xml_signature:
        changes.append(
            Change(
                kind="custom_xml_changed",
                message="Stored custom XML payload changed.",
                details={
                    "before_custom_xml_part_count": before.custom_xml_part_count,
                    "after_custom_xml_part_count": after.custom_xml_part_count,
                },
            )
        )
    if before.macro_signature != after.macro_signature:
        changes.append(
            Change(
                kind="macro_payload_changed",
                message="Stored macro payload changed.",
                details={
                    "before_macro_present": before.macro_present,
                    "after_macro_present": after.macro_present,
                },
            )
        )
    if before.embedded_objects.signature != after.embedded_objects.signature:
        changes.append(
            Change(
                kind="embedded_object_inventory_changed",
                message="Stored embedded object inventory changed.",
                details={
                    "before": before.embedded_objects.public_dict(),
                    "after": after.embedded_objects.public_dict(),
                },
            )
        )
    if (
        before.alternative_format_imports.signature
        != after.alternative_format_imports.signature
    ):
        changes.append(
            Change(
                kind="alternative_format_import_inventory_changed",
                message="Stored alternative-format import inventory changed.",
                details={
                    "before": before.alternative_format_imports.public_dict(),
                    "after": after.alternative_format_imports.public_dict(),
                },
            )
        )
    if before.document_properties.signature != after.document_properties.signature:
        changes.append(
            Change(
                kind="document_property_inventory_changed",
                message="Stored document property inventory changed.",
                details={
                    "before": before.document_properties.public_dict(),
                    "after": after.document_properties.public_dict(),
                },
            )
        )
    if before.package_thumbnails.signature != after.package_thumbnails.signature:
        changes.append(
            Change(
                kind="package_thumbnail_inventory_changed",
                message="Stored OPC package thumbnail inventory changed.",
                details={
                    "before": before.package_thumbnails.public_dict(),
                    "after": after.package_thumbnails.public_dict(),
                },
            )
        )
    if before.markup_compatibility.signature != after.markup_compatibility.signature:
        changes.append(
            Change(
                kind="markup_compatibility_inventory_changed",
                message="Stored OOXML markup-compatibility inventory changed.",
                details={
                    "before": before.markup_compatibility.public_dict(),
                    "after": after.markup_compatibility.public_dict(),
                },
            )
        )
    if before.sensitivity_labels.signature != after.sensitivity_labels.signature:
        changes.append(
            Change(
                kind="sensitivity_label_inventory_changed",
                message="Stored Office sensitivity-label metadata inventory changed.",
                details={
                    "before": before.sensitivity_labels.public_dict(),
                    "after": after.sensitivity_labels.public_dict(),
                },
            )
        )
    if (
        before.package_digital_signatures.signature
        != after.package_digital_signatures.signature
    ):
        changes.append(
            Change(
                kind="package_digital_signature_inventory_changed",
                message="Stored OPC package digital-signature inventory changed.",
                details={
                    "before": before.package_digital_signatures.public_dict(),
                    "after": after.package_digital_signatures.public_dict(),
                },
            )
        )
    if before.word_protection.signature != after.word_protection.signature:
        changes.append(
            Change(
                kind="word_protection_inventory_changed",
                message="Stored Word editing and write-protection inventory changed.",
                details={
                    "before": before.word_protection.public_dict(),
                    "after": after.word_protection.public_dict(),
                },
            )
        )
    if (
        before.word_document_variables.signature
        != after.word_document_variables.signature
    ):
        changes.append(
            Change(
                kind="word_document_variable_inventory_changed",
                message="Stored Word document-variable inventory changed.",
                details={
                    "before": before.word_document_variables.public_dict(),
                    "after": after.word_document_variables.public_dict(),
                },
            )
        )
    if (
        before.word_document_variable_fields.signature
        != after.word_document_variable_fields.signature
    ):
        changes.append(
            Change(
                kind="word_document_variable_field_inventory_changed",
                message="Stored DOCVARIABLE field-reference inventory changed.",
                details={
                    "before": before.word_document_variable_fields.public_dict(),
                    "after": after.word_document_variable_fields.public_dict(),
                },
            )
        )
    if before.word_hyperlink_fields.signature != after.word_hyperlink_fields.signature:
        changes.append(
            Change(
                kind="word_hyperlink_field_inventory_changed",
                message="Stored HYPERLINK field-reference inventory changed.",
                details={
                    "before": before.word_hyperlink_fields.public_dict(),
                    "after": after.word_hyperlink_fields.public_dict(),
                },
            )
        )
    if before.word_hyperlink_markup.signature != after.word_hyperlink_markup.signature:
        changes.append(
            Change(
                kind="word_hyperlink_markup_inventory_changed",
                message="Stored WordprocessingML hyperlink markup inventory changed.",
                details={
                    "before": before.word_hyperlink_markup.public_dict(),
                    "after": after.word_hyperlink_markup.public_dict(),
                },
            )
        )
    if (
        before.word_drawing_hyperlinks.signature
        != after.word_drawing_hyperlinks.signature
    ):
        changes.append(
            Change(
                kind="word_drawing_hyperlink_inventory_changed",
                message="Stored DrawingML hyperlink-action inventory changed.",
                details={
                    "before": before.word_drawing_hyperlinks.public_dict(),
                    "after": after.word_drawing_hyperlinks.public_dict(),
                },
            )
        )
    if (
        before.word_drawing_visibility.signature
        != after.word_drawing_visibility.signature
    ):
        changes.append(
            Change(
                kind="word_drawing_visibility_inventory_changed",
                message=("Stored DrawingML nonvisual visibility inventory changed."),
                details={
                    "before": before.word_drawing_visibility.public_dict(),
                    "after": after.word_drawing_visibility.public_dict(),
                },
            )
        )
    if (
        before.word_drawing_linked_pictures.signature
        != after.word_drawing_linked_pictures.signature
    ):
        changes.append(
            Change(
                kind="word_drawing_linked_picture_inventory_changed",
                message="Stored DrawingML linked-picture inventory changed.",
                details={
                    "before": before.word_drawing_linked_pictures.public_dict(),
                    "after": after.word_drawing_linked_pictures.public_dict(),
                },
            )
        )
    if before.word_vml_hyperlinks.signature != after.word_vml_hyperlinks.signature:
        changes.append(
            Change(
                kind="word_vml_hyperlink_inventory_changed",
                message="Stored VML hyperlink markup inventory changed.",
                details={
                    "before": before.word_vml_hyperlinks.public_dict(),
                    "after": after.word_vml_hyperlinks.public_dict(),
                },
            )
        )
    if (
        before.word_vml_external_images.signature
        != after.word_vml_external_images.signature
    ):
        changes.append(
            Change(
                kind="word_vml_external_image_inventory_changed",
                message="Stored VML external-image inventory changed.",
                details={
                    "before": before.word_vml_external_images.public_dict(),
                    "after": after.word_vml_external_images.public_dict(),
                },
            )
        )
    if (
        before.word_vml_image_hyperlinks.signature
        != after.word_vml_image_hyperlinks.signature
    ):
        changes.append(
            Change(
                kind="word_vml_image_hyperlink_inventory_changed",
                message="Stored VML image-data hyperlink inventory changed.",
                details={
                    "before": before.word_vml_image_hyperlinks.public_dict(),
                    "after": after.word_vml_image_hyperlinks.public_dict(),
                },
            )
        )
    if (
        before.word_vml_linked_ole_objects.signature
        != after.word_vml_linked_ole_objects.signature
    ):
        changes.append(
            Change(
                kind="word_vml_linked_ole_object_inventory_changed",
                message="Stored VML linked-OLE-object inventory changed.",
                details={
                    "before": before.word_vml_linked_ole_objects.public_dict(),
                    "after": after.word_vml_linked_ole_objects.public_dict(),
                },
            )
        )
    if (
        before.word_vml_embedded_ole_objects.signature
        != after.word_vml_embedded_ole_objects.signature
    ):
        changes.append(
            Change(
                kind="word_vml_embedded_ole_object_inventory_changed",
                message="Stored VML embedded-OLE-object inventory changed.",
                details={
                    "before": before.word_vml_embedded_ole_objects.public_dict(),
                    "after": after.word_vml_embedded_ole_objects.public_dict(),
                },
            )
        )
    if before.word_object_links.signature != after.word_object_links.signature:
        changes.append(
            Change(
                kind="word_object_link_inventory_changed",
                message=(
                    "Stored WordprocessingML linked-object-property inventory changed."
                ),
                details={
                    "before": before.word_object_links.public_dict(),
                    "after": after.word_object_links.public_dict(),
                },
            )
        )
    if (
        before.word_embedded_controls.signature
        != after.word_embedded_controls.signature
    ):
        changes.append(
            Change(
                kind="word_embedded_control_inventory_changed",
                message=(
                    "Stored WordprocessingML embedded-control-anchor inventory changed."
                ),
                details={
                    "before": before.word_embedded_controls.public_dict(),
                    "after": after.word_embedded_controls.public_dict(),
                },
            )
        )
    if (
        before.word_permission_ranges.signature
        != after.word_permission_ranges.signature
    ):
        changes.append(
            Change(
                kind="word_permission_range_inventory_changed",
                message="Stored Word editable-range permission inventory changed.",
                details={
                    "before": before.word_permission_ranges.public_dict(),
                    "after": after.word_permission_ranges.public_dict(),
                },
            )
        )
    if before.mail_merge.signature != after.mail_merge.signature:
        changes.append(
            Change(
                kind="mail_merge_inventory_changed",
                message="Stored mail-merge inventory changed.",
                details={
                    "before": before.mail_merge.public_dict(),
                    "after": after.mail_merge.public_dict(),
                },
            )
        )
    if before.save_through_xslt.signature != after.save_through_xslt.signature:
        changes.append(
            Change(
                kind="save_through_xslt_inventory_changed",
                message="Stored XSLT-on-single-XML-save inventory changed.",
                details={
                    "before": before.save_through_xslt.public_dict(),
                    "after": after.save_through_xslt.public_dict(),
                },
            )
        )
    if (
        before.attached_custom_xml_schemas.signature
        != after.attached_custom_xml_schemas.signature
    ):
        changes.append(
            Change(
                kind="attached_custom_xml_schema_inventory_changed",
                message="Attached custom XML schema inventory changed.",
                details={
                    "before": before.attached_custom_xml_schemas.public_dict(),
                    "after": after.attached_custom_xml_schemas.public_dict(),
                },
            )
        )
    if before.field_updates_on_open.signature != after.field_updates_on_open.signature:
        changes.append(
            Change(
                kind="field_update_on_open_inventory_changed",
                message="Stored automatic field-update-on-open inventory changed.",
                details={
                    "before": before.field_updates_on_open.public_dict(),
                    "after": after.field_updates_on_open.public_dict(),
                },
            )
        )
    if (
        before.template_style_updates_on_open.signature
        != after.template_style_updates_on_open.signature
    ):
        changes.append(
            Change(
                kind="template_style_update_on_open_inventory_changed",
                message=(
                    "Stored automatic template-style-update-on-open inventory changed."
                ),
                details={
                    "before": before.template_style_updates_on_open.public_dict(),
                    "after": after.template_style_updates_on_open.public_dict(),
                },
            )
        )
    if (
        before.personal_information_removal_on_save.signature
        != after.personal_information_removal_on_save.signature
    ):
        changes.append(
            Change(
                kind="personal_information_removal_on_save_inventory_changed",
                message=(
                    "Stored personal-information-removal-on-save inventory changed."
                ),
                details={
                    "before": before.personal_information_removal_on_save.public_dict(),
                    "after": after.personal_information_removal_on_save.public_dict(),
                },
            )
        )
    if before.data_bindings.signature != after.data_bindings.signature:
        changes.append(
            Change(
                kind="data_binding_inventory_changed",
                message="Stored content-control data-binding inventory changed.",
                details={
                    "before": before.data_bindings.public_dict(),
                    "after": after.data_bindings.public_dict(),
                },
            )
        )
    if before.external_fields.signature != after.external_fields.signature:
        changes.append(
            Change(
                kind="external_field_inventory_changed",
                message="Stored external-source Word field inventory changed.",
                details={
                    "before": before.external_fields.public_dict(),
                    "after": after.external_fields.public_dict(),
                },
            )
        )
    if (
        before.modern_comment_metadata.signature
        != after.modern_comment_metadata.signature
    ):
        changes.append(
            Change(
                kind="modern_comment_metadata_inventory_changed",
                message="Stored modern Word comment metadata inventory changed.",
                details={
                    "before": before.modern_comment_metadata.public_dict(),
                    "after": after.modern_comment_metadata.public_dict(),
                },
            )
        )
    if before.document_tasks.signature != after.document_tasks.signature:
        changes.append(
            Change(
                kind="document_task_inventory_changed",
                message="Stored Word document task inventory changed.",
                details={
                    "before": before.document_tasks.public_dict(),
                    "after": after.document_tasks.public_dict(),
                },
            )
        )
    if (
        before.taskpane_web_extensions.signature
        != after.taskpane_web_extensions.signature
    ):
        changes.append(
            Change(
                kind="taskpane_web_extension_inventory_changed",
                message="Stored task-pane Office web extension inventory changed.",
                details={
                    "before": before.taskpane_web_extensions.public_dict(),
                    "after": after.taskpane_web_extensions.public_dict(),
                },
            )
        )
    if (
        before.external_document_dependencies.signature
        != after.external_document_dependencies.signature
    ):
        changes.append(
            Change(
                kind="external_document_dependency_inventory_changed",
                message="Stored external Word document dependency inventory changed.",
                details={
                    "before": before.external_document_dependencies.public_dict(),
                    "after": after.external_document_dependencies.public_dict(),
                },
            )
        )
    if (
        before.settings_signature != after.settings_signature
        and before.track_revisions_enabled == after.track_revisions_enabled
    ):
        changes.append(
            Change(
                kind="document_settings_changed",
                message="Stored document settings changed.",
                details={},
            )
        )
    if before.unclassified_signature != after.unclassified_signature:
        changes.append(
            Change(
                kind="unclassified_package_payload_changed",
                message=(
                    "Stored package payload outside the specialized inventories "
                    "changed."
                ),
                details={
                    "before_unclassified_part_count": before.unclassified_part_count,
                    "after_unclassified_part_count": after.unclassified_part_count,
                },
            )
        )


def _append_count_change(
    changes: list[Change],
    kind: str,
    message: str,
    count_name: str,
    before: int,
    after: int,
) -> None:
    if before != after:
        changes.append(
            Change(
                kind=kind,
                message=message,
                details={f"before_{count_name}": before, f"after_{count_name}": after},
            )
        )
