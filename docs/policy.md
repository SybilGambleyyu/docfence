# Policy reference

DocFence policies are an intentionally narrow subset of YAML. A policy contains
exactly `version: 1` and a `rules` mapping. Rule values are booleans; absent
rules are disabled.

```yaml
version: 1
rules:
  no_external_relationship_changes: true
  no_macro_payload_changes: true
  require_no_comments: true
```

JSON with the same shape is also accepted. YAML must use spaces, a two-space
`rules` indentation, and the literal values `true` or `false`. Anchors,
aliases, sequences, quoted scalars, nested mappings, duplicate keys, and
unknown rules fail closed. `docfence init path.yml` writes the conservative
starter policy.

## Rule catalog

| Rule | SARIF ID | Trigger | Scope |
| --- | --- | --- | --- |
| `no_external_relationship_changes` | `DFP001` | External relationship signature differs | Comparison |
| `no_macro_payload_changes` | `DFP002` | Macro payload signature differs | Comparison |
| `no_custom_xml_changes` | `DFP003` | Custom XML payload signature differs | Comparison |
| `require_no_unresolved_revisions` | `DFP004` | Candidate has stored revision markup | Candidate |
| `require_no_comments` | `DFP005` | Candidate has stored comments | Candidate |
| `require_no_hidden_text` | `DFP006` | Candidate has direct stored hidden-text runs | Candidate |
| `no_relationship_changes` | `DFP007` | Any stored package relationship differs | Comparison |
| `no_document_settings_changes` | `DFP008` | Stored document settings signature differs | Comparison |
| `no_unclassified_package_payload_changes` | `DFP009` | Payload outside specialized inventories differs | Comparison |
| `require_track_revisions_disabled` | `DFP010` | Candidate enables Track Changes | Candidate |
| `require_no_field_codes` | `DFP011` | Candidate has stored field-code markup | Candidate |
| `require_no_content_controls` | `DFP012` | Candidate has stored content controls | Candidate |
| `require_no_hidden_text_style_declarations` | `DFP013` | Candidate has stored style/default declarations that can hide text | Candidate |
| `require_no_hidden_paragraph_marks` | `DFP014` | Candidate has direct hidden paragraph-mark markup | Candidate |
| `require_no_embedded_objects` | `DFP015` | Candidate has stored embedded OLE/package/control relationships or payload parts | Candidate |
| `require_no_alternative_format_imports` | `DFP016` | Candidate has stored `aFChunk` import relationships, payloads, or `w:altChunk` anchors | Candidate |
| `no_embedded_object_payload_changes` | `DFP017` | Embedded OLE/package/control inventory differs | Comparison |
| `no_alternative_format_import_changes` | `DFP018` | Alternative-format import inventory or anchor count differs | Comparison |
| `no_document_property_changes` | `DFP019` | Core, extended, or custom document-property inventory differs | Comparison |
| `require_no_custom_document_properties` | `DFP020` | Candidate has stored custom document-property definitions | Candidate |
| `require_no_mail_merge` | `DFP021` | Candidate has stored mail-merge configuration, source, or recipient-data state | Candidate |
| `no_mail_merge_changes` | `DFP022` | Mail-merge inventory differs | Comparison |
| `require_no_data_bindings` | `DFP023` | Candidate has stored content-control data bindings | Candidate |
| `no_data_binding_changes` | `DFP024` | Content-control data-binding inventory differs | Comparison |
| `require_no_external_document_dependencies` | `DFP025` | Candidate has stored attached-template, subdocument, or frameset-source dependency state | Candidate |
| `no_external_document_dependency_changes` | `DFP026` | External Word document dependency inventory differs | Comparison |
| `require_no_external_fields` | `DFP027` | Candidate has a recognized external-source Word field instruction | Candidate |
| `no_external_field_changes` | `DFP028` | External-source Word field inventory differs | Comparison |
| `require_no_modern_comment_metadata` | `DFP029` | Candidate has stored modern Word comment contact, threading, identifier, or reaction metadata | Candidate |
| `no_modern_comment_metadata_changes` | `DFP030` | Modern Word comment metadata inventory differs | Comparison |
| `require_no_document_tasks` | `DFP031` | Candidate has stored Word document-task workflow state | Candidate |
| `no_document_task_changes` | `DFP032` | Word document-task inventory differs | Comparison |
| `require_no_taskpane_web_extensions` | `DFP033` | Candidate has stored task-pane Office web-extension state | Candidate |
| `no_taskpane_web_extension_changes` | `DFP034` | Task-pane Office web-extension inventory differs | Comparison |
| `require_no_sensitivity_label_metadata` | `DFP035` | Candidate has stored Office sensitivity-label metadata | Candidate |
| `no_sensitivity_label_metadata_changes` | `DFP036` | Office sensitivity-label metadata inventory differs | Comparison |
| `require_no_package_digital_signatures` | `DFP037` | Candidate has stored OPC package digital-signature material | Candidate |
| `no_package_digital_signature_changes` | `DFP038` | OPC package digital-signature inventory differs | Comparison |
| `require_no_word_protection` | `DFP039` | Candidate has stored Word editing or write-protection state | Candidate |
| `no_word_protection_changes` | `DFP040` | Word editing/write-protection inventory differs | Comparison |
| `require_no_word_permission_ranges` | `DFP041` | Candidate has stored Word editable-range permission markup | Candidate |
| `no_word_permission_range_changes` | `DFP042` | Word editable-range permission inventory differs | Comparison |
| `require_no_word_document_variables` | `DFP043` | Candidate has stored Word document-variable state | Candidate |
| `no_word_document_variable_changes` | `DFP044` | Word document-variable inventory differs | Comparison |
| `require_no_word_document_variable_fields` | `DFP045` | Candidate has stored `DOCVARIABLE` field references | Candidate |
| `no_word_document_variable_field_changes` | `DFP046` | Word `DOCVARIABLE` field-reference inventory differs | Comparison |
| `require_no_word_hyperlink_fields` | `DFP047` | Candidate has stored `HYPERLINK` field references | Candidate |
| `no_word_hyperlink_field_changes` | `DFP048` | Word `HYPERLINK` field-reference inventory differs | Comparison |
| `require_no_word_hyperlink_markup` | `DFP049` | Candidate has direct WordprocessingML `w:hyperlink` markup | Candidate |
| `no_word_hyperlink_markup_changes` | `DFP050` | WordprocessingML hyperlink-markup inventory differs | Comparison |
| `require_no_word_drawing_hyperlinks` | `DFP051` | Candidate has direct DrawingML hyperlink-action markup | Candidate |
| `no_word_drawing_hyperlink_changes` | `DFP052` | DrawingML hyperlink-action inventory differs | Comparison |
| `require_no_word_vml_hyperlinks` | `DFP053` | Candidate has direct legacy VML shape-link markup | Candidate |
| `no_word_vml_hyperlink_changes` | `DFP054` | VML hyperlink-markup inventory differs | Comparison |
| `require_no_word_drawing_linked_pictures` | `DFP055` | Candidate has direct DrawingML linked-picture markup | Candidate |
| `no_word_drawing_linked_picture_changes` | `DFP056` | DrawingML linked-picture inventory differs | Comparison |
| `require_no_word_vml_external_images` | `DFP057` | Candidate has direct VML external-image markup | Candidate |
| `no_word_vml_external_image_changes` | `DFP058` | VML external-image inventory differs | Comparison |
| `require_no_word_vml_image_hyperlinks` | `DFP059` | Candidate has direct VML image-data hyperlink markup | Candidate |
| `no_word_vml_image_hyperlink_changes` | `DFP060` | VML image-data hyperlink inventory differs | Comparison |
| `require_no_word_vml_linked_ole_objects` | `DFP061` | Candidate has direct VML linked-OLE markup | Candidate |
| `no_word_vml_linked_ole_object_changes` | `DFP062` | VML linked-OLE-object inventory differs | Comparison |
| `require_no_word_object_links` | `DFP063` | Candidate has direct WordprocessingML linked-object-property markup | Candidate |
| `no_word_object_link_changes` | `DFP064` | WordprocessingML linked-object-property inventory differs | Comparison |
| `require_no_word_embedded_controls` | `DFP065` | Candidate has direct WordprocessingML embedded-control anchors | Candidate |
| `no_word_embedded_control_changes` | `DFP066` | WordprocessingML embedded-control-anchor inventory differs | Comparison |
| `require_no_word_vml_embedded_ole_objects` | `DFP067` | Candidate has direct VML embedded-OLE markup | Candidate |
| `no_word_vml_embedded_ole_object_changes` | `DFP068` | VML embedded-OLE-object inventory differs | Comparison |
| `require_no_save_through_xslt` | `DFP069` | Candidate has stored XSLT-on-single-XML-save configuration | Candidate |
| `no_save_through_xslt_changes` | `DFP070` | XSLT-on-single-XML-save inventory differs | Comparison |
| `require_no_attached_custom_xml_schemas` | `DFP071` | Candidate has attached custom XML schema declarations | Candidate |
| `no_attached_custom_xml_schema_changes` | `DFP072` | Attached custom XML schema inventory differs | Comparison |
| `require_no_field_updates_on_open` | `DFP073` | Candidate requests automatic field recalculation on open | Candidate |
| `no_field_update_on_open_changes` | `DFP074` | Automatic field-update-on-open inventory differs | Comparison |
| `require_no_template_style_updates_on_open` | `DFP075` | Candidate enables automatic template-style updates on open | Candidate |
| `no_template_style_update_on_open_changes` | `DFP076` | Automatic template-style-update-on-open inventory differs | Comparison |
| `require_personal_information_removal_on_save` | `DFP077` | Candidate does not store an enabled personal-information-removal-on-save request | Candidate |
| `no_personal_information_removal_on_save_changes` | `DFP078` | Personal-information-removal-on-save inventory differs | Comparison |
| `require_no_custom_xml_data` | `DFP079` | Candidate has stored conventional custom-XML package parts | Candidate |
| `require_no_package_thumbnails` | `DFP080` | Candidate has relationship-bound OPC package thumbnail images | Candidate |
| `no_package_thumbnail_changes` | `DFP081` | OPC package thumbnail inventory differs | Comparison |
| `require_no_markup_compatibility` | `DFP082` | Candidate has stored OOXML Markup Compatibility markup | Candidate |
| `no_markup_compatibility_changes` | `DFP083` | OOXML Markup Compatibility inventory differs | Comparison |
| `require_no_hidden_drawing_objects` | `DFP084` | Candidate has a stored hidden or invalid supported DrawingML nonvisual declaration | Candidate |
| `no_drawing_object_visibility_changes` | `DFP085` | DrawingML nonvisual visibility inventory differs | Comparison |
| `require_no_save_forms_data` | `DFP086` | Candidate requests form-data-only saving | Candidate |
| `no_save_forms_data_changes` | `DFP087` | Form-data-only-save inventory differs | Comparison |
| `require_no_save_preview_picture` | `DFP088` | Candidate requests preview-thumbnail generation on save | Candidate |
| `no_save_preview_picture_changes` | `DFP089` | Preview-thumbnail-on-save inventory differs | Comparison |
| `require_content_control_locks` | `DFP090` | Candidate has a content control without a direct non-`unlocked` lock declaration | Candidate |
| `no_content_control_lock_changes` | `DFP091` | Content-control lock inventory differs | Comparison |
| `require_complete_package_signature_coverage` | `DFP092` | Candidate lacks complete static declared OPC package-signature coverage in DocFence's bounded Word scope | Candidate |
| `no_package_signature_coverage_changes` | `DFP093` | Static declared OPC package-signature coverage inventory differs | Comparison |

All current findings have `high` severity except macro payload changes, which
are `critical`. SARIF deliberately contains no locations: a package member path
or paragraph reference can itself reveal information about a confidential file.

## Comparison rules vs candidate rules

Comparison rules protect a delta. For example,
`no_external_relationship_changes` permits a pre-existing external relationship
when it is unchanged. Candidate rules inspect the after document independently
of the baseline. For example, `require_no_comments` fails whenever the
candidate contains comments, even when both documents contain the same number.

Choose policies based on the handoff boundary. A publishing gate often uses the
starter policy and candidate-state rules. A controlled template workflow might
also enable `no_document_settings_changes` and
`no_unclassified_package_payload_changes`; those are intentionally stricter and
can flag a media, metadata, or other opaque package mutation. A template that
intentionally embeds a known payload can instead enable
`no_embedded_object_payload_changes` or
`no_alternative_format_import_changes` to preserve that baseline while blocking
a later mutation. `no_document_property_changes` gives the same controlled
baseline option for document metadata, and `no_mail_merge_changes` does so for
mail-merge configuration and recipient state. `no_save_through_xslt_changes`
does so for an approved custom XSLT-on-single-XML-save configuration.
`no_attached_custom_xml_schema_changes` does so for an approved set of attached
custom XML schema declarations.
`no_field_update_on_open_changes` does so for an approved automatic-field-
recalculation-on-open setting.
`no_save_forms_data_changes` does so for an approved form-data-only-save
setting, while `require_no_save_forms_data` rejects a candidate that stores an
enabled request. These rules inspect only the direct stored Settings leaf; they
do not discover or evaluate form fields, read field values, save a document,
emit a delimited record, or predict client behavior.
`no_save_preview_picture_changes` does so for an approved stored
preview-thumbnail-on-save setting, while `require_no_save_preview_picture`
rejects an enabled request. These rules inspect only the direct stored Settings
leaf: they do not prove a thumbnail is absent, decode or render an image, save
a document, generate a thumbnail, or predict client behavior. An existing
package thumbnail remains separately covered by the OPC package-thumbnail
inventory.
`no_content_control_lock_changes` protects the direct lock-declaration
baseline of a controlled template, while `require_content_control_locks`
requires each discovered direct `w:sdt` to carry one direct
non-`unlocked` `w:sdtPr/w:lock` declaration. Missing declarations stay
distinct from explicit `unlocked` values because OOXML gives omitted locks
type-specific behavior for group controls. These rules inspect stored direct
markup only: they do not identify a control, read its contents, determine a
control type, evaluate a data binding, modify a document, apply document
protection, or predict client behavior.
`no_package_thumbnail_changes` does so for a controlled template that retains
an approved relationship-bound OPC thumbnail image.
`no_markup_compatibility_changes` does so for a controlled template that
retains approved OOXML Markup Compatibility markup.
`no_data_binding_changes` does
the same for a controlled template whose content controls intentionally map to
custom XML. `no_external_document_dependency_changes` does the same for a
controlled template that deliberately retains an attached template, master
subdocument, or frameset source. `no_external_field_changes` does the same for
a controlled template that deliberately retains known external-source Word
fields. `no_modern_comment_metadata_changes` does the same for a controlled
template that intentionally retains modern comment review state.
`no_document_task_changes` and `no_taskpane_web_extension_changes` provide
the equivalent controlled-baseline gates for stored workflow and document-borne
add-in state. `no_sensitivity_label_metadata_changes` does the same for a
controlled template that intentionally retains an approved Office sensitivity
label or its related legacy metadata.
`no_package_digital_signature_changes` provides the controlled-baseline gate for
a template that intentionally retains approved OPC package-signature material.
`no_word_protection_changes` provides the equivalent gate when an approved
template intentionally retains Word editing/write-protection state.
`no_word_permission_range_changes` provides the equivalent gate when an
approved template intentionally retains editable-range permission markup.
`no_word_document_variable_changes` provides the equivalent gate when an
approved document or template intentionally retains document-variable state.
`no_word_document_variable_field_changes` provides the equivalent gate when it
intentionally retains stored `DOCVARIABLE` field references.
`no_word_hyperlink_field_changes` provides the equivalent gate when it
intentionally retains stored `HYPERLINK` field references.
`no_word_hyperlink_markup_changes` provides the equivalent gate when it
intentionally retains direct WordprocessingML hyperlink markup.
`no_word_drawing_hyperlink_changes` provides the equivalent gate when it
intentionally retains direct DrawingML hyperlink-action markup.
`no_drawing_object_visibility_changes` provides the equivalent gate when it
intentionally retains direct DrawingML nonvisual visibility declarations.
`no_word_drawing_linked_picture_changes` provides the equivalent gate when it
intentionally retains direct DrawingML linked-picture markup.
`no_word_vml_hyperlink_changes` provides the equivalent gate when it
intentionally retains direct legacy VML shape-link markup.
`no_word_vml_external_image_changes` provides the equivalent gate when it
intentionally retains direct legacy VML external-image markup.
`no_word_vml_image_hyperlink_changes` provides the equivalent gate when it
intentionally retains direct legacy VML image-data hyperlink markup.
`no_word_vml_linked_ole_object_changes` provides the equivalent gate when it
intentionally retains direct legacy VML linked-OLE markup.
`no_word_object_link_changes` provides the equivalent gate when it
intentionally retains direct WordprocessingML linked-object-property markup.
`no_word_embedded_control_changes` provides the equivalent gate when it
intentionally retains direct WordprocessingML embedded-control anchors.
`no_word_vml_embedded_ole_object_changes` provides the equivalent gate when it
intentionally retains direct Office VML embedded-OLE markup.
`word/styles.xml` is handled by the dedicated style inventory instead.

## Hidden-text scope

`require_no_hidden_text` covers direct `w:vanish` on ordinary runs only.
`require_no_hidden_paragraph_marks` covers direct `w:vanish` or `w:specVanish`
under `w:pPr/w:rPr`; it is separate because a paragraph mark is not a text run.

`require_no_hidden_text_style_declarations` fails when an enabled `w:vanish`
appears in a stored style's text-run properties or in document-default run
properties. Word treats hidden text as a toggle property in styles, and its
effective state also depends on style inheritance and application. Therefore,
this rule is a conservative declaration gate, not a claim that DocFence has
resolved which candidate runs Word will display as hidden. A `w:vanish` with an
explicit false value does not trigger the declaration count.

## Embedded and imported-content scope

`require_no_embedded_objects` is a candidate-state gate for OLE/package and
control payload evidence. It recognizes the standard `oleObject`, `package`,
`control`, and ActiveX-control-binary relationship types, as well as payloads
stored in the conventional `word/embeddings/` and `word/activeX/` folders. The
comparison rule fingerprints those recognized relationship semantics and the
payload bytes privately, so relationship-ID renumbering alone remains quiet.
That broader boundary still captures orphaned relationships and payloads.
`require_no_word_embedded_controls` is deliberately narrower: it inventories
only direct `w:control` anchors and complements, rather than replaces, this
package-level gate.
`require_no_word_vml_embedded_ole_objects` is another complementary direct
markup gate: it inventories only Office VML `o:OLEObject Type="Embed"` markers,
not arbitrary OLE relationships or payloads.

`require_no_alternative_format_imports` fails when a candidate has an `aFChunk`
relationship, its resolved internal payload, or a direct Word `w:altChunk`
anchor. The comparison counterpart detects a relationship/payload mutation and
an anchor-count mutation. An unanchored `aFChunk` relationship is still
inventoried because it is stored imported-content state.

For every encountered `w:altChunk`, DocFence requires a matching internal
standard `aFChunk` relationship whose target safely resolves to a stored package
member. It does not decode, import, render, execute, or assess the safety of
that payload. Findings and reports expose only aggregate counts.

## Document-property scope

DocFence inventories the standard core, extended, and custom document-property
relationship types, plus the canonical `docProps/core.xml`, `docProps/app.xml`,
and `docProps/custom.xml` paths when present. The property root must match the
expected OOXML vocabulary. Reports contain only part and aggregate-value counts;
property names, values, paths, relationship targets, and fingerprints remain
private.

Core and extended value counts cover direct property elements with stored text,
including automatically maintained timestamps, statistics, application details,
and template metadata. `no_document_property_changes` therefore detects a
normal resave that updates one of those fields. Use it when that exact metadata
delta matters at a controlled boundary, not as an assertion that every metadata
change is suspicious.

`require_no_custom_document_properties` is narrower: it fails only when the
candidate contains stored custom-property definitions. It does not classify
core or extended metadata as personal, confidential, user-authored, or safe.
Custom property names and values are never emitted.

## Custom XML data scope

DocFence's general custom-XML inventory counts each non-relationship package
member below the conventional `customXml/` folder and privately fingerprints
its stored bytes. This includes unbound custom XML: a content control or a
known `w:dataBinding` is not required for the package part to remain
review-visible. It also includes associated custom-XML properties parts.
Relationship parts below `customXml/_rels/` stay in the separate generic
relationship inventory rather than inflating the data-part count.

`require_no_custom_xml_data` is a candidate-state gate for a clean handoff. It
fails whenever that aggregate count is nonzero, including when the stored XML
is not referenced by a visible control. It does not expose XML values, classify
a part as personal or confidential, decide that a host will use it, or remove
or rewrite any part. The boundary is deliberately conventional: it does not
claim to discover arbitrary application-defined XML stored outside
`customXml/`.

## Mail-merge scope

DocFence inventories direct `w:mailMerge` settings plus recognized external
`mailMergeSource` and `mailMergeHeaderSource` relationships from
`word/settings.xml`. It also inventories recognized internal recipient-data
relationships and their stored target parts. Both documented recipient-data
relationship spellings are accepted, so conventional and Strict OOXML packages
can be compared without treating a relationship-ID rewrite as a mutation.

For an encountered direct `w:dataSource`, `w:headerSource`, ODSO `w:src`, or
ODSO `w:recipientData` reference, the relationship must have the required type
and target mode. Internal recipient targets must resolve to a stored package
part. Invalid recognized mail-merge markup fails closed. A recognized stored
source or recipient relationship is still inventoried even when no current
`w:mailMerge` element references it, because a residual relationship remains a
review surface.

`require_no_mail_merge` is a candidate-state gate: it fails if any of those
aggregate counts is nonzero. `no_mail_merge_changes` compares a private
signature of the mail-merge configuration, relationship semantics, and
recipient-data payloads. Public reports show only aggregate counts. Connection
strings, queries, table names, field mappings, source/header targets,
relationship IDs, recipient bytes, and fingerprints are never emitted.

These rules do not connect to a data source, execute a query, select a
recipient, or decide whether mail-merge state is expected, safe, personal, or
malicious. Use a clean-handoff gate when no mail merge should remain; use a
controlled baseline when an approved template legitimately retains it.

## XSLT-on-single-XML-save scope

`w:saveThroughXslt` stores a custom XSLT location for an application to use
when saving a document as a single XML file. `w:useXSLTWhenSaving` controls
whether that save-time behavior is enabled. DocFence inventories direct copies
of both settings from every discovered Document Settings part. A direct
relationship-backed anchor must resolve to a standard `transform` relationship
with `TargetMode="External"`; a local `w:solutionID`-only anchor is retained as
stored application-defined configuration without resolving it. A recognized
standard transform relationship is also inventoried when residual and
unanchored.

Public output is limited to enabled-setting, disabled-setting, anchor,
relationship, and solution-identifier counts. Transform locations, relationship IDs, solution
identifiers, Settings-part paths, and private fingerprints never leave the
process. The comparison signature includes the full recognized configuration,
so a same-count target or local-identifier rewrite remains visible while a
relationship-ID renumbering with unchanged semantics remains quiet.

`require_no_save_through_xslt` fails whenever the candidate has any stored
inventory evidence. `no_save_through_xslt_changes` protects an approved
baseline. Neither rule resolves, fetches, parses, executes, validates, or
otherwise applies an XSLT; it is not a prediction that Word will save a single
XML file or contact a transform target.

## Attached custom XML schema scope

DocFence inventories direct `w:attachedSchema` children of every discovered
Document Settings part. Each standard `CT_String` leaf must contain exactly its
required Word-namespace `w:val` attribute. The value names a custom XML schema
target namespace that a host may associate with a document when loading it if a
matching schema is available; it is not a package path or a fetch request.

Public output reports only `attached_custom_xml_schema_count`. Namespace
identifiers, Settings-part paths, and private fingerprints never leave the
process. The comparison signature retains each declaration, so a same-count
namespace rewrite remains review-visible.

`require_no_attached_custom_xml_schemas` fails whenever the candidate has at
least one declaration. `no_attached_custom_xml_schema_changes` protects an
approved baseline. Neither rule resolves, retrieves, loads, or validates
against a schema, nor does it claim that a host has the schema available or
will validate custom markup.

## Automatic field-recalculation-on-open scope

DocFence inventories direct `w:updateFields` children of every discovered
Document Settings part. The standard `CT_OnOff` leaf may omit `w:val`, which
means enabled when the element is present; an explicitly supplied
Word-namespace `w:val` must be one of `true`, `false`, `on`, `off`, `1`, or
`0`. The inventory rejects duplicate direct leaves, child markup, nonblank
text, unsupported attributes, and unsupported boolean values.

Public output contains only
`field_update_on_open_enabled_setting_count` and
`field_update_on_open_disabled_setting_count`. Settings-part paths and private
fingerprints never leave the process. The comparison signature records the
canonical enabled/disabled state: a state change is review-visible, while
equivalent enabled spellings do not add noise.

`require_no_field_updates_on_open` fails only when the candidate has an enabled
direct setting. An absent leaf or an explicitly disabled leaf does not request
automatic recalculation. `no_field_update_on_open_changes` protects any
approved stored baseline. Neither rule parses field instructions, recalculates
or updates a field, opens a document client, accesses a field source, follows a
link, starts an application, or claims that a particular client will honor the
stored request.

## Automatic template-style-update-on-open scope

DocFence inventories direct w:linkStyles children of every discovered Document
Settings part. The standard CT_OnOff leaf may omit w:val, which means enabled
when the element is present; an explicitly supplied Word-namespace w:val must
be one of true, false, on, off, 1, or 0. The inventory rejects duplicate direct
leaves, child markup, nonblank text, unsupported attributes, and unsupported
boolean values.

Microsoft's [LinkStyles reference](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.linkstyles?view=openxml-3.0.1)
describes this setting as automatically updating styles from a document
template. The inventory records the stored leaf only.

Public output contains only
template_style_update_on_open_enabled_setting_count and
template_style_update_on_open_disabled_setting_count. Settings-part paths and
private fingerprints never leave the process. The comparison signature records
the canonical enabled/disabled state: a state change is review-visible, while
equivalent enabled spellings do not add noise.

The direct setting records a request to update a document's styles from an
attached template when a capable host opens it. The external-document-dependency
inventory separately reports stored attached-template anchor and relationship
evidence. This setting inventory does not resolve, retrieve, load, open,
validate, authenticate, or follow an attached-template relationship, and it
does not perform style resolution or propagation.

require_no_template_style_updates_on_open fails only when the candidate has an
enabled direct setting. An absent leaf or an explicitly disabled leaf does not
store that request. no_template_style_update_on_open_changes protects any
approved stored baseline. Neither rule opens a document client or claims that a
particular client can locate a template, will update any style, or will honor
the stored request.

## Personal-information-removal-on-save scope

DocFence inventories direct `w:removePersonalInformation` children of every
discovered Document Settings part. The standard `CT_OnOff` leaf may omit
`w:val`, which means enabled when the element is present; an explicitly
supplied Word-namespace `w:val` must be one of `true`, `false`, `on`, `off`,
`1`, or `0`. The inventory rejects duplicate direct leaves, child markup,
nonblank text, unsupported attributes, and unsupported boolean values.

Microsoft's [RemovePersonalInformation reference](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.removepersonalinformation?view=openxml-3.0.1)
describes the stored setting as a request for a hosting application to remove
personal information of document authors when it saves the document, while the
definition and extent of personal information remain undefined. DocFence
records the direct stored leaf only.

Public output contains only
`personal_information_removal_on_save_enabled_setting_count` and
`personal_information_removal_on_save_disabled_setting_count`. Settings-part
paths and private fingerprints never leave the process. The comparison
signature records canonical enabled/disabled state: a state change is
review-visible, while equivalent enabled spellings do not add noise.

`require_personal_information_removal_on_save` fails only when the candidate
has no enabled direct setting. It gates a stored future-save request, not a
claim that the current package lacks personal information or that a client will
remove it. `no_personal_information_removal_on_save_changes` protects any
approved stored baseline. Neither rule opens or saves a document, identifies
authors, inspects or rewrites document properties, removes comments or
revisions, or claims that a particular client will honor the request.

## Content-control data-binding scope

DocFence inventories direct standard `w:dataBinding` children of a structured
document tag's `w:sdtPr` properties in recognized Word document stories. This
is a mapping declaration, not an XPath engine: DocFence does not select a node,
replace visible content, resolve a rich-text mapping, or claim what Word will
display.

For a nonempty `w:storeItemID`, DocFence discovers the associated custom XML
data part only through a standard internal `customXmlProps` relationship and a
valid `ds:datastoreItem` storage identifier. Conventional and Strict OOXML
relationship forms are accepted, along with both the current
`customXmlDataProps` root vocabulary and the established Word `customXml`
spelling found in real packages. A malformed recognized properties relationship
or properties root fails closed. A storage ID with no discovered associated part
is not treated as a parser error: it is counted as unmatched review evidence.

A binding can omit `w:storeItemID`; Word may then search custom XML parts using
the XPath. DocFence counts it as a binding without a storage ID but deliberately
does not guess the selected part. The public inventory reports aggregate binding
counts, referenced custom-XML part counts, and unmatched-ID counts only.
XPath expressions, prefix mappings, storage IDs, part names, data values, and
private fingerprints are never emitted.

`require_no_data_bindings` fails whenever the candidate has at least one
recognized mapping declaration. `no_data_binding_changes` compares the private
mapping declarations and, for bindings with a discovered storage ID, the paired
custom XML data/properties payloads. It therefore catches a mutation to an
identified bound data part without exposing it. It does not make an unscoped
XPath binding identify a target; use `no_custom_xml_changes` as well when every
custom XML mutation must block a handoff.

## External-source field scope

`require_no_external_fields` and `no_external_field_changes` cover a small,
explicit field-family boundary rather than every possible Word field. The
inventory recognizes `DATABASE`; legacy `DATA`; `DDE` and `DDEAUTO`;
`INCLUDE` and `INCLUDETEXT`; `INCLUDEPICTURE` and legacy `IMPORT`; `LINK`; and
`RD`. These are the families whose instructions can name a data source, file,
linked object, DDE source, or referenced document. A field is counted from
either a simple `w:fldSimple` `w:instr` attribute or a complete complex
begin-to-end field whose pre-separator `w:instrText` runs concatenate to one of
those keywords. Complex fields can nest and need not store a result separator.
When revision markup retains a deleted field code in `w:delInstrText`, its
deleted instruction is assembled separately from the current `w:instrText`
instruction. `w:moveFrom`/`w:moveTo` field-code ranges receive the same
deleted/current treatment. Consequently, a field-code replacement can
contribute both its current and its deleted stored instruction to this
inventory.

An `instrText` or `delInstrText` outside a complete complex field's instruction
portion remains ordinary text for this inventory. An unclosed complex field is
not counted. DocFence does not parse a field's arguments after classifying its
initial keyword, so it does not infer a path, connection, query, application,
item, or whether a source actually exists. It never evaluates a field, updates
a result, opens an object, starts a DDE conversation, follows a source, or runs
a query.

The candidate-state rule fails when any of the eight aggregate category counts
is nonzero. The comparison rule privately hashes each complete matching
instruction with its story context and category. Thus a changed source remains
visible to the policy without entering a report, while splitting the same
complex instruction across a different number of `instrText` runs remains
quiet. These rules intentionally do not claim to inventory arbitrary field
expressions or every field that might display a URL; generic external
relationships remain a separate inventory.

## Modern Word comment metadata scope

Modern Word comments can retain review data in parts outside the ordinary
comments story. DocFence inventories standard `people`, `commentsExtended`,
`commentsIds`, and `commentsExtensible` parts found through their standard
content types, relationship types, or conventional paths. It accepts the
established Office 15 `2010/11` and current `2012` vocabulary for the people
and comments-extended roots, while the identifier and extensible parts require
their documented `2016` and `2018` vocabularies. Every recognized root is
validated. A recognized metadata relationship must be internal and resolve to
a stored member; malformed recognized state fails closed.

Public output exposes only aggregate part, person, presence, comment-extension,
threaded/reply, resolved, identifier-record, extensible-record, reaction, and
reaction-user counts. The comparison signature retains author/contact/provider
data, comment paragraph and durable IDs, timestamps, extensions, and reactions
privately, so an identifier-only mutation remains visible without disclosing the
identifier. Relationship-ID renumbering alone remains quiet.

`require_no_modern_comment_metadata` fails when any public count is nonzero.
Use it for a clean handoff that must contain no retained modern-comment review
state. `no_modern_comment_metadata_changes` instead protects an approved
baseline. Neither rule renders a thread, resolves an account, contacts a
service, evaluates a notification, interprets extension payloads, or changes a
comment.

## Word document-task scope

Document tasks are stored workflow state in the Office document-tasks part;
they are not live task-service objects. DocFence discovers the standard part by
content type, relationship type, or conventional `word/tasks` / `word/tasks.xml`
path and validates the documented `Tasks` root. A recognized document-task
relationship must be internal and resolve to a stored package member; malformed
recognized state fails closed.

Public output contains only aggregate task-part, task, history-event,
user-reference, comment-anchor, assignment, unassignment, creation, title,
schedule, progress, priority, deletion, restoration, unassign-all, and undo
event counts. Task and event IDs, times, task titles, user IDs/names/providers,
dates, progress, priorities, and comment IDs remain inside private signatures.
An identifier-only or same-count task mutation remains visible to the
comparison rule without exposing the changed value. Relationship-ID renumbering
alone remains quiet.

`require_no_document_tasks` fails whenever any document-task public count is
nonzero, including an empty recognized task part. Use it for a handoff that must
not retain document task state. `no_document_task_changes` protects a known
baseline instead. Neither rule assigns, completes, creates, synchronizes,
notifies, or otherwise performs a task action.

## Task-pane Office web-extension scope

DocFence treats document-borne task-pane Office web-extension configuration as
stored state, not executable software. It discovers standard task-pane and
web-extension parts by content type, relationship type, and conventional
`word/webextensions/` paths; validates the `taskpanes` and `webextension`
roots; and requires every direct task-pane `webextensionref` or established
`webextension` reference to resolve through the expected internal
web-extension relationship. A malformed recognized relationship or direct
reference fails closed. A recognized web-extension part can still be counted
when it is not attached to a current task pane.

Public output reports only task-pane-part, task-pane, visible-pane,
locked-pane, web-extension-part, store-reference, property, binding, enabled
`Office.AutoShowTaskpaneWithDocument` property, and enabled Word content-control
binding-marker counts. Extension IDs, stores, versions, property names and
values, binding IDs/application references, pane dimensions, relationship IDs,
content-control marker values, and part paths are private. The comparison
signature keeps that material private, so a same-count configuration mutation
is visible but a relationship-ID renumbering alone remains quiet.

For a `w:sdtPr`, enabled `w15:webExtensionCreated` takes precedence over
enabled `w15:webExtensionLinked` when DocFence counts a bound content control.
An auto-show count is evidence of a stored enabled property only; it is not a
claim that Word will install, load, display, or auto-open an add-in.
`require_no_taskpane_web_extensions` fails when any public count is nonzero;
`no_taskpane_web_extension_changes` protects an approved baseline. These rules
do not install or execute an add-in, retrieve a manifest, access a store,
authenticate, resolve a binding, or assess add-in safety.

## Office sensitivity-label metadata scope

Office can retain sensitivity-label state in two compatible storage forms.
Modern Office 2021 packages use a `LabelInfo` part whose `labelList` records
label ID, enabled/removed state, method, tenant site ID, and optional label
details. Older or compatibility paths retain MIP key/value metadata in custom
document properties, including `MSIP_Label_<GUID>_<attribute>`, the legacy
`Sensitivity` property, and documented Word header/footer/watermark marking
properties. MIP also permits custom attributes and versioned attribute names.

DocFence discovers a modern LabelInfo part through the standard
classification-label package relationship, the Office SDK content type, or the
canonical `docMetadata/LabelInfo` / `docMetadata/LabelInfo.xml` paths. It
requires a recognized classification-label relationship to originate in the
package relationship item, use an internal target, and resolve to a stored
member. At most one LabelInfo part is accepted. The `labelList` root, direct
`label` records, required label attributes, boolean state, tenant-site GUID,
and extension-list structure are validated; malformed recognized state fails
closed. Legacy MIP evidence is found in canonical or standard
relationship-linked custom-property parts, including noncanonical property-part
targets.

Public output reports only aggregate LabelInfo part, label, enabled-label,
removed-label, extension, legacy MIP-label, legacy MIP-property, legacy
`Sensitivity`-property, and Word content-marking-property counts. Label IDs,
tenant IDs, label names, methods, dates, action IDs, extension payloads, MIP
custom attribute names and values, property names and values, relationship IDs,
paths, and fingerprints remain private. The comparison signature retains that
material, so a label-ID/name/value or same-count custom-property mutation is
review-visible without disclosure; relationship-ID renumbering alone remains
quiet.

`require_no_sensitivity_label_metadata` fails whenever any sensitivity-label
public count is nonzero, including an empty recognized LabelInfo part. Use it
for a handoff that must not retain label/tenant/governance metadata.
`no_sensitivity_label_metadata_changes` protects an approved baseline instead.
Neither rule decrypts IRM-protected content, reads a LabelInfo stream from an
encrypted storage, resolves a label policy, determines permissions, applies or
removes labels, or predicts whether an Office client will display a marking.

## OPC package thumbnail scope

An OPC thumbnail is an image part identified by the standard thumbnail
relationship from either the package or another package part. It is a separate
handoff and privacy surface from an ordinary embedded document image or a
filename that merely looks like a thumbnail.

DocFence recognizes only the exact Transitional or Strict standard thumbnail
relationship types. A recognized relationship can originate in the package
relationship item or a relationship part for a stored package member. It must
be internal, resolve to a stored member with an `image/` content type, and be
the only thumbnail relationship from that source. A recognized thumbnail target
must not own a Relationships part. Malformed recognized topology fails closed.
An unreferenced member such as `docProps/thumbnail.png` remains in the generic
unclassified-payload inventory; DocFence does not infer thumbnail status from a
path or inspect arbitrary image files.

Public output contains only `thumbnail_relationship_count` and
`thumbnail_part_count`. Image bytes, relationship IDs, sources and targets,
content types, part paths, and private fingerprints never leave the process.
The comparison signature retains the normalized relationship semantics and raw
image bytes, so a same-count image rewrite remains review-visible while a
relationship-ID renumbering alone remains quiet.

`require_no_package_thumbnails` fails whenever the candidate has a recognized
thumbnail part. `no_package_thumbnail_changes` protects an approved stored
baseline instead. Neither rule decodes, renders, classifies, or otherwise
interprets an image; searches arbitrary filenames for lookalikes; or predicts
whether Word, Explorer, or another client will display a thumbnail.

## OOXML Markup Compatibility scope

Markup Compatibility and Extensibility (MCE) can retain stored alternatives
that a consumer selects according to its own supported features and processing
settings. That makes `mc:AlternateContent` and its compatibility-rule
attributes a separate review boundary from ordinary text, rendering, or a
generic XML payload change.

DocFence scans stored non-relationship `word/*.xml` members that use the
standard `http://schemas.openxmlformats.org/markup-compatibility/2006`
namespace. It reports only aggregate MCE-part, `AlternateContent`, `Choice`,
and `Fallback` counts plus whitespace-token counts for `Choice/@Requires` and
the `mc:Ignorable`, `mc:MustUnderstand`, `mc:ProcessContent`,
`mc:PreserveElements`, and `mc:PreserveAttributes` attributes. Branch bodies,
feature-prefix and qualified-name values, compatibility-rule values, XML
member paths, and private fingerprints never leave the process. The private
signature retains the recognized elements and attributes, so a same-count
branch or rule-value rewrite remains review-visible.

This inventory is limited to Word members. Separately, OPC §10.5.2 makes MCE
namespace elements or attributes malformed anywhere in a recognized package
Digital Signature XML Signature part; DocFence rejects such a signature. That
structural signature rule does not select, preprocess, or otherwise interpret
MCE markup.

`require_no_markup_compatibility` fails whenever a candidate contains a
recognized MCE element or attribute. `no_markup_compatibility_changes` protects
an approved baseline instead. Neither rule validates MCE conformance, selects
a branch, resolves a feature prefix, applies a target Office version,
preprocesses or saves a package, or predicts what any document client will
load or render.

## DrawingML nonvisual visibility scope

The DrawingML hidden attribute is a direct stored nonvisual-property setting:
an object can be present while its declaration specifies hidden. This is
separate from Word hidden-text runs, style resolution, MCE branch processing,
layout, and renderer behavior.

DocFence inventories direct unqualified hidden attributes only on supported
nonvisual-property elements in supported Word stories: DrawingML main and
picture cNvPr, WordprocessingDrawing docPr, and Word 2010 w14, wpg, and wps
cNvPr. It supports Transitional and Strict forms of the standard DrawingML
namespaces. Every stored supported declaration is retained, including
duplicates and declarations in MCE branches. Other elements named cNvPr, any
inferred or inherited visibility, and arbitrary drawing markup are outside
this boundary.

Public output contains only declaration and story counts plus hidden,
explicitly-shown, and invalid-value counts. Object names, descriptions, titles,
IDs, raw invalid values, story paths, and fingerprints never leave the process.
The private signature canonicalizes valid XML Boolean spellings, so true and 1
are treated as the same stored hidden state and false and 0 as the same
explicitly-shown state. A same-count hidden-to-shown swap remains
review-visible; a rewrite limited to object metadata does not become a
visibility-inventory change. Invalid raw values remain private but retain a
separate signature state.

`require_no_hidden_drawing_objects` rejects any recognized hidden declaration
or invalid Boolean value. It permits an explicit false or 0 declaration.
`no_drawing_object_visibility_changes` protects an approved baseline of any
recognized declarations. Neither rule validates full DrawingML conformance,
selects an MCE branch, resolves a drawing object or its identity, calculates
effective visibility, lays out or renders an object, or asserts how a client
will display it.

## OPC package digital-signature scope

An OPC package can retain a Digital Signature Origin part, one or more XML
Signature parts, and optional Digital Signature Certificate parts. XML signature
markup can also contain certificate, signing-time, comment, provider, and
reference material. That is a meaningful handoff and privacy surface, but it is
not by itself a trust verdict.

DocFence recognizes the standard root-package digital-signature-origin
relationship, exact OPC origin/XML-signature/certificate content types
(including content-type defaults), and the conventional
`_xmlsignatures/origin.sigs` residue path. A recognized origin relationship must
be root-package scoped, internal, resolvable, and unique. Recognized signature
relationships must originate at a recognized origin; recognized certificate
relationships must originate at a recognized XML-signature part. Each must be
internal, resolve to a stored member, and have the expected content type.
Recognized XML signature parts must have the XMLDSIG `Signature` root, exactly
one direct `SignedInfo` then `SignatureValue`, optional `KeyInfo`, and zero
or more direct `Object` elements. `SignedInfo` must directly contain
`CanonicalizationMethod`, `SignatureMethod`, then one or more `Reference`
elements. `SignedInfo` and `SignatureValue` permit only their optional
`Id` attribute at this boundary; `SignatureValue` cannot have child XML, and
`SignatureMethod/@Algorithm` must be nonblank. The one direct
`CanonicalizationMethod/@Algorithm` must be exactly one of OPC's two permitted
XML Canonicalization URIs (with or without comments); a missing or other URI is
malformed. Each direct `SignedInfo/Reference` must carry an explicit XMLDSIG
same-document URI: the empty URI or a local fragment beginning with `#`.
An omitted, relative, or absolute URI is malformed. Malformed recognized
topology or XMLDSIG shape fails closed. Every XMLDSIG
`DigestMethod/@Algorithm` in a recognized package signature must also avoid
OPC's expressly forbidden MD5 URI; this is not a broad cryptographic judgment
about SHA-1 or other algorithms. Every XMLDSIG `ds:Transform/@Algorithm` must
be OPC's Relationship Transform URI or one of its two XML Canonicalization
URIs; a missing or other URI is malformed anywhere in the recognized
Signature. Every XMLDSIG `ds:XPath` element is also rejected anywhere in the
recognized Signature because OPC says it shall not be present. OPC §10.5.2
also forbids MCE-namespace elements and attributes anywhere in that Signature.
DocFence does not parse or execute XPath or a transform, validate MCE
conformance, or select an MCE branch. The direct-sequence rule is not full
XMLDSIG schema validation: it does not validate base64 lexical content, method
parameters, KeyInfo/Object payloads, a digest, or a signature.

XMLDSIG `ds:Transforms` elements must be attribute-free, and every
`ds:Transform` must have exactly its required `Algorithm` attribute. An
extra attribute is malformed anywhere in the recognized Signature, including
outside the bounded coverage chain. This direct grammar check does not validate
arbitrary transform parameters or child markup.

When that permitted URI is a Relationship Transform, its local package syntax
is also mandatory: it must be a direct `ds:Transform` under a direct
`ds:Manifest/ds:Reference/ds:Transforms` chain. The reference URI must declare a
`.rels` part with OPC's case-insensitive relationships content type; the transform must
contain at least one direct `opc:RelationshipReference` or
`opc:RelationshipsGroupReference`; and its immediate following transform must
be XML Canonicalization (with or without comments). One recognized XML
signature may name a declared relationships part with that transform only once.
A violation is malformed even when it occurs outside the bounded coverage
chain. This validates stored placement only: it does not resolve the declared
target, interpret selectors, or execute a transform.

Every OPC relationship selector is also globally constrained. A direct
`opc:RelationshipReference` child of a Relationship Transform must have exactly
`SourceId` and no child elements; a direct
`opc:RelationshipsGroupReference` must have exactly `SourceType` and no child
elements. The selectors may not occur anywhere else in the recognized
Signature. Missing, wrong, or extra attributes and nested selector markup are
malformed before coverage is considered. This does not resolve or interpret a
selector value.

Public output reports only aggregate origin-part, XML-signature-part,
certificate-part, SignedInfo-reference, manifest-reference,
relationship-reference, inline-X.509-certificate, and signature-property
counts. Signer and certificate contents, signature values, algorithms, signing
times, comments, provider data, reference URIs, relationship IDs, part paths,
and fingerprints remain private. The comparison digest retains all recognized
part bytes and relationship semantics, so a same-count signature or certificate
rewrite is review-visible.

`require_no_package_digital_signatures` fails whenever any package-signature
public count is nonzero, including a recognized empty origin part. Use it for a
handoff that must not retain signer, certificate, or signature residue.
`no_package_digital_signature_changes` protects an approved stored baseline.
Neither rule verifies a cryptographic signature or digest, validates a
certificate or chain, checks revocation or a timestamp, establishes signer
identity, assesses an Office client, or decides whether a signature should be
trusted.

### Static declared package-signature coverage

`require_complete_package_signature_coverage` is an opt-in gate for a
handoff that is expected to retain package signatures. It is not a softer form
of `require_no_package_digital_signatures`: the latter rejects all stored
package-signature material, while this rule requires a bounded declaration
surface to be present and complete.

For every recognized XML signature part, DocFence requires exactly one direct
`ds:Object` whose only attribute is `Id="idPackageObject"` and whose direct
children are exactly one `ds:Manifest` followed by one
`ds:SignatureProperties`. It also requires exactly one direct `SignedInfo`
`Reference` whose local fragment is `#idPackageObject`. Only that one
package-specific manifest is resolved; a missing or nonstandard object,
missing or extra direct child, duplicate object, or duplicate binding is not
combined with another manifest and leaves declaration coverage unavailable.
Its direct `ds:SignatureProperties` must in turn contain exactly one direct
`ds:SignatureProperty` whose only attributes are
`Id="idSignatureTime"` and `Target`. The target is either empty or the
direct local fragment for the root `ds:Signature/@Id`. That property must
contain exactly one attribute-free `opc:SignatureTime`, which in turn has
only attribute-free `opc:Format` then `opc:Value` children. DocFence
recognizes the six OPC schema time precisions and requires the stored value's
lexical shape to match its declared precision. Missing, duplicate,
mis-targeted, or malformed timestamp declarations leave coverage unavailable;
this recognizes only the claimed timestamp's stored syntax, not its accuracy,
source, authority, or cryptographic validity.
Within this bounded binding/manifest chain, each `Reference` may carry only
XMLDSIG's `Id`, `URI`, and `Type` attributes and each direct
`DigestMethod` must carry exactly `Algorithm`. An unknown attribute leaves a
binding without declared coverage or reports a manifest reference as aggregate
unsupported; it cannot lend coverage. This is not full XMLDSIG schema
validation: DocFence does not validate method-parameter child markup, base64,
or any digest or signature value.
The binding reference itself must retain XMLDSIG's direct child order:
optional `ds:Transforms`, one `ds:DigestMethod` with a nonblank
`Algorithm`, then one direct, attribute-free, child-free, nonempty
`ds:DigestValue`, without non-whitespace direct text. It may omit transforms
or carry one direct nonempty list of only OPC's two XML Canonicalization
algorithms. The global recognized-signature boundary rejects every non-OPC
`ds:Transform/@Algorithm`, every malformed Relationship Transform or
relationship selector, every XMLDSIG `ds:XPath` element, and MCE-namespace
markup before this bounded
binding/manifest chain is evaluated; a Relationship Transform in this binding
position is therefore malformed. A `DigestMethod` in the recognized
signature cannot use OPC's expressly forbidden MD5 URI; the audit does not
otherwise judge an algorithm's cryptographic suitability.
DocFence checks that stored syntax and the binding URI fragment, but
does not decode or recompute its digest, execute transforms, verify a signature,
or establish trust. Supported manifest part references use an exact local part
URI plus ASCII-case-insensitive matching content type. Before a manifest reference is
credited, its direct XMLDSIG children must be optional
`ds:Transforms`, one `ds:DigestMethod` with a nonblank `Algorithm`, and one
direct, attribute-free, child-free, nonempty `ds:DigestValue`, with no
non-whitespace direct text, in that order; the globally rejected exact MD5 URI,
any non-OPC `ds:Transform/@Algorithm`, malformed Relationship Transform or
relationship selector, any `ds:XPath` element, or MCE-namespace markup cannot
lend coverage.
Supported
relationship references use one direct `ds:Transforms` list containing only
the supported OPC relationship and XML Canonicalization algorithms, with
exactly one relationship transform immediately followed by XML Canonicalization
(with or without comments), and `RelationshipReference/@SourceId` or
`RelationshipsGroupReference/@SourceType` selector shapes. Within this bounded
coverage audit, selector values match stored relationship IDs and types
ASCII-case-insensitively and cover every match, as OPC §10.6 requires. A
selector-placement or
selector-shape failure has already rejected the recognized Signature. An
unsupported URI, reference-child shape, or other relationship-transform syntax
is not assumed to be coverage: it is reported only as an aggregate unsupported
reference; a missing member, content-type mismatch, missing relationship item,
or selector that selects no stored relationship is reported as an aggregate
unresolved reference.

OPC permits no more than one Relationship Transform for a particular
relationships part in one XML signature. DocFence enforces that limit across
every manifest in the recognized Signature, before coverage is evaluated; a
duplicate is malformed rather than an aggregate unsupported reference.

For a non-relationship package part, supported references have no
`ds:Transforms` element or exactly one direct nonempty `ds:Transforms` list of
XMLDSIG `Transform` elements using only the two OPC-supported XML
Canonicalization algorithms. Empty or duplicate lists are unsupported rather
than being treated as part coverage. A Relationship Transform in this context
is malformed; a non-OPC transform algorithm has already rejected the
recognized signature.

The bounded Word scope comprises every non-relationship member under `word/`,
root-package relationships whose type is `officeDocument`, and every stored
relationship sourced by a Word part. `DFP092` fails if no recognized XML
signature is present, any recognized XML signature lacks this
package-specific-object topology, any of those bounded parts or relationships
is undeclared, or its one manifest contains unresolved or unsupported
references. It does not assert that this scope is all package content or all
content an Office client may use. Coverage is the union of each qualifying
signature's one package-specific manifest; the rule does not require each
individual signature to select every bounded item itself.

```yaml
version: 1
rules:
  require_complete_package_signature_coverage: true
```

`no_package_signature_coverage_changes` compares a private semantic signature
of the declaration resolution. It catches same-count coverage reassignment
without exposing part paths, object identifiers, reference URIs, relationship
identifiers or types, selectors, or digest material.

```yaml
version: 1
rules:
  no_package_signature_coverage_changes: true
```

The audit does not parse, decode, or recompute reference digests or
canonicalization, evaluate
arbitrary XMLDSIG transforms, verify a signature value, validate a certificate
or trust chain, check revocation or timestamps, establish signer identity, or
determine the effective coverage, rendering, or trust decision of an Office
consumer. It resolves only this local, static declaration subset.

## Word editing and write-protection scope

Word can retain two related but independent settings: direct
`w:documentProtection` editing restrictions and `w:writeProtection` state. The
former can restrict editing to read-only, comments, tracked changes, or forms
and can include an enforcement/formatting setting; the latter can represent a
read-only recommendation or stored write-protection material. Neither is file
encryption or a general security verdict.

DocFence discovers document Settings parts through the conventional
`word/settings.xml` fallback and Transitional or Strict `settings` relationships
from the main or glossary document. The generic settings signature covers every
discovered part. The protection inventory recognizes direct Word-namespace
`documentProtection` and `writeProtection` leaves in those parts, allows at most
one of each per part, and rejects recognized elements with children, nonblank
text, unknown/non-Word attributes, duplicate attributes, invalid edit values,
or invalid boolean values.

Public output reports only aggregate document-protection element,
explicitly-enabled enforcement, formatting-restriction, read-only, comments,
tracked-changes, forms, document password-material, write-protection,
read-only-recommendation, and write password-material counts. A
password-material count means one recognized element stores at least one
`hash`, `salt`, `hashValue`, or `saltValue` attribute; it does not mean that
DocFence judged a verifier complete, valid, or strong. The enforcement-enabled
count similarly means only that the stored attribute is explicitly true—not
that a particular Word client will enforce it.

Hashes, salts, verifier values, provider and algorithm fields, all other
protection attributes, part paths, and fingerprints are never emitted. The
full recognized direct elements are privately fingerprinted, so a same-count
verifier or configuration rewrite remains visible. `require_no_word_protection`
fails whenever any protection element is present, including an otherwise empty
recognized element. Use it for a handoff that must carry no stored Word
editing/write-protection state. `no_word_protection_changes` protects an
approved baseline instead.

These rules do not validate password construction, derive or recover a
password, estimate password or algorithm strength, bypass a restriction,
determine the effective settings after Word/compatibility behavior, decrypt a
file, or decide whether the stored protection is a security control.

## Word editable-range permission scope

Word stores editable-region boundaries as `w:permStart` and `w:permEnd` in
document stories. A start can carry an individual editor value in `w:ed`, a
predefined application group in `w:edGrp`, and optional table-column selectors;
an end is paired with its start by the stored ID. Individual editor values can
be email addresses, aliases, or domain identities, so DocFence never emits
them. It also never emits marker IDs, exact column values, story paths, or
fingerprints.

The inventory scans every supported body, header, footer, footnote, endnote,
comment, and glossary story in either Word namespace. Recognized markers must
be Word-namespace leaves with a required ID and only their standard attributes.
Predefined groups must use the `none`, `everyone`, `administrators`,
`contributors`, `editors`, `owners`, or `current` vocabulary. Column selectors
must be nonnegative decimal syntax, custom-XML placement must be `next` or
`prev`, and duplicate start or end IDs within one story fail closed because the
pairing would be ambiguous.

Public output contains only aggregate start/end, paired/unpaired,
individual-editor, predefined-group, table-column-selector, and
custom-XML-placement counts. A raw marker change, including a same-count editor
identity rewrite, remains visible through the private inventory signature.
Unmatched boundaries are reported as stored markup; DocFence does not infer
that they grant an effective permission. Similarly, it records both `w:ed` and
`w:edGrp` when both are stored but does not resolve Word's client-specific
precedence behavior.

`require_no_word_permission_ranges` fails whenever a candidate stores any
recognized range-permission marker, including an unmatched marker.
`no_word_permission_range_changes` compares the private inventory signature to
protect a controlled baseline. Neither rule authenticates a person, resolves a
group, proves that an editor is currently authorized, calculates editable
content, or makes an encryption or security claim.

## Word document-variable scope

Word can retain arbitrary name/value state as direct `w:docVar` leaves inside a
`w:docVars` container in a document Settings part. Word's automation APIs use
this store for document or template state, and the values are normally invisible
unless a matching `DOCVARIABLE` field is inserted. Since either name or value
can be sensitive, DocFence never emits either, nor a Settings-part path or
private fingerprint.

The inventory discovers document Settings parts through the conventional
`word/settings.xml` fallback and Transitional or Strict relationships from the
main or glossary document. It accepts at most one direct `w:docVars` container
per Settings part. A recognized container must have no attributes or nonblank
text and may contain only direct Word-namespace `w:docVar` leaves. Each leaf
must have only required Word-namespace `w:name` and `w:val` attributes; names
must be 1–255 UTF-16 code units and values at most 65,280 UTF-16 code units,
matching the Open XML SDK schema constraints. The OOXML schema permits an empty
value, so it is counted rather than rejected even though normal Word automation
writes do not use one.

Public output contains only aggregate container, variable, and empty-value
counts. Every full recognized container is privately fingerprinted, so a
same-count name or value rewrite is visible to a comparison gate without
placing either string in a report.

`require_no_word_document_variables` fails whenever a candidate contains a
recognized `w:docVars` container, including an empty one. Use it for a handoff
that must carry no stored document-variable state.
`no_word_document_variable_changes` protects an approved baseline instead.

DocFence separately recognizes complete `DOCVARIABLE` field instructions in
both `w:fldSimple/@w:instr` and complete complex-field pre-separator
instruction sequences across supported stories. Current and deleted revision
variants are retained as separate stored evidence. The public field inventory
contains only reference and story counts plus literal/nonliteral and
exact-literal association counts; it never emits a field instruction or
variable name. A literal is deliberately narrow: it must be a leading plain
argument with no whitespace or one complete quoted string, optionally followed
by Word field-switch material. A nested or compound field expression is counted
as nonliteral rather than parsed.

For a literal, DocFence can only report whether its exact string appears in a
validated `w:docVar` associated with the same main or glossary package document
scope. This is stored-package association evidence, not field resolution. In
particular, a literal with no same-scope match may be supplied by an attached
template, and a matching stored name does not prove that Word will render it or
that a client will choose that value.

`require_no_word_document_variable_fields` fails whenever a candidate contains
a complete stored `DOCVARIABLE` reference. Use it for a handoff that must carry
no such field codes. `no_word_document_variable_field_changes` protects an
approved field-reference baseline. None of these rules evaluates a field, runs
a macro, resolves a template, determines whether a Word client will display a
value, or judges the safety or meaning of the stored state.

## Word HYPERLINK field scope

`HYPERLINK` field codes are distinct from direct `w:hyperlink` markup and from
the external-source field inventory: a complete field code can
carry a URL, file location, or bookmark directly without a relationship. The
inventory scans direct `w:fldSimple/@w:instr` instructions and complete complex
pre-separator instruction sequences across supported stories, retaining current
and deleted field-code variants separately. Loose instruction text and
unclosed complex fields do not count.

Public output reports only reference/story counts plus mutually exclusive
lexical classifications: a literal leading destination, a literal `\l`
internal-location-only target, or a dynamic/unparseable field. A literal is a
leading plain token or one wholly quoted string with optional trailing
field-switch material. Nested or compound expressions remain dynamic or
unparseable. The primary literal is not labeled external because Word permits a
bookmark there; `\l` is the documented in-file location switch.

`require_no_word_hyperlink_fields` fails whenever a candidate contains a
complete stored `HYPERLINK` reference. Use it for a handoff that must carry no
such field codes. `no_word_hyperlink_field_changes` compares the private
inventory signature to protect an approved baseline. Neither rule emits,
resolves, follows, evaluates, or renders a destination, nor does it establish
that a target is reachable, safe, or displayed by Word.

## WordprocessingML hyperlink markup scope

Direct `w:hyperlink` elements are a separate stored surface from both
`HYPERLINK` field codes and the generic package relationship inventory. The
inventory scans direct elements across every supported Word story. It counts an
element only when the element itself is present: an unreferenced hyperlink
relationship remains visible to the generic relationship inventory but does
not become a direct-markup count.

The inventory follows the element's documented stored precedence without
following a target. If `r:id` is present, its relationship is the target and
supersedes any `w:anchor`; that shadowed anchor is counted only as a stored
attribute. Without `r:id`, a `w:anchor` produces an anchor-only count; with
neither attribute, the element is counted as the documented current-document
start form. Relationship-backed elements are classified from the referenced
relationship's stored type and target mode: recognized hyperlink relationships
with `External` and `Internal` target modes receive separate counts. A resolved
relationship with another type or target mode is counted as unsupported
relationship-backed markup rather than treated as a standard link target.

The private signature includes the resolved relationship semantics and the
complete direct element, so target, anchor, `w:docLocation`, tooltip, target
frame, history, and display-markup rewrites remain visible even with unchanged
public counts. Relationship IDs are normalized through their resolved semantics,
so an ID renumbering alone remains quiet. Targets, anchors, locations, tooltips,
frame names, history values, display text, relationship IDs, story paths, and
fingerprints never appear in reports.

`require_no_word_hyperlink_markup` fails whenever a candidate contains a
direct `w:hyperlink` element. Use it for a handoff that must carry no such
markup. `no_word_hyperlink_markup_changes` compares the private inventory
signature for a controlled baseline. Neither rule resolves, retrieves, follows,
validates, evaluates, or renders a target, and neither establishes that a
relationship target is reachable, safe, or honored by a Word client.

## DrawingML hyperlink-action scope

Direct DrawingML `a:hlinkClick`, `a:hlinkHover`, and `a:hlinkMouseOver`
elements are a separate stored surface from direct WordprocessingML
`w:hyperlink` markup, `HYPERLINK` field codes, and broad package relationship
totals. The inventory scans those elements across supported Word stories and
counts each stored marker rather than deduplicating by relationship or visual
object. It does not select a Markup Compatibility branch or determine whether a
client renders a marker.

Markers with an `r:id` are classified from the referenced relationship's stored
type and target mode as external, internal, or unsupported. A marker with no
`r:id` remains separately visible as malformed stored evidence. An unreferenced
hyperlink relationship remains visible to the generic relationship inventory
but does not become a DrawingML marker count. The public aggregate also records
the presence of `action` and `invalidUrl` attributes; targets, URL values,
actions, tooltips, frame names, history settings, relationship IDs, story
paths, and fingerprints never appear in reports.

`require_no_word_drawing_hyperlinks` fails whenever a candidate contains a
stored DrawingML hyperlink-action marker. Use it for a handoff that must carry
no such stored action markup. `no_word_drawing_hyperlink_changes` compares the
private inventory signature for a controlled baseline; same-count target or
attribute rewrites remain visible, while relationship-ID renumbering with
unchanged semantics remains quiet. Neither rule resolves, retrieves, follows,
validates, evaluates, renders, or executes an action, nor does either establish
that a target is reachable, safe, or honored by a Word client.

## DrawingML linked-picture scope

Direct `a:blip/@r:link` markup is distinct from a `w:hyperlink` element,
`HYPERLINK` field, DrawingML hyperlink-action element, `r:embed` embedded-image
reference, and generic relationship total. DocFence scans each direct
`a:blip/@r:link` marker in supported Word stories, including duplicate markers
and markers in Markup Compatibility branches. It does not deduplicate visual
objects or select a client-rendered branch.

Each marker is privately fingerprinted with its relationship resolved by
semantics and publicly classified as a standard image relationship with stored
external mode, stored internal mode, or unsupported relationship. An orphaned
image relationship and an `a:blip` with `r:embed` but no `r:link` do not create
marker counts. Targets, relationship IDs, surrounding drawing markup, story
paths, and fingerprints never appear in reports. Same-count target or direct
markup changes remain visible; an ID-only renumbering with unchanged semantics
remains quiet.

`require_no_word_drawing_linked_pictures` fails whenever a candidate contains a
stored direct linked-picture marker. Use it for a handoff that must carry no
such markup. `no_word_drawing_linked_picture_changes` compares a private
inventory signature for a controlled baseline. Neither rule retrieves an image,
selects a Markup Compatibility branch, resolves an image target, renders a
picture, updates a link, or establishes that Word will load, reach, or honor a
target.

## Legacy VML shape-link scope

Direct legacy VML `href` attributes form a fifth stored link surface, separate
from `HYPERLINK` field codes, direct WordprocessingML `w:hyperlink` markup,
DrawingML action markup, and broad package relationship totals. The inventory
scans supported Word stories for a direct unqualified `href` only on the
documented VML shape-family elements: `arc`, `curve`, `image`, `line`, `oval`,
`polyline`, `rect`, `roundrect`, `shape`, `group`, and `shapetype`. It counts a
present attribute even when its value is empty, so malformed or inert-looking
stored markup remains reviewable.

Public output reports aggregate marker and story counts, separates concrete
geometry, group, and shape-template markers, and reports only the presence of a
direct `target` attribute. Raw `href` values, frame targets, titles, alternate
text, shape IDs, story paths, and fingerprints never appear in reports. The
private signature retains the complete direct element, so same-count `href`,
target, or other markup rewrites remain visible.

The inventory does not calculate an effective link inherited from a group or
shape template, select a Markup Compatibility branch, inspect arbitrary VML
elements, resolve, retrieve, follow, validate, evaluate, render, or execute an
action. It is stored-markup evidence only, not proof that Word or another
client will render, honor, safely reach, or follow a target.

`require_no_word_vml_hyperlinks` fails whenever a candidate contains a
supported direct VML shape-link marker. Use it for a handoff that must carry no
such legacy markup. `no_word_vml_hyperlink_changes` compares the private
inventory signature for a controlled baseline.

## Legacy VML external-image scope

Direct legacy VML `v:imagedata/@r:id` markup is a separate relationship-backed
surface from VML shape `href`, `HYPERLINK` fields, direct WordprocessingML
`w:hyperlink` markup, DrawingML linked pictures, and generic relationship
totals. DocFence scans each direct `r:id` marker in supported Word stories, but
records it only when its resolved relationship has stored `TargetMode=External`.
Duplicate markers and markers in Markup Compatibility branches remain separate
stored evidence; the inventory does not select a client-rendered branch or
deduplicate a visual image.

A standard image relationship with stored external mode is classified as an
external image relationship. Another externally stored relationship type is
counted as unsupported evidence. Ordinary internal image relationships do not
become external-image markers, nor does an orphaned external image
relationship. Raw VML `src`, `r:pict`, `r:href`, and `o:relid` are deliberately
excluded rather than treated as alternate forms of this marker.

Public output reports only aggregate marker/story and relationship-classification
counts. Image targets, relationship IDs, source values, VML attributes, story
paths, and fingerprints never appear in reports. The private signature is
limited to the reviewed external `r:id` relationship semantics: a same-count
external target rewrite stays visible, an unchanged-semantics relationship-ID
renumbering remains quiet, and a raw-`src` rewrite does not become an
external-image inventory change.

`require_no_word_vml_external_images` fails whenever a candidate contains a
stored direct VML external-image marker. Use it for a handoff that must carry no
such markup. `no_word_vml_external_image_changes` compares the private
inventory signature for a controlled baseline. Neither rule resolves,
retrieves, renders, updates, or validates an image target, or establishes that
Word will load, reach, or honor it.

## Legacy VML image-data hyperlink scope

Direct legacy VML `v:imagedata/@r:href` markup is separate from image-data
`r:id` external-image markers, VML shape `href`, `HYPERLINK` fields, direct
WordprocessingML `w:hyperlink` markup, DrawingML linked pictures, and generic
relationship totals. DocFence records every direct `r:href` marker in supported
Word stories, including duplicate markers and markers in Markup Compatibility
branches. It does not select a client-rendered branch or associate a marker
with a visual image.

The Open XML SDK identifies this attribute as an explicit relationship to a
hyperlink target. A standard hyperlink relationship with stored external or
internal target mode is counted in the respective class. Any other resolved
relationship type or mode remains reviewable as unsupported stored evidence;
it is not silently treated as a conventional hyperlink. An orphaned hyperlink
relationship does not create a marker count. Image-data `r:id`, `r:pict`, raw
`src`, and `o:relid` are deliberately excluded rather than treated as alternate
forms of this marker.

Public output reports only aggregate marker/story and relationship-classification
counts. Targets, relationship IDs, VML attributes, story paths, and fingerprints
never appear in reports. The private signature is limited to the reviewed
`r:href` relationship semantics: a same-count target rewrite stays visible, an
unchanged-semantics relationship-ID renumbering remains quiet, and changes to
the excluded image-data attributes do not become an image-data-hyperlink
inventory change.

`require_no_word_vml_image_hyperlinks` fails whenever a candidate contains a
stored direct VML image-data hyperlink marker. Use it for a handoff that must
carry no such markup. `no_word_vml_image_hyperlink_changes` compares the
private inventory signature for a controlled baseline. Neither rule resolves,
retrieves, follows, validates, evaluates, renders, or executes a target, or
establishes that Word will honor it.

## Legacy VML linked-OLE-object scope

Direct legacy Office VML `o:OLEObject` markup with unqualified `Type="Link"`
is separate from the separately inventoried `Type="Embed"`, WordprocessingML
`w:objectLink`, VML image data, VML shape links, DrawingML linked pictures,
fields, and generic embedded OLE/package/control relationship or payload
totals. DocFence records every direct marker in supported Word stories,
including duplicates and markers in Markup Compatibility branches. It does not
select a client-rendered branch, associate a marker with a visual object, or
deduplicate an object.

When a marker carries `r:id`, a standard OLE-object relationship with stored
external or internal target mode is counted in the respective class. Any other
resolved relationship type or mode is reviewable as unsupported evidence. A
direct marker without `r:id` is retained as a distinct stored-evidence class;
the schema permits the attribute to be absent. `UpdateMode="Always"` is
reported only as a stored automatic-update marker. Every other or absent
`UpdateMode` value is counted as nonautomatic-or-unspecified; neither class is
a claim that a Word client will actually update an object.

Public output exposes only aggregate marker/story, update-mode, and
relationship-classification counts. Source locations, monikers, program,
shape, and object IDs, relationship IDs and targets, field codes, VML markup,
story paths, and fingerprints never appear in reports. The full direct marker
is privately fingerprinted, so a same-count source, program, field-code,
update-mode, or target rewrite remains visible. Relationship-ID renumbering
with unchanged relationship semantics remains quiet.

`require_no_word_vml_linked_ole_objects` fails whenever a candidate contains a
stored direct VML linked-OLE marker. Use it for a handoff that must carry no
such markup. `no_word_vml_linked_ole_object_changes` compares the private
inventory signature for a controlled baseline. Neither rule resolves, retrieves,
opens, updates, activates, evaluates, renders, or executes an OLE object, or
establishes that Word will honor it.

## Legacy VML embedded-OLE-object scope

Direct legacy Office VML `o:OLEObject` markup with unqualified `Type="Embed"`
is separate from `Type="Link"`, WordprocessingML `w:objectEmbed` and
`w:objectLink`, VML image data and shape links, fields, and generic embedded
OLE/package/control relationship or payload totals. DocFence records every
direct marker in supported Word stories, including duplicates and markers in
Markup Compatibility branches. The Office VML contract permits several parent
forms, so the boundary follows the stored marker without selecting a rendering
position, client-rendered branch, or visual-object association.

When a marker carries `r:id`, a standard OLE-object relationship with stored
external or internal target mode is counted in the respective class. Any other
resolved relationship type or mode is reviewable as unsupported evidence. A
direct marker without optional `r:id` is retained as a distinct stored-evidence
class. `UpdateMode` is described for the Link form, so this embedded-object
inventory deliberately does not report a client-update category. It does not
retrieve a target or decode a payload, open, activate, render, or execute an
OLE object, and no count establishes that a Word client will honor it.

Public output exposes only aggregate marker/story and relationship-
classification counts. Program, shape, and object identifiers, relationship IDs
and targets, update metadata, field codes, VML markup, story paths, and
fingerprints never appear in reports. The full direct marker is privately
fingerprinted, so same-count program, field-code, update-metadata, or target
rewrites remain visible. Relationship-ID renumbering with unchanged
relationship semantics remains quiet.

`require_no_word_vml_embedded_ole_objects` fails whenever a candidate contains
a stored direct VML embedded-OLE marker. Use it for a handoff that must carry no
such markup. `no_word_vml_embedded_ole_object_changes` compares the private
inventory signature for a controlled baseline.

## WordprocessingML linked-object-property scope

Direct WordprocessingML `w:objectLink` markers are counted only when they are
direct children of `w:object` in a supported Word story. The boundary is
separate from `w:objectEmbed`, legacy Office VML `o:OLEObject Type="Link"` and
`Type="Embed"`, VML image data and shape links, DrawingML linked pictures,
fields, and broad embedded OLE/package/control relationship or payload totals.
DocFence retains duplicates and markers in Markup Compatibility branches, but
does not select a client-rendered branch, associate a marker with a visual
object, or deduplicate an object.

When a direct marker carries `r:id`, a standard OLE-object relationship with
stored external or internal target mode is counted in the respective class. Any
other resolved type or mode is preserved as unsupported evidence. Although the
schema requires `r:id` and `w:updateMode`, an absent relationship ID remains
its own stored-evidence class and an absent or unexpected mode remains
unsupported-or-missing evidence. Only exact `w:updateMode="always"` and
`w:updateMode="onCall"` receive the named automatic and on-call stored-mode
classes; none is a statement that a client will perform an update.

Public output exposes only aggregate marker/story, update-mode, and
relationship-classification counts. Program and shape identifiers,
relationship IDs and targets, field codes, locking metadata, markup, story
paths, and fingerprints never appear in reports. The full direct marker is
privately fingerprinted, so same-count program, field-code, locking,
update-mode, or target rewrites remain visible. A relationship-ID renumbering
with unchanged relationship semantics stays quiet.

`require_no_word_object_links` fails whenever a candidate contains a stored
direct WordprocessingML linked-object-property marker. Use it for a handoff
that must carry no such markup. `no_word_object_link_changes` compares the
private inventory signature for a controlled baseline. Neither rule resolves,
retrieves, opens, updates, activates, evaluates, renders, or executes an OLE
object, or establishes that Word will honor it.

## WordprocessingML embedded-control-anchor scope

Direct WordprocessingML `w:control` markers are counted only when they are
direct children of `w:object` or `w:pict` in a supported Word story. The public
inventory reports those two direct parent positions separately; arbitrary
`w:control` elements elsewhere are not treated as anchors. This is separate
from `w:objectLink`, `w:objectEmbed`, legacy Office VML linked- and
embedded-OLE markup, VML image data and shapes, fields, ActiveX-binary
relationships, and generic embedded OLE/package/control relationship or
payload totals. DocFence retains duplicates and markers in Markup Compatibility
branches, but does not select a client-rendered branch, associate a marker with
a visual control, or deduplicate an object.

`r:id` is optional. An anchor without one remains a separate stored-evidence
class; when it exists, only a standard control relationship is recognized. An
internal target mode receives the internal-standard class. An external target
mode is shown separately because the OOXML Embedded Control Persistence Part
requires an internal target; any other resolved type or mode is unsupported
evidence. These classifications report stored relationship semantics, not
whether a control is installed, enabled, loaded, safe, or honored by Word.

Public output exposes only aggregate anchor/story, direct-parent, and
relationship-classification counts. Control names and shape identifiers,
relationship IDs and targets, markup, story paths, and fingerprints never
appear in reports. The full direct marker is privately fingerprinted, so
same-count name, shape, or target rewrites remain visible while a
relationship-ID renumbering with unchanged semantics remains quiet.

`require_no_word_embedded_controls` fails whenever a candidate contains a
stored direct WordprocessingML embedded-control anchor. Use it for a handoff
that must carry no such markup. `no_word_embedded_control_changes` compares
the private inventory signature for a controlled baseline. Neither rule
resolves, retrieves, opens, instantiates, loads, activates, evaluates, renders,
or executes a control, or establishes that Word will honor it.

## External Word document dependency scope

DocFence distinguishes three standardized external-document dependencies from
the generic external-relationship inventory:

- an `attachedTemplate` relationship from a Document Settings part and its
  direct `w:attachedTemplate` anchor;
- a `subDocument` relationship from the main document and its `w:subDoc`
  anchor; and
- a `frame` relationship from a linked Web Settings part and its direct
  `w:frame/w:sourceFileName` anchor.

It discovers Settings and Web Settings parts through conventional or Strict
relationships from the main or glossary document, with the conventional
`word/settings.xml` Settings path retained as a compatibility fallback. The
expected relationship type and `TargetMode="External"` are required for every
recognized relationship and direct anchor. Invalid recognized state fails
closed. A recognized relationship is still counted without an anchor, because
it retains a stored external target that a handoff may need to review.

`require_no_external_document_dependencies` fails when any of the six public
counts is nonzero. `no_external_document_dependency_changes` compares private
relationship semantics and anchor markup. For a linked Web Settings part with
frame dependency state, it fingerprints the complete part privately so an
otherwise cosmetic frame change remains review-visible while relationship-ID
renumbering alone stays quiet. Public reports never show a target, relationship
ID, part path, frame name, or fingerprint.

These rules do not retrieve a template, open a subdocument, resolve a frame,
render any external content, or determine whether Word will reach the target.
They are stored-state controls, not a network or malware-analysis engine.

## CI example

```bash
docfence init docfence.yml
docfence check approved.docx candidate.docx \
  --policy docfence.yml \
  --format sarif \
  --output docfence.sarif
```

Exit status `1` means the report was produced and contains findings. Exit status
`2` means DocFence could not establish bounded evidence, so a CI workflow should
treat it as an infrastructure or input-validation failure rather than a clean
comparison.
