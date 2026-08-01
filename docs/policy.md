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
mail-merge configuration and recipient state. `no_data_binding_changes` does
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
one direct `SignedInfo` and `SignatureValue`, and the basic direct
canonicalization/signature-method/reference shape. Malformed recognized
topology or XMLDSIG shape fails closed.

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
identity, determines what content is covered, assesses an Office client, or
decides whether a signature should be trusted.

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

`HYPERLINK` field codes are distinct from ordinary `w:hyperlink` relationship
markup and from the external-source field inventory: a complete field code can
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
