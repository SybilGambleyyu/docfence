# DocFence

DocFence is a local-first change-assurance CLI for Word `.docx`, `.docm`,
`.dotx`, and `.dotm` files. It turns an opaque document diff into a
reviewable, privacy-safe account of stored content-block changes and the review
surfaces that often stay hidden:
tracked revisions, comments, hidden runs and paragraph marks, stored
style/default declarations, field codes, external-source field instructions,
`HYPERLINK` field references, direct `w:hyperlink` markup, external
relationships, custom XML, macros,
core/extended/custom document
properties, Microsoft Purview sensitivity-label metadata, mail-merge
configuration and recipient-data state, OPC package digital-signature material,
Word editing/write-protection state, document-variable state, editable-range
permission markup, document-variable field references, and
password-verifier material,
data-bound
content controls and their referenced custom XML state, external document
dependencies, modern comment contact/thread/identifier/reaction metadata,
document-task workflow state, task-pane Office web-extension configuration,
embedded OLE/package/control
payloads, OOXML alternative-format imports, headers, footers, notes, document
settings, and otherwise unclassified package payloads.

It never opens Word, executes macros, follows links, renders a document, uploads
source material, or writes a redline. By default, reports contain counts,
story categories, and change categories—not document text, reviewer
names, comments, URLs, relationship targets, field instructions, custom XML,
macro bytes, or modern-comment contact records and identifiers.
It also never emits document-task identities, user records, titles, dates, or
comment anchor IDs; Office web-extension IDs, stores, properties, bindings, or
part paths stay private as well. Sensitivity-label and tenant IDs, names, dates,
action IDs, extension data, custom MIP attributes, and content-marking strings
are private too.
Package signer/certificate material, signing times and comments, signature
values, signed-reference targets, and signature-part paths are private too.
Word protection hashes, salts, verifier values, cryptographic provider and
algorithm fields, and Settings-part paths are private too.
Word document-variable names, values, and Settings-part paths are private too.
DOCVARIABLE field instructions, literal field arguments, and story-part paths
are private too.
HYPERLINK field instructions, destinations, internal locations, ScreenTips,
frame targets, and story-part paths are private too.
Direct WordprocessingML hyperlink targets, anchors, locations, tooltips, frame
names, history values, display text, relationship IDs, and story-part paths are
private too.
Editable-range marker IDs, individual editor identities, and exact table-column
selectors are private too.

```bash
python -m pip install docfence

docfence diff approved.docx candidate.docx --format markdown
docfence init docfence.yml
docfence check approved.docx candidate.docx --policy docfence.yml --format sarif --output results.sarif
```

`profile` inventories one package. `diff` emits a report but never fails merely
because a change exists. `check` applies a policy and exits `1` when it finds a
violation, `0` when the policy passes, and `2` for an invalid policy, unreadable
input, or a safety-boundary failure. Output is JSON by default; Markdown and
SARIF are also available for comparisons. `--output` writes atomically.

## Why it exists

Word can make a legal blackline and its Document Inspector can identify hidden
content, but both are interactive review workflows. In CI or a controlled
handoff, teams also need a reproducible statement of whether a document gained
unresolved revisions, hidden text, comments, external targets, macros, or other
review-sensitive package state.

DocFence treats that state as a local, versionable boundary. Its `check`
command can make selected changes fail closed, starting with `docfence init`.

## Current boundary

Version 0.18 focuses on Office Open XML Word documents and templates and
deliberately keeps a small, inspectable contract:

- bounded `.docx` / `.docm` / `.dotx` / `.dotm` ZIP packages;
- body, header, footer, footnote, endnote, comment, and glossary stories;
- paragraph/table block fingerprints that ignore Word's volatile `rsid`
  bookkeeping while retaining stored text and formatting semantics privately;
- revision markup, comments, direct hidden-text runs, direct hidden
  paragraph-mark markup, stored style/default hidden-text declarations,
  field-code, `HYPERLINK` field-reference, direct `w:hyperlink` markup,
  external-source field,
  content-control, external-relationship,
  custom-XML, macro,
  core/extended/custom document-property, sensitivity-label metadata, OPC
  package digital-signature material, Word editing/write-protection,
  document-variable state and `DOCVARIABLE` field references,
  editable-range permission markup,
  mail-merge, content-control
  data-binding, modern-comment metadata,
  document-task workflow state, task-pane Office web-extension configuration
  and content-control binding markers,
  attached-template/master-subdocument/frameset-source external document
  dependency, embedded OLE/package/control, alternative-format-import,
  and Track Changes inventories;
- `w:altChunk` anchors paired with an internal OOXML alternative-format-import
  relationship and its stored payload;
- a generic alert for changed package payload outside those specialized
  inventories (for example, media, thumbnails, or other opaque parts);
- JSON, Markdown, and SARIF output with content redaction by design.

The reader rejects symlinks, encrypted ZIP entries, duplicate or colliding
member names, unsupported compression, oversized/over-expanded packages, DTDs,
entity declarations, excessively deep XML, and malformed required OOXML
structures. Default ceilings are 128 MiB source size, 4,096 members, 64 MiB per
expanded member, 512 MiB total expansion, 16 MiB XML payload, 200,000 XML
elements, and 256 XML levels. See the [threat model](docs/threat-model.md) for
the exact trust boundary.

DocFence is not a Word renderer, a document-calculation engine, an authoring
tool, or a replacement for legal review. It does not accept or reject changes,
evaluate fields, resolve or retrieve targets, authenticate external content, or
perform style resolution. `hidden_text_run_count` covers direct `w:vanish`
markup on ordinary runs. `hidden_paragraph_mark_count` covers direct
`w:vanish` or `w:specVanish` markup under a paragraph's mark properties. The
style inventory reports how many stored style definitions contain enabled
text-run `w:vanish` declarations and whether document-default run properties do
so. It is deliberately not an effective-format calculation: it does not say
which styles are used, resolve `basedOn`, table/numbering styles or toggle
semantics, or claim that any individual run will render hidden. Stored markup is
fingerprinted rather than rendered. Treat any report as bounded evidence about
the package, not a statement that Word would render an identical view.

DocFence also never opens, executes, imports, renders, or judges the safety of
an embedded or alternative-format payload. It inventories standard OLE,
package, control, ActiveX-control-binary, and `aFChunk` relationship types and
the conventional `word/embeddings/` and `word/activeX/` payload folders. An
encountered `w:altChunk` must reference an internal standard `aFChunk`
relationship; the target must resolve to a stored package member or parsing
fails closed. The payload itself remains opaque and private.

Document properties are likewise recorded without exposing their names or
values. DocFence recognizes the standard core, extended, and custom property
relationship types plus the canonical `docProps/core.xml`, `docProps/app.xml`,
and `docProps/custom.xml` paths. Core and extended value counts report direct
property elements containing stored text, including automatic dates, statistics,
and application metadata. The custom count reports stored custom-property
definitions even when a value is empty. It does not decide whether any value is
personal, confidential, user-authored, or safe to share.

Sensitivity-label metadata has a separate inventory because a general custom
property count does not describe the governance state that Office can retain.
DocFence recognizes the Office 2021 `LabelInfo` part by its standard package
classification-label relationship, its SDK content type, or conventional
`docMetadata/LabelInfo` paths. A recognized relationship must originate in the
package relationship item, be internal, and resolve to a stored member; its
`labelList` root and required label state are validated. It also recognizes
legacy `MSIP_Label_<GUID>_<attribute>` custom properties, the legacy
`Sensitivity` property, and documented Word content-marking property names.
Reports expose only aggregate LabelInfo part/label/enabled/removed/extension,
legacy MIP label/property, legacy sensitivity-property, and Word
content-marking-property counts. Label and tenant IDs, label names, methods,
dates, action IDs, extension content, custom MIP attributes, property names and
values, and part paths are privately fingerprinted only. DocFence does not
decrypt a protected file, read an encrypted LabelInfo stream, resolve a label
policy, determine effective permissions, apply or remove a label, or predict
which markings an Office client will display.

OPC package digital signatures have their own inventory because a generic
opaque-payload change cannot explain whether a package retained signature
origin, XML signature, or certificate material. DocFence recognizes the
standard root-package digital-signature-origin relationship, exact OPC content
types (including default content types), and the conventional
`_xmlsignatures/origin.sigs` residue path. A recognized origin relationship must
be root-package scoped, internal, resolve to a stored member, and occur at most
once. Recognized signature and certificate relationships must originate at the
expected preceding part, be internal, resolve to a stored member, and carry the
expected content type. XML signature parts must have the expected XMLDSIG root
and basic SignedInfo shape.

Reports expose only aggregate origin-part, XML-signature-part, certificate-part,
SignedInfo-reference, manifest-reference, relationship-reference, inline-X.509
certificate, and signature-property counts. Signer/certificate data, signature
values, algorithms, signing times, comments, provider data, reference URIs,
relationship IDs, part paths, and fingerprints remain private. The full
recognized material is privately digested, so a same-count signature or
certificate mutation remains review-visible. This is deliberately not a
cryptographic verifier: DocFence does not validate a signature value,
certificate chain, revocation or timestamp, signer identity, signing policy,
what a signature covers, or whether a consumer should trust it.

Word editing and write-protection state has its own inventory because a generic
Settings-part fingerprint cannot tell a reviewer whether a document retained an
editing restriction, a read-only recommendation, or password-verifier material.
DocFence discovers document Settings parts through Word's conventional
`word/settings.xml` fallback and recognized Transitional or Strict settings
relationships from the main or glossary document. It recognizes direct
`w:documentProtection` and `w:writeProtection` leaves, permits at most one of
each per Settings part, and rejects malformed recognized element shape,
attributes, edit modes, or boolean values.

Reports expose only aggregate protection-element, explicitly-enabled
enforcement, formatting-restriction, edit-mode, read-only-recommendation, and
password-material counts. Password hashes, salts, verifier values, provider and
algorithm fields, all other attributes, paths, and fingerprints remain private.
The full direct protection elements are privately fingerprinted, so a
same-count verifier rewrite remains review-visible. This is stored package
state—not encryption or an assurance that Word will enforce a restriction.
DocFence does not validate password construction or strength, try or recover a
password, determine effective enforcement, bypass a restriction, or make a
security claim about either protection feature.

Editable-range permission markup has a separate inventory because a protected
Word document can retain exceptions for selected text or table columns. A
`w:permStart` marker can store an individual editor in `w:ed` or a predefined
group in `w:edGrp`, with the corresponding `w:permEnd` linked by its marker ID.
DocFence recognizes these markers across every supported document story,
including Transitional and Strict Word namespaces. It validates their leaf
shape, required IDs, known attributes, predefined group values, nonnegative
column-selector syntax, placement values, and unambiguous IDs within a story.

Reports expose only aggregate start/end, paired/unpaired, individual-editor,
predefined-group, table-column-selector, and custom-XML-placement counts. Raw
editor identities, marker IDs, exact columns, part paths, and fingerprints stay
private. The full marker shape is privately fingerprinted so a same-count
identity or range rewrite remains review-visible. `w:ed` and `w:edGrp` are
stored attributes, not evidence that an identity is authenticated or presently
authorized; when both appear, Word behavior can prefer the individual editor.
Unmatched markers are counted as stored review state, not presented as an
effective permission. `require_no_word_permission_ranges` fails for any stored
range-permission marker, while `no_word_permission_range_changes` protects an
approved baseline.

This inventory does not authenticate editors, resolve groups, determine whether
a client will honor a marker, calculate the exact editable text or table cells,
or decide whether a document is secure.

Word document variables are another hidden Settings-part surface. A direct
`w:docVars` container stores `w:docVar` name/value pairs that Word can retain
for document or template automation. DocFence reports only aggregate container,
variable, and empty-value counts; variable names, values, paths, and
fingerprints stay private. The full recognized container is privately
fingerprinted, so a same-count name or value rewrite remains review-visible.

`require_no_word_document_variables` fails for any recognized `w:docVars`
container, including an empty one. `no_word_document_variable_changes` protects
an approved baseline.

`DOCVARIABLE` fields are independently inventoried from those stored values.
DocFence recognizes complete simple `w:fldSimple` instructions and complete
complex pre-separator field instructions across every supported story, retaining
current and deleted revision variants as separate stored evidence. Public output
contains only total/reference-story counts plus conservative literal/nonliteral
and exact-literal association counts. A literal association means the leading
field argument was a plain token or one complete quoted string (with optional
trailing Word field switches) and exactly matches a discovered `w:docVar` in
the same main or glossary package document scope. It does not mean Word will
render a value: an unmatched literal can be provided by an attached template,
and a nested or compound field expression is reported as nonliteral rather than
interpreted.

`require_no_word_document_variable_fields` fails for every stored
`DOCVARIABLE` reference, while `no_word_document_variable_field_changes`
protects an approved field-reference baseline. DocFence does not evaluate a
field, run macros, resolve a template, or claim that a stored variable is used,
visible, or safe.

`HYPERLINK` fields are independently inventoried from relationship hyperlinks
and external-source field families. A complete field code can carry a URL, file
location, or bookmark directly, so an OOXML relationship inventory cannot see
every stored target. Public output reports only total/reference-story counts
and one mutually exclusive lexical class per complete field: a literal leading
destination, a literal `\l` internal-location-only target, or a dynamic or
unparseable field. A leading literal accepts a plain token or a wholly quoted
string with optional trailing field-switch material; nested or compound
expressions remain dynamic or unparseable.

A literal destination is deliberately not labeled external: Word permits a
bookmark in the primary field argument, while `\l` denotes the documented
in-file location form. Destinations, bookmarks, ScreenTips, frame targets,
field instructions, story paths, and fingerprints stay private. The inventory
does not resolve, follow, evaluate, or render a field; it does not establish
that a target is reachable, safe, or shown by Word.

`require_no_word_hyperlink_fields` fails for every stored complete `HYPERLINK`
reference, while `no_word_hyperlink_field_changes` protects an approved
field-reference baseline.

Direct WordprocessingML `w:hyperlink` markup is inventoried separately from
both field codes and the broad relationship totals. A direct element can refer
to a relationship through `r:id`, name a local anchor through `w:anchor`, or
omit both (the documented default is the start of the current document). When
both attributes occur, `r:id` takes precedence, so the anchor is counted as
stored shadowed evidence rather than a second target. Relationship-backed
elements are classified by the stored relationship mode as external, internal,
or unsupported; a relationship is not counted merely because it exists without
a direct `w:hyperlink` reference.

The markup inventory keeps relationship targets, anchors, `w:docLocation`,
tooltips, frame names, display text, relationship IDs, story paths, and
fingerprints private. Its private signature catches same-count target or
attribute changes while treating a relationship-ID renumbering with identical
semantics as unchanged. It never resolves, retrieves, follows, validates,
evaluates, or renders a link, and an external relationship count is not a
claim that a target is reachable or safe.

`require_no_word_hyperlink_markup` fails for every stored direct
`w:hyperlink` element, while `no_word_hyperlink_markup_changes` protects an
approved direct-markup baseline.

Mail-merge state is also recorded without exposing connection strings, SQL
queries, field mappings, source/header targets, or recipient data. DocFence
counts stored `w:mailMerge` configuration, external data and header source
relationships, and internal recipient-data relationships and parts. It checks
that direct source/header/recipient references use the expected relationship
types and target modes, but does not connect to a data source, run a query, or
interpret recipient records.

Data-bound content controls are recorded separately from the general content
control and custom-XML inventories. DocFence recognizes direct standard
`w:sdtPr/w:dataBinding` declarations and safely associates a nonempty
`w:storeItemID` with an in-package custom XML data part only through its
standard internal custom-XML-properties relationship. Reports show aggregate
binding, identifier-presence, referenced-part, and unmatched-identifier counts
only. They never include XPath expressions, namespace-prefix mappings, storage
IDs, part names, or custom XML values. A binding without a storage ID is still
reported, but DocFence deliberately does not guess which part Word might choose
from its XPath. It does not evaluate XPath, update a control, or render the
result.

External-source Word fields are recorded separately from the general field-code
count. DocFence recognizes `DATABASE`, legacy `DATA`, `DDE`, `DDEAUTO`,
`INCLUDE`/`INCLUDETEXT`, `INCLUDEPICTURE`/`IMPORT`, `LINK`, and `RD` field
instructions in every supported story. It handles both OOXML encodings: a
simple field's `w:instr` attribute and a complete complex field's concatenated
pre-separator `w:instrText` runs. When tracked revisions retain a deleted field
code in `w:delInstrText`, DocFence inventories the current and deleted complex
instruction variants separately. The public inventory contains only category
counts; field instructions, source paths, connection strings, queries,
application names, item references, and private fingerprints never appear in a
report. Instruction text outside a complete complex field's instruction portion
is not treated as a field instruction, and an unclosed complex field is not
counted. The field keyword is a review signal, not a parser or evaluator:
DocFence does not interpret its arguments, update the field, locate a source,
run a query, start DDE, open an OLE object, or fetch any content.

Modern Word comments can retain review metadata outside the ordinary
`word/comments.xml` story. DocFence recognizes the standard `people`,
`commentsExtended`, `commentsIds`, and `commentsExtensible` parts by their
standard content types, relationship types, or conventional paths. It accepts
the established Office 15 `2010/11` and current `2012` root vocabularies for
people/thread metadata, validates every recognized root, and requires a
recognized metadata relationship to be internal and resolve to a stored part.
Reports contain aggregate part, contact-record, thread/reply, resolved-state,
identifier-record, reaction, and reaction-user counts only. Author names,
provider and user IDs, paragraph and durable IDs, dates, extension data,
reaction details, part paths, and private fingerprints never leave the process.
The inventory does not render comments, resolve a person, synchronize with a
cloud service, infer notification behavior, or modify comment state.

Document tasks are likewise stored review/workflow state, not live task
objects. DocFence discovers the Office document-tasks part through its standard
content type, relationship type, or conventional `word/tasks` / `word/tasks.xml`
path, validates
its `Tasks` root, and reports aggregate task, history-event, user-reference,
comment-anchor, and event-category counts. Task IDs, event IDs and dates,
user identities, task titles, schedules, progress, priorities, and comment IDs
are privately fingerprinted only. It does not assign, synchronize, notify,
complete, create, or otherwise operate on a task.

Task-pane Office web-extension state is a separate inventory. DocFence
recognizes task-pane and web-extension parts through their standard content
types, relationship types, and conventional `word/webextensions/` paths;
validates each root; and requires a task pane's direct web-extension reference
to resolve through the expected internal relationship. It reports only
aggregate task-pane, visible/locked-pane, extension-part, store-reference,
property, binding, enabled auto-show-property, and Word content-control binding
counts. It accepts the established `webextension` task-pane reference spelling
alongside `webextensionref`. A `w15:webExtensionCreated` marker takes precedence
over `w15:webExtensionLinked` for the content-control count. DocFence does not
install, execute, retrieve, authenticate, validate a manifest for, or claim
that Word will open an Office add-in. An enabled
`Office.AutoShowTaskpaneWithDocument` property is stored-state evidence only.

External Word document dependencies are recorded separately from the generic
external-relationship count. DocFence recognizes an attached template in a
Document Settings part, `w:subDoc` anchors in the main document, and
`w:sourceFileName` anchors inside frames in a linked Web Settings part. It
discovers Settings and Web Settings parts from either the main or glossary
document, retaining Word's conventional `word/settings.xml` fallback. Reports
show only paired anchor and relationship counts for attached templates,
subdocuments, and frame sources; they never include targets, relationship IDs,
part paths, frame names, or fingerprints. Recognized dependency relationships
and their direct anchors must use the expected type and `TargetMode="External"`
or parsing fails closed. A residual recognized relationship remains visible even
when no current anchor names it. DocFence never retrieves, opens, imports,
renders, authenticates to, or judges the safety of any dependency target.

## Policy

Policies are a deliberately small YAML subset (or equivalent JSON), with one
boolean switch per rule. The generated starter policy is suitable for a clean
handoff, while the complete supported rule set is documented in
[docs/policy.md](docs/policy.md).

```yaml
version: 1
rules:
  no_external_relationship_changes: true
  no_macro_payload_changes: true
  no_custom_xml_changes: true
  require_no_unresolved_revisions: true
  require_no_comments: true
  require_no_hidden_text: true
```

For stricter template or publishing gates, add the following under `rules:`
rather than assuming the run count resolves Word's style hierarchy:

```yaml
  require_no_hidden_text_style_declarations: true
  require_no_hidden_paragraph_marks: true
  require_no_embedded_objects: true
  require_no_alternative_format_imports: true
  require_no_custom_document_properties: true
  require_no_mail_merge: true
  require_no_data_bindings: true
  require_no_external_fields: true
  require_no_modern_comment_metadata: true
  require_no_document_tasks: true
  require_no_taskpane_web_extensions: true
  require_no_sensitivity_label_metadata: true
  require_no_package_digital_signatures: true
  require_no_word_protection: true
  require_no_word_permission_ranges: true
  require_no_word_document_variables: true
  require_no_word_document_variable_fields: true
  require_no_word_hyperlink_fields: true
  require_no_word_hyperlink_markup: true
  require_no_external_document_dependencies: true
```

When an established template intentionally contains one of those stored states,
use the corresponding comparison gates to permit the known baseline but fail a
later mutation:

```yaml
  no_embedded_object_payload_changes: true
  no_alternative_format_import_changes: true
  no_document_property_changes: true
  no_mail_merge_changes: true
  no_data_binding_changes: true
  no_external_field_changes: true
  no_modern_comment_metadata_changes: true
  no_document_task_changes: true
  no_taskpane_web_extension_changes: true
  no_sensitivity_label_metadata_changes: true
  no_package_digital_signature_changes: true
  no_word_protection_changes: true
  no_word_permission_range_changes: true
  no_word_document_variable_changes: true
  no_word_document_variable_field_changes: true
  no_word_hyperlink_field_changes: true
  no_word_hyperlink_markup_changes: true
  no_external_document_dependency_changes: true
```

YAML anchors, aliases, sequences, nested mappings, duplicate keys, unknown
rules, and non-boolean values are rejected. That keeps a policy reviewable and
avoids making the CLI's safety contract depend on a broad YAML loader.

## Privacy contract

DocFence keeps document material in memory only long enough to create private
SHA-256 fingerprints. Public models and renderers intentionally omit package
part names, paragraph content, reviewer identity, dates, comment content,
relationship targets, field instructions, style identifiers and names, custom
XML values, document-property names and values, mail-merge configuration and
recipient data, data-binding XPath expressions, prefix mappings, storage IDs,
referenced custom XML values, macro bytes, embedded and imported payload bytes,
external template/subdocument/frame-source targets, and all fingerprints.
Modern-comment author/contact/provider identifiers, paragraph and durable IDs,
timestamps, thread associations, and reaction identities receive the same
private treatment.
Document-task IDs, event IDs and times, user identities, titles, dates,
progress, priorities, and comment anchors are private too. So are task-pane
layout details and Office web-extension IDs, stores, reference versions,
property names and values, binding identifiers and application references,
content-control markers, and part paths.
Sensitivity-label IDs, tenant site IDs, label names, methods, set dates, action
IDs, label-extension data, legacy MIP custom attributes, and Word content
marking values remain private as well.
Package-signature signer and certificate material, signature values, reference
URIs, signing times, comments, provider data, relationship IDs, and part paths
remain private as well.
Word protection hashes, salts, verifier values, cryptographic provider and
algorithm fields, and Settings-part paths remain private as well.
Editable-range marker IDs, individual editor identities, exact table-column
selectors, and part paths remain private as well.
Regression tests place unique sensitive markers in each of those surfaces and
assert that JSON, Markdown, and SARIF never reproduce them.

This contract applies to DocFence's own reports. It cannot prevent a caller from
printing a source path, retaining a source document, or independently logging
process arguments.

## Evidence and sources

Microsoft documents both the [legal blackline comparison workflow](https://support.microsoft.com/en-us/office/compare-document-differences-using-the-legal-blackline-option-dbfc7351-4022-43a2-a0c4-54d1898702a0)
and the [hidden data surfaces found by Document Inspector](https://support.microsoft.com/en-us/office/collab-files/remove-hidden-data-and-personal-information-by-inspecting-documents-presentations-or-workbooks).
The rule boundary also follows the Open XML SDK's guidance on [WordprocessingML
revisions](https://learn.microsoft.com/en-us/office/open-xml/word/how-to-accept-all-revisions-in-a-word-processing-document)
and [markup compatibility](https://learn.microsoft.com/en-us/office/open-xml/general/introduction-to-markup-compatibility), plus its
[hidden-text semantics](https://learn.microsoft.com/en-us/office/open-xml/word/how-to-remove-hidden-text-from-a-word-processing-document)
and [`specVanish` paragraph-mark semantics](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.specvanish?view=openxml-3.0.1).
Microsoft also calls out [embedded files and objects](https://support.microsoft.com/en-us/excel/embedded-files-or-objects-found)
as an inspectable hidden-data surface. The alternative-import boundary follows
the Open XML SDK's [`w:altChunk` contract](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.altchunk?view=openxml-3.0.1)
and its package-part model for [embedded and import parts](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.packaging.maindocumentpart?view=openxml-3.0.1).
Microsoft's [document-property guidance](https://support.microsoft.com/en-us/office/view-or-change-the-properties-for-an-office-file-21d604c2-481e-4379-8e54-1dd4622c6b75)
and Document Inspector coverage inform the metadata boundary.
The mail-merge boundary follows the Open XML SDK's
[`w:odso` model](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.datasourceobject?view=openxml-3.0.1)
and Microsoft's documentation that accepting a linked mail-merge source can
[run its SQL query](https://support.microsoft.com/en-us/word/you-receive-the-opening-this-will-run-the-following-sql-command-message-when-you-open-a-word-mail-me).
The data-binding boundary follows the Open XML SDK's
[`w:dataBinding` contract](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.databinding?view=openxml-3.0.1)
and Word's [`XMLMapping` model](https://learn.microsoft.com/en-us/office/vba/api/word.xmlmapping),
which maps content-control text to document XML data. An independent
[OOXML reference-corpus fixture](https://loadfix.github.io/ooxml-reference-corpus/case/docx__custom-xml-part.html)
was profiled as a compatibility smoke test.
The external-source-field boundary follows OOXML's [simple and complex field
representations](https://ooxml.info/docs/17/17.16/17.16.2/) and the Open XML
SDK's [`instrText` rule](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.fieldcode?view=openxml-3.0.1),
which treats an instruction-text element outside a complex field's instruction
portion as ordinary text. The SDK also documents
[`delInstrText`](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.deletedfieldcode?view=openxml-3.0.1)
as the deleted complex-field-code representation. Microsoft's current field
documentation describes [database queries](https://support.microsoft.com/en-us/word/field-codes-database-field),
[included documents](https://support.microsoft.com/en-us/word/field-codes-includetext-field),
[linked pictures and the legacy IMPORT alias](https://support.microsoft.com/en-us/word/field-codes-includepicture-field),
[OLE links](https://support.microsoft.com/en-us/word/field-codes-link-field),
and [referenced documents](https://support.microsoft.com/en-us/word/field-codes-rd-referenced-document-field).
Microsoft's interoperability specifications document [DDE](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/a2c3a25a-1dba-40da-be7a-47cf63c78d55?redirectedfrom=MSDN),
`DDEAUTO`, and the legacy `DATA`, `INCLUDE`, and `IMPORT` aliases in its
[field-type catalog](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/28a8d2c2-6107-409d-8f6a-e345ab6d4179).
The external-document-dependency boundary follows the OOXML
[Document Template](https://c-rex.net/samples/ooxml/e1/Part1/OOXML_P1_Fundamentals_Document_topic_ID0E1IDK.html),
[Master Documents and Subdocuments](https://c-rex.net/samples/ooxml/e1/Part1/OOXML_P1_Fundamentals_Master_topic_ID0E2SDK.html),
and [Framesets](https://ooxml.info/docs/11/11.5/) contracts. Microsoft also
documents the attached-template relationship behavior, while
[MITRE ATT&CK T1221](https://attack.mitre.org/techniques/T1221/) describes
template injection as an abuse of document template references. A reconstructed
package from the open-source [XJTU thesis Office
template](https://github.com/obster-y/XJTU-thesis-Office/tree/master/%E6%A8%A1%E6%9D%BF%E4%BD%BF%E7%94%A8%E8%AF%B4%E6%98%8E-docx)
was profiled as an attached-template compatibility smoke test.
The modern-comment boundary follows Microsoft's [modern comments
overview](https://support.microsoft.com/en-US/Word/using-modern-comments-in-word),
the [people-part contract](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-docx/f461e6b7-7a35-4bc4-8153-b60f5d925539),
the [commentsExtended](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-docx/31f689cd-4192-4c2d-8d2f-202b1f8f20e9),
[commentsIds](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-docx/22977b5a-5bb5-4f27-b7a1-c6d216c2bb94),
and [commentsExtensible](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-docx/62c16828-8131-4d1f-99f8-afd7560a1c78)
part contracts. Microsoft's [reaction extension
example](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oreactxml/24d7d1f3-568d-4df9-89ef-42867776d742)
shows that reactions can carry user and time metadata. The Open XML SDK's
[Wordprocessing document types](https://github.com/dotnet/Open-XML-SDK/blob/main/src/DocumentFormat.OpenXml/WordprocessingDocumentType.cs)
also enumerate document, template, macro-enabled document, and macro-enabled
template packages.
The document-task boundary follows Microsoft's [document-task OOXML
contract](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-otaskxml/652d0608-31b8-4e90-a83a-98d6957b7fed).
The task-pane boundary follows the [Office web-extension OOXML
contract](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-owexml/29f59f30-b835-461a-bd8a-ca400a7bc717)
and Word's [`RemoveDocInfoType`](https://learn.microsoft.com/en-us/javascript/api/word/word.removedocinfotype?view=word-js-preview),
which lists both document tasks and task-pane web extensions as removable
document information. Microsoft's [task-pane auto-open
guidance](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/automatically-open-a-task-pane-with-a-document)
explains the stored auto-show setting; DocFence reports that setting without
predicting runtime behavior.
The sensitivity-label boundary follows Microsoft's
[Sensitivity Label Information part](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/c0599e21-b77f-475e-99e0-bd647f60bcbb),
[LabelInfo/custom-property precedence](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-offcrypto/13939de6-c833-44ab-b213-e0088bf02341),
and [sensitivity-label property contract](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/85388ac6-fb55-4017-828c-2680e3ab22ba).
Microsoft's [MIP SDK metadata guidance](https://learn.microsoft.com/en-us/information-protection/develop/concept-mip-metadata)
also documents legacy `MSIP_Label_` attributes and custom extensions; DocFence
keeps them private while comparing their stored state.
The package-signature boundary follows the OPC
[digital-signature model](https://learn.microsoft.com/en-us/previous-versions/windows/desktop/opc/open-packaging-conventions-overview)
and the [ECMA-376 Open Packaging Conventions standard](https://ecma-international.org/publications-and-standards/standards/ecma-376/).
Microsoft explicitly leaves signer identity and trust decisions to the package
consumer; DocFence therefore inventories structure without claiming signature
validity. Independent [OOXML-signature security research](https://www.usenix.org/conference/usenixsecurity23/presentation/rohlmann)
is further reason to keep that distinction explicit.
The Word-protection boundary follows the Open XML SDK's
[`w:documentProtection` contract](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.documentprotection?view=openxml-3.0.1)
and [`w:writeProtection` contract](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.writeprotection?view=openxml-3.0.1).
Both make the critical boundary explicit: editing/write protection is not
encryption or a security verdict. Microsoft also documents stored
[password-verifier salt](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.writeprotection.salt?view=openxml-3.0.1)
behavior, which is why DocFence keeps those fields out of reports.
The Word document-variable and `DOCVARIABLE`-field boundary follows the Open XML SDK's
[`w:docVar` contract](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.documentvariable?view=openxml-3.0.1)
and Word's [Variable object documentation](https://learn.microsoft.com/en-us/office/vba/api/word.variable):
the values persist with a document or template and are normally invisible until
a matching field is inserted. The OOXML [`DOCVARIABLE` field
definition](https://c-rex.net/samples/ooxml/e1/Part4/OOXML_P4_DOCX_DOCVARIABLEDOCVARIAB_topic_ID0EIEE1.html)
describes its field argument as the designated document variable. DocFence
therefore inventories the stored field code and only a conservative exact
literal association with discovered same-scope storage; it never emits,
resolves, or evaluates a variable name or value. Word's
[field-formatting guidance](https://support.microsoft.com/en-gb/office/format-field-results-baa61f5a-5636-4f11-ab4f-6c36ae43508c)
notes that `\* MERGEFORMAT` is inserted by default through the Field dialog;
DocFence retains the leading literal association while keeping that switch and
the whole field code private.
The `HYPERLINK` field boundary follows Microsoft's current [Word field
guidance](https://learn.microsoft.com/en-us/office/dev/add-ins/word/fields-guidance),
which describes links to same-document and external locations, and the OOXML
[`HYPERLINK` field definition](https://c-rex.net/samples/ooxml/e1/Part4/OOXML_P4_DOCX_HYPERLINKHYPERLINK_topic_ID0EFYG1.html),
which documents the primary argument and `\l` location switch. DocFence keeps
all those arguments private and reports stored lexical evidence only; a
literal primary argument is not treated as proof of an external target.
The direct WordprocessingML hyperlink-markup boundary follows the Open XML
SDK's [`w:hyperlink` contract](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.hyperlink?view=openxml-3.0.1)
and the ECMA-376 [`hyperlink` element definition](https://c-rex.net/samples/ooxml/e1/Part4/OOXML_P4_DOCX_hyperlink_topic_ID0EIMX1.html).
Those specifications distinguish an `r:id` relationship target from a local
`w:anchor`, give `r:id` precedence when both are present, and permit hyperlink
relationship targets inside or outside the package. DocFence reports only that
stored mechanism and mode evidence; it keeps target and markup values private.
The editable-range boundary follows the Open XML SDK's
[`w:permStart` contract](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.permstart?view=openxml-3.0.1)
and [`w:permEnd` contract](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.permend?view=openxml-3.0.1),
plus Microsoft's guidance on [allowing changes to selected protected-document
parts](https://support.microsoft.com/en-us/word/allow-changes-to-parts-of-a-protected-word-document),
which can grant editing to everyone or named individuals. The individual
identity is therefore review-sensitive package data, not report output.

## Development

```bash
python -m venv .venv
.venv/bin/python -m pip install -e '.[dev]'
.venv/bin/python -m pytest
.venv/bin/ruff check .
```

The CI workflow runs the test suite and linter on Python 3.11 and 3.13, then
builds and checks distribution artifacts. The current validation notes are in
[docs/validation.md](docs/validation.md).

MIT licensed. Contributions should preserve the privacy-first output contract.
