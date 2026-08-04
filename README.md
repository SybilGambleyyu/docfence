# DocFence

DocFence is a local-first change-assurance CLI for Word `.docx`, `.docm`,
`.dotx`, and `.dotm` files. It turns an opaque document diff into a
reviewable, privacy-safe account of stored content-block changes and the review
surfaces that often stay hidden:
tracked revisions, comments, hidden runs and paragraph marks, stored
style/default declarations, field codes, external-source field instructions,
`HYPERLINK` field references, direct `w:hyperlink` markup, DrawingML
click/hover/mouse-over hyperlink-action markup, direct DrawingML nonvisual
hidden-state declarations, direct DrawingML linked-picture `a:blip/@r:link`
markup, legacy VML shape `href` markup, legacy VML external
image-data `v:imagedata/@r:id` markup, legacy VML image-data hyperlink
`v:imagedata/@r:href` markup, legacy Office VML linked-OLE
`o:OLEObject Type="Link"` and embedded-OLE `o:OLEObject Type="Embed"` markup,
direct WordprocessingML linked-object-property `w:objectLink` markup, direct
WordprocessingML embedded-control-anchor
`w:control` markup, external
relationships, custom XML, macros,
core/extended/custom document
properties, Microsoft Purview sensitivity-label metadata, mail-merge
configuration and recipient-data state, custom XSLT-on-single-XML-save
configuration, attached custom XML schema declarations, automatic
field-recalculation-on-open settings, form-data-only-save settings,
preview-thumbnail-on-save settings, OPC package digital-signature material and
static declared package-signature coverage,
relationship-bound OPC package thumbnail images,
OOXML Markup Compatibility `mc:AlternateContent` branches and compatibility
rule attributes,
Word editing/write-protection state, document-variable state, editable-range
permission markup, document-variable field references, and
password-verifier material,
data-bound
content controls, their direct lock declarations, and their referenced custom
XML state, external document
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
OPC package thumbnail image bytes, relationship sources and targets, content
types, and part paths are private too.
OOXML Markup Compatibility branch bodies, feature-prefix and qualified-name
values, compatibility-rule values, and part paths are private too.
Word protection hashes, salts, verifier values, cryptographic provider and
algorithm fields, and Settings-part paths are private too.
Word document-variable names, values, and Settings-part paths are private too.
Attached custom XML schema namespace identifiers and Settings-part paths are
private too.
DOCVARIABLE field instructions, literal field arguments, and story-part paths
are private too.
HYPERLINK field instructions, destinations, internal locations, ScreenTips,
frame targets, and story-part paths are private too.
Direct WordprocessingML hyperlink targets, anchors, locations, tooltips, frame
names, history values, display text, relationship IDs, and story-part paths are
private too.
DrawingML hyperlink-action targets, invalid URLs, actions, tooltips, frame
names, history settings, relationship IDs, and story-part paths are private
too.
DrawingML nonvisual object names, descriptions, titles, IDs, exact hidden
attribute spellings, story-part paths, and fingerprints are private too.
DrawingML linked-picture targets, relationship IDs, surrounding drawing markup,
and story-part paths are private too.
Legacy VML shape-link URLs, target frames, titles, alternate text, shape
identifiers, and story-part paths are private too.
Legacy VML image and image-data-hyperlink targets, relationship IDs, raw
image-source values, other VML image-data attributes, and story-part paths are
private too.
Legacy VML linked- and embedded-OLE sources, monikers, program, shape, and
object identifiers, relationship IDs, update metadata, field codes, markup,
and story-part paths are private too.
WordprocessingML linked-object-property program and shape identifiers,
relationship IDs and targets, field codes, locking and update metadata, markup,
and story-part paths are private too.
WordprocessingML embedded-control names and shape identifiers, relationship IDs
and targets, markup, and story-part paths are private too.
Content-control IDs, aliases, tags, titles, placeholder text, current values,
and story-part paths are private too.
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

Version 0.45 focuses on Office Open XML Word documents and templates and
deliberately keeps a small, inspectable contract:

- bounded `.docx` / `.docm` / `.dotx` / `.dotm` ZIP packages;
- body, header, footer, footnote, endnote, comment, and glossary stories;
- paragraph/table block fingerprints that ignore Word's volatile `rsid`
  bookkeeping while retaining stored text and formatting semantics privately;
- revision markup, comments, direct hidden-text runs, direct hidden
  paragraph-mark markup, stored style/default hidden-text declarations,
  field-code, `HYPERLINK` field-reference, direct `w:hyperlink` markup,
  DrawingML click/hover/mouse-over hyperlink-action markup,
  direct DrawingML nonvisual hidden-state declarations,
  DrawingML `a:blip/@r:link` linked-picture markup,
  legacy VML shape/group/shape-template `href` markup,
  legacy VML `v:imagedata/@r:id` markup backed by an external relationship,
  legacy VML `v:imagedata/@r:href` hyperlink-target markup,
  legacy Office VML `o:OLEObject Type="Link"` linked-OLE markup,
  legacy Office VML `o:OLEObject Type="Embed"` embedded-OLE markup,
  direct WordprocessingML `w:object/w:objectLink` linked-object-property
  markup,
  direct WordprocessingML `w:object/w:control` and `w:pict/w:control`
  embedded-control-anchor markup,
  external-source field,
  content-control and direct content-control-lock declaration,
  external-relationship,
  custom-XML, macro,
  core/extended/custom document-property, relationship-bound OPC package
  thumbnail image, OOXML Markup Compatibility branch and compatibility-rule,
  sensitivity-label metadata, OPC package digital-signature material and
  static declared package-signature coverage, Word editing/write-protection,
  document-variable state and `DOCVARIABLE` field references,
  editable-range permission markup,
  mail-merge, XSLT-on-single-XML-save transform configuration, attached custom
  XML schema declarations, automatic field-recalculation-on-open configuration,
  automatic template-style-update-on-open configuration,
  personal-information-removal-on-save configuration,
  form-data-only-save configuration,
  preview-thumbnail-on-save configuration,
  content-control data-binding, modern-comment
  metadata,
  document-task workflow state, task-pane Office web-extension configuration
  and content-control binding markers,
  attached-template/master-subdocument/frameset-source external document
  dependency, embedded OLE/package/control, alternative-format-import,
  and Track Changes inventories;
- `w:altChunk` anchors paired with an internal OOXML alternative-format-import
  relationship and its stored payload;
- a generic alert for changed package payload outside those specialized
  inventories (for example, media or other opaque parts);
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

Direct Word embedded-control anchors are a separate stored-markup boundary from
those generic relationship and payload totals. DocFence records only direct
`w:control` children of `w:object` or `w:pict` in supported stories and does
not equate a package-level ActiveX/control relationship with a displayed
control anchor.

Direct Office VML embedded-OLE markers are likewise a separate stored-markup
boundary from generic embedded-object relationships and payloads. DocFence
records only Office VML `o:OLEObject Type="Embed"` evidence in supported Word
stories; it neither decodes a payload nor equates an embedded-object
relationship with a client-rendered object.

Document properties are likewise recorded without exposing their names or
values. DocFence recognizes the standard core, extended, and custom property
relationship types plus the canonical `docProps/core.xml`, `docProps/app.xml`,
and `docProps/custom.xml` paths. Core and extended value counts report direct
property elements containing stored text, including automatic dates, statistics,
and application metadata. The custom count reports stored custom-property
definitions even when a value is empty. It does not decide whether any value is
personal, confidential, user-authored, or safe to share.

OPC package thumbnails have their own inventory because a generic opaque-payload
change cannot distinguish a relationship-bound thumbnail image from unrelated
package residue. DocFence recognizes only the exact standard thumbnail
relationship type, in Transitional or Strict OOXML, from either the package or
a stored part. Each recognized relationship must be internal, resolve to a
stored target with an `image/` content type, and be the sole thumbnail
relationship for its source; a thumbnail target cannot have a relationship part
of its own. Malformed recognized topology fails closed. A filename such as
`docProps/thumbnail.png` without the standard relationship stays a generic
unclassified payload rather than becoming a guessed thumbnail.

Reports expose only thumbnail relationship and distinct-part counts. Image
bytes, relationship sources and targets, content types, and part paths remain
inside a private digest, so a same-count image mutation remains review-visible.
DocFence does not decode, render, classify, or otherwise interpret an image,
and it does not predict whether Word, Explorer, or another client will show it.

OOXML Markup Compatibility (MCE) has its own inventory because a generic
payload change cannot distinguish a stored runtime branch boundary from other
XML changes. Across stored `word/*.xml` members that use the standard MCE
namespace, DocFence reports only aggregate parts, `AlternateContent`, `Choice`,
and `Fallback` counts plus the token counts in `Requires`, `Ignorable`,
`MustUnderstand`, `ProcessContent`, `PreserveElements`, and
`PreserveAttributes`. A private signature retains the recognized MCE elements
and attributes, so an equal-count branch or compatibility-rule rewrite remains
review-visible without exposing branch material, prefix values, qualified names,
or part paths.

This is stored markup evidence only. DocFence does not validate MCE
conformance, select an `AlternateContent` branch, resolve a feature prefix,
apply a target Office version, preprocess or save a package, or predict what
any document client will load or render.

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
certificate mutation remains review-visible.

The companion static declared package-signature coverage inventory follows a
small, inspectable declaration chain. For each recognized XML signature, it
looks for a direct `SignedInfo` local-fragment reference to a direct
`ds:Object` with exactly one direct package `ds:Manifest`. From those bound
manifests it resolves exact local part URI/content-type references with
case-sensitive content-type matching. Before a manifest reference can be
credited, its direct XMLDSIG children must be the optional `ds:Transforms`, one
`ds:DigestMethod` with a nonblank `Algorithm`, and one direct, attribute-free,
child-free, nonempty `ds:DigestValue`, with no non-whitespace direct text, in
that order. It then recognizes standard OPC
relationship-transform declarations: one direct `ds:Transforms` list
containing only the supported OPC relationship and XML Canonicalization
algorithms, with exactly one relationship transform immediately followed by
XML Canonicalization (with or without comments), then exact
`RelationshipReference/@SourceId` and
`RelationshipsGroupReference/@SourceType` selectors. Within one XML signature,
it also rejects every transform-bearing declaration for a relationships part
when more than one relationship transform targets that same part across bound
manifests. Public output
contains only aggregate counts for signatures with and without that declaration
chain; covered and uncovered `word/` non-relationship parts; covered and
uncovered root `officeDocument` relationships; covered and uncovered
relationships sourced by Word parts; and unresolved or unsupported manifest
references. Object identifiers, manifest-reference URIs, part paths,
relationship identifiers and types, and digest material remain private. A
private semantic signature keeps a same-count selection reassignment
review-visible.

For a non-relationship package part, the audit accepts no transform list or one
direct nonempty list of only OPC-supported XML Canonicalization transforms. An
empty, duplicate, relationship, or unknown transform list is unsupported and
does not credit that part with coverage.

This remains deliberately narrower than cryptographic or client-effective
coverage. DocFence does not parse, decode, or recompute reference digests or
canonicalization,
evaluate arbitrary transforms, validate a signature value, certificate chain,
revocation or timestamp, establish signer identity, decide what an Office
consumer will validate or render, or decide whether a signature should be
trusted. It reports only bounded static declarations present in the package.

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

DrawingML hyperlink-action elements are a separate direct surface again: a
Word image or shape can contain `a:hlinkClick`, `a:hlinkHover`, or
`a:hlinkMouseOver` even when it has no `w:hyperlink` element or `HYPERLINK`
field. DocFence scans those direct DrawingML elements in every supported Word
story and counts stored click, hover, and mouse-over markers individually. It
does not collapse markers that share a relationship or a visual object, and it
does not choose a Markup Compatibility branch; the count is evidence of stored
markup, not a rendered-link count.

Markers with `r:id` are classified from the referenced relationship as
external, internal, or unsupported. A marker with no `r:id` is separately
counted as malformed stored evidence. An unreferenced hyperlink relationship is
not counted as DrawingML markup. Public output also notes the presence of
`action` and `invalidUrl` attributes, while keeping values, relationship
targets, tooltips, frame names, history settings, relationship IDs, story
paths, and fingerprints private. The private signature catches same-count
target or attribute rewrites and normalizes relationship-ID renumbering with
unchanged semantics.

`require_no_word_drawing_hyperlinks` fails for every stored direct DrawingML
hyperlink-action marker, while `no_word_drawing_hyperlink_changes` protects an
approved marker baseline. This inventory never resolves, retrieves, follows,
validates, evaluates, or renders an action, and no count establishes that a
target is reachable, safe, or honored by a Word client.

DrawingML nonvisual visibility is a separate stored review surface. A direct
hidden attribute on a supported nonvisual-property element can keep a DrawingML
object present while specifying that it is hidden. DocFence scans only direct
unqualified hidden attributes on supported Word-story forms: DrawingML main
and picture cNvPr, WordprocessingDrawing docPr, and Word 2010 w14, wpg, and
wps cNvPr. Transitional and Strict forms are both supported where OOXML
defines them. It counts each declaration, its story presence, canonical hidden
or explicitly-shown state, and invalid Boolean spellings. Object names,
descriptions, titles, IDs, exact invalid spellings, and story paths stay
private.

The private signature is intentionally limited to the supported element kind
and canonical hidden state, so equivalent XML Boolean spellings do not churn a
baseline while a same-count hidden-to-shown swap remains review-visible.
Malformed values remain private but are separately counted and rejected by the
candidate gate. The scan includes duplicate markers and stored MCE branches,
but never selects a branch, resolves object identity, calculates effective
visibility, applies layout, renders a drawing, or predicts a client view.
`require_no_hidden_drawing_objects` rejects stored hidden and malformed
supported declarations; `no_drawing_object_visibility_changes` protects an
approved declaration baseline. Neither rule says an object is effectively
visible or invisible in a particular Office client.

DrawingML linked pictures are a separate direct relationship surface again: an
`a:blip` with `r:link` stores a linked-picture reference, distinct from an
embedded-picture `r:embed` reference, a hyperlink action, a `w:hyperlink`
element, a `HYPERLINK` field, or broad relationship totals. DocFence scans each
direct `a:blip/@r:link` marker in every supported Word story. It does not
collapse markers that share a relationship or visual object and does not select
a Markup Compatibility branch; the count is stored-markup evidence, not a
rendered-picture or network-access count.

Each marker is classified from its resolved relationship as a standard image
relationship with stored external or internal target mode, or as an unsupported
relationship. An unreferenced image relationship and an `a:blip` carrying only
`r:embed` do not create linked-picture marker counts. Public output reports
only aggregate marker/story and relationship-classification counts. Image
targets, relationship IDs, surrounding drawing markup, story paths, and
fingerprints remain private. The private signature catches same-count target or
attribute rewrites and normalizes relationship-ID renumbering with unchanged
semantics. It neither retrieves an image nor determines that Word will load,
render, update, or honor a linked picture.

`require_no_word_drawing_linked_pictures` fails for every stored direct
DrawingML linked-picture marker, while
`no_word_drawing_linked_picture_changes` protects an approved marker baseline.

Legacy VML shape markup is a fifth direct link surface: a `v:shape`,
`v:roundrect`, `v:group`, or related VML geometry element can carry an
unqualified `href` directly, without a relationship, `w:hyperlink`,
`HYPERLINK` field, or DrawingML action. DocFence scans direct `href` attributes
only on documented VML geometry elements in supported Word stories: `arc`,
`curve`, `image`, `line`, `oval`, `polyline`, `rect`, `roundrect`, `shape`,
`group`, and `shapetype`. It reports aggregate element/story, concrete-shape,
group, shape-template, and `target`-attribute-presence counts. An empty direct
`href` remains stored-markup evidence.

This is deliberately a legacy-markup inventory, not an effective-link engine.
It does not infer a link inherited from a group or shape template, inspect
arbitrary VML attributes, select a Markup Compatibility branch, resolve,
retrieve, follow, validate, evaluate, render, or execute an action. URLs,
frame targets, titles, alternate text, IDs, story paths, and fingerprints stay
private. Its private signature catches same-count `href` or markup rewrites.
`require_no_word_vml_hyperlinks` fails for every supported direct VML marker,
while `no_word_vml_hyperlink_changes` protects an approved marker baseline; no
count establishes that a target is reachable, safe, or honored by a Word
client.

Legacy VML image data is a separate relationship-backed surface. A direct
`v:imagedata/@r:id` stores an explicit relationship to image data, and DocFence
records it only when the resolved relationship has stored `TargetMode=External`.
It counts a standard image relationship separately from another externally
stored relationship type, which remains reviewable as unsupported evidence.
Duplicate direct markers and markers in Markup Compatibility branches each
remain stored evidence; DocFence does not select a branch or deduplicate a
visual image.

Ordinary embedded VML images with internal relationships are intentionally not
external-image markers. Neither are an unreferenced image relationship, raw VML
`src`, `r:pict`, `r:href`, or `o:relid`; those are distinct legacy surfaces, not
fallback forms for this inventory. Public output exposes only aggregate
marker/story and relationship-classification counts. Image targets,
relationship IDs, source values, other VML image-data attributes, story paths,
and fingerprints remain private. Its private signature tracks the reviewed
external `r:id` relationship semantics: a same-count external target rewrite is
visible, relationship-ID renumbering with unchanged semantics is quiet, and a
raw-`src` rewrite does not become an external-image inventory change.

`require_no_word_vml_external_images` fails for every stored direct VML
external-image marker, while `no_word_vml_external_image_changes` protects an
approved external-image marker baseline. The inventory never resolves,
retrieves, renders, or updates an image, and no count establishes that a client
will load, reach, or honor a target.

VML image-data hyperlink markup is one more, deliberately separate
relationship-backed surface. The Open XML SDK names `v:imagedata/@r:href` an
explicit relationship to a hyperlink target. DocFence records every direct
stored `r:href` marker in supported Word stories, including duplicate markers
and markers in Markup Compatibility branches; it does not choose a branch or
associate a marker with a rendered image.

When the resolved relationship is a standard hyperlink relationship, the
inventory classifies its stored target mode as external or internal. A resolved
relationship of any other type or mode remains reviewable as unsupported
stored evidence rather than being assumed to be a conventional hyperlink. That
is intentional: public Word XML examples show the marker in real legacy VML,
while a public document-processing report shows an `r:href` tied to an external
standard image relationship. Neither case establishes how a particular Word
client will act on it.

This boundary does not treat image-data `r:id`, `r:pict`, raw `src`, or
`o:relid` as alternate hyperlink markers. Public output reports only aggregate
marker/story and external/internal/unsupported relationship-classification
counts; targets, relationship IDs, VML markup, story paths, and fingerprints
remain private. Its private signature tracks only the reviewed `r:href`
relationship semantics: same-count target changes remain visible,
relationship-ID renumbering with unchanged semantics is quiet, and rewrites of
the excluded image-data attributes do not become image-data-hyperlink inventory
changes.

`require_no_word_vml_image_hyperlinks` fails for every stored direct VML
image-data hyperlink marker, while `no_word_vml_image_hyperlink_changes`
protects an approved marker baseline. The inventory never resolves, retrieves,
follows, validates, evaluates, renders, or executes a target, and no count
establishes that a client will honor it.

Legacy Office VML linked-OLE markup is another deliberately separate direct
surface. DocFence records each direct `o:OLEObject` whose unqualified `Type`
is `Link` in supported Word stories, including duplicates and markers in
Markup Compatibility branches. It does not select a branch, associate the
marker with a rendered shape, deduplicate a visual object, retrieve a source,
or activate or update an OLE object.

If the marker has `r:id`, DocFence classifies a standard OLE-object relationship
by its stored external or internal target mode; every other resolved type or
mode remains reviewable as unsupported evidence. The OOXML schema makes `r:id`
optional, so a direct `Type="Link"` marker without it remains a separate
stored-evidence class. Public output also aggregates `UpdateMode="Always"` as
a stored automatic-update marker; all other or absent update values are grouped
as nonautomatic-or-unspecified. This is markup evidence, not proof that any
Word client will retrieve a source or perform an update.

The separately inventoried `Type="Embed"`, WordprocessingML `w:objectLink`,
VML image data, VML shape links, and broad embedded OLE/package/control
relationship or payload totals remain distinct. Public output reports only
aggregate marker/story, update, and relationship-classification counts.
Sources, monikers, program, shape, and object IDs, relationship IDs and
targets, field codes, VML markup, story paths, and fingerprints remain private.
The full direct marker is privately fingerprinted, so same-count source,
program, field-code, update-mode, or target rewrites remain visible;
relationship-ID renumbering with unchanged semantics is quiet.

`require_no_word_vml_linked_ole_objects` fails for every stored direct VML
linked-OLE marker, while `no_word_vml_linked_ole_object_changes` protects an
approved marker baseline. The inventory never resolves, retrieves, opens,
updates, activates, evaluates, renders, or executes an OLE object, and no count
establishes that a client will honor it.

Legacy Office VML embedded-OLE markup is a complementary direct surface.
DocFence records each Office VML `o:OLEObject` whose unqualified `Type` is
`Embed` in supported Word stories, including duplicates and markers in Markup
Compatibility branches. The Office VML contract permits several parent forms,
so the inventory follows the stored marker rather than assigning it a rendered
object position. It does not select a branch, associate a marker with a visual
object, deduplicate a visual object, decode a payload, retrieve a target, or
load or activate an OLE object.

If the marker has `r:id`, DocFence classifies a standard OLE-object
relationship by its stored external or internal target mode; every other
resolved type or mode remains reviewable as unsupported evidence. The OOXML
schema makes `r:id` optional, so a direct `Type="Embed"` marker without it
remains a separate stored-evidence class. `UpdateMode` is defined for the
`Type="Link"` form, so embedded-OLE output deliberately has no update-behavior
class. This is stored markup evidence, not proof that any Word client will open
or render an embedded object.

`Type="Link"`, WordprocessingML `w:objectEmbed`/`w:objectLink`, VML image data,
VML shape links, fields, and broad embedded OLE/package/control relationship
or payload totals remain separate inventories. Public output reports only
aggregate marker/story and relationship-classification counts. Program, shape,
and object IDs, relationship IDs and targets, update metadata, field codes,
VML markup, story paths, and fingerprints remain private. The full direct
marker is privately fingerprinted, so same-count program, field-code,
update-metadata, or target rewrites remain visible; relationship-ID renumbering
with unchanged semantics is quiet.

`require_no_word_vml_embedded_ole_objects` fails for every stored direct VML
embedded-OLE marker, while `no_word_vml_embedded_ole_object_changes` protects
an approved marker baseline. The inventory never resolves, retrieves, opens,
updates, activates, evaluates, renders, or executes an OLE object, and no count
establishes that a client will honor it.

WordprocessingML linked-object-property markup is a distinct, narrower
compatibility surface. DocFence records every direct `w:objectLink` child of a
`w:object` in supported Word stories, including duplicates and markers in
Markup Compatibility branches. It does not select a branch, associate a marker
with a rendered object, deduplicate a visual object, retrieve a source, or
activate or update an OLE object. `w:objectEmbed`, legacy Office VML
`o:OLEObject Type="Link"`/`Type="Embed"`, VML image data and shape links,
DrawingML linked pictures, fields, and generic embedded-object relationship or
payload totals remain separate inventories.

When an `objectLink` carries `r:id`, DocFence classifies a standard
OLE-object relationship by its stored external or internal target mode; every
other resolved type or mode remains reviewable as unsupported evidence. The
schema requires `r:id` and `w:updateMode`, but a direct marker missing the
former remains a separate stored-evidence class and an absent or unexpected
stored mode remains an unsupported-or-missing class. Only exact schema tokens
`w:updateMode="always"` and `w:updateMode="onCall"` receive the named
automatic and on-call aggregate counts. None of those counts predicts a client
update, retrieval, activation, or render outcome.

Public output exposes only aggregate marker/story, update-mode, and
relationship-classification counts. Program and shape identifiers, relationship
IDs and targets, field codes, locking metadata, markup, story paths, and
fingerprints remain private. The full direct marker is privately fingerprinted,
so same-count program, field-code, locking, update-mode, or target rewrites
remain visible; relationship-ID renumbering with unchanged relationship
semantics is quiet.

`require_no_word_object_links` fails for every stored direct
WordprocessingML linked-object-property marker, while
`no_word_object_link_changes` protects an approved marker baseline. The
inventory never resolves, retrieves, opens, updates, activates, evaluates,
renders, or executes an OLE object, and no count establishes that a client will
honor it.

WordprocessingML embedded-control anchors are a separate, direct compatibility
surface. DocFence records every direct `w:control` child of `w:object` or
`w:pict` in supported Word stories, including duplicates and markers in Markup
Compatibility branches. The two parent positions are reported separately. It
does not select a branch, associate a marker with a rendered shape, deduplicate
a visual control, instantiate or load a control, inspect its persistence data,
or predict a client outcome. Arbitrary `w:control` elements outside those two
direct parent positions, `w:objectLink`, `w:objectEmbed`, legacy Office VML
linked- and embedded-OLE markup, VML image data and shapes, ActiveX-binary
relationships, fields, and generic embedded-control relationship/payload totals
remain separate inventories.

When an anchor carries `r:id`, DocFence recognizes only a standard control
relationship. An internal target mode receives the standard internal class; an
external target mode is surfaced separately as stored nonconforming evidence,
because an Embedded Control Persistence Part must be internal. Another
resolved type or mode remains unsupported. `r:id` is optional, so an anchor
without it stays in a separate stored-evidence class rather than being
discarded. None of the classes says that a control is available, enabled, safe,
or loaded by a client.

Public output exposes only aggregate marker/story, direct-parent, and
relationship-classification counts. Control names and shape identifiers,
relationship IDs and targets, markup, story paths, and fingerprints remain
private. The full direct marker is privately fingerprinted, so same-count
control-name, shape, or target rewrites remain visible; a relationship-ID
renumbering with unchanged relationship semantics stays quiet.

`require_no_word_embedded_controls` fails for every stored direct
WordprocessingML embedded-control anchor, while
`no_word_embedded_control_changes` protects an approved marker baseline. The
inventory never resolves, retrieves, opens, instantiates, loads, activates,
evaluates, renders, or executes a control, and no count establishes that a
client will honor it.

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

Content-control lock declarations are separately inventoried across every
supported Word story. For each direct standard `w:sdt`, DocFence records only
which of five aggregate states applies: no direct `w:sdtPr/w:lock` declaration,
or a single exact `w:lock/@w:val` declaration of `unlocked`, `sdtLocked`,
`contentLocked`, or `sdtContentLocked`. It validates the direct leaf shape and
the four schema values, rejects duplicate direct properties or lock leaves, and
ignores lookalikes outside that direct parent-child path. The private signature
keeps each state associated with a story-local ordinal, so a same-count state
reassignment remains review-visible while control IDs, aliases, tags, titles,
placeholder text, values, and story paths remain absent from reports.

No direct declaration is deliberately not normalized to `unlocked`: the OOXML
contract gives omitted locks type-specific behavior for group controls. The
inventory therefore reports stored markup rather than an effective editing
decision. `require_content_control_locks` is an opt-in template gate that
requires each discovered content control to carry a direct non-`unlocked`
declaration. `no_content_control_lock_changes` protects an approved baseline.
Neither rule opens Word, changes a control, identifies a control owner, reads a
control's content, evaluates a data binding, applies document protection, or
claims that any Word-compatible client will enforce a declaration.

The generic custom-XML boundary is useful even when no content control maps to
it. `require_no_custom_xml_data` is an opt-in candidate gate for a clean
handoff: it fails when the package contains one or more non-relationship
members below its conventional `customXml/` folder. That includes both stored
data and associated data-properties XML parts, but not the relationship parts,
which remain covered by the separate package-relationship inventory. It does
not expose XML values, classify, remove, or decide whether custom XML is
sensitive, visible, used by Word, or safe to share.

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

XSLT-on-single-XML-save configuration is recorded separately from generic
Settings and external-relationship changes. `w:saveThroughXslt` can store a
custom transform for an application to use only when it saves a document as a
single XML file; `w:useXSLTWhenSaving` controls that optional behavior.
DocFence inventories direct settings leaves and standard `transform`
relationships from every discovered Document Settings part. A
relationship-backed anchor must use the expected relationship type and
`TargetMode="External"`; a local `w:solutionID`-only anchor is stored as
application-defined configuration evidence without any lookup. Reports contain
only enabled-setting, disabled-setting, anchor, relationship, and
solution-identifier counts;
transform targets, relationship IDs, solution identifiers, Settings-part paths,
and fingerprints remain private. `require_no_save_through_xslt` provides a
clean-handoff gate, while `no_save_through_xslt_changes` protects an approved
baseline. DocFence never resolves, retrieves, parses, executes, or applies an
XSLT, and it makes no claim that Word will save a single XML file, contact a
target, or produce any particular output.

Attached custom XML schema declarations are recorded separately from generic
Settings and custom-XML payload changes. A direct
`w:attachedSchema/@w:val` declaration identifies the target namespace of a
custom XML schema that a host may associate with the document when loading it,
if the schema is available to that host. Reports contain only an aggregate
declaration count; namespace values, Settings-part paths, and fingerprints stay
private. `require_no_attached_custom_xml_schemas` provides a clean-handoff
gate, while `no_attached_custom_xml_schema_changes` protects an approved
baseline, including same-count namespace rewrites. DocFence never resolves,
retrieves, loads, or validates against a declared schema, and it does not claim
that a host has a schema available or will validate any document markup.

Automatic field recalculation on open is recorded separately from generic
Settings and field-code changes. A direct `w:updateFields` `CT_OnOff` leaf asks
an application that supports field calculations to recalculate field results
when the document opens. DocFence accepts one direct leaf per discovered
Settings part in either Word namespace, validates its optional Word-namespace
`w:val` as a standard on/off token, and reports enabled and explicitly disabled
setting counts only. Settings-part paths and private fingerprints remain local.
The private signature retains the canonical stored state, so an enabled-to-
disabled transition remains review-visible while equivalent enabled spellings
such as omitted `w:val`, `on`, and `true` remain quiet.

`require_no_field_updates_on_open` fails only when a candidate explicitly
requests automatic recalculation. `no_field_update_on_open_changes` protects a
controlled baseline, including an explicitly disabled setting. Neither rule
parses or evaluates field instructions, updates field results, opens Word,
locates a field source, starts an application, follows a link, or claims that a
client will honor the request.

Automatic template-style updates on open are recorded separately from generic
Settings and external-document-dependency changes. A direct w:linkStyles
CT_OnOff leaf requests that a supporting document host update document styles
from an attached template when the document opens. DocFence accepts one direct
leaf per discovered Settings part in either Word namespace, validates its
optional Word-namespace w:val as a standard on/off token, and reports enabled
and explicitly disabled setting counts only. Settings-part paths and private
fingerprints remain local. The private signature retains canonical stored
state, so an enabled-to-disabled transition remains review-visible while
equivalent enabled spellings such as omitted w:val, on, and true remain quiet.

require_no_template_style_updates_on_open fails only when a candidate
explicitly requests automatic template-style updates. The separately reported
attached-template dependency inventory remains the evidence for a template
anchor and relationship; this setting inventory does not resolve either.
no_template_style_update_on_open_changes protects a controlled baseline,
including an explicitly disabled setting. Neither rule resolves, retrieves,
loads, opens, validates, or authenticates an attached template; performs style
resolution or propagation; opens a document client; or claims that a particular
client will honor the stored request.

Personal-information removal on save is recorded separately from generic
Settings and document-property changes. A direct
`w:removePersonalInformation` `CT_OnOff` leaf asks a capable document host to
remove personal information when it saves the document. DocFence accepts one
direct leaf per discovered Settings part in either Word namespace, validates
its optional Word-namespace `w:val` as a standard on/off token, and reports
enabled and explicitly disabled setting counts only. Settings-part paths and
private fingerprints remain local. The private signature retains canonical
stored state, so an enabled-to-disabled transition remains review-visible while
equivalent enabled spellings such as omitted `w:val`, `on`, and `true` remain
quiet.

`require_personal_information_removal_on_save` requires only an enabled stored
request. It does not inspect whether the package currently contains personal
information, define what counts as personal information, or prove that any
host will remove anything on a later save.
`no_personal_information_removal_on_save_changes` protects a controlled stored
baseline. Neither rule opens Word, saves a document, identifies authors,
rewrites properties, removes comments or revisions, or claims that a client
will honor the request.

Form-data-only save configuration is recorded separately from generic Settings
and field-code changes. A direct `w:saveFormsData` `CT_OnOff` leaf stores a
request for a capable document host to save only legacy form-field content as a
delimited record on a later save. DocFence accepts one direct leaf per
discovered Settings part in either Word namespace, validates its optional
Word-namespace `w:val` as a standard on/off token, and reports enabled and
explicitly disabled setting counts only. Settings-part paths and private
fingerprints remain local. The private signature retains canonical stored
state, so an enabled-to-disabled transition remains review-visible while
equivalent enabled spellings such as omitted `w:val`, `on`, and `true` remain
quiet.

`require_no_save_forms_data` fails only when a candidate explicitly requests
form-data-only saving. `no_save_forms_data_changes` protects a controlled
stored baseline, including an explicitly disabled setting. Neither rule finds
or evaluates form fields, reads field values, opens Word, saves a document,
emits a delimited record, determines a delimiter, or claims that a particular
client will honor the request.

Preview-thumbnail-on-save configuration is recorded separately from both
generic Settings changes and a currently stored OPC package thumbnail. A direct
`w:savePreviewPicture` `CT_OnOff` leaf stores a request for a capable document
host to generate a thumbnail of the first page when it saves the document.
DocFence accepts one direct leaf per discovered Settings part in either Word
namespace, validates its optional Word-namespace `w:val` as a standard on/off
token, and reports enabled and explicitly disabled setting counts only.
Settings-part paths and private fingerprints remain local. The private
signature retains canonical stored state, so an enabled-to-disabled transition
remains review-visible while equivalent enabled spellings such as omitted
`w:val`, `on`, and `true` remain quiet.

`require_no_save_preview_picture` rejects only an explicit enabled stored
request. `no_save_preview_picture_changes` protects a controlled stored
baseline, including an explicitly disabled setting. Neither rule proves that a
thumbnail is absent, prevents an application from choosing a thumbnail when
the setting is absent or disabled, decodes or renders an image, opens Word,
saves a document, creates a thumbnail, or claims that a particular client will
honor the request. The existing package-thumbnail inventory is the separate
evidence for an image already stored in a package.

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
  require_no_custom_xml_data: true
  require_no_package_thumbnails: true
  require_no_markup_compatibility: true
  require_no_embedded_objects: true
  require_no_alternative_format_imports: true
  require_no_custom_document_properties: true
  require_no_mail_merge: true
  require_no_save_through_xslt: true
  require_no_attached_custom_xml_schemas: true
  require_no_field_updates_on_open: true
  require_no_template_style_updates_on_open: true
  require_personal_information_removal_on_save: true
  require_no_save_forms_data: true
  require_no_save_preview_picture: true
  require_content_control_locks: true
  require_no_data_bindings: true
  require_no_external_fields: true
  require_no_modern_comment_metadata: true
  require_no_document_tasks: true
  require_no_taskpane_web_extensions: true
  require_no_sensitivity_label_metadata: true
  require_no_package_digital_signatures: true
  require_complete_package_signature_coverage: true
  require_no_word_protection: true
  require_no_word_permission_ranges: true
  require_no_word_document_variables: true
  require_no_word_document_variable_fields: true
  require_no_word_hyperlink_fields: true
  require_no_word_hyperlink_markup: true
  require_no_word_drawing_hyperlinks: true
  require_no_word_drawing_linked_pictures: true
  require_no_word_vml_hyperlinks: true
  require_no_word_vml_embedded_ole_objects: true
  require_no_external_document_dependencies: true
```

When an established template intentionally contains one of those stored states,
use the corresponding comparison gates to permit the known baseline but fail a
later mutation:

```yaml
  no_embedded_object_payload_changes: true
  no_alternative_format_import_changes: true
  no_document_property_changes: true
  no_package_thumbnail_changes: true
  no_markup_compatibility_changes: true
  no_mail_merge_changes: true
  no_save_through_xslt_changes: true
  no_attached_custom_xml_schema_changes: true
  no_field_update_on_open_changes: true
  no_template_style_update_on_open_changes: true
  no_personal_information_removal_on_save_changes: true
  no_save_forms_data_changes: true
  no_save_preview_picture_changes: true
  no_content_control_lock_changes: true
  no_data_binding_changes: true
  no_external_field_changes: true
  no_modern_comment_metadata_changes: true
  no_document_task_changes: true
  no_taskpane_web_extension_changes: true
  no_sensitivity_label_metadata_changes: true
  no_package_digital_signature_changes: true
  no_package_signature_coverage_changes: true
  no_word_protection_changes: true
  no_word_permission_range_changes: true
  no_word_document_variable_changes: true
  no_word_document_variable_field_changes: true
  no_word_hyperlink_field_changes: true
  no_word_hyperlink_markup_changes: true
  no_word_drawing_hyperlink_changes: true
  no_word_drawing_linked_picture_changes: true
  no_word_vml_hyperlink_changes: true
  no_word_vml_embedded_ole_object_changes: true
  no_external_document_dependency_changes: true
```

YAML anchors, aliases, sequences, nested mappings, duplicate keys, unknown
rules, and non-boolean values are rejected. That keeps a policy reviewable and
avoids making the CLI's safety contract depend on a broad YAML loader.

`require_complete_package_signature_coverage` is the opt-in alternative for a
handoff that is expected to retain package signatures: it requires every
recognized XML signature to provide a bound manifest and requires their
combined declarations to cover the bounded Word scope described above. It is
intentionally incompatible in purpose with
`require_no_package_digital_signatures`, which instead rejects all stored
package-signature material. `no_package_signature_coverage_changes` protects
an approved declaration-coverage baseline without claiming signature validity.

## Privacy contract

DocFence keeps document material in memory only long enough to create private
SHA-256 fingerprints. Public models and renderers intentionally omit package
part names, paragraph content, reviewer identity, dates, comment content,
relationship targets, field instructions, style identifiers and names, custom
XML values, document-property names and values, mail-merge configuration and
recipient data, XSLT transform targets and solution identifiers, data-binding
XPath expressions, prefix mappings, storage IDs,
referenced custom XML values, macro bytes, embedded and imported payload bytes,
relationship-bound OPC thumbnail image bytes, relationship sources and targets,
content types, and part paths,
OOXML Markup Compatibility branch bodies, feature-prefix and qualified-name
values, compatibility-rule values, and part paths,
external template/subdocument/frame-source targets, and all fingerprints.
Attached custom XML schema namespace identifiers and Settings-part paths remain
private too.
Modern-comment author/contact/provider identifiers, paragraph and durable IDs,
timestamps, thread associations, and reaction identities receive the same
private treatment.
Document-task IDs, event IDs and times, user identities, titles, dates,
progress, priorities, and comment anchors are private too. So are task-pane
layout details and Office web-extension IDs, stores, reference versions,
property names and values, binding identifiers and application references,
content-control IDs, aliases, tags, titles, placeholder text, current values,
markers, and part paths.
Sensitivity-label IDs, tenant site IDs, label names, methods, set dates, action
IDs, label-extension data, legacy MIP custom attributes, and Word content
marking values remain private as well.
Package-signature signer and certificate material, signature values, reference
URIs, signing times, comments, provider data, relationship IDs, part paths,
manifest object identifiers, static coverage selectors, and digest material
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
The DrawingML visibility boundary follows the Open XML SDK's
[nonvisual hidden-property contract](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.nonvisualdrawingproperties.hidden?view=openxml-3.0.1),
which describes a stored object that can remain present while hidden and notes
that a client may expose it through application settings.
Microsoft also calls out [embedded files and objects](https://support.microsoft.com/en-us/excel/embedded-files-or-objects-found)
as an inspectable hidden-data surface. The alternative-import boundary follows
the Open XML SDK's [`w:altChunk` contract](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.altchunk?view=openxml-3.0.1)
and its package-part model for [embedded and import parts](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.packaging.maindocumentpart?view=openxml-3.0.1).
Microsoft's [document-property guidance](https://support.microsoft.com/en-us/office/view-or-change-the-properties-for-an-office-file-21d604c2-481e-4379-8e54-1dd4622c6b75)
and Document Inspector coverage inform the metadata boundary.
The OPC thumbnail boundary follows the Open XML SDK's
[`AddThumbnailPart` contract](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.packaging.wordprocessingdocument.addthumbnailpart?view=openxml-2.20.0)
and the OOXML [Thumbnail Part contract](https://ooxml.info/docs/15/15.2/15.2.16/),
which defines relationship-bound, internal image parts and the per-source
thumbnail relationship limit. DocFence records that stored topology without
decoding an image or asserting how a document client will display it.
The mail-merge boundary follows the Open XML SDK's
[`w:odso` model](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.datasourceobject?view=openxml-3.0.1)
and Microsoft's documentation that accepting a linked mail-merge source can
[run its SQL query](https://support.microsoft.com/en-us/word/you-receive-the-opening-this-will-run-the-following-sql-command-message-when-you-open-a-word-mail-me).
The XSLT-on-single-XML-save boundary follows the Open XML SDK's
[`w:saveThroughXslt` contract](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.savethroughxslt?view=openxml-3.0.1)
and [`w:useXSLTWhenSaving` contract](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.usexsltwhensaving?view=openxml-3.0.1),
plus the OOXML [XSL Transformation relationship
contract](https://ooxml.info/docs/11/11.9/). Those sources define the
settings-only, single-XML-save behavior and the required external transform
relationship; DocFence records that stored configuration without following it.
The attached-schema boundary follows the Open XML SDK's
[`w:attachedSchema` contract](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.attachedschema?view=openxml-3.0.1),
which specifies a target-namespace association when a matching schema is
available to the host. DocFence records that declaration without locating or
using a schema.
The automatic-field-recalculation-on-open boundary follows the Open XML SDK's
[`w:updateFields` contract](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.updatefieldsonopen?view=openxml-3.0.1),
which specifies a request to recalculate field results on open in a supporting
application, and the Office interoperability definition of
[`ST_OnOff`](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/815fcc39-dc6c-44f8-ad27-87b1d9f21571).
DocFence records the direct stored setting without evaluating a field or
emulating a document client.
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
and the [ECMA-376 Open Packaging Conventions standard](https://ecma-international.org/publications-and-standards/standards/ecma-376/),
including its [relationship transform](https://c-rex.net/samples/ooxml/e1/Part2/OOXML_P2_Open_Packaging_Conventions_Digital_topic_ID0EHROM.html).
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
The personal-information-removal-on-save boundary follows the Open XML SDK's
[`w:removePersonalInformation` contract](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.removepersonalinformation?view=openxml-3.0.1),
which specifies a request to remove authors' personal information when a
document is saved while leaving the definition and extent of that information
undefined. DocFence records the direct stored request without identifying,
removing, or rewriting any document material.
The form-data-only-save boundary follows the Open XML SDK's
[`w:saveFormsData` contract](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.saveformsdata?view=openxml-3.0.1).
It describes a stored request for form-field-content-only saving; DocFence
records only the direct declaration and never attempts that save or inspects a
field value.
The preview-thumbnail-on-save boundary follows the Open XML SDK's
[`w:savePreviewPicture` contract](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.savepreviewpicture?view=openxml-3.0.1).
It describes a stored request for a capable host to generate a first-page
thumbnail when saving. DocFence records only the direct declaration; it does
not create, decode, render, or infer the presence of a thumbnail image.
The content-control lock boundary follows the Open XML SDK's
[`w:lock` contract](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.lock?view=openxml-3.0.1)
and its
[`LockingValues` enum](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.lockingvalues?view=openxml-3.0.1).
Those documents distinguish deletion locking, content locking, both, and an
explicit unlocked state; the element contract also explains why a missing leaf
cannot safely be collapsed into a general runtime "unlocked" claim. DocFence
therefore reduces only direct stored declarations to aggregate evidence.
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
The DrawingML hyperlink-action boundary follows ECMA-376's
[`a:hlinkClick` definition](https://c-rex.net/samples/ooxml/e1/Part4/OOXML_P4_DOCX_hlinkClick_topic_ID0ENF2KB.html),
which lists drawing object properties among its parents and documents the
relationship-ID target plus action, invalid-URL, tooltip, frame, and history
attributes; the corresponding [`a:hlinkHover` definition](https://c-rex.net/samples/ooxml/e1/part4/OOXML_P4_DOCX_hlinkHover_topic_ID0EF62IB.html)
and Open XML SDK [`a:hlinkMouseOver` contract](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.hyperlinkonmouseover?view=openxml-3.0.1)
complete the stored action family. DocFence counts direct stored markers only,
keeps all values private, and makes no client-execution or target-safety claim.
The DrawingML linked-picture boundary follows the Open XML SDK's
[`a:blip/@r:link` contract](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.blip.link?view=openxml-3.0.1),
which defines a linked-picture reference as an image that does not reside in
the file, and ECMA-376's [Image Part
contract](https://c-rex.net/samples/ooxml/e1/Part1/OOXML_P1_Fundamentals_Image_topic_ID0EGXDO.html),
which permits a standard image relationship to have an internal or external
target mode. The release check also profiles an open-source Word package from
the [`abspath2relpath-docx` investigation](https://github.com/pea-sys/shell-experiments/blob/91386d4de9e499a21bbb2e54743eb63a63727bfb/powershell/survey/abspath2relpath-docx/survey/abs/1.docx): it contains two direct markers backed by external image relationships. DocFence reports only the two marker and relationship-classification counts, never their paths or targets.
The legacy VML external-image boundary follows the Open XML SDK's
[`v:imagedata/@r:id` contract](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.vml.imagedata.relationshipid?view=openxml-3.0.1),
which identifies it as the explicit relationship to image data, together with
the ECMA-376 [Image Part
contract](https://c-rex.net/samples/ooxml/e1/Part1/OOXML_P1_Fundamentals_Image_topic_ID0EGXDO.html)
for the standard image relationship and stored target mode. The boundary does
not treat raw VML `src` as an alternate relation: Microsoft's Office
compatibility notes identify that image-data attribute as unsupported
([MS-OE376 §2.1.202](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oe376/7e506612-7a40-4d4e-95f4-e1f36173fe14)). The release check also profiles the paired open-source
[`rel/1.docx` package](https://github.com/pea-sys/shell-experiments/blob/91386d4de9e499a21bbb2e54743eb63a63727bfb/powershell/survey/abspath2relpath-docx/survey/rel/1.docx): it stores two direct `v:imagedata/@r:id` markers backed by external standard image relationships. DocFence reports only the two marker and relationship-classification counts, never their paths or targets.
The VML image-data hyperlink boundary follows the Open XML SDK's
[`ImageData` contract](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.vml.imagedata?view=openxml-3.0.1),
which identifies `r:href` as an explicit relationship to a hyperlink target.
A public [Word XML fragment](https://stackoverflow.com/questions/52124509/read-images-from-docx-file-with-python-docx)
shows the exact `v:imagedata r:id="…" r:href="…"` form in legacy VML. A
separate public [document-processing report](https://forum.aspose.com/t/corrupted-targetmode-attribute-value-in-relationship-tag/22502)
shows that its `r:href` can instead resolve to an external standard image
relationship. DocFence therefore records direct stored markers and classifies
only recognized hyperlink relationships by mode; all other resolved
relationship types or modes stay reviewable as unsupported evidence. It makes
no rendering, link-following, or target-safety claim.
The legacy VML linked-OLE boundary follows the Open XML SDK's
[`o:OLEObject` contract](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.vml.office.oleobject?view=openxml-3.0.1)
and the ECMA-376 [`OLEObject` definition](https://c-rex.net/samples/ooxml/e1/Part4/OOXML_P4_DOCX_OLEObject_topic_ID0EXGUXB.html),
which identify the direct Office VML element, its optional `r:id`, `Type`, and
`UpdateMode` attributes. The standard specifies that update mode describes
automatic versus on-demand new data only when `Type` is `Link`; the
[Embedded Object Part contract](https://c-rex.net/samples/ooxml/e1/Part1/OOXML_P1_Fundamentals_Embedded_topic_ID0EA5BO.html)
permits the standard OLE-object relationship target to be internal or external.
Microsoft's [linked-versus-embedded-object guidance](https://support.microsoft.com/en-US/Word/linked-objects-and-embedded-objects)
explains why the distinction matters for source-backed, independently maintained
data. Public Word XML fragments show the direct stored form in practice: an
[Excel-linked object](https://stackoverflow.com/questions/60565712/insert-ole-object-into-ms-word-document-and-keep-the-underlying-format-wmf-intac)
with `Type="Link"` and `UpdateMode="Always"`, and a
[linked OLE shape](https://forum.aspose.com/t/cannot-found-includepicture-field-type-via-doc-range-fields/277035)
with `UpdateMode="OnCall"`. DocFence reports only bounded stored evidence from
that markup, never a client update or retrieval outcome.
The complementary VML embedded-OLE boundary follows the same Open XML SDK
[`o:OLEObject` contract](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.vml.office.oleobject?view=openxml-3.0.1),
ECMA-376 [`OLEObject` definition](https://c-rex.net/samples/ooxml/e1/Part4/OOXML_P4_DOCX_OLEObject_topic_ID0EXGUXB.html),
and [Embedded Object Part contract](https://c-rex.net/samples/ooxml/e1/Part1/OOXML_P1_Fundamentals_Embedded_topic_ID0EA5BO.html).
They identify the direct Office VML element, optional `r:id`, `Type`, several
parent forms, and the OLE-object relationship's stored target mode; the
standard describes `UpdateMode` for the Link form. A public Open XML SDK
[`ole.docx` fixture at `cd2b359`](https://github.com/dotnet/Open-XML-SDK/blob/cd2b359ef824737edb93f1c6157c19551aae1e52/test/DocumentFormat.OpenXml.Tests.Assets/assets/TestDataStorage/v2FxTestFiles/wordprocessing/ole/ole.docx)
contains one direct `Type="Embed"` marker with an internal standard
relationship. Five public [docx4j OLE fixtures at
`74ea743`](https://github.com/plutext/docx4j/tree/74ea74323a33d92769fdbd3e6d5fe730bbfd8ffb/docx4j-core-tests/src/test/resources/OLE)
each contain the same direct stored form. In pinned scans, 22 of 503 readable
Open XML SDK package candidates (two of 505 produced read errors) and five of
141 docx4j Word packages contained a direct marker. Those test corpora establish
stored-package coverage, not prevalence, client interoperability, payload
safety, or runtime behavior.
The distinct WordprocessingML linked-object-property boundary follows the Open
XML SDK's
[`w:objectLink` contract](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.objectlink?view=openxml-3.0.1),
which identifies it as an Office 2007+ leaf element whose parent is
`w:object` and exposes the `r:id`, program, field-code, locking, and
update-mode attributes. The ISO 29500 schema's
[`CT_ObjectLink` definition](https://github.com/dolanmiu/docx/blob/309b972e107b8cc12c65027d5161b7045e404b6c/ooxml-schemas/ISO-IEC29500-4_2016/wml.xsd#L1188-L1228)
requires `r:id` and `updateMode` and enumerates `always` and `onCall`;
the [Embedded Object Part contract](https://c-rex.net/samples/ooxml/e1/Part1/OOXML_P1_Fundamentals_Embedded_topic_ID0EA5BO.html)
permits the standard OLE-object relationship target to be internal or external.
The Microsoft API page's illustrative `updateMode="user"` is not promoted to a
standard mode: DocFence keeps an absent or unexpected stored mode only as
unsupported evidence, never as a client-behavior claim.
The direct embedded-control-anchor boundary follows the Open XML SDK's
[`w:control` contract](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.control?view=openxml-3.0.1),
which documents `w:object` and `w:pict` parent forms and the optional property
relationship. The ISO 29500 schema's
[`CT_Control`, `CT_Object`, and `CT_Picture` definitions](https://github.com/dolanmiu/docx/blob/309b972e107b8cc12c65027d5161b7045e404b6c/ooxml-schemas/ISO-IEC29500-4_2016/wml.xsd#L1146-L1199)
confirm the optional `r:id` and the two direct parent positions. The ECMA-376
[Embedded Control Persistence Part contract](https://c-rex.net/samples/ooxml/e1/Part1/OOXML_P1_Fundamentals_Embedded_topic_ID0EQDBO.html)
requires its target relationship to be internal. Microsoft lists ActiveX
controls among active content that Trust Center can block or warn about, which
makes an anchored-control review signal useful without equating every anchor
with executable behavior ([Microsoft Support](https://support.microsoft.com/en-us/office/collab-files/active-content-types-in-your-files)).
Release validation also profiles docx4j's public
[`LegacyForms.docx` fixture at `74ea743`](https://github.com/plutext/docx4j/blob/74ea74323a33d92769fdbd3e6d5fe730bbfd8ffb/docx4j-core-tests/src/test/resources/LegacyForms.docx): it contains five direct `w:object/w:control` anchors, each backed by an internal standard control relationship. The release reports only aggregate anchor and relationship classes, never the fixture's control names, shape IDs, or payload data.
The legacy VML boundary follows Microsoft's [`HRef` shape
attribute](https://learn.microsoft.com/en-us/windows/win32/vml/href-attribute--shape--vml),
which defines the URL used when a shape is clicked, and the W3C's
[VML note](https://www.w3.org/TR/NOTE-VML), which documents `href` and
`target` on the VML shape-family vocabulary. DocFence limits itself to those
direct stored shape/group/shape-template attributes in Word stories and does
not infer an effective rendered or inherited link.
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
