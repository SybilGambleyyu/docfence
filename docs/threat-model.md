# Threat model

DocFence is designed for a repository, build agent, or handoff workstation that
needs a small amount of dependable evidence about an untrusted Word OOXML
package. It is not a general-purpose malware sandbox or a renderer.

## Assets and boundaries

The source document is treated as sensitive. Its body text, hidden text,
comments, reviewer metadata, relationship targets, field instructions, style
identifiers, custom XML, macro bytes, embedded-object bytes, and
alternative-format-import bytes, document-property names, and document-property
values, mail-merge configuration, source/header targets, connection strings,
queries, field mappings, recipient data, XSLT-on-single-XML-save transform
targets and solution identifiers, data-binding XPath expressions,
namespace-prefix mappings, custom XML storage IDs, and referenced custom XML
values, attached custom XML schema namespace identifiers and Settings-part
paths, external-source field paths, connections, queries, application/item
references, external template/subdocument/frame-source targets, and frame names
must not become DocFence report content. Modern-comment author names, contact
providers and user IDs, comment paragraph and durable IDs, timestamps, thread
associations, reaction users, and reaction timestamps are sensitive package
material under the same rule. The report may be stored in CI artifacts, pasted
into an issue, or uploaded to a SARIF consumer, so it is intentionally
restricted to counts, fixed categories, booleans, and generic story kinds.
Document-task IDs, event IDs/times, users, titles, schedules, progress,
priorities, and comment anchors are sensitive under the same rule. So are
task-pane layout values and Office web-extension IDs, store/reference data,
property names and values, binding IDs/application references, content-control
marker values, control IDs, aliases, tags, titles, placeholder text, current
values, and part paths. Sensitivity-label IDs, tenant site IDs, label
names, assignment methods, set dates, action IDs, extension payloads, legacy
MIP custom attributes, Word content-marking text, property names and values,
and LabelInfo part paths are sensitive package material too.
Package-signature signer and certificate material, signature values, algorithm
identifiers, reference URIs, signing times, comments, provider data,
relationship IDs, and signature-part paths are sensitive package material too.
OPC package thumbnail image bytes, relationship IDs, sources and targets,
content types, and part paths are sensitive package material too.
OOXML Markup Compatibility branch bodies, feature-prefix and qualified-name
values, compatibility-rule values, and member paths are sensitive package
material too.
Word protection hashes, salts, verifier values, cryptographic provider and
algorithm fields, and Settings-part paths are sensitive package material too.
Word document-variable names, values, Settings-part paths, and `DOCVARIABLE`
field arguments are sensitive package material too.
Word `HYPERLINK` field destinations, internal locations, ScreenTips, frame
targets, field instructions, and story-part paths are sensitive package
material too.
Direct WordprocessingML hyperlink targets, local anchors, locations, tooltips,
frame names, history values, display text, relationship IDs, and story-part
paths are sensitive package material too.
Direct DrawingML linked-picture targets, relationship IDs, surrounding drawing
markup, and story-part paths are sensitive package material too.
Direct DrawingML nonvisual object names, descriptions, titles, IDs, raw hidden
attribute spellings, and story-part paths are sensitive package material too.
Direct legacy VML shape-link URLs, target frames, titles, alternate text, shape
identifiers, and story-part paths are sensitive package material too.
Direct legacy VML image targets, relationship IDs, raw source values, other
image-data attributes, and story-part paths are sensitive package material too.
Direct legacy VML linked- and embedded-OLE sources, monikers, program, shape,
and object identifiers, relationship IDs and targets, update metadata, field
codes, markup, and story-part paths are sensitive package material too.
Direct WordprocessingML linked-object-property program and shape identifiers,
relationship IDs and targets, field codes, locking and update metadata, markup,
and story-part paths are sensitive package material too.
Direct WordprocessingML embedded-control names and shape identifiers,
relationship IDs and targets, markup, and story-part paths are sensitive package
material too.
Word editable-range marker IDs, individual editor values, exact table-column
selectors, custom-XML placement, and story-part paths are sensitive package
material too. An editor value can be an email address, alias, or domain
identity.

DocFence does not invoke Word or Office automation. It does not execute macro
code, calculate fields, resolve a hyperlink, retrieve a relationship target,
open an embedded object, import alternative-format content, render images, or
send package bytes over the network. It operates on a local copy of the ZIP
container with the Python standard library.

## Package and parser defenses

Before parsing, the reader rejects a source path that is not a regular file or
is a symlink. It reads only `.docx`, `.docm`, `.dotx`, and `.dotm` paths and
applies these default limits:

| Boundary | Default |
| --- | ---: |
| Source package bytes | 128 MiB |
| ZIP members | 4,096 |
| Expanded member bytes | 64 MiB |
| Total expanded bytes | 512 MiB |
| Compression ratio per member | 1,000:1 |
| Parsed XML bytes | 16 MiB |
| XML elements | 200,000 |
| XML depth | 256 |

Encrypted entries, unsupported compression, symbolic-link entries, duplicate
or Unicode/case-colliding names, unsafe member paths, DTDs, entity declarations,
and malformed required OOXML structures are rejected. Parsed XML uses bounded
`xml.etree.ElementTree` traversal; a parser failure produces a generic error
without embedding package material in the message.

These controls reduce common ZIP-bomb and XML-expansion risk. They are not a
substitute for container isolation, filesystem quotas, process timeouts, or a
dedicated malware scanner when those are required by the environment.

## Privacy design

All sensitive comparison material is transformed into SHA-256 fingerprints
inside the process. Fingerprints are held privately for comparison and are not
present in public report dictionaries. Relationship IDs are normalized to their
private relationship semantics before story fingerprints are made, so an ID
renumbering alone does not create report churn. Word's volatile `rsid` metadata
and common volatile revision author/date metadata are also excluded from story
fingerprints.

Embedded OLE/package/control and alternative-format-import inventories use the
same private-digest approach. For the recognized relationship types, an internal
target is resolved only as a normalized package-member name and must exist in
the already validated ZIP member map. A `w:altChunk` anchor must name a matching
internal `aFChunk` relationship. These checks never interpret the target's
payload bytes as an application document, script, image, or HTML.

Core, extended, and custom property XML follows the same boundary. DocFence
validates recognized property roots and fingerprints the full stored structure
privately. It does not disclose a property name, value, path, relationship
target, or digest. A count is evidence that stored metadata exists, not a
classification of the metadata as personal or safe.

Relationship-bound OPC package thumbnails receive a dedicated private-digest
treatment. DocFence recognizes only the standard Transitional or Strict
thumbnail relationship type from the package or a stored source part, requires
an internal stored target with an `image/` content type, limits each source to
one thumbnail relationship, and rejects a thumbnail target that owns a
Relationships part. Public output contains only relationship and distinct-part
counts; bytes, sources, targets, content types, paths, and fingerprints remain
private. A same-count image rewrite stays review-visible without emitting image
data. A conventional filename alone is not enough to make a member a thumbnail.
DocFence never decodes, renders, classifies, or otherwise interprets the image.

OOXML Markup Compatibility (MCE) receives a dedicated private-digest
treatment. DocFence scans stored non-relationship `word/*.xml` members that
use the standard MCE namespace and reports only aggregate part, branch, and
compatibility-rule token counts. Recognized MCE elements and attributes are
fingerprinted privately, so an equal-count branch body or rule-value rewrite
stays review-visible without placing branch text, prefix values, qualified
names, attribute values, or member paths in a report. This boundary does not
validate MCE conformance or infer whether a stored branch can be selected.
DocFence never resolves a feature prefix, applies a target version, preprocesses
or saves a package, chooses a branch, or renders an MCE result.

This Word-part MCE inventory does not relax OPC's package-signature syntax:
under §10.5.2, a recognized Digital Signature XML Signature part with an MCE
namespace element or attribute is rejected. That structural check does not
select or interpret an MCE branch and does not cause signature parts to enter
the Word-part inventory.

DrawingML nonvisual visibility receives a separate private-digest treatment.
DocFence reads only direct unqualified hidden attributes on its supported
DrawingML nonvisual-property QNames in supported Word stories. It reports
aggregate declaration/story and hidden/explicitly-shown/invalid counts, while
object names, descriptions, titles, IDs, raw invalid Boolean values, paths,
and fingerprints remain private. Valid XML Boolean spellings are canonicalized
before signature construction; invalid raw values remain private signature
material. This bounds the signal to the stored direct marker: DocFence does
not validate arbitrary DrawingML, choose an MCE branch, identify an object,
calculate effective visibility, apply layout, render, or claim what an Office
client will show.

The general custom-XML inventory also privately fingerprints each
non-relationship package member under the conventional `customXml/` folder.
Only the aggregate count can leave the process. This keeps unbound custom XML
review-visible without emitting a member name, XML content, namespace, or
digest. Custom-XML relationship parts remain in the generic relationship
inventory. The boundary does not infer that the data is visible, sensitive,
referenced, or used by an Office host.

Sensitivity-label metadata receives a dedicated private-digest treatment rather
than relying on the broad custom-property inventory. DocFence recognizes the
Office 2021 LabelInfo root, standard package relationship, SDK content type,
and canonical LabelInfo paths. A recognized classification-label relationship
must be root-package scoped, internal, and resolvable; multiple LabelInfo parts
or malformed label state fail closed. It also scans recognized custom-property
parts for legacy MIP `MSIP_Label_<GUID>_<attribute>` state, the legacy
`Sensitivity` property, and documented Word marking properties. Public output
is aggregate-only, while label IDs, tenant IDs, names, dates, methods, action
IDs, extension content, MIP custom attributes, and marking values remain in the
private signature. Thus a same-count label change remains visible without
placing governance metadata in a report.

OPC package digital-signature material receives the same dedicated treatment.
DocFence recognizes the standard root origin relationship, the exact origin,
XML-signature, and certificate content types (including defaults), and the
conventional origin path. Recognized relationships are constrained to the
expected source, internal target mode, and stored target; the root origin is
unique. A recognized XML signature receives only bounded structural XMLDSIG
checks before its full bytes are privately digested. Public output is limited
to topology, reference, inline-certificate, and signature-property counts.
It never emits signer/certificate contents, values, algorithms, reference URIs,
times, comments, provider data, relationship IDs, paths, or digests. A
same-count signature or certificate rewrite is therefore visible without
copying signature material into a report.

The direct XMLDSIG grammar is constrained before inventory: `Signature` must
directly contain `SignedInfo`, `SignatureValue`, optional `KeyInfo`, then
zero or more `Object` elements. `SignedInfo` must directly contain
`CanonicalizationMethod`, `SignatureMethod`, then one or more `Reference`
elements. `SignedInfo` and `SignatureValue` may have only their optional
`Id` attribute at this boundary; `SignatureValue` cannot carry child XML,
and the direct `CanonicalizationMethod` and `SignatureMethod` elements must
each carry exactly their required `Algorithm` attribute.
`SignatureMethod/@Algorithm` must be nonblank. XMLDSIG permits parameter
child markup on those method elements; this bounded check leaves that markup
outside its scope. This is a stored-sequence check, not full XMLDSIG schema
validation: it does not validate base64 lexical content, method parameters,
KeyInfo/Object payloads, a digest, or a signature.

Every XMLDSIG `DigestMethod/@Algorithm` in a recognized package signature is
checked for OPC's exact MD5 prohibition. This stored-syntax check neither
endorses nor broadly rejects SHA-1 or other algorithms, and does not recompute
any digest.

Every XMLDSIG `ds:Transform/@Algorithm` in a recognized package signature
must be OPC's Relationship Transform URI or one of OPC's two XML
Canonicalization URIs. A missing or other URI is rejected anywhere in the
Signature, even outside the bounded declaration chain. DocFence does not
resolve or execute a transform.

XMLDSIG's `TransformsType` declares no attributes, while `TransformType`
declares only its required `Algorithm`. DocFence therefore rejects an
attribute-bearing `ds:Transforms` or a `ds:Transform` with a missing, wrong,
or extra attribute anywhere in a recognized package signature. This is a
direct stored-grammar check, not validation of arbitrary transform parameters
or child markup.

A permitted Relationship Transform must also occupy OPC's required local
package context: a direct `ds:Transform` under a direct
`ds:Manifest/ds:Reference/ds:Transforms` chain whose URI declares a `.rels`
part with the case-insensitive relationships content type. It must contain at least one
direct OPC relationship selector and be immediately followed by one of OPC's
two XML Canonicalization transforms. A declared relationships part may occur
only once as a Relationship Transform target in a recognized XML signature.
This is a stored-placement check, not target resolution, selector
interpretation, or transform execution.

Every `opc:RelationshipReference` and `opc:RelationshipsGroupReference` in a
recognized signature is also checked globally. It must be a direct child of a
Relationship Transform; the former has exactly `SourceId`, the latter exactly
`SourceType`, and neither may have child elements. A standalone selector,
another parent, a missing, wrong, or extra attribute, or nested selector markup
is malformed. This is not selector-value resolution or interpretation.

Every XMLDSIG `ds:XPath` element in a recognized package signature is also
rejected because OPC says the XPath element shall not be present. This is a
global stored-markup check: DocFence neither parses nor evaluates an XPath
expression or transform.

OPC §10.5.2 likewise says a Digital Signature XML Signature part shall not
contain MCE-namespace elements or attributes. DocFence rejects either form
anywhere in a recognized package signature. This is a structural namespace
check, not MCE conformance validation, branch selection, or preprocessing.

A separate static declaration audit follows only one direct
`ds:Object Id="idPackageObject"` with no other attributes and exactly direct
`ds:Manifest` then `ds:SignatureProperties` children, plus exactly one
direct `SignedInfo` local-fragment reference to that object. That binding
must carry XMLDSIG's direct optional `Transforms`, `DigestMethod`, then
`DigestValue` shape; when present, transforms can use only OPC's two XML
Canonicalization algorithms. Within this bounded binding/manifest chain, each
`Reference` may carry only XMLDSIG's `Id`, `URI`, and `Type` attributes and
each direct `DigestMethod` must carry exactly `Algorithm`. Unknown attributes
leave a binding unavailable or a manifest reference aggregate unsupported,
rather than lending coverage. This is not full XMLDSIG schema validation:
DocFence does not validate method-parameter child markup, base64, or any
digest or signature value. OPC's global `ds:Transform/@Algorithm` restriction,
Relationship Transform local context, relationship-selector shape, and
`ds:XPath` prohibition are enforced before this bounded declaration chain is
evaluated, including references outside the chain, as is OPC §10.5.2's
MCE-namespace element/attribute prohibition. It does not decode or recompute
the digest, execute transforms, or verify XMLDSIG. It resolves only that
package-specific manifest using the binding reference's direct URI fragment. Its direct
`SignatureProperties` must have one `SignatureProperty` with the fixed
`idSignatureTime` ID and an empty or root-`Signature/@Id` fragment target.
That property must contain only an attribute-free `opc:SignatureTime` with
attribute-free `opc:Format` then `opc:Value` children; the value must match
one of OPC's six declared time-precision forms. This is a bounded syntax check
for a claimed timestamp, not an assertion about timestamp accuracy, authority,
or signature validity. From that bounded link, it resolves exact part URIs
with ASCII-case-insensitive content-type matching. Before a
manifest reference is credited, its direct XMLDSIG children must be optional
`ds:Transforms`, one `ds:DigestMethod` with a nonblank `Algorithm`, and one
direct, attribute-free, child-free, nonempty `ds:DigestValue`, with no
non-whitespace direct text, in that order. The global recognized-signature
boundary rejects OPC's expressly forbidden MD5 URI and every `ds:XPath`
element, every non-OPC transform algorithm, every malformed Relationship
Transform or relationship selector, and MCE-namespace markup; DocFence
does not otherwise judge an algorithm's cryptographic suitability. It then
recognizes standard relationship-transform declarations: one direct
`ds:Transforms` list containing only the supported OPC relationship and XML
Canonicalization algorithms, with exactly one relationship transform
immediately followed by XML Canonicalization (with or without comments), then
`RelationshipReference/@SourceId` and
`RelationshipsGroupReference/@SourceType` selector shapes. Within this bounded
declaration audit, selector values match stored relationship IDs and types
ASCII-case-insensitively and cover every match, as OPC §10.6 requires. Its
public surface
is aggregate-only: signatures with or without a declaration link; covered and
uncovered Word parts and relationships; and unresolved or unsupported
references. Object identifiers, raw reference URIs, selectors, relationship
IDs/types, part paths, and digest material remain private. A same-count
coverage reassignment is retained in the private semantic signature.

Within one XML signature, the recognized-signature boundary also rejects more
than one Relationship Transform for the same declared relationships part across
every manifest. That preserves the OPC per-relationships-part constraint
without exposing the part, selector, or relationship identity.

For a non-relationship package part, the audit accepts no transform list or one
direct nonempty list of only the two OPC XML Canonicalization algorithms. Empty
and duplicate lists are unsupported and do not credit the part; a Relationship
Transform in that context is malformed, and a non-OPC transform algorithm has
already rejected the signature.
This is structural filtering of a small declared subset, not transform execution.

Word editing/write-protection state receives the same dedicated treatment.
DocFence discovers document Settings parts through Word's conventional
`word/settings.xml` path and Transitional or Strict settings relationships from
main or glossary documents, then privately fingerprints each discovered
Settings root. It inventories direct `w:documentProtection` and
`w:writeProtection` leaves only after validating their bounded structure,
known Word attributes, edit values, and boolean values. Public output is
aggregate-only; verifier fields, provider/algorithm data, paths, and
fingerprints remain private. A same-count verifier rewrite is therefore
review-visible without putting password-verifier material in a CI artifact.
This treats all cryptographic-looking fields as opaque package data: DocFence
does not decode, interpret, or emit them as password verifiers.

Word document-variable state receives a separate private-digest treatment.
DocFence discovers document Settings parts through the same conventional and
Transitional/Strict relationship paths, then recognizes direct `w:docVars`
containers and their direct `w:docVar` leaves. It validates the standard
container/leaf shape, Word-namespace attribute vocabulary, required name/value
attributes, Open XML SDK length constraints, and at-most-one container per
Settings part. Names, values, and paths remain private; public output is
limited to container, variable, and empty-value counts. Full recognized
containers are privately fingerprinted, so a same-count name or value rewrite
remains review-visible without placing automation state into a CI artifact.

XSLT-on-single-XML-save configuration receives a separate private-digest
treatment. DocFence inventories direct `w:useXSLTWhenSaving` and
`w:saveThroughXslt` settings from every discovered Document Settings part, as
well as standard `transform` relationships from those parts. A
relationship-backed anchor must name a standard external transform
relationship; a local `w:solutionID`-only anchor is retained only as stored
configuration evidence. Reports expose aggregate enabled-setting,
disabled-setting, anchor, relationship, and solution-identifier counts while targets, relationship IDs,
solution identifiers, Settings-part paths, and fingerprints remain private.
DocFence does not resolve, retrieve, parse, execute, or apply an XSLT.

Attached custom XML schema declarations receive a separate private-digest
treatment. DocFence recognizes direct `w:attachedSchema` leaves from every
discovered Document Settings part after validating the standard required
Word-namespace `w:val` attribute and leaf shape. Public output is limited to a
declaration count; namespace identifiers, Settings-part paths, and fingerprints
remain private. A same-count namespace rewrite therefore remains
review-visible without copying a private vocabulary identifier into a report.
DocFence does not resolve, retrieve, load, or validate against a declared
schema, and it does not claim that a host has a matching schema available.

Automatic field-recalculation-on-open settings receive a separate private-
digest treatment. DocFence recognizes at most one direct `w:updateFields`
`CT_OnOff` leaf per discovered Document Settings part and validates its leaf
shape, Word-namespace attribute vocabulary, and standard on/off token. Public
output is limited to enabled and explicitly disabled setting counts;
Settings-part paths and fingerprints remain private. The private signature
retains canonical enabled/disabled state so a state transition stays
review-visible without treating equivalent token spellings as distinct events.
DocFence does not parse, evaluate, or update a field, open Word, access a
source, follow a link, start an application, or claim a host will recalculate a
field.

Automatic template-style-update-on-open settings receive a separate private-
digest treatment. DocFence recognizes at most one direct w:linkStyles CT_OnOff
leaf per discovered Document Settings part and validates its leaf shape,
Word-namespace attribute vocabulary, and standard on/off token. Public output
is limited to enabled and explicitly disabled setting counts; Settings-part
paths and fingerprints remain private. The private signature retains canonical
enabled/disabled state so a state transition stays review-visible without
treating equivalent token spellings as distinct events. The separate
external-document-dependency inventory remains responsible for stored
attached-template anchors and relationships. DocFence does not resolve,
retrieve, load, open, validate, authenticate, or follow a template; perform
style resolution or propagation; start a document client; or claim a host will
update document styles.

Personal-information-removal-on-save settings receive a separate private-digest
treatment. DocFence recognizes at most one direct
`w:removePersonalInformation` `CT_OnOff` leaf per discovered Document Settings
part and validates its leaf shape, Word-namespace attribute vocabulary, and
standard on/off token. Public output is limited to enabled and explicitly
disabled setting counts; Settings-part paths and fingerprints remain private.
The private signature retains canonical enabled/disabled state so a state
transition stays review-visible without treating equivalent token spellings as
distinct events. DocFence does not open or save a document, identify authors,
inspect or rewrite document properties, remove comments or revisions, or claim
that a host will honor the stored request.

Form-data-only-save settings receive a separate private-digest treatment.
DocFence recognizes at most one direct `w:saveFormsData` `CT_OnOff` leaf per
discovered Document Settings part and validates its leaf shape, Word-namespace
attribute vocabulary, and standard on/off token. Public output is limited to
enabled and explicitly disabled setting counts; Settings-part paths and
fingerprints remain private. The private signature retains canonical
enabled/disabled state so a state transition stays review-visible without
treating equivalent token spellings as distinct events. DocFence does not find
or evaluate a legacy form field, read a field value, open Word, save a
document, emit a delimited record, determine a delimiter, or claim that a host
will honor the stored request.

Preview-thumbnail-on-save settings receive a separate private-digest treatment.
DocFence recognizes at most one direct `w:savePreviewPicture` `CT_OnOff` leaf
per discovered Document Settings part and validates its leaf shape,
Word-namespace attribute vocabulary, and standard on/off token. Public output
is limited to enabled and explicitly disabled setting counts; Settings-part
paths and fingerprints remain private. The private signature retains canonical
enabled/disabled state so a state transition stays review-visible without
treating equivalent token spellings as distinct events. This is configuration
evidence only: DocFence does not infer whether a thumbnail is already stored,
decode or render an image, open Word, save a document, generate a thumbnail,
or claim that a host will honor the stored request. The existing
relationship-bound OPC package-thumbnail inventory remains the separate static
boundary for an actual stored thumbnail image.

Complete `DOCVARIABLE` field instructions receive a companion private-digest
treatment. DocFence scans direct simple-field instructions and complete complex
pre-separator instructions across supported stories, retaining current and
deleted field-code variants separately. The raw instruction is reduced to a
private hash. Public output contains only reference/story counts and a
conservative literal/nonliteral classification. A literal name is associated
only when a leading plain or wholly quoted argument (with optional trailing
field-switch material) exactly matches a validated `w:docVar` from the same
main or glossary package document scope. Nested or compound expressions remain
nonliteral. This association is never output as a name, never crosses into an
attached template, and is not an evaluation of a Word field.

Complete `HYPERLINK` field instructions receive a separate private-digest
treatment because their primary argument can carry a URL, file location, or
bookmark without an OOXML relationship. DocFence scans direct simple-field
instructions and complete complex pre-separator instructions across supported
stories, retaining current and deleted field-code variants separately. Public
output contains only reference/story counts and a mutually exclusive lexical
class: literal leading destination, literal `\l` internal-location-only target,
or dynamic/unparseable. A literal is deliberately narrow—a leading plain token
or wholly quoted string with optional trailing switch material—and a nested or
compound expression stays dynamic or unparseable. All arguments, locations,
ScreenTips, targets, paths, and fingerprints remain private. A primary literal
is not reported as external because Word permits a bookmark there.

Direct `w:hyperlink` markup receives a separate private-digest treatment from
both field instructions and broad relationship totals. DocFence scans direct
elements in supported Word stories, privately normalizing an `r:id` through its
resolved relationship semantics. The element's `r:id` relationship takes
precedence over a local `w:anchor`; without an ID, an anchor-only and a
no-anchor current-document-start form remain distinct stored categories.
Recognized hyperlink relationships with internal and external target modes are
counted separately, while another resolved relationship type or target mode is
counted as unsupported markup. Relationship targets, anchors, locations,
tooltips, frame names, history, display text, relationship IDs, paths, and
fingerprints remain private. A relationship-ID rewrite with unchanged semantics
does not add report churn; a same-count target or markup rewrite remains
review-visible in the private signature.

Direct DrawingML `a:hlinkClick`, `a:hlinkHover`, and `a:hlinkMouseOver` markup
receives a separate private-digest treatment from field instructions,
`w:hyperlink` markup, and broad relationship totals. DocFence scans direct
elements in supported Word stories and retains each stored marker separately;
it does not deduplicate markers sharing a relationship or visual object, nor
does it choose a Markup Compatibility branch. A marker with `r:id` is privately
normalized through its resolved relationship semantics and publicly classified
as external, internal, or unsupported. A missing `r:id` remains a distinct
malformed stored-evidence count. Public output also reports only whether
`action` or `invalidUrl` is present. Targets, URL values, actions, tooltips,
frame names, history settings, relationship IDs, paths, and fingerprints remain
private. A same-count target or attribute rewrite remains review-visible in the
private signature; an ID rewrite with identical semantics does not add report
churn.

Direct DrawingML `a:blip/@r:link` markup receives a separate private-digest
treatment from field instructions, `w:hyperlink` markup, DrawingML
hyperlink-action markup, `r:embed` image references, and broad relationship
totals. DocFence scans each direct marker in supported Word stories and retains
every stored marker separately; it does not deduplicate markers that share a
relationship or visual object, or choose a Markup Compatibility branch. The
`r:link` relationship is privately normalized through its resolved semantics
and publicly classified only as a standard image relationship with external or
internal stored target mode, or an unsupported relationship. An `a:blip` with
only `r:embed` and an unreferenced image relationship do not create marker
counts. Targets, relationship IDs, surrounding drawing markup, paths, and
fingerprints remain private. A same-count target or markup rewrite remains
review-visible in the private signature; an ID rewrite with identical semantics
does not add report churn.

Direct legacy VML shape `href` markup receives a separate private-digest
treatment from field instructions, `w:hyperlink` markup, DrawingML actions,
and broad relationship totals. DocFence scans direct unqualified `href`
attributes on the documented VML `arc`, `curve`, `image`, `line`, `oval`,
`polyline`, `rect`, `roundrect`, `shape`, `group`, and `shapetype` elements in
supported Word stories. It reports only aggregate marker/story, concrete-shape,
group, shape-template, and direct-`target`-attribute-presence counts. Raw URLs,
frame targets, titles, alternate text, shape IDs, paths, and fingerprints remain
private. The full direct element is privately fingerprinted so same-count
attribute rewrites remain review-visible. This does not calculate an inherited
or rendered link, inspect arbitrary VML elements, or follow an `href`.

Direct legacy VML `v:imagedata/@r:id` markup receives a separate private-digest
treatment from VML shape links, DrawingML linked pictures, field instructions,
`w:hyperlink` markup, and broad relationship totals. It retains a marker only
when the resolved relationship has stored external target mode, classifying a
standard image relationship separately from another external relationship type.
Internal image relationships and orphaned external image relationships do not
become marker counts. Raw VML `src`, `r:pict`, `r:href`, and `o:relid` are
intentionally excluded from both the marker definition and its private
signature. Targets, relationship IDs, source values, VML attributes, paths,
and fingerprints remain private. A same-count external target rewrite remains
review-visible; relationship-ID renumbering with unchanged semantics and a
raw-`src` rewrite do not create external-image inventory churn.

Direct legacy VML `v:imagedata/@r:href` markup receives a separate
private-digest treatment from image-data `r:id` external-image markers, VML
shape links, DrawingML linked pictures, field instructions, `w:hyperlink`
markup, and broad relationship totals. DocFence records every direct marker in
supported Word stories and classifies a resolved standard hyperlink relationship
by stored external or internal target mode. Another resolved relationship type
or mode remains visible only as unsupported stored evidence. Image-data `r:id`,
`r:pict`, raw `src`, and `o:relid` are intentionally excluded from both the
marker definition and its private signature. Targets, relationship IDs, VML
attributes, paths, and fingerprints remain private. A same-count reviewed
target rewrite remains review-visible; unchanged-semantics relationship-ID
renumbering and excluded-attribute rewrites do not create image-data-hyperlink
inventory churn.

Direct legacy Office VML `o:OLEObject` markup with `Type="Link"` receives a
separate private-digest treatment from broad embedded-object relationship and
payload totals, the separately inventoried `Type="Embed"`, WordprocessingML
`w:objectLink`, VML image data, VML shape links, DrawingML linked pictures, and
field instructions.
DocFence retains every direct marker in supported Word stories, including
duplicates and Markup Compatibility branches. A standard OLE-object
relationship is classified by stored external or internal target mode; another
resolved type or mode is unsupported, and a direct marker without `r:id`
remains its own stored-evidence class. Public output reports only aggregate
marker/story, stored-automatic-update, nonautomatic-or-unspecified-update, and
relationship-classification counts. Sources, monikers, program, shape, and
object IDs, relationship IDs and targets, field codes, markup, paths, and
fingerprints remain private. The full direct marker is privately fingerprinted,
so same-count source, program, field-code, update-mode, or target rewrites stay
review-visible while relationship-ID renumbering with unchanged semantics stays
quiet.

Direct legacy Office VML `o:OLEObject` markup with `Type="Embed"` receives a
separate private-digest treatment from `Type="Link"`, WordprocessingML
`w:objectEmbed`/`w:objectLink`, VML image data and shape links, fields, and
broad embedded-object relationship/payload totals. DocFence retains every
direct marker in supported Word stories, including duplicates and Markup
Compatibility branches; the Office VML contract permits several parent forms,
so it does not assign a rendered-object position. A standard OLE-object
relationship is classified by stored external or internal target mode; another
resolved type or mode is unsupported, and a direct marker without optional
`r:id` remains its own stored-evidence class. Public output reports only
aggregate marker/story and relationship-classification counts. Program, shape,
and object IDs, relationship IDs and targets, update metadata, field codes,
markup, paths, and fingerprints remain private. The full direct marker is
privately fingerprinted, so same-count program, field-code, update-metadata,
or target rewrites stay review-visible while relationship-ID renumbering with
unchanged semantics stays quiet. `UpdateMode` is retained only in that private
marker because the standard describes it for the Link form, not as embedded
object behavior.

Direct WordprocessingML `w:control` markup receives a separate private-digest
treatment only when it is a direct child of `w:object` or `w:pict` in a
supported Word story. The direct parent position is retained as an aggregate
class because the standard describes the two forms separately; arbitrary
`w:control` elements elsewhere do not become anchors. This boundary is distinct
from `w:objectLink`, `w:objectEmbed`, legacy VML linked- and embedded-OLE
markup, VML image data and shapes, ActiveX-binary relationships, fields, and
broad embedded-control relationship/payload totals. A direct marker without
optional `r:id` remains reviewable evidence. When it has an ID, only a standard control
relationship is classified: internal is the conforming persistence-part mode,
while external remains separately visible as nonconforming stored evidence;
another resolved type or mode is unsupported. Public output exposes only
aggregate anchor/story, parent-position, and relationship-classification counts.
Control names and shape identifiers, relationship IDs and targets, markup,
paths, and fingerprints remain private. The full direct marker is privately
fingerprinted, so same-count name, shape, or target rewrites stay review-visible
while a relationship-ID renumbering with unchanged semantics stays quiet.

Word editable-range permission markup receives a separate private-digest
treatment. DocFence scans supported Word stories for `w:permStart` and
`w:permEnd`, accepts Transitional and Strict Word namespaces, and validates the
standard leaf shape, required IDs, attribute vocabulary, predefined editor
groups, decimal column-selector syntax, and custom-XML placement values.
Duplicate start or end IDs in one story fail closed because they would make a
pair ambiguous. Marker IDs, individual editor values, exact columns, placement,
and story paths stay private. Public output exposes only aggregate marker,
paired/unpaired, individual-editor, predefined-group, table-column-selector,
and custom-XML-placement counts. Full marker shape is privately fingerprinted,
so an identity-only or same-count range rewrite remains review-visible.
Unmatched boundaries are retained as stored markup rather than discarded or
presented as effective authorization.

Mail-merge configuration and recipient-data parts use the same private-digest
approach. Recognized data and header sources must be external relationships;
recognized recipient data must be an internal stored target. DocFence validates
that direct recognized references use those modes but never opens a source,
executes a query, interprets a connection string, or parses recipient values
for public output. Orphaned recognized relationships remain visible as aggregate
state because they can still retain an external target or recipient payload.

Content-control data-binding declarations and the custom XML data/properties
parts associated through an in-package storage ID use the same private-digest
approach. The public inventory exposes only aggregate counts. DocFence does not
emit or evaluate an XPath expression, namespace-prefix mapping, storage ID,
part name, or custom XML value. A standard custom XML-properties relationship
used to identify a bound data part must be internal and point to a valid storage
properties root; malformed recognized state fails closed. A binding whose
storage ID cannot be associated with a discovered part remains a counted review
signal rather than causing DocFence to guess a target.

External-source Word field instructions receive the same private-digest
treatment. DocFence classifies only the initial keyword of complete simple
`w:fldSimple` instructions and complete complex `w:fldChar`/`w:instrText`
sequences for the documented `DATABASE`, legacy `DATA`, `DDE`, `DDEAUTO`,
`INCLUDE`/`INCLUDETEXT`, `INCLUDEPICTURE`/`IMPORT`, `LINK`, and `RD` families.
The instruction—including a potential file path, connection string, query, DDE
application or item, OLE class, or bookmark—is reduced to a private hash before
reporting. Public output exposes only category counts. It never parses a source
argument, evaluates a field, starts an external application, opens a file,
connects to a data source, or retrieves content. Complex-field instruction runs
are concatenated before hashing, so run fragmentation alone remains quiet. A
tracked `w:delInstrText` sequence is assembled as a separate deleted variant
rather than being combined with the current `w:instrText` sequence. Instruction
text outside a complete field or an unclosed complex field is not classified as
an external source.

Modern Word comment metadata receives the same private-digest treatment. The
recognized `people`, `commentsExtended`, `commentsIds`, and
`commentsExtensible` parts are root-validated before their content is reduced to
private signatures. Public output contains aggregate counts only. In
particular, it never emits author/contact/provider data, paragraph or durable
IDs, dates, extension values, reaction values, or part paths. The private
signature deliberately retains comment identifiers so a same-count thread or
identifier rewrite remains review-visible. A recognized metadata relationship
must be internal and resolve to a validated package member. DocFence does not
render, resolve, synchronize, notify, or modify comments or reaction state.

Document-task and task-pane Office web-extension state receive the same
private-digest treatment. Recognized document-task parts must have the
documented `Tasks` root. Recognized task-pane and web-extension parts must have
their documented roots; a task pane's direct web-extension reference must name
the expected internal relationship and a stored target. Recognized external,
unavailable, or malformed relationships fail closed. Public output contains
only aggregate workflow and configuration counts, never values from an event,
task user, add-in store reference, property, binding, pane layout, or
content-control marker. Relationship IDs are normalized to their private relationship
semantics, so ID renumbering alone remains quiet while a same-count stored
configuration change remains review-visible. DocFence does not install,
execute, retrieve, authenticate to, or assess the safety of an add-in; it does
not contact a task service or perform a task action.

External Word document dependencies receive the same private-digest treatment.
DocFence recognizes attached-template relationships from discovered Document
Settings parts, subdocument relationships from the main document, and frame
relationships from discovered Web Settings parts. It checks their direct Word
anchors and requires the standard relationship type and `External` target mode.
Settings and Web Settings parts are discovered through conventional or Strict
relationships from the main or glossary document; the conventional
`word/settings.xml` Settings path is also retained for compatibility. Every
discovered document Settings part is privately fingerprinted as settings state;
a linked Web Settings part with frame dependency state is fingerprinted
privately and removed from the generic opaque-payload inventory, so
relationship-ID rewriting does not create report churn while a stored layout or
source change remains visible. DocFence never resolves, retrieves, opens,
authenticates to, imports, or renders a template, subdocument, or frame target.

This protects DocFence-controlled report surfaces, not arbitrary caller logs.
Shell history, paths provided on the command line, operating-system audit logs,
and external tools are outside this contract.

## Known semantic limits

The tool compares stored OOXML, not Word's rendered result. It intentionally
does not resolve styles, fields, tracked-change acceptance, markup-compatibility
choices, or external content. Direct `w:vanish` run markup is counted. Direct
`w:vanish` and `w:specVanish` paragraph-mark markup is counted separately.
The dedicated styles inventory records enabled `w:vanish` declarations in
text-run style properties and document defaults, but does not decide whether a
style is used or calculate inherited/toggled effective formatting. Unsupported
or unclassified package parts are still fingerprinted as an opaque category so
their mutation is visible, but DocFence does not claim to explain their visual
or business effect. Embedded OLE/package/control evidence and alternative-format
imports are separate inventories, but they are not decoded, rendered, imported,
or scanned for malware. Recognized conventional folders make otherwise opaque
payloads review-visible; they do not prove the payload is valid, safe, or used
by Word. Relationship-bound package thumbnail counts are stored-topology
evidence, not proof that any client will display a thumbnail or that its image
reflects the current document. DocFence does not decode it or infer a thumbnail
from an unreferenced filename. Markup Compatibility counts are stored XML
evidence, not MCE validation or proof that a client will select, preserve,
render, or save any branch; DocFence does not resolve prefixes, target versions,
or preprocessing rules. Core and extended document-property counts include automatic data such
as timestamps, statistics, and application information. Their changes can be
expected on a normal save; DocFence records them without interpreting their
provenance or sensitivity. The custom-property candidate gate is limited to
stored custom definitions and is not a general PII detector. The custom-XML
candidate gate is limited to conventional `customXml/` package members and is
not a PII detector, an XML parser, or evidence that a host will use the data.
Mail-merge counts
are stored-state evidence, not proof that Word will access a source, run a
query, or merge recipients. DocFence does not identify, classify, or disclose
the recipient records or connection details a package may retain.
XSLT-on-single-XML-save counts are stored-state evidence, not proof that Word
will save a document as a single XML file, apply a transform, retrieve a
target, or create any particular output. DocFence does not resolve a local
solution identifier or assess transform behavior or safety.
Attached custom XML schema declaration counts are stored-state evidence, not
proof that a host has a referenced schema available, will associate it with the
document, or will validate any custom markup. DocFence does not locate, fetch,
load, or apply a schema.
Automatic-field-recalculation-on-open counts are likewise stored-state
evidence, not proof that a host supports field calculations, will recalculate a
particular field, access an external source, or produce a particular result.
DocFence does not emulate a document open or any field runtime.
Content-control
data-binding counts are stored-state evidence, not proof that a given XPath
selects a node, that Word will update the visible control, or that an unscoped
mapping uses any particular custom XML part. DocFence does not evaluate XPath,
resolve XML namespaces, or calculate rich-text mappings. External-source field
counts are likewise stored-state evidence, not proof that Word will update a
field, access a source, use a particular connection, or accept a field's
argument grammar. The boundary classifies only its named field families, not
arbitrary field expressions or every field that can display a URL. Attached-template,
subdocument, and frameset counts are likewise stored-state evidence, not proof
that Word will retrieve a target, that an external document exists, or that any
external content is safe. DocFence does not validate a target's scheme, follow
a redirected resource, determine its effective contents, or emulate Word's
template, master-document, or frameset behavior. Modern-comment metadata counts
are also stored-state evidence: they do not prove that a comment is visible,
resolved in a service, associated with a live account, or that a reaction will
be shown. DocFence does not interpret unknown comment extensions or synchronize
comment state with Word or a cloud service. Document-task counts are not proof
that a task is assigned, actionable, synchronized, visible, or connected to a
service. Task-pane web-extension counts are not proof that an add-in is
installed, valid, loaded, displayed, safe, or able to auto-open. An enabled
`Office.AutoShowTaskpaneWithDocument` property is recorded as stored state, not
as a prediction of runtime behavior. Content-control markers are not resolved
to a specific extension or rendered control.
Sensitivity-label counts are stored-state evidence, not a statement that a
particular label policy exists, that the label is effective for a recipient, or
that a file is encrypted or accessible. DocFence does not decrypt IRM-protected
files, read a LabelInfo stream from an encrypted storage, interpret label
policy, calculate permissions, apply/remove a label, or render a label's
header, footer, or watermark.
Package-signature counts and static declaration-coverage counts are stored
state evidence, not a statement that a signature is cryptographically valid,
that a certificate is trusted or current, that a signer is who the package
claims, that all relevant content is effectively covered, or that an Office
client will make a particular trust decision. DocFence does not verify XMLDSIG
values, parse, decode, or recompute reference digests or canonicalization, evaluate arbitrary
transforms, build or validate certificate chains, check revocation or
timestamps, establish signer identity, evaluate a signing policy, or calculate
cryptographic or client-effective signature coverage. It resolves only the
bounded static declarations described above.
Word-protection counts are stored-state evidence, not a statement that a
restriction is encrypted, difficult to circumvent, currently effective in a
particular client, or appropriate for a document. DocFence does not validate
password construction or verifier completeness, derive or recover passwords,
estimate password/algorithm strength, bypass a restriction, resolve the
standard-versus-Word behavior of omitted enforcement, or make a security
decision from `w:documentProtection` or `w:writeProtection`.
Word document-variable and `DOCVARIABLE` field-reference counts are likewise
stored-state evidence, not proof that a macro will consume a value, a field
will resolve or render it, a template is attached, or an automation client will
preserve the state. A same-scope literal association is not variable resolution,
and a missing association does not make a field broken. DocFence does not run
macros, evaluate fields, resolve variable names or templates, or interpret the
content or purpose of a stored value.
Word `HYPERLINK` field-reference counts are likewise stored-state evidence, not
proof that Word will render a link, a literal primary argument is external, a
destination is reachable or safe, or a client will follow it. DocFence does not
resolve, retrieve, follow, evaluate, or render a `HYPERLINK` field.
Direct `w:hyperlink` markup counts are likewise stored-state evidence, not proof
that a relationship target is external in the everyday web sense, reachable,
safe, permitted, or honored by a Word client. DocFence does not resolve,
retrieve, follow, validate, evaluate, or render a direct hyperlink target.
Direct DrawingML hyperlink-action counts are likewise stored-state evidence, not
proof that Word selects a stored marker, a relationship target is reachable,
safe, permitted, or honored, or an `action` executes. DocFence does not select
Markup Compatibility branches, associate markers with rendered objects,
deduplicate a visual link, resolve, retrieve, follow, validate, evaluate,
render, or execute an action.
Direct DrawingML linked-picture counts are likewise stored-state evidence, not
proof that Word selects a marker, reaches a relationship target, retrieves or
updates an image, renders a picture, or honors a stored relationship. DocFence
does not select Markup Compatibility branches, associate markers with rendered
objects, deduplicate a visual picture, resolve, retrieve, validate, evaluate,
render, or update a linked picture.
Direct DrawingML nonvisual visibility counts are likewise stored-state evidence,
not proof that a client selects a marker or MCE branch, associates it with a
rendered object, calculates effective visibility, applies layout, or presents
the object as hidden or visible. DocFence does not resolve object identity,
deduplicate visual objects, render DrawingML, or make a client-behavior claim
from a hidden attribute.
Direct legacy VML shape-link counts are likewise stored-state evidence, not
proof that a client inherits a group/template link, renders a shape, honors an
`href`, reaches or safely follows a target, or opens a frame. DocFence does not
select a Markup Compatibility branch, calculate effective VML inheritance,
inspect arbitrary VML attributes, resolve, retrieve, follow, validate,
evaluate, render, or execute an `href` action.
Direct legacy VML external-image counts are likewise stored-state evidence, not
proof that a client selects a marker, retrieves or updates an image, renders a
picture, or honors an external relationship. DocFence does not select Markup
Compatibility branches, associate a marker with a rendered image, deduplicate a
visual picture, resolve, retrieve, validate, evaluate, render, or update an
image target.
Direct legacy VML image-data hyperlink counts are likewise stored-state
evidence, not proof that a client selects a marker, associates it with a visual
image, treats a relationship as a conventional hyperlink, reaches or safely
follows a target, or executes an action. DocFence does not select Markup
Compatibility branches, associate markers with rendered images, deduplicate a
visual picture, resolve, retrieve, follow, validate, evaluate, render, or
execute a target.
Direct legacy VML linked-OLE counts are likewise stored-state evidence, not
proof that a client selects a marker, associates it with a visual shape,
retrieves a source, updates or activates an OLE object, or honors an automatic
update flag. DocFence does not select Markup Compatibility branches, associate
markers with rendered objects, deduplicate a visual object, resolve, retrieve,
open, update, activate, evaluate, render, or execute an OLE object.
Direct legacy VML embedded-OLE counts are likewise stored-state evidence, not
proof that a client selects a marker, associates it with a visual shape,
retrieves, opens, or activates an OLE object, or renders embedded content.
DocFence does not select Markup Compatibility branches, associate markers with
rendered objects, deduplicate a visual object, decode payload bytes, resolve,
retrieve, open, update, activate, evaluate, render, or execute an OLE object.
Direct WordprocessingML linked-object-property counts are likewise stored-state
evidence, not proof that a client selects a marker, associates it with a visual
object, retrieves a source, updates or activates an OLE object, or honors a
stored update mode. DocFence does not select Markup Compatibility branches,
associate markers with rendered objects, deduplicate a visual object, resolve,
retrieve, open, update, activate, evaluate, render, or execute an OLE object.
Direct WordprocessingML embedded-control-anchor counts are likewise stored-state
evidence, not proof that a client selects a marker, associates it with a visual
control, has an applicable control installed, loads persisted data, enables
active content, or renders or honors it. DocFence does not select Markup
Compatibility branches, associate markers with rendered controls, deduplicate a
visual control, resolve, retrieve, open, instantiate, load, activate, evaluate,
render, or execute a control.
Word editable-range counts are likewise stored-state evidence, not a statement
that an individual is authenticated, a group resolves, a client will honor a
boundary, an exact text/table region is editable, or a restriction is secure.
DocFence does not authenticate editors, resolve `w:ed` or `w:edGrp`, calculate
editable content, emulate Word precedence when both attributes are present, or
make an authorization, encryption, or security decision from range markup.

For a consequential legal, medical, financial, or publishing decision, use
DocFence as a controlled review signal alongside an appropriate rendering and
domain review process.
