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
queries, field mappings, recipient data, data-binding XPath expressions,
namespace-prefix mappings, custom XML storage IDs, and referenced custom XML
values, external-source field paths, connections, queries, application/item
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
marker values, and part paths. Sensitivity-label IDs, tenant site IDs, label
names, assignment methods, set dates, action IDs, extension payloads, legacy
MIP custom attributes, Word content-marking text, property names and values,
and LabelInfo part paths are sensitive package material too.
Package-signature signer and certificate material, signature values, algorithm
identifiers, reference URIs, signing times, comments, provider data,
relationship IDs, and signature-part paths are sensitive package material too.
Word protection hashes, salts, verifier values, cryptographic provider and
algorithm fields, and Settings-part paths are sensitive package material too.
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
by Word. Core and extended document-property counts include automatic data such
as timestamps, statistics, and application information. Their changes can be
expected on a normal save; DocFence records them without interpreting their
provenance or sensitivity. The custom-property candidate gate is limited to
stored custom definitions and is not a general PII detector. Mail-merge counts
are stored-state evidence, not proof that Word will access a source, run a
query, or merge recipients. DocFence does not identify, classify, or disclose
the recipient records or connection details a package may retain. Content-control
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
Package-signature counts are stored-state evidence, not a statement that a
signature is cryptographically valid, that a certificate is trusted or current,
that a signer is who the package claims, that all relevant content is covered,
or that an Office client will make a particular trust decision. DocFence does
not verify XMLDSIG values or reference digests, build or validate certificate
chains, check revocation or timestamps, establish signer identity, evaluate a
signing policy, or calculate signature coverage.
Word-protection counts are stored-state evidence, not a statement that a
restriction is encrypted, difficult to circumvent, currently effective in a
particular client, or appropriate for a document. DocFence does not validate
password construction or verifier completeness, derive or recover passwords,
estimate password/algorithm strength, bypass a restriction, resolve the
standard-versus-Word behavior of omitted enforcement, or make a security
decision from `w:documentProtection` or `w:writeProtection`.
Word editable-range counts are likewise stored-state evidence, not a statement
that an individual is authenticated, a group resolves, a client will honor a
boundary, an exact text/table region is editable, or a restriction is secure.
DocFence does not authenticate editors, resolve `w:ed` or `w:edGrp`, calculate
editable content, emulate Word precedence when both attributes are present, or
make an authorization, encryption, or security decision from range markup.

For a consequential legal, medical, financial, or publishing decision, use
DocFence as a controlled review signal alongside an appropriate rendering and
domain review process.
