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
marker values, and part paths.

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
`word/settings.xml` Settings path is also retained for compatibility. A linked
Web Settings part with frame dependency state is fingerprinted privately and
removed from the generic opaque-payload inventory, so relationship-ID rewriting
does not create report churn while a stored layout or source change remains
visible. DocFence never resolves, retrieves, opens, authenticates to, imports,
or renders a template, subdocument, or frame target.

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

For a consequential legal, medical, financial, or publishing decision, use
DocFence as a controlled review signal alongside an appropriate rendering and
domain review process.
