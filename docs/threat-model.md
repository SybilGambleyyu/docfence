# Threat model

DocFence is designed for a repository, build agent, or handoff workstation that
needs a small amount of dependable evidence about an untrusted Word OOXML
package. It is not a general-purpose malware sandbox or a renderer.

## Assets and boundaries

The source document is treated as sensitive. Its body text, hidden text,
comments, reviewer metadata, relationship targets, field instructions, style
identifiers, custom XML, macro bytes, embedded-object bytes, and
alternative-format-import bytes, document-property names, and document-property
values must not become DocFence report content. The report may be stored in CI
artifacts, pasted into an issue, or uploaded to a SARIF consumer, so it is
intentionally restricted to counts, fixed categories, booleans, and generic
story kinds.

DocFence does not invoke Word or Office automation. It does not execute macro
code, calculate fields, resolve a hyperlink, retrieve a relationship target,
open an embedded object, import alternative-format content, render images, or
send package bytes over the network. It operates on a local copy of the ZIP
container with the Python standard library.

## Package and parser defenses

Before parsing, the reader rejects a source path that is not a regular file or
is a symlink. It reads only `.docx` and `.docm` paths and applies these default
limits:

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
stored custom definitions and is not a general PII detector.

For a consequential legal, medical, financial, or publishing decision, use
DocFence as a controlled review signal alongside an appropriate rendering and
domain review process.
