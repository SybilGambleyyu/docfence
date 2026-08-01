# Threat model

DocFence is designed for a repository, build agent, or handoff workstation that
needs a small amount of dependable evidence about an untrusted Word OOXML
package. It is not a general-purpose malware sandbox or a renderer.

## Assets and boundaries

The source document is treated as sensitive. Its body text, hidden text,
comments, reviewer metadata, relationship targets, field instructions, style
identifiers, custom XML, and macro bytes must not become DocFence report
content. The report may be stored in CI artifacts, pasted into an issue, or
uploaded to a SARIF consumer, so it is intentionally restricted to counts,
fixed categories, booleans, and generic story kinds.

DocFence does not invoke Word or Office automation. It does not execute macro
code, calculate fields, resolve a hyperlink, retrieve a relationship target,
render images, or send package bytes over the network. It operates on a local
copy of the ZIP container with the Python standard library.

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
or business effect.

For a consequential legal, medical, financial, or publishing decision, use
DocFence as a controlled review signal alongside an appropriate rendering and
domain review process.
