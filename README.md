# DocFence

DocFence is a local-first change-assurance CLI for Word `.docx` and `.docm`
files. It turns an opaque document diff into a reviewable, privacy-safe account
of stored content-block changes and the review surfaces that often stay hidden:
tracked revisions, comments, hidden runs and paragraph marks, stored
style/default declarations, field codes, external relationships, custom XML,
macros, headers, footers, notes, document settings, and otherwise unclassified
package payloads.

It never opens Word, executes macros, follows links, renders a document, uploads
source material, or writes a redline. By default, reports contain counts,
story categories, and change categories—not document text, reviewer
names, comments, URLs, relationship targets, field instructions, custom XML,
or macro bytes.

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

Version 0.2 focuses on Office Open XML Word documents and deliberately keeps a
small, inspectable contract:

- bounded `.docx` / `.docm` ZIP packages;
- body, header, footer, footnote, endnote, comment, and glossary stories;
- paragraph/table block fingerprints that ignore Word's volatile `rsid`
  bookkeeping while retaining stored text and formatting semantics privately;
- revision markup, comments, direct hidden-text runs, direct hidden
  paragraph-mark markup, stored style/default hidden-text declarations,
  field-code, content-control, external-relationship, custom-XML, macro, and
  Track Changes inventories;
- a generic alert for changed package payload outside those specialized
  inventories (for example, media, embedded, or metadata parts);
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
```

YAML anchors, aliases, sequences, nested mappings, duplicate keys, unknown
rules, and non-boolean values are rejected. That keeps a policy reviewable and
avoids making the CLI's safety contract depend on a broad YAML loader.

## Privacy contract

DocFence keeps document material in memory only long enough to create private
SHA-256 fingerprints. Public models and renderers intentionally omit package
part names, paragraph content, reviewer identity, dates, comment content,
relationship targets, field instructions, style identifiers and names, custom
XML values, macro bytes, and all fingerprints. Regression tests place unique
sensitive markers in each of those surfaces and assert that JSON, Markdown, and
SARIF never reproduce them.

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
