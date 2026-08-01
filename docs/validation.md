# Validation notes

DocFence 0.9 is validated as a parser and reporting boundary, not as a Word
renderer. The test suite constructs small OOXML packages with controlled body,
header, footer, footnote, endnote, comment, and glossary stories and checks the
following properties:

- stored paragraph/table changes are summarized without source text;
- revision markup, comments, modern comment contact/thread/identifier/reaction
  metadata, direct `w:vanish` runs, direct hidden paragraph
  marks (including `w:specVanish`), stored style/default hidden-text
  declarations, fields, content controls, Track Changes, external
  relationships, custom XML, macros, mail-merge configuration, source/header
  relationships, recipient-data payloads, direct content-control XML mappings,
  referenced custom XML data/properties payloads, embedded OLE/package/control
  payloads, ActiveX control chains, alternative-format imports, core/extended/
  custom document properties (including Strict OOXML property variants),
  attached templates, master-document subdocuments, frameset sources
  (including Strict relationships and glossary-linked settings), and
  unclassified payload changes are detected by the intended inventories;
- `DATABASE`, legacy `DATA`, `DDE`, `DDEAUTO`, `INCLUDE`/`INCLUDETEXT`,
  `INCLUDEPICTURE`/`IMPORT`, `LINK`, and `RD` field instructions are
  separately categorized from simple and complex field encodings, including
  split, nested, resultless, Strict, and header-story cases; ordinary
  instruction text, post-separator text, and unclosed complex fields do not
  produce external-source field counts;
- tracked field-code replacements retain separately scanned current
  `w:instrText` and deleted `w:delInstrText` variants, including a changed
  argument with shared field-code text; loose deleted instruction text does not
  produce a count;
- standard `people`, `commentsExtended`, `commentsIds`, and
  `commentsExtensible` parts are separately inventoried from the ordinary
  comments story, including direct/noncanonical relationship targets, unlinked
  conventional paths, legacy Office 15 `2010/11` and current `2012` roots,
  thread/reply and resolved state, durable identifiers, and reaction/user
  records; malformed roots and external metadata relationships fail closed;
- modern-comment metadata changes with the same public counts—including a
  paragraph/durable identifier-only rewrite—produce a private-inventory change,
  while a relationship-ID renumbering alone remains quiet;
- direct false-valued hidden declarations and `w:specVanish` outside paragraph
  marks do not create false text-run findings;
- volatile Word `rsid` updates and relationship-ID renumbering that preserve the
  underlying target, embedded/import payload, document-property state,
  mail-merge state, data-binding storage association, and external-document
  dependency semantics remain quiet;
- generated JSON, Markdown, and SARIF do not reproduce unique markers placed
  in visible/hidden text, reviewer metadata, comments, URLs, field instructions,
  custom XML, document-property names or values, mail-merge connection/query/
  source/recipient values, data-binding XPath/prefix/storage/payload values,
  external-source field paths, connections, queries, application/item names,
  modern-comment authors, providers, user IDs, paragraph/durable IDs, dates,
  thread state, reaction identities,
  external template/subdocument/frame targets and frame names, macros,
  embedded/control payloads, alternative-format imports, or opaque package
  parts;
- policy failures return a nonzero CI status and SARIF uses no source location;
- DTD/entity markup, unsafe ZIP member names, and `w:altChunk` markup without a
  matching internal import relationship, malformed recognized document-property
  roots, malformed recognized mail-merge relationship references, and malformed
  recognized data-binding custom-XML-properties relationships, and malformed
  recognized external-document dependency relationship/anchor state, malformed
  modern-comment metadata roots, and external modern-comment metadata
  relationships are rejected before reporting.

The release check is:

```bash
.venv/bin/python -m pytest
.venv/bin/ruff check .
SOURCE_DATE_EPOCH="$(git log -1 --format=%ct)" .venv/bin/python -m build
.venv/bin/twine check dist/*
```

The package job uses the release commit's timestamp as `SOURCE_DATE_EPOCH`, so
the wheel and source distribution can be reproduced byte-for-byte from that
commit.

As compatibility smoke tests, the development environment generated a
conventional package using `python-docx` and profiled an independent
[custom-XML data-binding fixture](https://loadfix.github.io/ooxml-reference-corpus/case/docx__custom-xml-part.html)
from the OOXML Reference Corpus. It also reconstructed and profiled the
unpacked DOCX representation in the open-source [XJTU thesis Office
template](https://github.com/obster-y/XJTU-thesis-Office/tree/master/%E6%A8%A1%E6%9D%BF%E4%BD%BF%E7%94%A8%E8%AF%B4%E6%98%8E-docx),
which contains a real external attached-template relationship. Those tools and
fixtures are not DocFence dependencies and are not required at runtime.

For the modern-comment and template boundary, the release check profiles an
Office 15 `.dotx` conformance asset from Microsoft's open-source
[Open XML SDK](https://github.com/dotnet/Open-XML-SDK) and a reconstructed
package from the MIT-licensed
[stormdown-docx Word template repository](https://github.com/HarrisburgUniversityPhd/stormdown-docx).
The former exercises the legacy `commentsExtended` root vocabulary in a real
template; the latter contains real `people`, `commentsExtended`, and
`commentsIds` parts. These are compatibility smoke tests, not runtime
dependencies.

For the field encoding boundary, the release check also profiled Apache POI's
independent [`FieldCodes.docx`](https://github.com/apache/poi/blob/trunk/test-data/document/FieldCodes.docx)
and [`FldSimple.docx`](https://github.com/apache/poi/blob/trunk/test-data/document/FldSimple.docx)
fixtures. They exercise real complex and simple Word field encodings without
creating a false external-source-field count. They are compatibility smoke
tests, not runtime dependencies.

## What this does not validate

The suite does not claim layout equivalence, Word calculation behavior,
style-inherited effective hiddenness, macro safety, embedded-payload safety,
malware detection, alternative-format rendering/import behavior, or
compatibility with every vendor extension. It also does not decide whether a
document-property value is personal, confidential, intended, or safe to share,
or whether stored mail-merge state will run, retrieve data, or select a
recipient. It does not assert that a recognized external-source field will
update, reach a target, use a particular argument as a source, or be accepted
by Word; it records the bounded stored field-family evidence only. The
style/default layer is a stored declaration inventory, not a renderer. Those
limits are explicit in the 0.9 contract; see the
[threat model](threat-model.md).
