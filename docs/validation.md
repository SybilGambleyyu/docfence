# Validation notes

DocFence 0.7 is validated as a parser and reporting boundary, not as a Word
renderer. The test suite constructs small OOXML packages with controlled body,
header, footer, footnote, endnote, comment, and glossary stories and checks the
following properties:

- stored paragraph/table changes are summarized without source text;
- revision markup, comments, direct `w:vanish` runs, direct hidden paragraph
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
  external template/subdocument/frame targets and frame names, macros,
  embedded/control payloads, alternative-format imports, or opaque package
  parts;
- policy failures return a nonzero CI status and SARIF uses no source location;
- DTD/entity markup, unsafe ZIP member names, and `w:altChunk` markup without a
  matching internal import relationship, malformed recognized document-property
  roots, malformed recognized mail-merge relationship references, and malformed
  recognized data-binding custom-XML-properties relationships, and malformed
  recognized external-document dependency relationship/anchor state are
  rejected before reporting.

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

## What this does not validate

The suite does not claim layout equivalence, Word calculation behavior,
style-inherited effective hiddenness, macro safety, embedded-payload safety,
malware detection, alternative-format rendering/import behavior, or
compatibility with every vendor extension. It also does not decide whether a
document-property value is personal, confidential, intended, or safe to share,
or whether stored mail-merge state will run, retrieve data, or select a
recipient. The style/default layer is a stored declaration inventory, not a
renderer. Those limits are explicit in the 0.7 contract; see the
[threat model](threat-model.md).
