# Validation notes

DocFence 0.2 is validated as a parser and reporting boundary, not as a Word
renderer. The test suite constructs small OOXML packages with controlled body,
header, footer, footnote, endnote, comment, and glossary stories and checks the
following properties:

- stored paragraph/table changes are summarized without source text;
- revision markup, comments, direct `w:vanish` runs, direct hidden paragraph
  marks (including `w:specVanish`), stored style/default hidden-text
  declarations, fields, content controls, Track Changes, external
  relationships, custom XML, macros, and unclassified payload changes are
  detected by the intended inventories;
- direct false-valued hidden declarations and `w:specVanish` outside paragraph
  marks do not create false text-run findings;
- volatile Word `rsid` updates and relationship-ID renumbering that preserve the
  underlying target remain quiet;
- generated JSON, Markdown, and SARIF do not reproduce unique markers placed
  in visible/hidden text, reviewer metadata, comments, URLs, field instructions,
  custom XML, macros, or opaque package parts;
- policy failures return a nonzero CI status and SARIF uses no source location;
- DTD/entity markup and unsafe ZIP member names are rejected before reporting.

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

As a compatibility smoke test, the development environment also generated a
conventional package using `python-docx` and successfully profiled it. That tool
is not a DocFence dependency and is not required at runtime.

## What this does not validate

The suite does not claim layout equivalence, Word calculation behavior,
style-inherited effective hiddenness, macro safety, malware detection, or
compatibility with every vendor extension. The style/default layer is a stored
declaration inventory, not a renderer. Those limits are explicit in the 0.2
contract; see the [threat model](threat-model.md).
