# Changelog

All notable changes are documented here.

## 0.3.0 — 2026-08-01

- Added separate privacy-safe inventories for embedded OLE/package payloads,
  embedded controls (including ActiveX control-binary relationships), and OOXML
  alternative-format imports.
- Added direct `w:altChunk` anchor inventory and fail-closed validation that an
  encountered anchor names an internal standard `aFChunk` relationship with a
  stored target.
- Added opt-in absence gates `DFP015` and `DFP016`, plus comparison gates
  `DFP017` and `DFP018` for controlled templates that allow known embedded or
  imported payloads but reject mutations.
- Removed recognized embedded and alternative-format payloads from the generic
  unclassified-payload inventory while retaining aggregate change reporting.
- Added regression coverage for a conventional ActiveX relationship chain,
  embedded/import payload changes, relationship-ID renumbering, malformed
  `altChunk` references, policy findings, and report redaction.

## 0.2.0 — 2026-08-01

- Added separate direct hidden paragraph-mark inventory for `w:vanish` and
  `w:specVanish` markup.
- Added privacy-safe `word/styles.xml` inventory for stored text-run hidden
  declarations and document-default run properties, plus dedicated style change
  reporting.
- Added opt-in policy rules `DFP013` for hidden text style/default declarations
  and `DFP014` for hidden paragraph marks.
- Clarified that `w:specVanish` belongs to paragraph-mark handling and is not
  treated as ordinary hidden-text run markup.
- Added regression coverage for style declaration scope, false values,
  paragraph-mark semantics, malformed styles parts, and report redaction.
- Made CI distribution builds reproducible from the release commit via
  `SOURCE_DATE_EPOCH`.

## 0.1.0 — 2026-08-01

- Initial local-first DOCX/DOCM change-assurance CLI.
- Added bounded OOXML package inspection and private semantic fingerprints.
- Added privacy-safe JSON, Markdown, and SARIF reports.
- Added policy gates for hidden review surfaces and package-state changes.
- Added regression coverage for report redaction, story families, volatile Word
  metadata, ZIP/XML safety boundaries, CLI status, and policy parsing.
