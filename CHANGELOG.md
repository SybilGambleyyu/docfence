# Changelog

All notable changes are documented here.

## 0.6.0 — 2026-08-01

- Added a separate privacy-safe inventory for direct Word content-control XML
  mappings, storage-ID presence, matched custom XML data parts, and unmatched
  storage identifiers.
- Added `DFP023` to require that a candidate contain no recognized data-binding
  declarations and `DFP024` to block data-binding inventory changes against a
  controlled baseline.
- Privately fingerprints mapping declarations and, when a storage ID can be
  associated safely, the paired custom XML data and properties payloads without
  exposing XPath expressions, prefix mappings, storage IDs, part names, or
  values.
- Validates recognized custom XML properties relationships and roots used for
  binding association; malformed internal associations fail closed while an
  unmatched storage ID remains explicit review evidence.
- Added conventional, Strict, established Word-legacy root-vocabulary, unscoped,
  unmatched, relationship-ID-stability, malformed-association, policy, SARIF,
  and redaction regression coverage. Profiled an independent public OOXML
  reference-corpus data-binding package as a compatibility smoke test.

## 0.5.0 — 2026-08-01

- Added a separate privacy-safe inventory for Word mail-merge configuration,
  external data/header-source relationships, and internal recipient-data parts.
- Added `DFP021` to require that a candidate contain no stored mail-merge state
  and `DFP022` to block mail-merge inventory changes against a controlled
  baseline.
- Validated direct mail-merge source, header-source, and recipient-data
  references against their required relationship types and target modes;
  malformed recognized state fails closed.
- Removed recognized recipient-data parts from the generic unclassified-payload
  inventory while retaining their private payload comparison.
- Added regression coverage for conventional and Strict relationships, both
  documented recipient-data relationship spellings, query and recipient changes,
  orphaned source relationships, relationship-ID renumbering, malformed
  references, policy results, and redaction.

## 0.4.0 — 2026-08-01

- Added separate privacy-safe inventories for core, extended, and custom OOXML
  document-property parts, including aggregate part/value counts only.
- Added `DFP019` to block document-property inventory changes and `DFP020` to
  require that a candidate have no stored custom-property definitions.
- Validated recognized document-property roots and removed recognized property
  parts from the generic unclassified-payload inventory.
- Added regression coverage for core/extended/custom property changes, physical
  canonical property paths, Strict OOXML property variants, malformed property
  roots, relationship-ID renumbering, policy results, and metadata-name/value
  redaction.
- Documented that core and extended counts include automatic metadata and do not
  classify a property as personal, confidential, intentional, or safe.

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
