# Changelog

All notable changes are documented here.

## 0.15.0 — 2026-08-01

- Added a first-class, privacy-safe inventory for stored Word document
  variables in direct Settings-part `w:docVars` / `w:docVar` markup.
- Added `DFP043` to require a handoff with no stored document-variable state
  and `DFP044` to protect an approved document-variable baseline.
- Reports now expose aggregate container, variable, and empty-value counts
  only. Variable names, values, Settings-part paths, and fingerprints remain
  private.
- Discovers Settings through the conventional path and Transitional or Strict
  relationships from main/glossary documents. It validates the standard
  container/leaf shape, required `name`/`val` attributes, Word namespace,
  SDK length limits, and at-most-one container per Settings part; malformed
  recognized variable state fails closed.
- Privately detects same-count name or value rewrites without evaluating a
  `DOCVARIABLE` field, running a macro, or asserting that any automation state
  will be used. Added regression coverage for privacy redaction, policy/SARIF
  output, discovery modes, Strict OOXML, empty values, and malformed markup.

## 0.14.0 — 2026-08-01

- Added a first-class, privacy-safe inventory for Word editable-range
  `w:permStart` / `w:permEnd` markup across supported document stories.
- Added `DFP041` to require a handoff with no stored editable-range permission
  markup and `DFP042` to protect an approved range-permission baseline.
- Reports now expose aggregate marker, paired/unpaired-range, individual-editor
  assignment, predefined-group, table-column-selector, and custom-XML-placement
  counts only. Individual editor values, marker IDs, exact column values, part
  paths, and fingerprints remain private.
- Accepts Word's predefined editing-group vocabulary and Transitional/Strict
  Word namespaces. It validates recognized marker leaf shape, required IDs,
  allowed attributes, group values, column syntax, placement values, and
  unambiguous per-story IDs. Unmatched start/end markers are inventoried as
  stored review state rather than presented as effective authorization.
- Privately detects same-count editor-identity or range-shape changes without
  exposing identities or claiming that an editor is authenticated, currently
  authorized, or able to edit a document. Added regression coverage for
  privacy redaction, policy/SARIF output, multiple stories, Strict OOXML,
  unmatched markers, and malformed markup.

## 0.13.0 — 2026-08-01

- Added a first-class, privacy-safe inventory for stored Word
  `documentProtection` editing restrictions and `writeProtection` state.
- Added `DFP039` to require a handoff with no stored Word protection state and
  `DFP040` to protect an approved protection-state baseline.
- Reports now expose only aggregate protection-element, explicitly-enabled
  enforcement, formatting-restriction, edit-mode, read-only-recommendation, and
  password-material counts. Hashes, salts, verifier values, provider and
  algorithm fields, Settings-part paths, and fingerprints remain private.
- Discovers document Settings through the conventional path and Transitional or
  Strict relationships from main/glossary documents. It validates direct
  protection-element shape, known attributes, edit modes, and booleans; duplicate
  or malformed recognized protection state fails closed.
- Privately detects same-count password-material changes without validating a
  password, estimating its strength, bypassing protection, inferring effective
  enforcement, or claiming document encryption/security. Added regression
  coverage for privacy redaction, policy/SARIF output, malformed markup, and
  public Open XML SDK protection-fixture compatibility smokes.

## 0.12.0 — 2026-08-01

- Added a first-class, privacy-safe inventory for OPC package digital-signature
  origin, XML-signature, and certificate material.
- Added `DFP037` to require that a handoff contain no stored package-signature
  material and `DFP038` to protect an approved signature-material baseline.
- Reports now expose aggregate origin-part, XML-signature-part, certificate-part,
  SignedInfo-reference, manifest-reference, relationship-reference,
  inline-X.509-certificate, and signature-property counts only. Signer and
  certificate data, signature values, algorithms, reference URIs, signing
  times, comments, provider data, IDs, paths, and fingerprints remain private.
- Detects standard origin/signature/certificate relationships and exact OPC
  content types (including defaults), plus conventional origin residue.
  Recognized relationships are constrained to their expected source, internal
  target, stored member, and content type; origin relationships are unique.
  Recognized XML-signature parts receive bounded XMLDSIG-shape validation and
  malformed state fails closed.
- Privately detects same-count signature and certificate rewrites, while
  explicitly not claiming cryptographic verification, certificate trust,
  signature coverage, signer identity, or Office trust behavior. Added
  regression coverage for privacy redaction, policy/SARIF output, discovery
  modes, malformed topology and XMLDSIG shape, and a public signed-package
  compatibility smoke test.

## 0.11.0 — 2026-08-01

- Added a first-class, privacy-safe inventory for Microsoft Purview sensitivity
  label metadata in both Office 2021 `LabelInfo` parts and legacy MIP custom
  document properties.
- Added `DFP035` to require that a handoff contain no stored sensitivity-label
  metadata and `DFP036` to protect an approved label-metadata baseline.
- Reports now expose aggregate LabelInfo part/label/enabled/removed/extension,
  legacy MIP label/property, legacy `Sensitivity` property, and Word
  content-marking-property counts only. Label and tenant IDs, names, dates,
  action IDs, extension data, custom MIP attributes, marking strings, paths,
  and fingerprints remain private.
- Detects modern label parts through the Office classification-label
  relationship, SDK content type, and canonical `docMetadata/LabelInfo` paths.
  Validates the `labelList` root and required label state, requires a recognized
  relationship to be internal, root-package scoped, and resolvable, and fails
  closed on malformed recognized metadata.
- Privately fingerprints same-count LabelInfo and legacy-property mutations,
  while relationship-ID renumbering alone remains quiet. Added adversarial
  coverage for discovery modes, malformed roots/attributes/site IDs,
  unavailable/external/non-root relationships, multiple LabelInfo parts,
  noncanonical custom-property parts, policy/SARIF output, and redaction.

## 0.10.0 — 2026-08-01

- Added first-class, privacy-safe inventories for Word document tasks and
  document-borne task-pane Office web-extension state.
- Added `DFP031` / `DFP032` to prohibit document-task state or protect an
  approved task baseline, and `DFP033` / `DFP034` to prohibit task-pane
  web-extension state or protect an approved add-in baseline.
- Reports now expose aggregate task/history/user-reference/comment-anchor/event
  counts, plus aggregate task-pane, extension, reference, property, binding,
  auto-show, and Word content-control binding-marker counts. IDs, identities,
  titles, dates, stores, properties, bindings, pane settings, and paths remain
  private.
- Detects recognized parts through standard content types, relationship types,
  and conventional extensionful/extensionless Word paths. Validates each root,
  requires direct task-pane web-extension references to resolve through the
  expected internal relationship, and fails closed on malformed recognized
  state.
- Privately fingerprints document-task and Office web-extension payload state
  so same-count changes remain review-visible. Relationship-ID renumbering alone
  remains quiet. Supports both `webextensionref` and the established
  `webextension` task-pane reference spelling; `webExtensionCreated` takes
  precedence over `webExtensionLinked` for bound Word content controls.
- Added regression coverage for semantic changes, policy/SARIF output,
  redaction, noncanonical and unlinked conventional paths, extensionless paths,
  invalid roots/references, external relationships, and relationship-ID churn.

## 0.9.0 — 2026-08-01

- Added first-class bounded scanning for Word `.dotx` and `.dotm` template
  packages alongside `.docx` and `.docm` documents.
- Added a separate privacy-safe inventory for modern Word comment metadata in
  the standard `people`, `commentsExtended`, `commentsIds`, and
  `commentsExtensible` parts. Public output contains aggregate contact, thread,
  resolution, identifier-record, reaction, and reaction-user counts only.
- Added `DFP029` to require that a candidate contain no modern-comment metadata
  and `DFP030` to block a modern-comment metadata inventory change against a
  controlled baseline.
- Validates each recognized part root, accepts both established Office 15
  `2010/11` and current `2012` people/comments-extended vocabularies, and
  requires a recognized metadata relationship to be internal and resolve to a
  stored part. Recognized metadata parts are removed from the generic opaque
  payload inventory.
- Privately fingerprints contact/provider data, paragraph and durable IDs,
  timestamps, extensions, and reaction data while preserving those identifiers
  for comparison. Relationship-ID renumbering remains quiet, but an
  identifier-only rewrite is detected without disclosing it.
- Added noncanonical-path, unlinked-conventional-path, legacy-vocabulary,
  template-format, same-count mutation, malformed-root, external-relationship,
  policy, SARIF, and redaction regression coverage. Profiled real modern
  comment/template packages from the Open XML SDK and an independent
  MIT-licensed Word template repository.

## 0.8.0 — 2026-08-01

- Added a separate privacy-safe inventory for stored Word field instructions
  that can source or query material outside the package: `DATABASE`, legacy
  `DATA`, `DDE`, `DDEAUTO`, `INCLUDE`/`INCLUDETEXT`,
  `INCLUDEPICTURE`/`IMPORT`, `LINK`, and `RD`.
- Added `DFP027` to require that a candidate contain no recognized
  external-source field instructions and `DFP028` to block external-field
  inventory changes against a controlled baseline.
- Handles simple `w:fldSimple` instructions and complete, nested complex-field
  begin/separate/end sequences without treating standalone or result-side
  `w:instrText` as a field instruction. Complex fields without a result
  separator are supported; an unclosed complex field is ignored. Tracked
  `w:delInstrText` field-code variants are separately inventoried rather than
  concatenated with current `w:instrText` variants.
- Privately fingerprints complete instructions and their story context while
  reporting only field-family counts. Source paths, connection strings, SQL,
  application names, item references, OLE details, and fingerprints are never
  emitted. Splitting one unchanged complex instruction over a different number
  of runs remains quiet.
- Added conventional, Strict, header-story, nested, resultless, split-run,
  target-change, tracked-deletion, non-field-text, unclosed-field, policy,
  SARIF, and redaction regression coverage.

## 0.7.0 — 2026-08-01

- Added a separate privacy-safe inventory for three standard external Word
  document dependency families: attached templates, master-document
  subdocuments, and frameset source files.
- Added `DFP025` to require that a candidate contain no recognized external
  Word document dependency state and `DFP026` to block dependency-inventory
  changes against a controlled baseline.
- Validates expected conventional and Strict relationship types, direct anchors,
  and `TargetMode="External"`; malformed recognized state fails closed while
  residual recognized relationships remain explicit review evidence.
- Discovers Settings and Web Settings parts from main or glossary documents,
  keeps the conventional Settings path as a compatibility fallback, and
  privately fingerprints an involved Web Settings part to make ID renumbering
  quiet without losing frame-layout/source changes.
- Added conventional, Strict, glossary-linked-settings, orphaned-relationship,
  target-change, relationship-ID-stability, malformed-state, policy, SARIF, and
  redaction regression coverage. Profiled a reconstructed open-source thesis
  template package with a real attached-template relationship as a compatibility
  smoke test.

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
