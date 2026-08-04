# Validation notes

DocFence 0.49 is validated as a parser and reporting boundary, not as a Word
renderer. The test suite constructs small OOXML packages with controlled body,
header, footer, footnote, endnote, comment, and glossary stories and checks the
following properties:

- stored paragraph/table changes are summarized without source text;
- revision markup, comments, modern comment contact/thread/identifier/reaction
  metadata, document tasks, task-pane Office web-extension configuration,
  Office sensitivity-label LabelInfo and legacy custom-property metadata, and
  relationship-bound OPC package thumbnail, OOXML Markup Compatibility branch/
  compatibility-rule, and digital-signature
  origin/XML-signature/certificate material,
  Word `w:documentProtection`/`w:writeProtection` state, and
  Word Settings `w:docVars`/`w:docVar` state and `DOCVARIABLE` field-reference
  state, Word Settings XSLT-on-single-XML-save, attached-custom-XML-schema,
  automatic-field-recalculation-on-open, form-data-only-save, and
  preview-thumbnail-on-save state, Word `HYPERLINK` field-reference state,
  direct WordprocessingML
  `w:hyperlink` markup state, and direct DrawingML click/hover/mouse-over
  hyperlink-action markup state, direct DrawingML nonvisual hidden-state
  declaration state, direct DrawingML `a:blip/@r:link`
  linked-picture markup state, direct legacy VML shape-link `href` markup
  state, direct legacy VML external-image `v:imagedata/@r:id` markup state,
  direct legacy VML image-data hyperlink `v:imagedata/@r:href` markup state,
  direct legacy Office VML linked-OLE `o:OLEObject Type="Link"` markup state,
  direct legacy Office VML embedded-OLE `o:OLEObject Type="Embed"` markup
  state, direct WordprocessingML `w:object/w:objectLink` linked-object-property
  markup state, direct WordprocessingML `w:object/w:control` and
  `w:pict/w:control` embedded-control-anchor markup state, and
  Word content-control web-extension markers, direct `w:vanish` runs, and
  direct hidden paragraph
  marks (including `w:specVanish`), stored style/default hidden-text
  declarations, fields, content controls, Track Changes, external
  relationships, custom XML, macros, mail-merge configuration, source/header
  relationships, recipient-data payloads, direct content-control XML mappings,
  direct content-control lock declarations, referenced custom XML
  data/properties payloads, embedded OLE/package/control
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
- document-task parts are discovered through standard relationships, content
  types, and conventional paths (including extensionless paths); public
  task/history/user/comment-anchor/event counts remain aggregate-only;
  identifier- or value-only task rewrites remain visible privately while
  relationship-ID renumbering alone remains quiet;
- task-pane and web-extension parts are discovered through standard
  relationships, content types, and conventional paths (including an
  extensionless task-pane path); both `webextensionref` and established
  `webextension` direct task-pane references are accepted; malformed roots,
  external recognized relationships, and missing direct references fail closed;
- same-count Office web-extension store/property/binding rewrites produce a
  private-inventory change; enabled `webExtensionCreated` takes precedence over
  `webExtensionLinked` when counting document-bound content controls;
- Office sensitivity-label metadata is separately inventoried from the general
  document-property count across modern LabelInfo and legacy MIP property
  storage. LabelInfo discovery covers standard root-package relationships,
  Office SDK content types, canonical extensionful/extensionless paths, and
  noncanonical relationship targets; legacy MIP discovery covers noncanonical
  custom-property parts reached through a standard relationship. Same-count
  label and legacy-property rewrites remain review-visible while relationship-ID
  renumbering alone remains quiet;
- OPC package digital-signature material is separately inventoried from generic
  opaque payloads. Tests cover origin relationships and default content types,
  noncanonical relationship targets, content-type-only signature residue,
  conventional orphan origins, same-count signature/certificate mutations,
  JSON/Markdown/SARIF redaction, policy findings, malformed XMLDSIG roots and
  SignedInfo shape, unavailable/external/non-root/duplicate origin
  relationships, and unavailable/external signature and certificate targets.
  The static declared-coverage boundary has a fully declared synthetic package
  fixture plus uncovered Word parts and relationships, an unbound package
  object, a missing `SignatureProperties` child, a nonstandard package-object
  ID, an extra package-object child, duplicate package objects and bindings,
  all six valid OPC `SignatureTime` precision forms and its permitted empty
  target, plus missing, duplicate, misidentified, mis-targeted, attribute-
  bearing, missing-child, extra-child, nonstandard-format, and
  format/value-mismatched timestamp declarations,
  direct package-object binding references with no transform or either
  canonicalization form, plus unsupported/relationship/empty/duplicate
  transform lists, `ds:XPath` parameters in binding, relationship, and part
  transforms, and missing, misordered, malformed, nested, extra, or
  text-bearing binding digest children,
  unresolved and unsupported manifest references, a case-mismatched content
  type, standard `RelationshipsGroupReference/@SourceType` selectors,
  malformed nonstandard `RelationshipReference/@SourceType` lookalikes,
  both OPC XML Canonicalization algorithms, missing or misordered
  canonicalization after a relationship transform, duplicate direct
  `ds:Transforms` lists, duplicate relationship transforms for one
  relationships part, direct part references with permitted C14N-only and
  unsupported empty/duplicate/relationship/unknown transform lists, an
  unsupported trailing transform, missing or misordered XMLDSIG digest
  children, a missing digest-method algorithm, an unexpected digest-value
  attribute or nested element, unexpected reference text, an extra direct
  digest child, same-count selector reassignment,
  private-output redaction,
  Markdown/JSON/SARIF projection, and both coverage policy modes;
- Word editing/write-protection state is separately inventoried from generic
  Settings-part changes. Tests cover conventional, Transitional, and Strict
  Settings discovery; aggregate editing/write/recommendation/password-material
  counts; same-count hash/verifier rewrites; JSON/Markdown/SARIF redaction;
  policy findings; duplicate leaves; malformed roots/leaves; unsupported
  attributes; invalid edit/boolean values; and external Settings relationships;
- Word document-variable state is separately inventoried from generic
  Settings-part changes. Tests cover conventional, Transitional, Strict, and
  glossary-linked Settings discovery; aggregate container/variable/empty-value
  counts; same-count name/value rewrites; JSON/Markdown/SARIF redaction; policy
  findings; and malformed containers/leaves, attributes, required names/values,
  SDK length limits, nonblank text, duplicate containers, and external Settings
  relationships;
- XSLT-on-single-XML-save configuration is separately inventoried from generic
  Settings and external-relationship changes. Tests cover enabled and disabled
  settings, relationship-backed and local-identifier-only anchors, residual
  relationships, Transitional and Strict namespaces, same-count target and
  solution-identifier rewrites, relationship-ID renumbering, privacy
  redaction, policy findings, and malformed relationship type, target mode,
  missing relationship, and duplicate-anchor state;
- attached custom XML schema declarations are separately inventoried from
  generic Settings and custom-XML payload changes. Tests cover direct
  `w:attachedSchema` leaves in conventional and Strict OOXML, multiple
  declarations, same-count namespace rewrites, JSON/Markdown/SARIF redaction,
  policy findings, and malformed missing/unsupported attributes, nonblank text,
  and nested markup;
- automatic field-recalculation-on-open state is separately inventoried from
  generic Settings and field-code changes. Tests cover absent, implicit-enabled,
  explicitly enabled, explicitly disabled, and Strict leaves; canonical
  equivalent enabled spellings; JSON/Markdown/SARIF redaction; policy findings;
  and malformed unsupported attributes/boolean values, nonblank text, nested
  markup, and duplicate direct leaves;
- automatic template-style-update-on-open state is separately inventoried from
  generic Settings and external-document-dependency changes. Tests cover absent,
  implicit-enabled, explicitly enabled, explicitly disabled, and Strict leaves;
  canonical equivalent enabled spellings; JSON/Markdown/SARIF redaction; policy
  findings; and malformed unsupported attributes/boolean values, nonblank text,
  nested markup, and duplicate direct leaves;
- personal-information-removal-on-save state is separately inventoried from
  generic Settings and document-property changes. Tests cover absent,
  implicit-enabled, explicitly enabled, explicitly disabled, and Strict leaves;
  canonical equivalent enabled spellings; JSON/Markdown/SARIF redaction;
  positive-request and baseline policy findings; and malformed unsupported
  attributes/boolean values, nonblank text, nested markup, and duplicate direct
  leaves;
- form-data-only-save state is separately inventoried from generic Settings
  and field-code changes. Tests cover absent, implicit-enabled, explicitly
  enabled, explicitly disabled, and Strict leaves; canonical equivalent enabled
  spellings; JSON/Markdown/SARIF redaction; candidate and baseline policy
  findings; and malformed unsupported attributes/boolean values, nonblank text,
  nested markup, and duplicate direct leaves. This validates stored
  configuration only: no test opens Word, evaluates a form field, reads a field
  value, saves a document, or emits a delimited record;
- preview-thumbnail-on-save state is separately inventoried from generic
  Settings and actual package-thumbnail state. Tests cover absent,
  implicit-enabled, explicitly enabled, explicitly disabled, and Strict
  leaves; canonical equivalent enabled spellings; JSON/Markdown/SARIF
  redaction; candidate and baseline policy findings; and malformed unsupported
  attributes/boolean values, nonblank text, nested markup, and duplicate direct
  leaves. This validates a stored future-save request only: no test opens Word,
  saves a document, creates, decodes, or renders a thumbnail, or asserts a
  host will generate one;
- direct content-control lock state is separately inventoried across supported
  Word stories. Tests cover absent direct declarations, all four exact
  `w:lock/@w:val` schema values, same-count state reassignment, Strict OOXML,
  out-of-scope lookalikes, JSON/Markdown/SARIF redaction, candidate and
  baseline policy findings, and malformed missing/unsupported attributes,
  nonblank text, nested markup, duplicate direct lock leaves, and duplicate
  direct properties. This validates direct stored declarations only: no test
  opens Word, reads a control value, identifies a control, evaluates a data
  binding, applies a lock, or asserts client enforcement;
- the conventional custom-XML candidate gate is separately tested from the
  baseline-comparison rule. A document with unbound `customXml/` data triggers
  `DFP079` even against itself, a clean document passes, and JSON, Markdown,
  and SARIF reports omit both visible text and custom-XML payload markers;
- relationship-bound OPC package thumbnails are separately inventoried from
  generic opaque payloads. Tests cover package and part relationships,
  Transitional and Strict relationship forms, same-count image payload and
  content-type changes,
  unreferenced filename lookalikes, JSON/Markdown/SARIF redaction, candidate
  and baseline policy findings, and rejection of external, duplicate, missing,
  non-image, source-unavailable, or relationship-owning thumbnail topology;

- OOXML Markup Compatibility is separately inventoried across stored Word XML
  parts. Tests cover aggregate `AlternateContent`/`Choice`/`Fallback` and
  compatibility-rule token counts, a same-count `Choice/@Requires` rewrite,
  a non-MCE text rewrite that does not churn the MCE inventory,
  JSON/Markdown/SARIF redaction, and candidate and baseline policy findings;

- DrawingML nonvisual visibility is separately inventoried in supported Word
  stories. Tests cover standard Transitional and Strict main/picture/docPr
  forms, Word 2010 w14/wpg/wps forms, direct hidden and explicitly-shown
  states, invalid Boolean values, same-count state swaps, equivalent Boolean
  spellings, unrelated object-metadata/text rewrites, JSON/Markdown/SARIF
  redaction, and candidate and baseline policy findings. This is a direct
  stored-marker test only; it does not calculate effective visibility, select
  an MCE branch, render, or assert client behavior;

- the validation pass profiles the public Open XML SDK Strict Word chart
  [`2D Column-O12-Word-Charts.docx`](https://github.com/dotnet/Open-XML-SDK/blob/main/test/DocumentFormat.OpenXml.Tests.Assets/assets/TestDataStorage/O14ISOStrict/Graphics/2D%20Column-O12-Word-Charts.docx).
  It reports nine MCE-bearing Word XML parts, three `AlternateContent` nodes,
  three `Choice` nodes, three `Fallback` nodes, three `Requires` prefixes, and
  eight `Ignorable` prefixes. This is a stored-package compatibility smoke test
  only; it does not select a branch or assert how any document client will
  process it;

- the release check also profiles the public Open XML SDK
  [`AcademicLetter` Word template](https://github.com/dotnet/Open-XML-SDK/blob/main/test/DocumentFormat.OpenXml.Tests.Assets/assets/TestDataStorage/v2FxTestFiles/wordprocessing/O12%20templates/AcademicLetter_TP10067035.dotx).
  It reports one thumbnail relationship and one thumbnail part. This is a
  package-compatibility smoke test only; it does not inspect the image or
  assert how any document client will display it;

- `DOCVARIABLE` field references are separately inventoried from generic field
  counts and stored-variable state. Tests cover simple and complete complex
  encodings, quoted names, trailing Word formatting switches, nested/dynamic
  expressions, current/deleted revision variants, headers, Strict Word
  namespaces, main-versus-glossary exact-literal association, unclosed/loose
  instruction exclusion, same-count changes, JSON/Markdown/SARIF redaction,
  and policy findings;
- `HYPERLINK` field references are separately inventoried from relationship
  hyperlinks and external-source field families. Tests cover simple and complete
  split complex encodings, quoted/plain leading arguments, `\l`
  internal-location-only fields, trailing field switches, nested/dynamic and
  compound expressions, current/deleted revision variants, headers, Strict Word
  namespaces, unclosed/loose/post-separator instruction exclusion, same-count
  destination changes, JSON/Markdown/SARIF redaction, and policy findings;
- direct WordprocessingML `w:hyperlink` markup is separately inventoried from
  both `HYPERLINK` fields and generic relationship totals. Tests cover external
  and internal recognized hyperlink relationships, a resolved unsupported
  relationship, anchor-only and no-attribute current-document-start forms,
  `r:id` precedence over a shadowed `w:anchor`, body/header stories,
  Transitional and Strict Word/relationship namespaces, orphaned relationship
  exclusion, same-count target changes, relationship-ID renumbering stability,
  JSON/Markdown/SARIF redaction, and policy findings;
- direct DrawingML `a:hlinkClick`, `a:hlinkHover`, and `a:hlinkMouseOver`
  action markup is separately inventoried from direct `w:hyperlink` markup,
  `HYPERLINK` fields, and broad relationship totals. Tests cover click/hover/
  mouse-over kinds, recognized external and internal relationships, a resolved
  unsupported relationship, missing `r:id` stored evidence, `action` and
  `invalidUrl` attribute presence, body/header stories, Transitional and Strict
  DrawingML/Word/relationship namespaces, orphaned relationship exclusion,
  same-count target and action changes, relationship-ID renumbering stability,
  JSON/Markdown/SARIF redaction, and policy findings;
- direct DrawingML `a:blip/@r:link` linked-picture markup is separately
  inventoried from direct `w:hyperlink` markup, `HYPERLINK` fields,
  DrawingML hyperlink actions, `r:embed`-only image references, and generic
  relationship totals. Tests cover standard image relationships with external
  and internal stored target modes, a resolved unsupported relationship,
  dual `r:link`/`r:embed` attributes, body/header stories, Transitional and
  Strict DrawingML/Word/relationship namespaces, orphaned image-relationship
  exclusion, same-count target and direct-markup changes, relationship-ID
  renumbering stability, JSON/Markdown/SARIF redaction, and policy findings;
- direct legacy VML `href` markup is separately inventoried from direct
  `w:hyperlink` markup, `HYPERLINK` fields, DrawingML actions, and generic
  relationship totals. Tests cover all supported direct VML element kinds
  (`arc`, `curve`, `image`, `line`, `oval`, `polyline`, `rect`, `roundrect`,
  `shape`, `group`, and `shapetype`), body/header stories, aggregate
  concrete-shape/group/shape-template and `target`-attribute-presence counts,
  an empty direct `href`, same-count `href` and target rewrites,
  JSON/Markdown/SARIF redaction, and
  policy findings;
- direct legacy VML `v:imagedata/@r:id` markup is separately inventoried from
  DrawingML linked pictures, VML shape `href`, `HYPERLINK` fields, direct
  `w:hyperlink` markup, and generic relationship totals. Tests cover external
  standard-image and unsupported relationship types, excluded internal image
  relationships, duplicate markers, body/header stories, Transitional and
  Strict Word/relationship namespaces, orphaned relationship exclusion,
  excluded raw-`src`, `r:pict`, and `r:href` attributes, same-count external
  target changes, raw-`src` change quietness, relationship-ID renumbering
  stability, JSON/Markdown/SARIF redaction, and policy findings;
- direct legacy VML `v:imagedata/@r:href` markup is separately inventoried from
  VML image-data `r:id` external-image markers, VML shape `href`, DrawingML
  linked pictures, `HYPERLINK` fields, direct `w:hyperlink` markup, and generic
  relationship totals. Tests cover standard hyperlink relationships with
  external and internal stored target modes, a resolved unsupported image
  relationship, duplicate markers, body/header stories, Transitional and
  Strict Word/relationship namespaces, orphaned relationship exclusion,
  excluded `r:id`, `r:pict`, raw-`src`, and `o:relid` attributes, same-count
  target changes, excluded-attribute change quietness, relationship-ID
  renumbering stability, JSON/Markdown/SARIF redaction, and policy findings;
- direct legacy Office VML `o:OLEObject Type="Link"` markup is separately
  inventoried from broad embedded-object relationship/payload totals,
  the separately inventoried `Type="Embed"`, WordprocessingML `w:objectLink`,
  VML image data, VML shape links, DrawingML linked pictures, fields, and
  generic relationship totals. Tests cover standard OLE-object relationships
  with external and internal stored target modes, a resolved unsupported
  non-OLE relationship, a direct marker without `r:id`, `UpdateMode="Always"`
  and other update modes, duplicate markers, body/header stories, Transitional
  and Strict Word/relationship namespaces, orphaned relationship exclusion,
  unavailable relationship failure, same-count target/source-marker changes,
  relationship-ID renumbering stability, JSON/Markdown/SARIF redaction, and
  policy findings;
- direct legacy Office VML `o:OLEObject Type="Embed"` markup is separately
  inventoried from `Type="Link"`, WordprocessingML `w:objectEmbed` and
  `w:objectLink`, VML image data and shape links, fields, and broad
  embedded-object relationship/payload totals. Tests cover standard OLE-object
  relationships with external and internal stored target modes, a resolved
  unsupported non-OLE relationship, a direct marker without optional `r:id`,
  duplicate markers, body/header stories, Transitional and Strict
  Word/relationship namespaces, orphaned relationship exclusion, unavailable
  relationship failure, same-count program/update/target-marker changes,
  relationship-ID renumbering stability, JSON/Markdown/SARIF redaction, and
  policy findings;
- direct WordprocessingML `w:objectLink` markup is separately inventoried
  only as a direct child of `w:object`, distinct from `w:objectEmbed`, legacy
  VML linked- and embedded-OLE markup, VML image data/shape links, DrawingML
  linked pictures, field inventories, and generic relationship totals. Tests cover
  standard external and internal OLE-object relationships, unsupported and
  missing relationships, exact `always` and `onCall` modes plus missing or
  unsupported modes, duplicate markers, body/header stories, Transitional and
  Strict Word/relationship namespaces, orphaned and unparented-marker
  exclusion, unavailable relationship failure, same-count field/target
  changes, relationship-ID renumbering stability, JSON/Markdown/SARIF
  redaction, and policy findings;
- direct WordprocessingML `w:control` markup is separately inventoried only
  as a direct child of `w:object` or `w:pict`, distinct from arbitrary
  unparented markers, `w:objectLink`, `w:objectEmbed`, legacy VML linked- and
  embedded-OLE markup, VML image data/shapes, ActiveX-binary relationships,
  field inventories, and generic control relationship/payload totals. Tests cover
  parent-position counts, standard internal and external control relationships,
  unsupported and missing relationships, duplicate markers, body/header
  stories, Transitional and Strict Word/relationship namespaces, orphaned and
  unparented-marker exclusion, unavailable relationship failure, same-count
  control-name/target rewrites, relationship-ID renumbering stability,
  JSON/Markdown/SARIF redaction, and policy findings;
- Word editable-range permission markup is separately inventoried across body,
  header, footer, note, comment, and glossary stories. Tests cover aggregate
  marker/pairing/individual-editor/predefined-group/table-column/custom-XML
  placement counts; same-count editor-identity rewrites; JSON/Markdown/SARIF
  redaction; policy findings; Transitional and Strict namespaces; unmatched
  boundaries; and malformed marker leaves, IDs, attributes, groups, columns,
  placement values, and duplicate per-story boundaries;
- direct false-valued hidden declarations and `w:specVanish` outside paragraph
  marks do not create false text-run findings;
- volatile Word `rsid` updates and relationship-ID renumbering that preserve the
  underlying target, embedded/import payload, document-property state,
  mail-merge state, data-binding storage association, content-control lock
  state, and external-document dependency semantics remain quiet;
- generated JSON, Markdown, and SARIF do not reproduce unique markers placed
  in visible/hidden text, reviewer metadata, comments, URLs, field instructions,
  custom XML, document-property names or values, mail-merge connection/query/
  source/recipient values, data-binding XPath/prefix/storage/payload values,
  external-source field paths, connections, queries, application/item names,
  modern-comment authors, providers, user IDs, paragraph/durable IDs, dates,
  thread state, reaction identities,
  document-task IDs, event times, users, titles, schedules, priorities,
  progress, comment anchors, Office web-extension IDs, stores, properties,
  bindings, content-control IDs, aliases, tags, titles, placeholder text,
  current values, markers, and pane values,
  sensitivity-label and tenant IDs, names, methods, dates, action IDs,
  extension payloads, legacy MIP custom attributes, and content-marking values,
  package-signature signer/certificate material, values, algorithms, reference
  URIs, signing times, comments, provider data, relationship IDs, and paths,
  Word protection hashes, salts, verifier values, cryptographic provider and
  algorithm fields, and Settings-part paths,
  Word document-variable names, values, Settings-part paths, `DOCVARIABLE`
  instructions, literal field arguments, and story paths,
  `HYPERLINK` instructions, destinations, internal locations, ScreenTips,
  frame targets, and story paths,
  direct `w:hyperlink` relationship targets, anchors, locations, tooltips,
  frame names, history values, display text, relationship IDs, and story paths,
  direct DrawingML linked-picture targets, relationship IDs, surrounding
  drawing markup, and story paths,
  direct VML `href` values, frame targets, titles, alternate text, shape IDs,
  and story paths,
  direct VML image targets, relationship IDs, raw source values, other
  image-data attributes, and story paths,
  direct VML linked-OLE sources, monikers, program, shape, and object IDs,
  relationship IDs and targets, update metadata, field codes, markup, and
  story paths,
  direct VML embedded-OLE program, shape, and object IDs, relationship IDs and
  targets, update metadata, field codes, markup, and story paths,
  direct WordprocessingML linked-object-property program and shape IDs,
  relationship IDs and targets, field codes, locking and update metadata,
  markup, and story paths,
  direct WordprocessingML embedded-control names and shape IDs, relationship
  IDs and targets, markup, and story paths,
  editable-range marker IDs, individual editor identities, exact table-column
  selectors, placement values, and story paths,
  external template/subdocument/frame targets and frame names, macros,
  embedded/control payloads, alternative-format imports, or opaque package
  parts;
- policy failures return a nonzero CI status and SARIF uses no source location;
- DTD/entity markup, unsafe ZIP member names, and `w:altChunk` markup without a
  matching internal import relationship, malformed recognized document-property
  roots, malformed recognized mail-merge relationship references, malformed
  recognized save-through-XSLT relationship/anchor state, malformed recognized
  data-binding custom-XML-properties relationships, and malformed
  recognized external-document dependency relationship/anchor state, malformed
  modern-comment metadata roots, external modern-comment metadata
  relationships, document-task roots/relationships, task-pane or
  web-extension roots, relationships, and direct task-pane references are
  rejected before reporting. Malformed sensitivity LabelInfo roots, required
  attributes, tenant-site IDs, unavailable/external/non-root classification
  relationships, and multiple LabelInfo parts are rejected before reporting.
  Malformed recognized package-signature XMLDSIG shape and origin/signature/
  certificate relationship topology are rejected before reporting. Malformed
  recognized Word protection element structure, attributes, edit values,
  booleans, duplicate leaves, and external Settings relationships are rejected
  before reporting. Malformed recognized Word document-variable containers,
  leaves, attributes, lengths, and duplicate containers are rejected before
  reporting. Malformed recognized Word editable-range leaves, IDs,
  attributes, groups, column selectors, placement values, and duplicate
  per-story boundaries are rejected before reporting.

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

For the document-variable-field boundary, the release check profiles
LibreOffice's public [`tdf150542.docx` regression
fixture](https://github.com/LibreOffice/core/blob/2553d132a7e95170568cb99920a7264fe5c8081d/sw/qa/extras/ooxmlexport/data/tdf150542.docx).
It contains one stored `DOCVARIABLE` reference, a standard `w:docVars`
container with three variables, and Word's common `\* MERGEFORMAT` field
switch. DocFence reports one literal same-scope association without emitting
the variable name or value. This is a compatibility smoke test, not a runtime
dependency or a statement about field rendering.

For the `HYPERLINK` field boundary, the release check profiles the public
[`doc.docx` attachment](https://github.com/jgm/pandoc/files/13645632/doc.docx)
from Pandoc issue [#9246](https://github.com/jgm/pandoc/issues/9246). It is a
real Word document with a complete complex `HYPERLINK \l` field pointing to an
in-document location. DocFence reports one literal internal-location-only
reference without emitting the bookmark or result text. This is a compatibility
smoke test, not a runtime dependency or a statement about field rendering.

For direct WordprocessingML hyperlink markup, the release check profiles
python-docx's public
[`par-hyperlinks.docx` fixture](https://github.com/python-openxml/python-docx/blob/e45454602b53e8e572b179ccf1c91093ec9f4ed7/features/steps/test_files/par-hyperlinks.docx).
It contains four conventional relationship-backed `w:hyperlink` elements with
external relationship targets. DocFence reports four direct elements and four
external relationship-backed elements without emitting the targets or display
text. This is a compatibility smoke test, not a runtime dependency or a
statement about link rendering.

For DrawingML hyperlink-action markup, the release check profiles LibreOffice's
public [`tdf78657_picture_hyperlink.docx` regression
fixture](https://github.com/LibreOffice/core/blob/100f43ab871bedda6b427645cbfc2c8083da98b5/sw/qa/extras/ooxmlexport/data/tdf78657_picture_hyperlink.docx).
Its commit identifies it as a Microsoft Word 2016 test file. The package stores
two direct `a:hlinkClick` markers around its image and one external hyperlink
relationship; DocFence reports two click markers and two external
relationship-backed marker references without emitting a target. This is a
compatibility smoke test, not a runtime dependency or a statement about link
rendering or action execution.

For DrawingML linked-picture markup, the release check profiles the public
[`abs/1.docx` package](https://github.com/pea-sys/shell-experiments/blob/91386d4de9e499a21bbb2e54743eb63a63727bfb/powershell/survey/abspath2relpath-docx/survey/abs/1.docx)
from pea-sys's open-source `abspath2relpath-docx` investigation. The pinned
package SHA-256 is
`8554932d5de0f1c81dbdcbf7b480c17f9008abca18c66a7f03b2e881cfe5147b`.
It stores two direct `a:blip/@r:link` markers backed by two external standard
image relationships. DocFence reports two direct markers and two
external-image classifications while keeping relationship IDs and paths out of
the output. This is a compatibility smoke test, not a runtime dependency or a
statement that any Word client will retrieve, update, or render the pictures.

For legacy VML external-image markup, release validation profiles the paired
public [`rel/1.docx` package](https://github.com/pea-sys/shell-experiments/blob/91386d4de9e499a21bbb2e54743eb63a63727bfb/powershell/survey/abspath2relpath-docx/survey/rel/1.docx)
from the same open-source `abspath2relpath-docx` investigation. The pinned
package SHA-256 is
`60708e292cd38cb9bee28886e91b2103b7d2ee43e963fa5e1cac4eccaaa71ed6`.
It stores two direct `v:imagedata/@r:id` markers backed by two external standard
image relationships. DocFence reports two direct markers and two external-image
classifications while keeping relationship IDs and paths out of the output.
This is a compatibility smoke test, not a runtime dependency or a statement
that any Word client will retrieve, update, or render the images.

For legacy VML image-data hyperlink markup, release validation combines the
controlled Word-story package constructed in the public regression test with a
public real Word XML fragment from a
[Stack Overflow question](https://stackoverflow.com/questions/52124509/read-images-from-docx-file-with-python-docx).
That fragment contains the exact `v:imagedata r:id="…" r:href="…"` form in
legacy VML. It is marker evidence rather than a downloadable fixture suitable
for an end-to-end package smoke test. The Open XML SDK documents `r:href` as an
explicit relationship to a hyperlink target, while a separate public
[Aspose report](https://forum.aspose.com/t/corrupted-targetmode-attribute-value-in-relationship-tag/22502)
shows the marker paired with an external standard image relationship. The
release therefore checks both recognized hyperlink modes and a preserved
unsupported relationship class; it does not claim that a Word client will
render, honor, or follow the target.

For legacy VML linked-OLE markup, release validation combines the controlled
Word-story package constructed in the public regression test with two public
real Word XML fragments. One [documents an Excel-linked object](https://stackoverflow.com/questions/60565712/insert-ole-object-into-ms-word-document-and-keep-the-underlying-format-wmf-intac)
stored as `o:OLEObject Type="Link"` with `UpdateMode="Always"`; another
[documents a linked OLE shape](https://forum.aspose.com/t/cannot-found-includepicture-field-type-via-doc-range-fields/277035)
with `UpdateMode="OnCall"`, plus private-looking program and field-code
metadata. The fragments are stored-marker evidence, not downloadable package
fixtures, so the release verifies both modes and relationship classifications
with a controlled package and separately confirms the real serialized form.
It retains all source, program, and field data privately and makes no claim
that Word will retrieve, update, activate, or honor the object.

For legacy VML embedded-OLE markup, release validation combines the controlled
Word-story package in the public regression test with public package fixtures
from two independently maintained OOXML test corpora. The Open XML SDK's
[`ole.docx` fixture at `cd2b359`](https://github.com/dotnet/Open-XML-SDK/blob/cd2b359ef824737edb93f1c6157c19551aae1e52/test/DocumentFormat.OpenXml.Tests.Assets/assets/TestDataStorage/v2FxTestFiles/wordprocessing/ole/ole.docx)
contains one direct `o:OLEObject Type="Embed"` marker with an internal standard
OLE-object relationship and payload. Five docx4j fixtures at
[`74ea743`](https://github.com/plutext/docx4j/tree/74ea74323a33d92769fdbd3e6d5fe730bbfd8ffb/docx4j-core-tests/src/test/resources/OLE)
each contain one direct marker in a `w:object` with the same stored internal
standard relationship class. A pinned scan of the Open XML SDK corpus found
direct markers in 22 of 503 readable package candidates (two of 505 candidates
produced read errors); the docx4j scan found them in five of 141 readable Word
packages. These are public test assets, not a prevalence survey or a claim of
client runtime behavior. The release profiles raw public fixtures without
emitting their program IDs, targets, payloads, or paths.

For WordprocessingML linked-object-property markup, release validation uses the
controlled Word-story package in the public regression test. It exercises direct
`w:object/w:objectLink` scoping, duplicate body/header markers, standard
internal and external OLE-object relationships, unsupported and missing
relationship/update-mode state, Strict encodings, unavailable relationships,
same-count field/target rewrites, and relationship-ID renumbering. The Open XML
SDK [documents the `w:objectLink`
element](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.objectlink?view=openxml-3.0.1)
as an Office 2007+ leaf under `w:object`; the ISO-schema
[`CT_ObjectLink` source](https://github.com/dolanmiu/docx/blob/309b972e107b8cc12c65027d5161b7045e404b6c/ooxml-schemas/ISO-IEC29500-4_2016/wml.xsd#L1188-L1228)
requires `r:id` and `updateMode` and enumerates `always` and `onCall`.

Public package coverage is intentionally reported as limited: among 505
nominal package files in the [Open XML SDK test-asset corpus at
`cd2b359`](https://github.com/dotnet/Open-XML-SDK/tree/cd2b359ef824737edb93f1c6157c19551aae1e52/test/DocumentFormat.OpenXml.Tests.Assets)
there were no direct `w:objectLink` markers in the 503 readable packages (two
files produced read errors). A separate scan of all 141 Word packages in
[docx4j's `VERSION_17_0_3`
tree](https://github.com/plutext/docx4j/tree/VERSION_17_0_3) likewise found no
direct marker, including the corpus's five legacy OLE samples. That absence is
not a claim of nonexistence or Word behavior. It means this
release is a narrow, standards-shaped stored-XML compatibility boundary, not an
assertion of widespread contemporary Word authoring behavior or runtime
interoperability.

For WordprocessingML embedded-control anchors, release validation uses the
controlled Word-story package in the public regression test. It exercises the
two direct parent positions, standard internal and external control
relationships, unsupported and missing-ID state, duplicates, body/header
stories, Strict encodings, orphan and unparented-marker exclusion, unavailable
relationships, same-count name/target rewrites, and relationship-ID renumbering.
The Open XML SDK [documents `w:control`](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.control?view=openxml-3.0.1)
as an Office 2007+ element with `w:object` and `w:pict` parent forms; the
ISO-schema [`CT_Control`, `CT_Object`, and `CT_Picture` source](https://github.com/dolanmiu/docx/blob/309b972e107b8cc12c65027d5161b7045e404b6c/ooxml-schemas/ISO-IEC29500-4_2016/wml.xsd#L1146-L1199)
shows the optional `r:id` and direct parent positions. The ECMA-376
[Embedded Control Persistence Part contract](https://c-rex.net/samples/ooxml/e1/Part1/OOXML_P1_Fundamentals_Embedded_topic_ID0EQDBO.html)
requires the persistence relationship target to be internal.

Public package coverage is deliberately modest but includes a real positive
fixture. The 503 readable packages from the 505 nominal Open XML SDK test-asset
corpus at `cd2b359` contain no direct `w:control` marker or `word/activeX/`
member. A scan of 141 Word packages at docx4j commit
[`74ea743`](https://github.com/plutext/docx4j/tree/74ea74323a33d92769fdbd3e6d5fe730bbfd8ffb)
found one positive package: [`LegacyForms.docx`](https://github.com/plutext/docx4j/blob/74ea74323a33d92769fdbd3e6d5fe730bbfd8ffb/docx4j-core-tests/src/test/resources/LegacyForms.docx)
has five direct `w:object/w:control` anchors, all with internal standard
control relationships, plus fifteen `word/activeX/` ZIP members. Profiling the
public raw fixture reports the five anchors while the existing generic inventory
reports ten control relationships, demonstrating why the direct-anchor boundary
is not interchangeable with package-level totals. This is stored-package
evidence, not a claim that every Word client will load or enable a control.

For legacy VML shape-link markup, the release check uses a controlled,
standards-shaped Word-story package that places direct `href` attributes on all
supported VML geometry kinds and keeps the package's relationship, field,
WordprocessingML-link, and DrawingML-action inventories empty. The package is
constructed inside the public regression test so its exact XML is reviewable
with the parser contract. Public Word-package fixtures carrying direct VML
`href` markup are scarce: a scan of the Open XML SDK Word test assets at
`cd2b359ef824737edb93f1c6157c19551aae1e52` found legacy VML shapes but no
direct VML `href`. The release therefore treats the case as a stored-XML
compatibility boundary validated against the documented VML attribute, not as
a claim of widespread contemporary Word authoring behavior.

For the package-signature boundary, the release check profiles the signed DOCX
fixture from the public [USENIX 2023 OOXML Signature Security
artifacts](https://github.com/RUB-NDS/OOXML_Signature_Security). Its unmodified
content-injection baseline satisfies the strict static declaration-coverage
gate. The published content-injection, universal-signature-forgery,
duplicate-document, and evil-type attacker variants are rejected for static
declaration gaps. This is a structural compatibility smoke test over a bounded
scope, not a claim that signatures are valid, trusted, exhaustive, or that a
client will handle any sample in a particular way; the fixtures are not runtime
dependencies.

For the Word-protection boundary, the release check profiles two independent
Wordprocessing protection assets from Microsoft's open-source
[Open XML SDK](https://github.com/dotnet/Open-XML-SDK/tree/main/test/DocumentFormat.OpenXml.Tests.Assets/assets/TestDataStorage/v2FxTestFiles/wordprocessing/protected).
The write-protected asset reports one write-protection element and one
read-only recommendation; the partially protected asset reports one
document-protection element with explicit enforcement and formatting
restriction. Neither carries counted password-material attributes. These are
compatibility smoke tests, not runtime dependencies.

## What this does not validate

The suite does not claim layout equivalence, Word calculation behavior,
style-inherited effective hiddenness, macro safety, embedded-payload safety,
malware detection, alternative-format rendering/import behavior, or
compatibility with every vendor extension. It also does not decide whether a
document-property value is personal, confidential, intended, or safe to share,
whether an OPC package thumbnail reflects the document or will be displayed by
any client,
or whether stored mail-merge state will run, retrieve data, or select a
recipient. It does not assert that a recognized external-source field will
update, reach a target, use a particular argument as a source, or be accepted
by Word; it records the bounded stored field-family evidence only. It does not
assert that a `HYPERLINK` field is valid, external, reachable, safe, rendered,
or followed by Word; it records bounded private-digest evidence only. It
likewise does not assert that direct `w:hyperlink` markup is valid,
reachable, safe, rendered, or followed by Word; it records bounded stored
relationship/markup evidence only. It likewise does not assert that a direct
DrawingML hyperlink-action marker is valid, selected by Word, reachable, safe,
rendered, followed, or executed; it records bounded stored marker evidence
only. It likewise does not assert that a direct DrawingML linked-picture marker
is valid, selected by Word, reachable, safe, rendered, updated, or retrieved;
it records bounded stored marker/relationship evidence only. It likewise does
not assert that a direct VML shape `href` is valid,
inherited, rendered, reachable, safe, followed, or honored by a client; it
records bounded direct-markup evidence only. It does not calculate effective
VML inheritance or inspect arbitrary VML elements. It likewise does not assert
that a direct VML `imagedata/@r:id` marker is selected, valid, retrievable,
rendered, updated, reachable, safe, or honored by a client; it records bounded
stored relationship evidence only and does not use raw VML `src` as an
alternate marker. It likewise does not assert that a direct VML
`o:OLEObject Type="Link"` marker is valid, selected, retrievable, source-backed
at review time, updated, activated, rendered, or honored by Word; it records
bounded stored marker/relationship evidence only. It also does not assert that
an `o:OLEObject Type="Embed"` marker is valid, selected, retrievable,
payload-backed, opened, activated, rendered, or honored by Word; it records
bounded stored marker/relationship evidence only. It also does not assert that
a direct WordprocessingML `w:objectLink` marker is schema-valid,
selected, source-backed at review time, updated, activated, rendered, or
honored by Word; it records bounded stored marker/relationship evidence only.
It also does not assert that a direct WordprocessingML `w:control` marker is
schema-valid, selected, installed, loaded, enabled, safe, activated, rendered,
or honored by Word; it records bounded stored marker/relationship evidence
only.
The style/default layer is a stored declaration inventory, not a renderer. It
does not decrypt an IRM-protected file, resolve a sensitivity-label policy, calculate
permissions, or predict label markings. It also does not verify a package
digital signature or digest, validate certificates or trust chains, check
revocation or timestamps, establish signer identity, calculate cryptographic
or client-effective coverage, or make an Office trust decision. It resolves
only a bounded static package-manifest declaration subset. It does not validate
Word password-verifier construction or strength, recover or test a password,
bypass a Word restriction, infer effective enforcement, or treat Word protection as
encryption/security. It does not authenticate a stored editable-range editor,
resolve a group, calculate an editable region, or infer effective range
authorization. It does not evaluate a `DOCVARIABLE` field, run a macro, resolve
a document-variable name or template, or infer whether a stored variable is
used or visible. Exact-literal same-scope association is stored-package evidence
only, not field evaluation. Those limits are explicit in the 0.49 contract; see
[threat model](threat-model.md).
