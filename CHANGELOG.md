# Changelog

All notable changes are documented here.

## 0.56.0 — 2026-08-06

- Hardened recognized OPC XML signature parts to reject every XMLDSIG
  `ds:Transform` whose `Algorithm` is not OPC's Relationship Transform URI or
  one of OPC's two permitted XML Canonicalization URIs, as required by OPC
  §10.5.8.1. A missing algorithm is rejected too.
- This global stored-syntax boundary applies before the bounded declaration
  coverage audit, including an additional same-document `SignedInfo` reference.
  It does not resolve or execute a transform, recompute a digest, verify XMLDSIG,
  or make a trust decision.
- Added regressions for an unsupported transform in both a non-coverage
  `SignedInfo` reference and otherwise fully declared coverage paths, plus a
  positive non-coverage canonicalization case. Corpus and release-profile
  evidence is recorded with the release.

## 0.55.0 — 2026-08-06

- Hardened recognized OPC XML signature parts to reject every element or
  attribute in the OOXML Markup Compatibility namespace, as required by OPC
  §10.5.2. MCE markup can no longer remain inside an otherwise fully declared
  package signature.
- This is a structural signature boundary distinct from DocFence's Word-part
  MCE inventory: it does not validate MCE conformance, select a branch,
  preprocess a package, or assess an Office client's behavior.
- Added fully declared synthetic regressions for both an MCE attribute and an
  MCE element in a package XML signature. Corpus and release-profile evidence
  is recorded with the release.

## 0.54.0 — 2026-08-06

- Hardened recognized OPC XML signature parts to reject XMLDSIG
  `ds:XPath` elements anywhere in the signature, rather than only in transforms
  used by the bounded declaration-coverage chain. An XPath-bearing
  application-object or other non-coverage reference can no longer leave the
  package recognized as an OPC XML signature.
- This is a stored-markup boundary for OPC's global XPath prohibition. DocFence
  does not parse or execute an XPath expression, evaluate a transform, verify
  XMLDSIG, or make a trust decision.
- Added a regression for XPath on an additional same-document
  `SignedInfo` reference. The public USENIX OOXML Signature Security corpus
  contained no XMLDSIG XPath elements in its 22 XML signature parts, so all 29
  DOCX profiles remain unchanged from 0.53.

## 0.53.0 — 2026-08-06

- Hardened recognized OPC XML signature parts to reject the one
  `DigestMethod/@Algorithm` URI OPC expressly forbids—MD5—anywhere in the
  XML signature, not only along the bounded declaration-coverage chain. An MD5
  application-object or other non-coverage reference can no longer leave the
  package recognized as an OPC XML signature.
- This remains an exact stored-syntax rule. DocFence does not decode or
  recompute a digest, execute a transform, verify XMLDSIG, or turn SHA-1
  guidance into a broader cryptographic policy or trust decision.
- Added a regression for MD5 on an additional same-document `SignedInfo`
  reference. The public USENIX OOXML Signature Security corpus still used 243
  SHA-256 digest methods and no MD5 declarations, so all 29 DOCX profiles
  remain unchanged from 0.52.

## 0.52.0 — 2026-08-06

- Hardened recognized OPC XML signature parts to require every direct
  `SignedInfo/Reference` to carry explicit XMLDSIG same-document syntax: the
  empty URI or a nonempty local `#` fragment. An omitted URI leaves object
  identity to application context; relative package-part and absolute URIs
  point outside the Signature document. All three now fail the recognized
  XML-signature shape closed.
- This is a narrow URI-location boundary: DocFence does not dereference a URI,
  resolve a fragment or XPointer, execute a transform, recompute a digest,
  verify XMLDSIG, or make a trust decision.
- Added positive empty-URI and local-fragment coverage plus missing,
  package-relative, and absolute URI regressions. Compared with 0.51, the
  public USENIX OOXML Signature Security corpus changes only one attacker DOCX:
  its eight ODF-style package-file references in `SignedInfo` now fail
  structurally; the other 28 DOCX profiles are unchanged.

## 0.51.0 — 2026-08-06

- Hardened recognized OPC XML signature parts to require their one direct
  `SignedInfo/CanonicalizationMethod/@Algorithm` to be exactly one of OPC's two
  permitted XML Canonicalization URIs: C14N or C14N with comments. A missing or
  other URI now fails the recognized XML signature shape closed instead of
  being inventoried as though it were an OPC signature.
- This is a narrow stored-syntax boundary: DocFence does not execute
  canonicalization, validate a `SignatureMethod`, recompute a digest, verify a
  signature, or make a trust decision.
- Added regression coverage for both permitted methods and missing or
  unsupported algorithms. The public USENIX OOXML Signature Security corpus
  had 22 XML signature parts: all 21 parseable declarations used standard C14N;
  the remaining attacker part was empty and already failed XML parsing.

## 0.50.0 — 2026-08-06

- Hardened the bounded OPC package-signature declaration audit to reject the
  one `DigestMethod/@Algorithm` URI OPC expressly forbids:
  `http://www.w3.org/2001/04/xmldsig-more#md5`.
- An MD5 package-object binding now leaves declaration coverage unavailable;
  an MD5 manifest part or relationship reference is aggregate unsupported and
  cannot lend part or selector coverage. The parser does not decode or
  recompute a digest. It deliberately retains legacy SHA-1 and other values as
  structural syntax because OPC only discourages SHA-1 and permits other
  algorithms.
- Added binding, part, and relationship MD5 regressions plus a SHA-1
  compatibility regression. The public USENIX OOXML Signature Security corpus
  used 243 SHA-256 `DigestMethod` declarations and no MD5 declarations during
  the compatibility scan.

## 0.49.0 — 2026-08-04

- Hardened the bounded OPC package-signature declaration audit to reject a
  `ds:XPath` element anywhere inside a transform used by its binding or
  manifest-reference chain. OPC forbids XPath filtering even though XMLDSIG's
  generic `Transform` schema permits an XPath parameter.
- A package-object binding with that parameter now leaves declaration coverage
  unavailable. A manifest part or relationship reference with it is aggregate
  unsupported and grants neither part nor partial relationship-selector
  coverage. This check does not parse or execute XPath expressions.
- Added regressions for XPath-bearing binding, part, and relationship
  transforms. The public USENIX OOXML Signature Security corpus had no XPath
  element in its 22 XML signature parts during the compatibility scan.

## 0.48.0 — 2026-08-04

- Hardened the one `SignedInfo` Reference that binds
  `idPackageObject` to the static declaration audit. It now requires the
  XMLDSIG direct child order of optional `ds:Transforms`, one `ds:DigestMethod`
  with a nonblank `Algorithm`, and one direct, attribute-free, child-free,
  nonempty `ds:DigestValue`, without non-whitespace direct text.
- The binding may omit transforms or use a direct nonempty list of OPC's two
  XML Canonicalization algorithms. Relationship, unknown, empty, or duplicate
  transform lists—and malformed binding digest shape—leave coverage
  unavailable rather than lending a malformed Object reference authority.
- Added regression coverage for both canonicalization forms and all rejected
  binding transform/digest shapes. This remains a bounded structural parser:
  DocFence does not decode or recompute a digest, execute a transform, verify
  XMLDSIG, validate certificates, establish trust, or predict Office client
  behavior.

## 0.47.0 — 2026-08-04

- Hardened static package-signature declaration coverage with OPC's required
  timestamp-property shape. A qualifying package object now requires one direct
  `ds:SignatureProperty Id="idSignatureTime"` with no attributes other than
  its required `Id` and `Target`, an empty or root-signature-fragment
  target, and exactly one attribute-free `opc:SignatureTime` child. That
  timestamp must contain
  only attribute-free `opc:Format` then `opc:Value` children; its declared
  precision and value must match one of OPC's six schema forms.
- Missing, duplicate, misidentified, mis-targeted, attribute-bearing,
  malformed, or format/value-mismatched timestamp declarations now leave
  coverage unavailable rather than lending a malformed package object
  authority. The check remains a bounded structural parser: DocFence does not
  verify timestamp accuracy or authority, XMLDSIG digests, signature values,
  certificates, trust, or Office client behavior.
- Added regression coverage for every accepted precision and empty target, and
  for each rejected timestamp topology and value shape. The current official
  ECMA-376 Part 2 schema bundle and the public USENIX OOXML Signature Security
  corpus were used for compatibility checks.

## 0.46.0 — 2026-08-04

- Hardened the static package-signature declaration boundary to OPC's exact
  package-specific topology. A recognized signature now receives declaration
  coverage only from one direct `ds:Object Id="idPackageObject"` with no other
  attributes and exactly direct `ds:Manifest` then `ds:SignatureProperties`
  children, plus exactly one direct `SignedInfo` reference to that object.
  Missing or nonstandard IDs, missing or extra package-object children,
  duplicate package objects, and duplicate bindings leave coverage unavailable
  rather than combining arbitrary manifests.
- Added regression coverage for every rejected topology above and updated the
  deterministic DCAB package-signature fixtures to use the OPC identifier.
  This remains a bounded structural audit: DocFence does not parse or validate
  the binding reference's digest or transforms, recompute manifest digests,
  validate signatures or certificates, establish trust, or predict Office
  client behavior.

## 0.45.0 — 2026-08-04

- Hardened the bounded XMLDSIG `Reference` shape required before a bound
  package-manifest declaration receives coverage. References now require the
  optional direct `ds:Transforms` list followed by one `ds:DigestMethod` with a
  nonblank `Algorithm` and one direct, attribute-free, child-free, nonempty
  `ds:DigestValue`, with no non-whitespace direct text. Missing, misordered,
  malformed, or extra direct children remain aggregate unsupported references
  rather than being credited.
- Added regression coverage for missing and misordered digest children, a
  missing digest-method algorithm, an unexpected digest-value attribute or
  nested element, unexpected reference text, and an extra direct child,
  including policy failures and private-report redaction. This remains a
  bounded structural audit: DocFence does not parse, decode, or recompute
  digest values, execute transforms, validate a signature or certificate,
  establish trust, or predict Office client behavior.

## 0.44.0 — 2026-08-04

- Hardened direct package-part manifest references in the static
  package-signature declaration audit. A non-relationships part now receives
  coverage only with no transform list or one direct nonempty list containing
  only OPC's supported XML Canonicalization algorithms. Relationship, unknown,
  empty, and duplicate transform lists remain aggregate unsupported references.
- Added regression coverage for permitted C14N-only part references and for
  unsupported empty, duplicate, relationship, and unknown lists, including
  policy failures and private-report redaction. DocFence still does not execute
  XMLDSIG transforms, recompute digests, validate a signature or certificate,
  establish trust, or predict Office client behavior.

## 0.43.0 — 2026-08-04

- Hardened static package-signature declaration coverage for OPC's
  one-relationship-transform-per-relationships-part constraint. DocFence now
  counts relationship transforms across all bound manifests in one XML
  signature. If more than one targets the same relationships part, none of
  those selectors is credited; each remains an aggregate unsupported reference.
- Added regression coverage for duplicate relationship-transform declarations,
  coverage removal, aggregate-only unsupported evidence, policy failures, and
  private-report redaction. This remains a bounded static audit, not XMLDSIG
  transform or digest evaluation, certificate validation, trust establishment,
  or Office client prediction.

## 0.42.0 — 2026-08-04

- Hardened the static package-signature declaration audit to require the OPC
  relationship-transform sequence: one direct `ds:Transforms` list containing
  only the supported relationship and XML Canonicalization algorithms, exactly
  one relationship transform, and a canonicalization transform immediately
  after it. Both OPC XML Canonicalization forms are recognized.
- Missing, misordered, or unsupported relationship-transform sequences now
  remain aggregate-only unsupported references rather than being credited with
  declaration coverage. This is still a bounded static parser: DocFence does
  not recompute transforms or digests, validate XMLDSIG values or certificates,
  establish trust, or predict Office client behavior.
- Added regression coverage for ordinary and comment-preserving XML
  canonicalization, missing and misordered canonicalization, unsupported
  trailing transforms, policy failures, and private-report redaction.

## 0.41.0 — 2026-08-04

- Corrected static package-signature declaration coverage for the standard OPC
  relationship-type selector: `RelationshipsGroupReference/@SourceType` now
  selects matching relationship types, is included in the structural signature
  inventory, and remains aggregate-only in public output.
- A malformed `RelationshipReference/@SourceType` lookalike is now treated as
  unsupported rather than being credited with declaration coverage. This keeps
  the bounded parser exact and fail-closed without changing policy IDs or
  making cryptographic claims.
- Added regression coverage for standard type selectors, duplicate selected
  relationship types, malformed lookalikes, inventory counts, privacy
  redaction, and same-count semantic changes. DocFence still does not recompute
  XMLDSIG transforms or digests, verify signature values or certificates,
  establish trust, or predict Office client behavior.

## 0.40.0 — 2026-08-04

- Added a first-class, privacy-safe static declared OPC package-signature
  coverage inventory. It reports only aggregate counts for signatures with or
  without a bound package manifest, covered/uncovered bounded Word parts and
  relationships, and unresolved/unsupported references; object identifiers,
  manifest URIs, relationship selectors, paths, and digest material remain
  private.
- Added DFP092 to require a recognized XML signature and complete static
  declaration coverage in the bounded Word scope, and DFP093 to protect an
  approved declaration-coverage baseline. These are opt-in structural gates,
  not substitutes for `require_no_package_digital_signatures`.
- Added regression coverage for direct manifest binding, case-sensitive part
  content-type and relationship declaration resolution, same-count selector
  reassignment, privacy redaction, Markdown/JSON/SARIF output, and both policy
  modes. A public USENIX OOXML
  Signature Security corpus smoke accepts the unmodified content-injection
  baseline and rejects the selected published attacker variants on declaration
  gaps. DocFence does not recompute XMLDSIG digests or canonicalization, verify
  signature values or certificates, establish trust, or predict Office client
  behavior.

## 0.39.0 — 2026-08-04

- Added a first-class, privacy-safe inventory for direct Word content-control
  `w:sdtPr/w:lock` declarations across every supported story. Public reports
  contain only aggregate counts for no direct declaration and the four exact
  schema states: `unlocked`, `sdtLocked`, `contentLocked`, and
  `sdtContentLocked`. Control IDs, aliases, tags, titles, placeholder text,
  current values, story paths, and private fingerprints remain local.
- Added DFP090 to require an explicit non-`unlocked` direct lock declaration
  for every discovered content control, and DFP091 to protect an approved
  lock-declaration baseline. A missing leaf remains distinct from explicit
  `unlocked`, because the OOXML contract gives omitted locks type-specific
  behavior for group controls.
- Added regression coverage for every schema state, absent declarations,
  same-count state reassignment, Strict OOXML, out-of-scope lookalikes,
  malformed leaves/properties, redaction, Markdown/JSON/SARIF output, and both
  policy modes. DocFence does not open Word, read a control value, mutate a
  control, apply document protection, evaluate a binding, or predict client
  enforcement.

## 0.38.0 — 2026-08-04

- Added a first-class, privacy-safe inventory for direct stored Word Settings
  `w:savePreviewPicture` `CT_OnOff` declarations. Public reports contain only
  enabled and explicitly disabled setting counts; Settings-part paths and
  private fingerprints remain local. Canonical on/off state prevents lexical
  equivalent declarations from adding review noise while an enabled/disabled
  transition remains review-visible.
- Added DFP088 to reject a candidate that requests preview-thumbnail generation
  on save and DFP089 to protect an approved stored configuration baseline.
  This setting is separately inventoried from an already stored OPC package
  thumbnail: it does not prove a thumbnail exists or that a host will create
  one.
- Added regression coverage for absent, implicit-enabled, explicit, disabled,
  and Strict leaves; malformed direct state; private output; Markdown, JSON,
  and SARIF output; and policy findings. DocFence does not decode or render an
  image, open Word, save a document, create a thumbnail, or predict Office
  client behavior.

## 0.37.0 — 2026-08-04

- Added a first-class, privacy-safe inventory for direct stored Word Settings
  `w:saveFormsData` `CT_OnOff` declarations. Public reports contain only
  enabled and explicitly disabled setting counts; Settings-part paths and
  private fingerprints remain local. Canonical on/off state prevents lexical
  equivalent declarations from adding review noise while an enabled/disabled
  transition remains review-visible.
- Added DFP086 to reject a candidate that requests form-data-only saving and
  DFP087 to protect an approved stored configuration baseline.
- Added regression coverage for absent, implicit-enabled, explicit,
  disabled, and Strict leaves; malformed direct state; private output;
  Markdown, JSON, and SARIF output; and policy findings. DocFence does not
  find or evaluate form fields, read field values, open Word, save a document,
  emit a delimited record, determine a delimiter, or predict Office client
  behavior.

## 0.36.0 — 2026-08-04

- Added a first-class, privacy-safe inventory for direct stored DrawingML
  nonvisual hidden declarations in supported Word stories. It recognizes
  standard DrawingML main, picture, and WordprocessingDrawing forms in
  Transitional and Strict packages plus Word 2010 w14, wpg, and wps forms.
  Public reports contain only aggregate declaration/story and
  hidden/explicitly-shown/invalid-value counts; object names, descriptions,
  titles, IDs, raw values, story paths, and digests remain private.
- Added DFP084 to reject candidate packages containing recognized hidden or
  malformed nonvisual hidden declarations and DFP085 to protect an approved
  visibility baseline. Equivalent XML Boolean spellings are canonicalized;
  same-count hidden-to-shown swaps remain review-visible.
- Added regression coverage for supported namespaces and forms, duplicate
  markers, MCE-branch scanning, invalid values, semantic same-count changes,
  privacy redaction, policy findings, Markdown, JSON, and SARIF output.
  DocFence does not validate full DrawingML conformance, resolve object
  identity, calculate effective visibility, choose a compatibility branch,
  lay out or render a drawing, or predict Office client behavior.

## 0.35.0 — 2026-08-04

- Added a first-class, privacy-safe inventory for stored OOXML Markup
  Compatibility and Extensibility (MCE) markup in non-relationship Word XML
  members that use the standard MCE namespace. Public reports contain only
  aggregate part, branch, and compatibility-rule token counts; branch bodies,
  feature-prefix and qualified-name values, member paths, and digests remain
  private. Equal-count branch or rule rewrites stay review-visible.
- Added `DFP082` to require a candidate with no recognized stored MCE markup
  and `DFP083` to protect an approved MCE baseline.
- Added regression coverage for `AlternateContent`/`Choice`/`Fallback` and MCE
  rule attributes, same-count `Choice/@Requires` rewrites, redacted JSON,
  Markdown, and SARIF output, policy findings, and a Strict Word package smoke
  profile. DocFence does not validate MCE conformance, resolve feature
  prefixes, choose a branch, preprocess or save a package, or predict client
  behavior.

## 0.34.0 — 2026-08-04

- Added a first-class, privacy-safe inventory for relationship-bound OPC
  package thumbnail images. It recognizes only the exact standard Transitional
  or Strict thumbnail relationship from the package or a stored package part,
  validates internal image-target topology, and does not infer thumbnails from
  a filename alone.
- Added `DFP080` to require a candidate with no recognized package thumbnail
  images and `DFP081` to protect an approved thumbnail baseline. Image bytes,
  relationship sources and targets, content types, paths, and digests remain
  private; same-count image changes stay review-visible.
- Added regression coverage for package and part relationship sources, Strict
  relationship forms, unreferenced lookalikes, malformed topology, private
  output, policy findings, and SARIF output. DocFence does not decode, render,
  classify, or predict client display of an image.

## 0.33.0 — 2026-08-04

- Added `DFP079`, an opt-in candidate-state gate for conventional custom XML
  package data. It fails a handoff when one or more non-relationship members
  remain below `customXml/`, including unbound data and associated properties
  parts.
- The finding reports only the existing aggregate part count. XML values,
  namespaces, member names, relationship targets, and fingerprints remain
  private; the gate does not expose values, classify, remove, or rewrite custom
  XML.

## 0.32.0 — 2026-08-04

- Added a first-class, privacy-safe inventory for direct Word Settings
  `w:removePersonalInformation` `CT_OnOff` declarations. The inventory records
  only a stored request for a capable host to remove personal information on a
  later save; it does not inspect, identify, redact, or rewrite document
  properties, comments, revisions, or other package material.
- Added DFP077 to require that a candidate store an enabled
  personal-information-removal-on-save request and DFP078 to protect an
  approved stored baseline. The positive request does not prove that the
  package is currently free of personal information or that a client will act
  on a future save.
- Reports expose enabled and explicitly disabled setting counts only.
  Settings-part paths and fingerprints remain private. Canonical state avoids
  noise for equivalent enabled spellings while an enabled/disabled transition
  remains review-visible.
- Added regression coverage for conventional and Strict leaves, implicit and
  explicit forms, malformed shape/value rejection, duplicate state, privacy
  redaction, policy findings, and SARIF output.

## 0.31.0 — 2026-08-03

- Added a first-class, privacy-safe inventory for direct Word Settings
  w:linkStyles CT_OnOff declarations. The inventory records a stored request
  to automatically update document styles from an attached template on open
  without resolving a template relationship, opening Word, loading a template,
  propagating styles, or making a client-behavior claim.
- Added DFP075 to require that a candidate not enable automatic
  template-style updates on open and DFP076 to protect an approved stored
  baseline.
- Reports expose enabled and explicitly disabled setting counts only.
  Settings-part paths and fingerprints remain private. Canonical state avoids
  noise for equivalent enabled spellings while an enabled/disabled transition
  remains review-visible.
- Added regression coverage for conventional and Strict leaves, implicit and
  explicit forms, malformed shape/value rejection, duplicate state, privacy
  redaction, policy findings, and SARIF output.

## 0.30.0 — 2026-08-03

- Added a first-class, privacy-safe inventory for direct Word Settings
  `w:updateFields` `CT_OnOff` declarations. The inventory records the stored
  automatic-field-recalculation-on-open request without parsing or evaluating a
  field, opening Word, accessing a source, or making a client-behavior claim.
- Added `DFP073` to require that a candidate not request automatic field
  recalculation on open and `DFP074` to protect an approved stored baseline.
- Reports expose enabled and explicitly disabled setting counts only.
  Settings-part paths and fingerprints remain private. Canonical state avoids
  noise for equivalent enabled spellings while an enabled/disabled transition
  remains review-visible.
- Added regression coverage for conventional and Strict leaves, implicit and
  explicit forms, malformed shape/value rejection, duplicate state, privacy
  redaction, policy findings, and SARIF output.

## 0.29.0 — 2026-08-02

- Added a first-class, privacy-safe inventory for Word's direct
  `w:attachedSchema` declarations. Each declaration names a custom XML schema
  target namespace that a host may associate when loading a document if that
  schema is available; DocFence records the stored declaration without locating
  or using a schema.
- Added `DFP071` to require a handoff with no attached custom XML schema
  declarations and `DFP072` to protect an approved declaration baseline.
- Reports expose only an aggregate declaration count. Namespace identifiers,
  Settings-part paths, and fingerprints remain private. Same-count namespace
  rewrites remain review-visible.
- Added regression coverage for conventional and Strict OOXML, multiple
  declarations, same-count namespace rewrites, malformed leaves, privacy
  redaction, and policy/SARIF output.

## 0.28.0 — 2026-08-02

- Added a first-class, privacy-safe inventory for Word's optional
  `w:useXSLTWhenSaving` / `w:saveThroughXslt` configuration. It detects the
  direct Settings leaves, standard external XSLT `transform` relationships, and
  application-defined local `w:solutionID` anchors without resolving or
  applying a transform.
- Added `DFP069` to require a handoff with no stored XSLT-on-single-XML-save
  configuration and `DFP070` to protect an approved configuration baseline.
- Reports expose aggregate enabled/disabled-setting, anchor, relationship, and
  solution-identifier counts only. Transform targets, solution identifiers,
  relationship IDs, Settings-part paths, and fingerprints remain private.
  Same-count configuration or target rewrites stay review-visible while a
  relationship-ID renumbering with unchanged semantics remains quiet.
- Added regression coverage for conventional and Strict OOXML, enabled and
  disabled settings, relationship-backed and local-only anchors, residual
  relationships, target rewrites, malformed topology, privacy redaction, and
  policy/SARIF output.

## 0.27.0 — 2026-08-01

- Added a first-class, privacy-safe inventory for direct legacy Office VML
  `o:OLEObject Type="Embed"` markers in supported Word stories. It is separate
  from the existing `Type="Link"` inventory, WordprocessingML
  `w:objectEmbed`/`w:objectLink` markup, VML image and shape links, fields, and
  broad embedded OLE/package/control relationship or payload totals. The VML
  contract permits several parent forms, so the boundary follows the stored
  Office VML marker rather than assigning it a rendered-object position.
- Added `DFP067` to require a handoff with no direct VML embedded-OLE marker
  and `DFP068` to protect an approved embedded-OLE marker baseline.
- Reports expose only aggregate marker/story and backing-relationship classes:
  standard OLE-object relationships with stored external or internal target
  mode, unsupported relationships, and markers without optional `r:id`.
  Program, shape, and object identifiers, relationship IDs and targets, update
  metadata, field codes, VML markup, story paths, and fingerprints remain
  private. `UpdateMode` is not reported because it is defined for the Link
  form, not as embedded-object behavior.
- The private signature retains every full direct marker with relationship IDs
  normalized to their stored semantics. Same-count program, field-code,
  update-metadata, or target rewrites therefore remain review-visible while a
  relationship-ID renumbering with unchanged semantics remains quiet.
- Added regression coverage for external/internal/unsupported/missing-ID
  classes, duplicate and `Type="Link"` separation, body/header stories,
  Strict encodings, orphan exclusion, unavailable relationships, same-count
  program/update/target rewrites, relationship-ID renumbering stability,
  privacy redaction, and policy/SARIF output. Release validation profiles a
  public Open XML SDK OLE fixture and five public docx4j OLE fixtures; a pinned
  corpus scan found direct markers in 22 of 503 readable Open XML SDK package
  candidates and five of 141 docx4j Word packages.

## 0.26.0 — 2026-08-01

- Added a first-class, privacy-safe inventory for direct WordprocessingML
  embedded-control anchors: `w:control` children of `w:object` or `w:pict` in
  supported Word stories. It is separate from arbitrary `w:control` markup,
  `w:objectLink`, `w:objectEmbed`, VML linked-OLE/image/shape markup, fields,
  ActiveX-binary relationships, and broad embedded-control relationship/payload
  totals.
- Added `DFP065` to require a handoff with no direct Word embedded-control
  anchor and `DFP066` to protect an approved anchor baseline.
- Reports expose only aggregate anchor/story, direct-parent, and backing
  relationship classes: standard control relationships with stored internal or
  external target mode, unsupported relationships, and anchors without `r:id`.
  The external class remains reviewable nonconforming evidence because the
  standard Embedded Control Persistence Part requires an internal target.
  Control names and shape identifiers, relationship IDs and targets, markup,
  story paths, and fingerprints remain private.
- The private signature retains each full direct marker, with relationship IDs
  normalized to their stored relationship semantics. Same-count name, shape, or
  target rewrites therefore remain review-visible while a relationship-ID
  renumbering with unchanged semantics remains quiet.
- Added regression coverage for `w:object`/`w:pict` direct-parent scoping,
  external/internal/unsupported/missing-ID relationship classes, duplicates,
  body/header stories, Strict encodings, unparented and orphan exclusion,
  unavailable relationships, same-count name/target rewrites, relationship-ID
  renumbering stability, privacy redaction, and policy/SARIF output. Release
  validation also profiles docx4j's public `LegacyForms.docx` fixture, which
  supplies five real direct anchors and their internal standard relationship
  chains.

## 0.25.0 — 2026-08-01

- Added a first-class, privacy-safe inventory for direct WordprocessingML
  `w:objectLink` children of `w:object` in supported Word stories. It is
  separate from `w:objectEmbed`, legacy Office VML `o:OLEObject Type="Link"`,
  VML image data and shape links, DrawingML linked pictures, field-code
  inventories, and broad embedded-object relationship/payload totals.
- Added `DFP063` to require a handoff with no stored WordprocessingML
  linked-object-property marker and `DFP064` to protect an approved marker
  baseline.
- Reports expose only aggregate marker/story, exact stored schema-mode
  (`always` / `onCall`) and unsupported-or-missing-mode, and backing
  relationship classes: standard OLE-object relationships with stored external
  or internal target mode, unsupported relationships, and markers without
  `r:id`. Program, shape, relationship IDs and targets, field codes, locking
  metadata, markup, story paths, and fingerprints remain private.
- The private signature retains each full direct marker, with relationship IDs
  normalized to their stored relationship semantics. Same-count program,
  field-code, update-mode, locking, or target rewrites therefore remain
  review-visible while a relationship-ID renumbering with unchanged semantics
  remains quiet.
- Added regression coverage for direct-parent scoping, external/internal/
  unsupported/missing-ID relationship classes, exact standard and unsupported
  or missing update modes, duplicate markers, body/header stories, Strict
  encodings, `w:objectEmbed` and unparented-marker exclusion, orphan exclusion,
  unavailable relationships, same-count field/target rewrites,
  relationship-ID renumbering stability, privacy redaction, and policy/SARIF
  output.

## 0.24.0 — 2026-08-01

- Added a first-class, privacy-safe inventory for direct legacy Office VML
  `o:OLEObject Type="Link"` markup in supported Word stories. It is separate
  from broad embedded-object relationship/payload totals, `Type="Embed"`,
  WordprocessingML `w:objectLink`, VML image data, VML shape links, DrawingML
  linked pictures, and field-code inventories.
- Added `DFP061` to require a handoff with no stored VML linked-OLE marker and
  `DFP062` to protect an approved linked-OLE marker baseline.
- Reports expose only aggregate marker/story, stored automatic-update versus
  nonautomatic-or-unspecified-update, and backing relationship classes:
  standard OLE-object relationships with stored external or internal target
  mode, unsupported relationships, and markers without `r:id`. Link sources,
  program IDs, relationship IDs, field codes, VML markup, story paths, and
  fingerprints remain private.
- The private signature retains each full direct linked-OLE marker, with
  relationship IDs normalized to their stored relationship semantics. Therefore
  same-count source, program, update-mode, field-code, or target rewrites stay
  review-visible while a relationship-ID renumbering with unchanged semantics
  remains quiet.
- Added regression coverage for external/internal/unsupported/missing-ID
  classes, `UpdateMode="Always"`, duplicate markers, body/header stories,
  Strict encodings, `Type="Embed"` and `w:objectLink` separation, orphan
  exclusion, unavailable relationships, same-count source/target rewrites,
  relationship-ID renumbering stability, privacy redaction, and policy/SARIF
  output.

## 0.23.0 — 2026-08-01

- Added a first-class, privacy-safe inventory for direct legacy VML
  `v:imagedata/@r:href` markers in supported Word stories. It is separate from
  VML image-data `r:id` external-image markup, VML shape `href`, DrawingML
  linked pictures, `HYPERLINK` fields, direct `w:hyperlink` markup, and broad
  package relationship totals.
- Added `DFP059` to require a handoff with no stored VML image-data hyperlink
  marker and `DFP060` to protect an approved marker baseline.
- Reports expose only aggregate marker/story and backing-relationship
  classifications: standard hyperlink relationships with stored external or
  internal target mode, plus a reviewable unsupported class for every other
  resolved relationship type or mode. Targets, relationship IDs, VML markup,
  story paths, and fingerprints remain private.
- The boundary deliberately excludes `r:id`, `r:pict`, raw VML `src`, and
  `o:relid` from both its direct-marker definition and private signature. Its
  signature tracks only reviewed `r:href` relationship semantics, so a
  same-count target rewrite remains visible while excluded-attribute changes
  and unchanged-semantics relationship-ID renumbering do not create this
  inventory's churn.
- Added regression coverage for standard external/internal hyperlink and
  unsupported relationship classes, duplicate markers, body/header stories,
  Strict encodings, orphan exclusion, excluded-attribute separation,
  same-count target changes, relationship-ID renumbering stability, privacy
  redaction, and policy/SARIF output.

## 0.22.0 — 2026-08-01

- Added a first-class, privacy-safe inventory for direct legacy VML
  `v:imagedata/@r:id` markers backed by externally stored relationships in
  supported Word stories. It is separate from DrawingML linked pictures,
  `HYPERLINK` fields, direct `w:hyperlink` markup, VML shape `href` markup, and
  broad package relationship totals.
- Added `DFP057` to require a handoff with no stored VML external-image marker
  and `DFP058` to protect an approved external-image marker baseline.
- Reports expose only aggregate marker/story and relationship classifications:
  a standard image relationship with stored external target mode, or another
  externally stored relationship type. Image targets, relationship IDs, VML
  markup, story paths, and fingerprints remain private.
- The boundary deliberately excludes ordinary internal image relationships,
  orphaned image relationships, raw VML `src`, and VML `r:pict`, `r:href`, and
  `o:relid` attributes. Its private signature tracks only the reviewed external
  `r:id` relationship semantics, so a same-count external target rewrite stays
  visible while excluded attributes and unchanged-semantics relationship-ID
  renumbering do not create this inventory's churn.
- Added regression coverage for external standard-image and unsupported
  relationship classes, embedded-image and excluded-attribute separation,
  duplicates, body/header stories, Strict encodings, orphan exclusion,
  same-count target changes, excluded raw-`src` changes, relationship-ID
  renumbering stability, privacy redaction, and policy/SARIF output.

## 0.21.0 — 2026-08-01

- Added a first-class, privacy-safe inventory for direct DrawingML
  `a:blip/@r:link` linked-picture markers in supported Word stories. It is
  separate from `HYPERLINK` fields, direct `w:hyperlink` markup, DrawingML
  hyperlink actions, `r:embed` embedded-picture references, and broad package
  relationship totals.
- Added `DFP055` to require a handoff with no direct DrawingML linked-picture
  markup and `DFP056` to protect an approved marker baseline.
- Reports now expose only aggregate marker/story and backing standard-image
  relationship classifications (external target mode, internal target mode, or
  unsupported). Image targets, relationship IDs, surrounding drawing markup,
  story paths, and fingerprints remain private.
- Each direct marker is counted independently, including duplicate markers and
  markers in Markup Compatibility branches. A standalone image relationship or
  `a:blip` carrying only `r:embed` is not counted. The inventory does not select
  a rendering branch, retrieve or resolve an image, render a picture, update a
  link, or claim a Word client will honor a target.
- Private inventory signatures catch same-count target and direct-markup
  changes while normalizing relationship-ID renumbering with unchanged
  semantics. Added regression coverage for privacy redaction, policy/SARIF
  output, external/internal/unsupported relationship classes, dual
  `r:link`/`r:embed` markup, orphan exclusion, header and Strict encodings,
  same-count target/markup changes, and relationship-ID renumbering stability.

## 0.20.0 — 2026-08-01

- Added a first-class, privacy-safe inventory for direct legacy VML shape-link
  `href` markup in supported Word stories. It is separate from `HYPERLINK`
  fields, direct `w:hyperlink` markup, DrawingML hyperlink actions, and broad
  package relationship totals because a VML geometry element stores its target
  directly.
- Added `DFP053` to require a handoff with no supported direct VML shape-link
  markup and `DFP054` to protect an approved marker baseline.
- Reports now expose only aggregate marker/story, concrete-shape/group/
  shape-template, and direct-`target`-attribute-presence counts. VML `href`
  values, frame targets, titles, alternate text, shape identifiers, story
  paths, and fingerprints remain private.
- Supports direct unqualified `href` attributes on VML `arc`, `curve`, `image`,
  `line`, `oval`, `polyline`, `rect`, `roundrect`, `shape`, `group`, and
  `shapetype` elements, including an empty direct `href` as stored evidence. It
  does not inspect arbitrary VML elements, calculate
  inherited group/template links, select a Markup Compatibility branch, or
  claim a client renders or honors a target.
- Private inventory signatures catch same-count `href`, frame-target, and other
  direct-markup changes. The implementation does not resolve, retrieve,
  follow, validate, evaluate, render, or execute an action.
- Added regression coverage for every supported VML geometry kind, body/header
  stories, privacy redaction, policy/SARIF output, separation from relationship
  totals, fields, `w:hyperlink`, and DrawingML inventories, and same-count
  `href` and target changes.

## 0.19.0 — 2026-08-01

- Added a first-class, privacy-safe inventory for direct DrawingML
  `a:hlinkClick`, `a:hlinkHover`, and `a:hlinkMouseOver` markers in supported
  Word stories, separate from stored `HYPERLINK` fields, direct `w:hyperlink`
  markup, and broad package relationship totals. An unreferenced hyperlink
  relationship is not counted as DrawingML markup.
- Added `DFP051` to require a handoff with no direct DrawingML hyperlink-action
  markup and `DFP052` to protect an approved marker baseline.
- Reports now expose only aggregate marker/story, click/hover/mouse-over,
  external/internal/unsupported/missing-relationship-ID, and `action`/
  `invalidUrl`-attribute-presence counts. Relationship targets, invalid URL
  values, action strings, tooltips, frame names, history settings,
  relationship IDs, story paths, and fingerprints remain private.
- Counts each stored direct marker rather than deduplicating by target,
  relationship, or visual object, and does not select Markup Compatibility
  branches. A missing `r:id` remains reviewable malformed stored evidence;
  recognized relationship modes may be internal or external. No count is a
  safety, reachability, rendering, or action-execution claim.
- Private inventory signatures catch same-count relationship-target and markup
  changes while normalizing relationship-ID renumbering with unchanged
  semantics. The implementation does not resolve, retrieve, follow, validate,
  evaluate, render, or execute an action.
- Added regression coverage for privacy redaction, policy/SARIF output,
  marker kinds, relationship modes/type mismatch, missing IDs, action and
  invalid-URL attributes, orphaned relationship exclusion, header and Strict
  encodings, same-count target/action changes, and relationship-ID renumbering
  stability.

## 0.18.0 — 2026-08-01

- Added a first-class, privacy-safe inventory for direct WordprocessingML
  `w:hyperlink` markup, separate from both stored `HYPERLINK` field codes and
  broad package relationship totals. An unreferenced hyperlink relationship is
  not counted as direct markup.
- Added `DFP049` to require a handoff with no direct hyperlink markup and
  `DFP050` to protect an approved direct-markup baseline.
- Reports now expose only aggregate element/story, relationship-backed,
  external/internal/unsupported relationship, anchor-only,
  current-document-start, and relationship-backed-anchor-attribute counts.
  Relationship targets, anchors, locations, tooltips, frame names, history,
  display text, relationship IDs, story paths, and fingerprints remain private.
- Applies `r:id` precedence over `w:anchor`, distinguishes the documented
  no-attribute current-document-start form, and supports Transitional and
  Strict WordprocessingML/relationship namespaces. A recognized hyperlink
  relationship may have an internal or external target mode; neither count is
  a safety, reachability, rendering, or URL classification claim.
- Private inventory signatures catch same-count relationship-target and markup
  changes while normalizing relationship-ID renumbering with unchanged
  semantics. The implementation does not resolve, retrieve, follow, validate,
  evaluate, or render a link.
- Added regression coverage for privacy redaction, policy/SARIF output,
  relationship modes/type mismatch, anchor precedence, orphaned relationship
  exclusion, header and Strict encodings, same-count target changes, and
  relationship-ID renumbering stability.

## 0.17.0 — 2026-08-01

- Added a first-class, privacy-safe inventory for stored Word `HYPERLINK`
  field references. It is separate from both relationship hyperlinks and the
  existing external-source-field families, because a field code can carry a
  destination directly.
- Added `DFP047` to require a handoff with no stored `HYPERLINK` field
  references and `DFP048` to protect an approved field-reference baseline.
- Reports now expose only aggregate reference/story counts and mutually
  exclusive lexical classes: a literal leading destination, a literal `\l`
  internal-location-only target, or dynamic/unparseable field code. Raw
  destinations, locations, ScreenTips, frame targets, field instructions,
  story paths, and fingerprints remain private.
- Recognizes direct simple-field instructions and complete complex pre-separator
  instructions across supported stories, including separate current/deleted
  revision variants. Loose instruction text and unclosed complex fields remain
  outside the inventory.
- A literal leading destination is not labeled external: Word permits a URL,
  file location, or bookmark there. The inventory never resolves, follows,
  evaluates, or renders a field, and it does not claim that a target is
  reachable or that Word will display a link.
- Added regression coverage for privacy redaction, policy/SARIF output, simple,
  split complex, nested/dynamic, revision, header, Strict, `\l`, formatting
  switch, same-count destination-change, and ignored-instruction encodings.

## 0.16.0 — 2026-08-01

- Added a first-class, privacy-safe inventory for stored Word `DOCVARIABLE`
  field references, separate from `w:docVars` / `w:docVar` storage.
- Added `DFP045` to require a handoff with no stored `DOCVARIABLE` field
  references and `DFP046` to protect an approved field-reference baseline.
- Reports now expose only aggregate reference/story, literal/nonliteral, and
  exact-literal same-scope stored-variable-association counts. Field
  instructions, literal arguments, variable names/values, story paths,
  Settings-part paths, and fingerprints remain private.
- Recognizes direct simple-field instructions and complete complex pre-separator
  instructions across supported stories, including separate current/deleted
  revision variants. Loose instruction text and unclosed complex fields remain
  outside the inventory.
- Associates only a leading plain or wholly quoted literal argument, with
  optional Word field switches, with an exact validated `w:docVar` in the same
  main or glossary package document scope. Nested or compound expressions remain
  nonliteral; an unmatched literal can still be supplied by an attached template.
  The inventory never evaluates a field, resolves a template, or claims that
  Word will display a value.
- Added regression coverage for privacy redaction, policy/SARIF output, simple,
  complex, nested, revision, header, Strict, and glossary-scope encodings.

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
