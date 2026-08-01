# Policy reference

DocFence policies are an intentionally narrow subset of YAML. A policy contains
exactly `version: 1` and a `rules` mapping. Rule values are booleans; absent
rules are disabled.

```yaml
version: 1
rules:
  no_external_relationship_changes: true
  no_macro_payload_changes: true
  require_no_comments: true
```

JSON with the same shape is also accepted. YAML must use spaces, a two-space
`rules` indentation, and the literal values `true` or `false`. Anchors,
aliases, sequences, quoted scalars, nested mappings, duplicate keys, and
unknown rules fail closed. `docfence init path.yml` writes the conservative
starter policy.

## Rule catalog

| Rule | SARIF ID | Trigger | Scope |
| --- | --- | --- | --- |
| `no_external_relationship_changes` | `DFP001` | External relationship signature differs | Comparison |
| `no_macro_payload_changes` | `DFP002` | Macro payload signature differs | Comparison |
| `no_custom_xml_changes` | `DFP003` | Custom XML payload signature differs | Comparison |
| `require_no_unresolved_revisions` | `DFP004` | Candidate has stored revision markup | Candidate |
| `require_no_comments` | `DFP005` | Candidate has stored comments | Candidate |
| `require_no_hidden_text` | `DFP006` | Candidate has direct stored hidden-text runs | Candidate |
| `no_relationship_changes` | `DFP007` | Any stored package relationship differs | Comparison |
| `no_document_settings_changes` | `DFP008` | Stored document settings signature differs | Comparison |
| `no_unclassified_package_payload_changes` | `DFP009` | Payload outside specialized inventories differs | Comparison |
| `require_track_revisions_disabled` | `DFP010` | Candidate enables Track Changes | Candidate |
| `require_no_field_codes` | `DFP011` | Candidate has stored field-code markup | Candidate |
| `require_no_content_controls` | `DFP012` | Candidate has stored content controls | Candidate |
| `require_no_hidden_text_style_declarations` | `DFP013` | Candidate has stored style/default declarations that can hide text | Candidate |
| `require_no_hidden_paragraph_marks` | `DFP014` | Candidate has direct hidden paragraph-mark markup | Candidate |
| `require_no_embedded_objects` | `DFP015` | Candidate has stored embedded OLE/package/control relationships or payload parts | Candidate |
| `require_no_alternative_format_imports` | `DFP016` | Candidate has stored `aFChunk` import relationships, payloads, or `w:altChunk` anchors | Candidate |
| `no_embedded_object_payload_changes` | `DFP017` | Embedded OLE/package/control inventory differs | Comparison |
| `no_alternative_format_import_changes` | `DFP018` | Alternative-format import inventory or anchor count differs | Comparison |
| `no_document_property_changes` | `DFP019` | Core, extended, or custom document-property inventory differs | Comparison |
| `require_no_custom_document_properties` | `DFP020` | Candidate has stored custom document-property definitions | Candidate |
| `require_no_mail_merge` | `DFP021` | Candidate has stored mail-merge configuration, source, or recipient-data state | Candidate |
| `no_mail_merge_changes` | `DFP022` | Mail-merge inventory differs | Comparison |
| `require_no_data_bindings` | `DFP023` | Candidate has stored content-control data bindings | Candidate |
| `no_data_binding_changes` | `DFP024` | Content-control data-binding inventory differs | Comparison |

All current findings have `high` severity except macro payload changes, which
are `critical`. SARIF deliberately contains no locations: a package member path
or paragraph reference can itself reveal information about a confidential file.

## Comparison rules vs candidate rules

Comparison rules protect a delta. For example,
`no_external_relationship_changes` permits a pre-existing external relationship
when it is unchanged. Candidate rules inspect the after document independently
of the baseline. For example, `require_no_comments` fails whenever the
candidate contains comments, even when both documents contain the same number.

Choose policies based on the handoff boundary. A publishing gate often uses the
starter policy and candidate-state rules. A controlled template workflow might
also enable `no_document_settings_changes` and
`no_unclassified_package_payload_changes`; those are intentionally stricter and
can flag a media, metadata, or other opaque package mutation. A template that
intentionally embeds a known payload can instead enable
`no_embedded_object_payload_changes` or
`no_alternative_format_import_changes` to preserve that baseline while blocking
a later mutation. `no_document_property_changes` gives the same controlled
baseline option for document metadata, and `no_mail_merge_changes` does so for
mail-merge configuration and recipient state. `no_data_binding_changes` does
the same for a controlled template whose content controls intentionally map to
custom XML. `word/styles.xml` is handled by the dedicated style inventory
instead.

## Hidden-text scope

`require_no_hidden_text` covers direct `w:vanish` on ordinary runs only.
`require_no_hidden_paragraph_marks` covers direct `w:vanish` or `w:specVanish`
under `w:pPr/w:rPr`; it is separate because a paragraph mark is not a text run.

`require_no_hidden_text_style_declarations` fails when an enabled `w:vanish`
appears in a stored style's text-run properties or in document-default run
properties. Word treats hidden text as a toggle property in styles, and its
effective state also depends on style inheritance and application. Therefore,
this rule is a conservative declaration gate, not a claim that DocFence has
resolved which candidate runs Word will display as hidden. A `w:vanish` with an
explicit false value does not trigger the declaration count.

## Embedded and imported-content scope

`require_no_embedded_objects` is a candidate-state gate for OLE/package and
control payload evidence. It recognizes the standard `oleObject`, `package`,
`control`, and ActiveX-control-binary relationship types, as well as payloads
stored in the conventional `word/embeddings/` and `word/activeX/` folders. The
comparison rule fingerprints those recognized relationship semantics and the
payload bytes privately, so relationship-ID renumbering alone remains quiet.

`require_no_alternative_format_imports` fails when a candidate has an `aFChunk`
relationship, its resolved internal payload, or a direct Word `w:altChunk`
anchor. The comparison counterpart detects a relationship/payload mutation and
an anchor-count mutation. An unanchored `aFChunk` relationship is still
inventoried because it is stored imported-content state.

For every encountered `w:altChunk`, DocFence requires a matching internal
standard `aFChunk` relationship whose target safely resolves to a stored package
member. It does not decode, import, render, execute, or assess the safety of
that payload. Findings and reports expose only aggregate counts.

## Document-property scope

DocFence inventories the standard core, extended, and custom document-property
relationship types, plus the canonical `docProps/core.xml`, `docProps/app.xml`,
and `docProps/custom.xml` paths when present. The property root must match the
expected OOXML vocabulary. Reports contain only part and aggregate-value counts;
property names, values, paths, relationship targets, and fingerprints remain
private.

Core and extended value counts cover direct property elements with stored text,
including automatically maintained timestamps, statistics, application details,
and template metadata. `no_document_property_changes` therefore detects a
normal resave that updates one of those fields. Use it when that exact metadata
delta matters at a controlled boundary, not as an assertion that every metadata
change is suspicious.

`require_no_custom_document_properties` is narrower: it fails only when the
candidate contains stored custom-property definitions. It does not classify
core or extended metadata as personal, confidential, user-authored, or safe.
Custom property names and values are never emitted.

## Mail-merge scope

DocFence inventories direct `w:mailMerge` settings plus recognized external
`mailMergeSource` and `mailMergeHeaderSource` relationships from
`word/settings.xml`. It also inventories recognized internal recipient-data
relationships and their stored target parts. Both documented recipient-data
relationship spellings are accepted, so conventional and Strict OOXML packages
can be compared without treating a relationship-ID rewrite as a mutation.

For an encountered direct `w:dataSource`, `w:headerSource`, ODSO `w:src`, or
ODSO `w:recipientData` reference, the relationship must have the required type
and target mode. Internal recipient targets must resolve to a stored package
part. Invalid recognized mail-merge markup fails closed. A recognized stored
source or recipient relationship is still inventoried even when no current
`w:mailMerge` element references it, because a residual relationship remains a
review surface.

`require_no_mail_merge` is a candidate-state gate: it fails if any of those
aggregate counts is nonzero. `no_mail_merge_changes` compares a private
signature of the mail-merge configuration, relationship semantics, and
recipient-data payloads. Public reports show only aggregate counts. Connection
strings, queries, table names, field mappings, source/header targets,
relationship IDs, recipient bytes, and fingerprints are never emitted.

These rules do not connect to a data source, execute a query, select a
recipient, or decide whether mail-merge state is expected, safe, personal, or
malicious. Use a clean-handoff gate when no mail merge should remain; use a
controlled baseline when an approved template legitimately retains it.

## Content-control data-binding scope

DocFence inventories direct standard `w:dataBinding` children of a structured
document tag's `w:sdtPr` properties in recognized Word document stories. This
is a mapping declaration, not an XPath engine: DocFence does not select a node,
replace visible content, resolve a rich-text mapping, or claim what Word will
display.

For a nonempty `w:storeItemID`, DocFence discovers the associated custom XML
data part only through a standard internal `customXmlProps` relationship and a
valid `ds:datastoreItem` storage identifier. Conventional and Strict OOXML
relationship forms are accepted, along with both the current
`customXmlDataProps` root vocabulary and the established Word `customXml`
spelling found in real packages. A malformed recognized properties relationship
or properties root fails closed. A storage ID with no discovered associated part
is not treated as a parser error: it is counted as unmatched review evidence.

A binding can omit `w:storeItemID`; Word may then search custom XML parts using
the XPath. DocFence counts it as a binding without a storage ID but deliberately
does not guess the selected part. The public inventory reports aggregate binding
counts, referenced custom-XML part counts, and unmatched-ID counts only.
XPath expressions, prefix mappings, storage IDs, part names, data values, and
private fingerprints are never emitted.

`require_no_data_bindings` fails whenever the candidate has at least one
recognized mapping declaration. `no_data_binding_changes` compares the private
mapping declarations and, for bindings with a discovered storage ID, the paired
custom XML data/properties payloads. It therefore catches a mutation to an
identified bound data part without exposing it. It does not make an unscoped
XPath binding identify a target; use `no_custom_xml_changes` as well when every
custom XML mutation must block a handoff.

## CI example

```bash
docfence init docfence.yml
docfence check approved.docx candidate.docx \
  --policy docfence.yml \
  --format sarif \
  --output docfence.sarif
```

Exit status `1` means the report was produced and contains findings. Exit status
`2` means DocFence could not establish bounded evidence, so a CI workflow should
treat it as an infrastructure or input-validation failure rather than a clean
comparison.
