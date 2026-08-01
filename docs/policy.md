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
can flag a media, metadata, or other opaque package mutation. `word/styles.xml`
is handled by the dedicated style inventory instead.

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
