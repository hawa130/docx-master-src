# OOXML Schemas — Provenance

XSD schemas used by `lib/docx-validate.ts` (xmllint-wasm-based OOXML
validator). Each subdirectory is sourced from the upstream listed below.
Files are unmodified except where noted.

## ISO-IEC29500-4_2016/

ECMA-376, 5th edition (December 2016), Part 4: Transitional Migration
Features. Downloaded from:

- https://ecma-international.org/wp-content/uploads/ECMA-376-4_5th_edition_december_2016.zip
- → inner archive `OfficeOpenXML-XMLSchema-Transitional.zip`

ECMA standards are freely available; redistribution permitted under ECMA's
standard terms (https://ecma-international.org/policies/by-ipr/ecma-text-copyright-policy/).

Plus `xml.xsd` (W3C XML namespace schema):

- https://www.w3.org/2001/xml.xsd
- W3C Software and Document Notice and License — permits redistribution.

## ecma/fouth-edition/

ECMA-376, 5th edition (December 2021), Part 2: Open Packaging Conventions
(OPC). Downloaded from:

- https://ecma-international.org/wp-content/uploads/ECMA-376-2_5th_edition_december_2021.zip
- → inner archive `OpenPackagingConventions-XMLSchema.zip`

(The directory name "fouth-edition" is a typo preserved from the OPC archive
itself.)

Plus two **hand-written stubs** for Dublin Core schemas that
opc-coreProperties.xsd imports but ECMA doesn't bundle:

- `dc.xsd` — minimal Dublin Core Elements stub.
- `dcterms.xsd` — minimal Dublin Core Terms stub.

Both are CC0-licensed (see file headers); they declare only the elements
opc-coreProperties.xsd references, with `xs:anyType` content. They're not
derived from the W3C/Dublin Core official schemas.

## mce/

`mc.xsd` — **hand-written** Markup Compatibility (MC) namespace schema. ECMA-376
Part 3 defines this namespace but the official download ships only the spec
PDF, not the XSD. Authored from the ECMA-376 Part 3 normative text. CC0.

## What's not bundled

Microsoft extension schemas (w14, w15, w16cid, w16cex, w16sdtdh, w16se,
w16) — defined in Microsoft's [MS-DOCX] Open Specifications. Not bundled
because:

1. The validator pre-strips MC-extension namespaces from main-content parts
   (`stripMcExtensions` in `lib/docx-validate.ts`), matching the official
   Anthropic docx skill's approach. After stripping, w14/w15/w16 attributes
   are gone from the XML before XSD validation, so their schemas aren't
   needed.
2. Parts that are themselves MS-extension-namespaced (`word/people.xml`,
   `word/commentsExtended.xml`, etc.) are skipped by `pickSchema`.
