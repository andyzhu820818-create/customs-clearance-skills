---
name: customs-clearance-ci-pl-header-footer
description: Use when updating CI/PL customs clearance templates from warehouse XLS files, a DFCR/FCR PDF, and PT-container mapping screenshots, especially to rewrite only header and footer callout fields while preserving the template formatting and leaving item details untouched.
metadata:
  short-description: Update CI/PL header and footer from DFCR and warehouse files
---

# Customs Clearance CI/PL Header/Footer Updates

Use this skill when the user asks to update a commercial invoice / packing list template using:

- warehouse `.xls` files containing PT, PO, SKU, prepack SKU, quantities, weights, carton details, and product rows
- a DFCR/FCR PDF containing shipment-level transport data
- a PT/seal/container mapping screenshot or table
- an existing CI/PL Excel template whose format must be preserved

## Scope

Default to editing a copy of the supplied template. Preserve the original workbook.

## Scripted Workflow

Prefer the bundled script for repeatable jobs instead of rewriting ad-hoc code:

```bash
python scripts/update_ci_pl_header_footer.py \
  --dfcr DFCR_NGB461001.pdf \
  --mapping mapping.json \
  --warehouse "仓库4-PT27045 ...xls" "仓库4-PT27046 ...xls" \
  --vendor-template VENDOR-template.xlsx \
  --factory-template FACTORY-template.xlsx \
  --output-dir <destination-folder> \
  --summary-json summary.json
```

`mapping.json` must be a UTF-8 JSON array in DFCR-independent form:

```json
[
  {"container": "ZCSU6676780", "seal": "A4260019073", "pt": "PT27045"},
  {"container": "JXLU6383951", "seal": "A4260019080", "pt": "PT27046"}
]
```

The script reads DFCR container order and reorders the mapping automatically. It converts `.xls` warehouse files through Excel COM, updates only header/footer callouts, preserves invoice-cell rich-text sizes via character-level Excel COM edits, marks invoice-number values red by default, preserves existing `TOTAL NET WEIGHT` text while coloring it red, and prints a JSON summary for verification.

For the workflow described by this skill, update only:

- CI header shipment and invoice fields
- CI footer `CONTAINER/SEAL NO.` SKU callout
- PL header PO/invoice/shipping summary fields
- PL footer `CONTAINER/SEAL NO.` SKU callout

Do not fill or rebuild middle item-detail rows unless the user explicitly asks.

Do not fill `TOTAL NET WEIGHT`; leave it for the person completing the final detail table.

## Source Priority

Use the DFCR/FCR PDF for shipment-level values:

- ETD
- ETA, if needed
- vessel and voyage
- container numbers
- seal numbers
- total gross weight
- port of loading, discharge, and destination, when present

Use warehouse files for PT-level and item-level values:

- PT number
- PO number
- prepack SKU
- SKU
- per-SKU quantity
- item grouping under each PT

Use the PT/seal mapping image or table to connect:

`PT number -> seal number -> container number`

## Ordering Rule

All ordering is controlled by the container order shown in the DFCR/FCR PDF.

Do not sort by PT number, Excel file order, screenshot order, PO number, or SKU number unless the user explicitly changes the rule.

Process as:

1. Read container order from the DFCR/FCR PDF.
2. Pair each container to the seal in exactly the same DFCR/FCR order.
3. Use the PT/seal mapping to identify which PT belongs to each seal.
4. Use that ordered PT list for all header/footer and PO/SKU callout order.

Example:

```text
DFCR order:
EITU8128178 / EMCDSN0655
EITU8132481 / EMCDSM6425

Mapping:
EMCDSN0655 -> PT27082
EMCDSM6425 -> PT27081

Final PT order:
PT27082
PT27081
```

## Header Rules

CI page:

- `CONTAINER/SEAL NO.`: write container/seal pairs in DFCR order, using the template's line style.
- `FINAL DESTINATION`: use DFCR `PLACE OF DELIVERY`, not `PORT OF DISCHARGE`.
- `VESSEL`: use vessel and voyage from DFCR.
- `ETD`: use DFCR ETD.
- `SAILING ON OR ABOUT`: equal to DFCR ETD.
- `TOTAL GROSS WEIGHT`: use DFCR gross weight directly; do not calculate it from warehouse Excel.
- `TOTAL NET WEIGHT`: do not actively fill and do not delete an existing template value. If the template already contains a net-weight value, preserve the text and change that cell's font color to red so it is visibly marked for human review.
- `SHIP DATE`: only update when the user has provided the intended rule or source. If not provided, preserve template value or ask briefly.
- `INVOICE DATE`: set to one calendar day after `SHIP DATE`; keep the template's date text format, for example `JUN 08, 2026`.
- When updating cells that contain multiple lines such as `INVOICE DATE` / `INVOICE No.`, preserve the existing rich-text styling and per-line font sizes. Do not flatten the cell into one uniform font size; edit only the needed text and keep the template's original font-size differences.
- For early/draft work where the invoice number is not final, replace the invoice-number value's last three characters with uppercase `XXX` and mark only that value in red. Preserve the existing `INVOICE No.` label style and the existing font size for that line.

PL page:

- In the header summary, list all PO numbers contained in the ordered PTs.
- Preserve first occurrence order by DFCR container/PT order.
- If a PO appears more than once, include it only once.
- Preserve `INVOICE NO.` unless the user explicitly provides a new invoice number.
- For early/draft work where the invoice number is not final, replace the `INVOICE NO.` value's last three characters with uppercase `XXX` on the PL page as well, then mark only that value in red; keep the label black.
- Set `INVOICE DATE` to one calendar day after `SHIP DATE`.
- Update `ETD` and `PORT OF DISCHARGE` only when their source values are available; `PORT OF DISCHARGE` remains separate from CI `FINAL DESTINATION`.

## Footer Callout Rules

Footer callout format follows the template:

```text
CONTAINER/SEAL NO.:
<container>/<seal> including sku# <prepack sku>,<prepack sku>;
<container>/<seal> including sku# <prepack sku>,<prepack sku>.
SKU# <prepack sku>: <n> pcs were loaded into container <container>; the remaining <n> pcs were loaded into container <container>.
```

Use the warehouse file's main/prepack SKU column for the footer callout unless the user's template clearly uses another SKU field.
If a source row has both a main/prepack SKU and a sub SKU, record only the main/prepack SKU in the callout.
If following rows belong to the same main/prepack SKU group and their main/prepack SKU cell is blank, continue using the previous main/prepack SKU; do not list child/sub SKUs from those continuation rows.
Use the sub SKU only when the main/prepack SKU for that product row is blank.

For each container:

1. Find the PT assigned to that container through the seal mapping.
2. Collect SKUs from that PT's warehouse file in source order, using main/prepack SKU first and falling back to sub SKU only when main/prepack SKU is blank.
3. Remove repeated prepack SKUs within the same container callout.
4. Write containers in DFCR order.

If the same prepack SKU appears in more than one container:

- Add a final split note after the container lines.
- Use the per-container quantities from the warehouse file.
- For two containers, use the template wording:

```text
SKU# <sku>: <qty1> pcs were loaded into container <container1>; the remaining <qty2> pcs were loaded into container <container2>.
```

For more than two containers, keep the same style but list each container and quantity clearly.

## Excel Editing Rules

- Work on a copied output workbook, not the template.
- Use the generation date (`YYYYMMDD`) for the output workbook filename date; do not use `SHIP DATE` or `INVOICE DATE` for the filename date.
- Use Excel/COM or another format-preserving method when the template contains merged cells, print settings, page layout, and precise borders.
- Preserve merged cells, fonts, borders, alignment, print area, fixed footer sections, banking details, vendor details, and manufacturer details.
- Set wrap text for long header/footer cells.
- Vertically center the CI and PL footer `CONTAINER/SEAL NO.` callout cells after updating them.
- Adjust row height only where needed so longer footer text remains readable.
- Do not overwrite middle detail rows, formulas, discount lines, banking details, shipping mark, signatures, or fixed legal text unless specifically instructed.

## Verification Checklist

Before final response, reopen the output workbook read-only and confirm:

- Original template file remains unchanged.
- DFCR container/seal order is preserved exactly.
- Seal-to-PT mapping is applied correctly.
- CI and PL footer callouts match the same container order.
- Duplicate SKU split notes appear when required.
- PL PO list is ordered by DFCR/PT order and deduplicated.
- `TOTAL GROSS WEIGHT` comes from DFCR.
- `TOTAL NET WEIGHT` has no filled numeric value unless the user explicitly requested it.
- Middle detail rows were not changed when outside scope.
