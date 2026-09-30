---
name: formkeep
description: Register, visualize, and fill explicitly marked DOCX, PPTX, or XLSX templates with bidirectional structure maps and layout-first fit policies, then block delivery when layout safety cannot be established. Use for fixed-layout resumes, presentations, forms, reports, and spreadsheet templates. Do not use for unmarked arbitrary templates without first completing template registration.
---

# FormKeep

Produce an editable Office file without silently clipping, deleting, or overflowing content. Treat layout safety and template fidelity as hard delivery gates. The template is the primary constraint; supplied copy is fitted into its registered layout contract, not pasted first and repaired later.

## Visual red lines (v2)

Before authoring, create a **template visual handbook** from the untouched source. Read [references/visual-contract.md](references/visual-contract.md). Every added or changed element must preserve the source's visual language: palette, color roles, typography, icon family, stroke width, fill, effects, spacing, anchors, and image crop. This applies to text, icons, charts, tables, shapes, images, connectors, and copied slides. Adding content does not authorize redesign.

- Do not invent a palette, choose arbitrary theme accents, or copy photographic colors into graphic styling. Use the page's same-role source object as the style donor.
- Prefer editing or duplicating an existing object. Every addition needs a source donor, purpose, target anchors and an entry in the pre-authoring change plan. New assets without a matching registered style require a separately reviewed template variant.
- Preserve logos, background art, headers, footers and master/layout elements. Do not overwrite them to make room.
- Inspect required fonts before drafting. A missing font triggers a documented candidate substitution, character-coverage and layout test; never silently substitute a decorative/brand font or assume an embedded font is usable. Recommendations are not automatic approvals.
- Require **each slide and the deck to reach at least 90/100 structural style fidelity**, plus zero red-line violations and native visual QA. A new off-palette icon blocks delivery even when the deck score exceeds 90. Never report this metric as 90% pixel or perceptual similarity.
- Lack of a renderer, unsupported styles, unresolved inheritance or missing evidence means a draft/unverified result; never invent a passing score or render evidence.

PPTX helper workflow (Python standard library only):

```bash
python3 scripts/pptx_style_guard.py extract TEMPLATE.pptx --to template-visual.json
python3 scripts/office_template.py fill TEMPLATE.pptx template-spec.json bindings.json draft.pptx --visual-book template-visual.json
python3 scripts/pptx_style_guard.py audit TEMPLATE.pptx final.pptx --plan style-plan.json --review render-review.json --to style-audit.json
```

Extraction also produces `template-visual.md`. Complete its role/anchor and font review before generation; freeze the source hash and change plan. Audit re-extracts the actual output: a handbook alone does not enforce style. See the reference for schemas, score meaning and limitations. For DOCX/XLSX apply the same handbook/red lines using native format-specific inspection; the PPTX helper does not claim to score those formats.

## Route the task

1. Determine whether the user is registering a template or filling an already registered template.
2. Read [references/format-support.md](references/format-support.md) for the requested file type.
3. Read [references/visual-mapping.md](references/visual-mapping.md) when registering a template or building the structure/template comparison UI.
4. Read [references/template-spec.md](references/template-spec.md) when creating or changing a template specification.
5. Read [references/quality-gates.md](references/quality-gates.md) before generating a deliverable.

## Register a template

Run `python3 scripts/office_template.py inspect TEMPLATE --to template-spec.json`.

The MVP recognizes only explicit slots:

- DOCX: Word content-control tag `ot.slot:<id>`.
- PPTX: shape name `ot.slot:<id>`.
- XLSX: defined name `ot_slot_<id>` or `ot.slot.<id>` pointing to one cell.

Do not infer an unmarked region and write into it. The generated specification is intentionally uncalibrated. Establish each slot's semantic role, geometry, protected neighbors, minimum useful content, recommended maximum, hard maximum, line-width budget, maximum lines, required facts, and permitted fit actions. Capacity must be calibrated for every permitted layout state, not guessed from character count alone. Mark calibration `verified` only after testing the slot with the target fonts and rendering engine.

Registration is incomplete until it produces both:

- a machine-readable structure map with stable slot ids, object locators, page/slide/sheet coordinates, z-order, style, lock state, adaptive bounds, and JSON paths;
- a rendered visual map in which selecting a region resolves to exactly one structure node and selecting a structure node highlights that region.

Use the structure map as the source of truth for replacement. Do not rely on visual coordinates alone when a stable Office object locator exists.

Preserve the original template. Store the specification beside a copied or versioned template, not inside the only original.

## Fill a registered template

1. Extract a complete fact model from the user's material before shortening it.
2. Separate facts into `must_keep`, `important`, and `optional` information.
3. Select the registered layout state for each slot before drafting. For titles, prefer the state that preserves the intended line count and hierarchy.
4. Draft each slot within that state's recommended budget. Never change numbers, names, dates, credentials, or other source facts merely to fit.
5. Save bindings as a JSON object from slot id to plain text.
6. Run `python3 scripts/office_template.py validate template-spec.json bindings.json`.
7. If validation passes, run `python3 scripts/office_template.py fill TEMPLATE template-spec.json bindings.json OUTPUT`; PPTX additionally requires `--visual-book template-visual.json` from the exact template.
8. Apply only the selected, registered geometry/font/spacing adaptations. The current filler writes text but does not apply these adaptations automatically; use the format-appropriate authoring tool and record every change in a fit report.
9. Re-run the style audit after ALL edits, chart changes and exports. Render that exact output with the intended Office engine and complete the visual gates in `quality-gates.md`. Bind evidence to the final output SHA-256; any subsequent edit invalidates it.

The fill command refusing a file is a correct outcome. Do not bypass missing calibration, a template hash mismatch, a hard capacity violation, a missing required slot, or missing must-keep text.

## Fit and repair overflow

Repair only the failing slot and any explicitly coupled objects. Use this layout-first order unless the registered specification says otherwise:

1. Remove accidental manual line breaks and use the registered wrap rule.
2. Expand the text box within its registered `adaptive_bounds`, first horizontally and then vertically, without entering protected regions or changing the visual grid.
3. Reposition only registered coupled objects within their allowed ranges.
4. Reduce font size uniformly within the registered minimum while preserving hierarchy; do not shrink isolated words or lines.
5. Tighten internal margins, paragraph spacing, or line spacing only within registered limits.
6. Use a registered compact layout variant.
7. Only after layout options are exhausted, remove redundancy, shorten sentence structure, and remove `optional` information in that order.
8. Add a page, slide, row, or continuation block only when the template contract permits it.
9. Ask the user to choose what to omit or whether to approve a new layout variant.

For a title whose contract is one line, wrapping is a failure unless the specification explicitly permits two lines. First widen the title region; then reduce the whole title's font size within limits; then use the registered compact title variant. Do not rewrite or truncate the title merely because it wrapped.

Never silently truncate text, hide content, shrink below the minimum type size, move locked elements, or convert a one-page document to multiple pages without authorization.

## Deliver only after QA

Structural capacity checks are necessary but do not prove visual fit. Deliver only when the intended renderer confirms page/slide count, fonts, text visibility, absence of overlaps, and locked-region stability. If native rendering is unavailable, label the output as a draft requiring visual review; do not claim it is layout-safe.
