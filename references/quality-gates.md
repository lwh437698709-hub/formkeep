# Quality gates

The build is fail-closed. A file is deliverable only when all applicable gates pass.

## Visual style gate (v2; before and after all editing)

- Follow [visual-contract.md](visual-contract.md): source-bound handbook and frozen intended-change plan precede authoring.
- Extract actual output styles and resources; do not accept an agent's declaration that it followed the template.
- Added/changed icons, shapes, charts, tables and text inherit a same-role template donor, including color roles, stroke, fill, font, opacity and effects.
- New arbitrary colors, default chart palettes, emoji icons and unrelated icon families are blocking errors.
- Every slide and the deck meet ≥90 structural style fidelity, with zero hard violations. Coverage of source layouts is not a fidelity score.
- Preserve background, masters, logos and protected decorations. Check edge/center anchors, title baseline, columns and safe zones.
- Missing fonts have measured substitutions; the native renderer verifies glyphs, wraps and hierarchy.
- Exact-source/output hashes bind the audit and rendered images. A later edit requires re-audit.
- The PPTX helper exits nonzero for both blocking errors and missing native-render evidence. Distinguish structural pass from deliverable status.

## Automated preflight

- Template SHA-256 matches the registered specification.
- Every required slot has a binding.
- No unknown binding is present.
- Every slot has verified calibration.
- Minimum useful content is met where applicable.
- Recommended maximum exceptions are reported.
- Hard character and estimated-line limits are not exceeded.
- Every `must_include` phrase remains in the binding.
- Every registered slot is found and filled in the output.
- Every used layout state and fit action is declared by the slot contract.
- One-line title slots contain no unexpected wrap or manual line break.
- The output is a readable ZIP/OOXML package.

## Structure-map review

- Every editable region has one stable slot id and one authoritative Office locator.
- Code selection highlights the exact visual object; visual selection resolves to the exact JSON path.
- Geometry, z-order, style, lock state, protected neighbors, and adaptive bounds match the template.
- No unregistered editable region or ambiguous many-to-one mapping remains.

## Native render review

- Expected page or slide count is unchanged unless change is permitted.
- No text is clipped, hidden, or outside its region.
- No text or shape overlaps another protected element.
- No unexpected blank page or slide appears.
- No font substitution changes wrapping or hierarchy.
- Titles preserve their registered line count. If a title was adapted, width expansion was attempted before font reduction, and font reduction was uniform.
- Every geometry change stays within adaptive bounds and outside protected regions.
- Baseline/output overlay or side-by-side comparison shows no unintended movement, scaling, crop, or style drift.
- Locked headers, footers, logos, redheads, signatures, and page furniture remain fixed.
- DOCX headings, tables, images, page breaks, headers, footers, and section boundaries remain valid.
- PPTX text, images, crop, alignment, safe areas, and master elements remain valid.
- XLSX visible values, formulas, row heights, column widths, print areas, and page breaks remain valid.

## Repair report

When a gate fails, report the exact slot, JSON path, actual measurement, active capacity state, permitted measurement, protected facts, attempted layout actions, remaining actions, and blocked protected regions. Do not return a generic “content too long” message.

If native rendering is unavailable, return the generated file only as a draft and state that visual safety has not been established.
