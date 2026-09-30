# Template specification

Use `office-template-spec@1`. A specification is bound to the exact template bytes by SHA-256. Re-register or deliberately update it when the template changes.

## Companion visual contract (v2)

Keep the existing slot schema; add the source-bound `template-visual.json` and `style-plan.json` described in [visual-contract.md](visual-contract.md). PPTX filling requires the visual handbook. The existing `visual` fields describe per-slot mapping; the handbook includes locked and added objects, palette/font evidence and page anchors. Cross-link stable IDs and JSON pointers. The source-bound change plan supplies exact source slide mapping, donor IDs and allowable deltas; it must agree with slot adaptive bounds. Character capacity and style fidelity are separate validations.

## Slot contract

Each slot requires:

```json
{
  "id": "resume.summary",
  "locator": {
    "kind": "docx-content-control-tag",
    "matches": [{"part": "word/document.xml", "value": "ot.slot:resume.summary"}]
  },
  "required": true,
  "role": "body",
  "content_type": "plain_text",
  "visual": {
    "page_or_slide": 1,
    "object_id": "stable-office-object-id",
    "geometry": {"x": 1.2, "y": 2.1, "w": 5.4, "h": 1.3, "unit": "in"},
    "z_order": 7,
    "json_pointer": "/slots/0"
  },
  "budget": {
    "min_chars": 60,
    "recommended_max_chars": 105,
    "hard_max_chars": 125,
    "max_width_units_per_line": 27,
    "max_lines": 5
  },
  "must_include": [],
  "fit_policy": {
    "preserve_line_count": false,
    "preferred_lines": 3,
    "min_font_pt": 9,
    "adaptive_bounds": {"max_x_delta": 0.0, "max_y_delta": 0.0, "max_w_delta": 0.5, "max_h_delta": 0.3},
    "protected_neighbors": ["resume.photo", "resume.skills"],
    "allowed_actions": ["remove_manual_breaks", "widen", "increase_height", "uniform_font_reduce", "tighten_spacing", "compact_variant", "edit_copy"]
  },
  "capacity_states": [
    {"id": "default", "recommended_max_chars": 105, "hard_max_chars": 125, "max_lines": 5},
    {"id": "wider", "recommended_max_chars": 118, "hard_max_chars": 138, "max_lines": 5}
  ],
  "overflow_policy": ["widen", "increase_height", "uniform_font_reduce", "tighten_spacing", "compact_variant", "remove_redundancy", "shorten_sentences", "request_user_choice"],
  "calibration": {
    "status": "verified",
    "renderer": "Microsoft Word",
    "font": "Aptos 10pt",
    "notes": "Verified with Chinese, Latin, digits, and mixed text."
  }
}
```

`min_chars` is a content-quality floor only when the slot needs one. Set it to `0` for names, dates, identifiers, and other naturally short values.

`recommended_max_chars` guides drafting. Exceeding it produces a warning. `hard_max_chars` is a build failure.

`max_width_units_per_line` is a conservative preflight estimate: full-width characters count as roughly 1 unit, Latin letters and digits as roughly 0.55, and spaces as roughly 0.33. It does not replace native rendering.

`must_include` contains exact facts or phrases that may not disappear during compression. Prefer generating it from the verified fact model rather than asking the model to invent it.

`role` controls the default fit strategy. Use at least `title`, `subtitle`, `body`, `caption`, `label`, `table-cell`, and `data-value` where applicable. A `title` should declare whether its intended line count is one or two; one-line titles must not wrap silently.

`visual` provides the bidirectional bridge between structure and rendering. The locator remains authoritative; coordinates support highlighting and hit-testing. A selection on the rendered side must resolve to the slot's `json_pointer`, while selecting that JSON node must highlight the same object.

`fit_policy` defines permitted layout mutations. All deltas are relative to the registered baseline. Empty bounds mean geometry is locked. `protected_neighbors` and page safe areas are collision barriers, not suggestions.

`capacity_states` describes measured capacity after each permitted adaptation. Do not advertise a capacity state unless that state has been rendered and verified. The active state's budget governs drafting and validation.

## Calibration

For every fixed-layout slot:

1. Test representative Chinese, Latin, numeric, URL, and mixed-script content.
2. Render with the target fonts and production renderer.
3. Measure the baseline geometry, font, margins, line spacing, intended line count, protected neighbors, and safe area.
4. Test permitted adaptations in layout-first order and determine the comfortable maximum and hard failure boundary for each state.
5. Record the smallest allowed font, geometry deltas, spacing limits, and any coupled-object movement.
6. Verify that every structure node maps to exactly one rendered region and that every editable region maps back to one node.
7. Verify the template's permitted page or slide count.

Do not set `calibration.status` to `verified` from character counting alone.
