# Bidirectional visual mapping

Use this contract for the registration UI and for visual QA. A flat screenshot is not sufficient because it cannot identify the underlying Office object.

## Data model

Each editable object must expose:

- template visual handbook pointer, source-style donor, color/font role, palette tokens and resolved/unresolved state;
- normalized left/right/top/bottom/center anchors, title baseline or text baseline when measured, group coordinate space and composed transform;

- stable `slot_id` and authoritative Office locator;
- JSON Pointer to its structure node;
- page, slide, or sheet index;
- geometry in document units plus normalized 0–1 coordinates for browser hit-testing;
- z-order, rotation, object type, style summary, lock state, and protected-neighbor ids;
- editable text, intended line count, capacity states, and permitted fit actions.

Nested groups may appear in the structure tree, but the leaf slot must remain directly selectable. Preserve source object ids where the format exposes them; do not use array position as the only identity.

## Interaction

- Clicking or keyboard-selecting a structure node scrolls to the correct page and highlights exactly one visual object.
- Clicking an editable visual object focuses and expands its structure node.
- Overlapping objects are resolved by z-order, with a cycling control for objects under the pointer.
- Hover may preview a match, but selection is persistent and shown on both sides.
- Zoom and responsive scaling must not change hit targets; transform source coordinates into viewport coordinates.
- Locked or decorative elements remain inspectable but cannot be treated as fill slots.
- Ambiguous, missing, or many-to-one mappings are registration failures.

## Comparison modes

Provide three synchronized views:

1. `baseline`: the untouched registered template;
2. `filled`: the current output;
3. `diff`: an overlay or pixel-difference view, with expected text changes separated from unexpected geometry/style movement.

Selecting a changed slot shows its baseline geometry, current geometry, active capacity state, font change, line count, attempted fit actions, and any collision. Use the same stable slot id in the structure map, fit report, and QA report.

Also show added/deleted objects, old/new colors, font substitution evidence, donor style, per-page fidelity scores and hard violations. A locked object remains selectable for inspection. Never hide foreground discrepancies behind a whole-page similarity number.

## Layout-first review

For title overflow, show candidate states in this order: wider box, taller box when allowed, uniformly smaller title font, tighter registered spacing, compact title variant, then copy editing. Never present silent wrapping or truncation as a successful state.

For body overflow, preview the smallest permitted local change first and keep protected regions visible. Reject a candidate when it collides, clips, crosses a safe area, changes a locked object, breaks the page/slide count contract, or creates a new overflow elsewhere.
