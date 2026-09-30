# MVP format support

## DOCX

Supported slot: Word content control whose tag is `ot.slot:<id>`. The filler replaces text nodes inside the content control while preserving the surrounding OOXML and the first text run's formatting.

MVP limitations: plain text only; no automatic repeated sections, rich text, lists, field refresh, TOC refresh, tracked changes, or arbitrary paragraph reconstruction. Use Microsoft Word for final rendering when Word fidelity matters.

## PPTX

Supported slot: a text-bearing shape whose shape name is `ot.slot:<id>`. Rename shapes through the Selection Pane. The filler replaces the shape's existing text nodes and keeps its shape geometry and existing first run formatting. A later layout-fit pass may widen or increase the height of the shape, uniformly reduce its font, or tighten spacing only when those actions and bounds are registered in the slot contract.

MVP limitations: plain text only; no automatic slide duplication, SmartArt, chart-data editing, animation editing, or automatic application of fit-policy geometry. Use a PowerPoint-capable authoring tool for the registered fit pass and verify by rendering. Register compact layout variants for changes that exceed a slot's adaptive bounds.

The v2 companion `pptx_style_guard.py` extracts an OOXML visual handbook and compares actual output against source. The filler requires `--visual-book` and runs a structural style audit before writing its output. All subsequent edits must run the final audit again. The helper checks preserved graphics and registered donor clones conservatively; it is not a native renderer, theme RGB resolver, installed-font detector, or general chart/SmartArt editor. See `visual-contract.md` for the score definition and remaining visual gates.

## XLSX

Supported slot: a workbook defined name `ot_slot_<id>` or `ot.slot.<id>` that points to exactly one cell. Dots in the id are represented as double underscores in the underscore form, for example `ot_slot_resume__name` becomes `resume.name`.

The filler writes an inline string and removes a formula from that target cell. Therefore never point an input slot at a formula cell.

MVP limitations: one-cell plain-text inputs only; no repeating tables, merged-range expansion, pivot refresh, chart update, macro execution, formula recalculation, or external-link refresh.

## Unsupported input

Legacy `.doc`, `.ppt`, and `.xls`, PDFs used as editable templates, password-protected packages, encrypted packages, and arbitrary unmarked regions are outside the MVP.
