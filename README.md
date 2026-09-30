<p align="right"><strong>English</strong> | <a href="README.zh-CN.md">简体中文</a></p>

# FormKeep · 版如初

### Fit content into your Office templates—without losing the original design

An agent skill for applying business content to existing PPTX, DOCX, and XLSX templates while preserving their layout, visual style, and editable structure.

![FormKeep: turn Excel data and a user-supplied PPT template into an on-brand analysis deck](docs/images/excel-to-ppt-hero.png)

> **MVP / experimental.** The hero image is a promotional composite based on a real template-fitting workflow, using entirely fictional business data. Directly rendered comparison pages appear below. FormKeep is a Skill plus helper scripts, not a standalone visual editor, and it does not promise that every arbitrary template will be free of layout issues.

## A real office workflow

FormKeep is designed for format-sensitive office deliverables: business reporting, presentations, resumes, formal documents, and data communication. It fits new content into a user-supplied Office template and supports PPTX, DOCX, and XLSX. The Excel-analysis-to-PPT workflow below is one representative example, not the boundary of the product.

> "Analyze this regional and channel performance workbook, then present the findings in our company PowerPoint template. Keep the template's colors, typography, cards, and icons."

FormKeep gives an agent a template-fitting contract: establish the facts and conclusions, register the editable regions, organize content within calibrated slot capacity, and finally check style and layout changes. Excel analysis and native Office rendering depend on tools available in the execution environment.

| User-supplied template | Template-fitted analysis page |
| --- | --- |
| ![Original template overview](docs/images/template-overview.png) | ![Generated analysis overview](docs/images/report-overview.png) |

These are directly rendered pages, not redraws from the promotional image. The output retains the top-left ornament, red-and-gold palette, four-card structure, and visual hierarchy. All business names and figures are fictional. These screenshots are not evidence of native PowerPoint rendering acceptance.

<details>
<summary>Another page: the original card and icon language reused for channel comparison</summary>

![Channel comparison example](docs/images/report-channels.png)

</details>

The demonstration figures are traceable in the [English case study](docs/case-study.en.md): budget 50.0 million, actual revenue 52.0 million, and attainment 104.0%. The original business workbook, complete template, and customer report are not published.

## How it works

- **Template visual handbook:** extract PPTX color references, fonts, object geometry, z-order, anchors, and asset fingerprints before authoring, so the template's visual language constrains the content.
- **Explicit content slots:** connect structure to the template with stable object identifiers. Unmarked templates must be registered before use; replacement regions are never guessed blindly.
- **Bidirectional capacity constraints:** define minimum content, recommended and hard limits, line limits, and approved layout states per slot. Character count is only a preflight signal; font metrics and actual rendering still require calibration.
- **Layout-first fitting:** within approved bounds, widen a text box before applying uniform font reduction or a compact layout state. Facts must not be rewritten merely to fit, and body copy must never be silently truncated.
- **Style checks and delivery gates:** every new element needs a same-role style donor. Off-palette content, changes to protected elements, or missing render evidence cannot be hidden by a high aggregate score.

## Scope and limitations

| Capability | Current scope |
| --- | --- |
| PPTX | Plain-text filling of named shape slots; visual handbook extraction and structural style audit |
| DOCX | Plain-text filling of content-control slots; pagination and layout still require validation in the target editor |
| XLSX | Plain-text filling of named slots that point to individual cells; not a general formula or chart writer |
| Bidirectional visual mapping | Provides a structure contract and implementation guidance; does not include a complete visual editor |
| Font substitution, box expansion, font reduction | The Skill requires these adaptations to be recorded and verified; current helper scripts do not automate every adjustment |
| Excel analysis and PowerPoint charts | Performed by the agent with external spreadsheet and presentation tools, not by the filling CLI alone |
| Native rendering | Requires a target engine such as PowerPoint, Word, or Excel in the user's environment; no renderer is bundled |

**What does 90/100 mean?** It is a configurable structural-style fidelity gate: palette 25, typography 20, layout 25, assets 20, and shape styling 10. Every slide and the overall deck must pass, with zero red-line violations. It is not a pixel-similarity score, a perceptual-similarity score, or a claim of cross-model reliability. Without native visual review, the audit returns `needs-native-render-review`; the output must not be described as guaranteed layout-safe.

Theme inheritance, complex group transforms, and font availability may still require human review or validation in the target engine. See [visual-contract.md](references/visual-contract.md) for the detailed rules.

## Usage

Place this repository in an agent environment that supports `SKILL.md`, using `formkeep` as the Skill folder name. Preserve the relative structure of `scripts/`, `references/`, `assets/`, and `agents/`. Discovery and installation details vary by host.

Example prompt:

```text
Use $formkeep to analyze the Excel performance data I uploaded
and create a report using my PowerPoint template.
Build the template visual handbook and content slots first, then confirm
capacity and adaptation limits. Reuse the template's palette, typography,
icons, and layout. If the content cannot fit safely, report the conflict;
do not truncate facts or redesign the template without approval.
```

### Helper CLI

The helper scripts use the Python 3.10+ standard library and require no model API key. The scripts themselves do not make network requests. The agent host may still process uploaded content according to its own configuration, so this does not imply that the entire workflow is offline.

```bash
# Generate a slot inventory for an already marked template.
# Capacity calibration is not complete at this stage.
python3 scripts/office_template.py inspect TEMPLATE.pptx --to template-spec.json

# Extract the visual handbook as JSON and Markdown.
python3 scripts/pptx_style_guard.py extract TEMPLATE.pptx --to template-visual.json

# Validate and fill after template registration and calibration.
python3 scripts/office_template.py validate template-spec.json bindings.json
python3 scripts/office_template.py fill TEMPLATE.pptx template-spec.json bindings.json draft.pptx --visual-book template-visual.json

# Audit the final file after all fitting, chart editing, and export work.
python3 scripts/pptx_style_guard.py audit TEMPLATE.pptx final.pptx --plan style-plan.json --review render-review.json --to style-audit.json
```

Prepare `style-plan.json` and `render-review.json` according to the [visual contract](references/visual-contract.md). Passing evidence must not be fabricated, and native-review records are bound to the final file hash. The examples in `assets/` have not been calibrated against a real template and must not be treated as delivery-ready configuration.

## Development and validation

```bash
python3 -m unittest discover -s scripts -p 'test_*.py' -v
```

The tests cover slot resolution, capacity constraints, and critical rejection paths in the style audit. They do not replace Office rendering or cross-model behavioral evaluation.

Key entry points: [Skill instructions](SKILL.md), [format support](references/format-support.md), [structure specification](references/template-spec.md), and [delivery gates](references/quality-gates.md).

## Privacy and publication scope

This repository contains only the Skill, helper scripts, fictional examples, and content-reviewed static demonstration images. Do not publish customer data, credentials, unauthorized templates, font files, or execution logs containing personal paths in issues, commits, or screenshots. Rights to complete templates and related visual assets remain with their respective owners; a public screenshot does not grant redistribution rights.

No open-source license is currently included. Do not interpret this public repository as a grant of permission to use or redistribute its contents.
