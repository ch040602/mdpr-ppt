# mdpr-ppt

`mdpr-ppt` is the PowerPoint bridge for MDPR. It captures user-selected
PowerPoint shapes, anchors, and style samples as approved JSON evidence that
MDPR and `mdpr-skill` can reference safely.

It is not an agent layout tool. It does not ask an LLM to invent coordinates,
colors, recipes, variants, z-order, or exact renderer objects.

## Role Boundary

| Repository | Responsibility |
| --- | --- |
| MDPR | Deterministic Markdown-to-editable-PPTX runtime, schema source of truth, final layout/style/render decisions |
| mdpr-skill | Optional hint, review, eval, design-analysis, and gate companion |
| mdpr-ppt | User-approved PowerPoint selection capture and JSON bridge |

## Rails

| Rail | Coordinates and style values | Consumer |
| --- | --- | --- |
| `mdpr-ppt-selection-v1` | Allowed because the user selected the PowerPoint object | `mdpr-ppt`, MDPR pack or override import |
| `mdpr-selection-context-v1` | Not allowed; contains only source/block/region context | `mdpr-skill` hint and review inputs |
| `mdpr-ppt-pack-candidate-v1` | Allowed after user approval and provenance checks | MDPR pack validation/import |
| `mdpr-user-override-candidate-v1` | Limited, approval-bound, semantic override preferred | MDPR override validation/apply |

## Repository Structure

```text
schemas/                 Synced MDPR source-of-truth bridge schema copies
packages/shared/         Shared TypeScript types, validation, mapping, and selection-context export
packages/cli/            Local validation and conversion commands
packages/addin/          Office Add-in boundary for selected shape capture
packages/ooxml-extractor/ Future fallback for grouped shapes and table internals
tests/                   Runtime tests for rail boundaries and mapping helpers
```

MDPR owns the schema source of truth. This repository keeps synced copies for
the bridge contracts it emits or consumes: `mdpr-ppt-selection`,
`mdpr-ppt-pack-candidate`, `mdpr-user-override-candidate`,
`mdpr-selection-context`, and `mdpr-pptx-object-map`.

## Usage

Validate a captured selection:

```bash
mdpr-ppt validate-selection .mdpresent/ppt/selection.json
```

Refresh a captured selection against an MDPR manifest or object-map export:

```bash
mdpr-ppt map-selection .mdpresent/ppt/selection.json \
  --object-map .mdpresent/mdpresent-manifest.json \
  --out .mdpresent/ppt/selection.mapped.json
```

Convert the approved selection into a weak selection context for `mdpr-skill`:

```bash
mdpr-ppt selection-context .mdpresent/ppt/selection.mapped.json \
  --out .mdpresent/review/selection-context.json
```

Export a semantic override candidate from the approved selection:

```bash
mdpr-ppt override-candidate .mdpresent/ppt/selection.mapped.json \
  --op pinBlock \
  --out .mdpresent/proposals/override.candidate.json
```

The generated selection context intentionally removes raw PowerPoint geometry
and style snapshots. Those values stay in the user-approved selection or pack
rail, not in agent hints.

Check synced schema copies against a local MDPR checkout:

```bash
mdpr-ppt validate-schema-sync --mdpr-path ../mdpr-skill/.cache/mdpr
```

## Development

```bash
npm install
npm run validate
```

The first milestone keeps Office.js integration behind a small adapter
boundary. `packages/addin/src/office/getSelectedShapes.ts` captures the current
PowerPoint selection through an Office.js-like context and converts each shape
through the shared snapshot mapper. Future milestones should add the taskpane
UI, approval panels, and OOXML fallback extraction without weakening the
no-agent deterministic boundary.
