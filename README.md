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

### Use inside PowerPoint

The repository now includes a sideloadable PowerPoint task pane UI at
`packages/addin/taskpane/index.html` and an Office add-in manifest at
`packages/addin/manifest.xml`. The manifest adds an `MDPR` PowerPoint ribbon
tab with an `Inspect Selection` button, so the bridge is opened from inside
PowerPoint instead of from an external JSON tool.

If the MDPR tab is not visible, use the direct PowerPoint sideload command
first. This is the closest development workflow to a normal in-PowerPoint
add-in:

```powershell
npm run start:ppt
```

This registers the manifest for a debugging session and opens PowerPoint with
the add-in loaded. When you are done testing, stop that session so Office clears
the development registration:

```powershell
npm run stop:ppt
```

Prepare the local Windows add-in catalog and copy the manifest:

```powershell
npm run install:addin:windows
```

For the closest IguanaTex-like Windows setup, run PowerShell as an
administrator and let the helper create a shared-folder catalog and register it
under the current user's Office Trusted Add-in Catalogs registry key:

```powershell
npm run install:addin:windows:register
```

This runs the installer with `-TryShare -RegisterTrustCatalog`. If you want to
run the options manually:

```powershell
npm run install:addin:windows -- -TryShare -RegisterTrustCatalog
```

The script writes `install-next-steps.txt` under
`%LOCALAPPDATA%\mdpr-ppt\AddinCatalog`. It contains the manifest path and, when
`-TryShare` succeeds, the UNC catalog path to add in PowerPoint. If you run
without `-TryShare`, share that folder manually or rerun the command from an
elevated PowerShell session with `-TryShare`. If `-RegisterTrustCatalog`
succeeds, the Office trusted catalog entry is created for the current Windows
user; otherwise add the catalog manually in PowerPoint Trust Center.

Office task pane manifests use HTTPS `SourceLocation` URLs. Start the task pane
asset server with your trusted development certificate and key:

```bash
npm run serve:addin -- --cert ./certs/localhost.crt --key ./certs/localhost.key
```

Then restart PowerPoint. If the MDPR tab is not visible yet, register or
confirm the catalog in PowerPoint:

1. Open PowerPoint.
2. Go to `File > Options > Trust Center > Trust Center Settings`.
3. Open `Trusted Add-in Catalogs`.
4. Confirm the catalog URL from `install-next-steps.txt` is listed.
5. Enable `Show in Menu` for that catalog if it is not already enabled.
6. Restart PowerPoint.
7. If needed, go to `Home > Add-ins > Advanced`.
8. Choose `SHARED FOLDER`.
9. Add `mdpr-ppt` once.

In PowerPoint:

1. Open an MDPR-generated PPTX.
2. Select one or more MDPR-created shapes.
3. Open the `MDPR` tab and choose `Inspect Selection`.
4. Enter the MDPR source SHA-256 and PPTX SHA-256.
5. Click `Capture Selected Shapes`.
6. Click `Approve Selection`.
7. Copy one of the approved rail outputs:
   - `Copy Selection JSON`
   - `Copy Object Info`
   - `Copy Selection Context`
   - `Copy Override Candidate`

`Copy Object Info` is the fastest way to inspect the selected PowerPoint
object. It includes shape ids, names, types, bounds, style snapshots, and any
MDPR metadata inferred from shape names.

The task pane keeps raw PowerPoint `bboxPt`, fill, line, and text style data in
the approved `mdpr-ppt-selection-v1` output. `Copy Selection Context` strips
geometry and style before producing weak mdpr-skill review context.

If a selected shape name follows MDPR renderer metadata such as
`mdpr:slide-4:region-main:b12`, the task pane automatically fills the MDPR
slide, region, and block mapping fields.

### Use from the CLI

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

The Office.js integration is split into a small adapter and a static task pane.
`packages/addin/src/office/getSelectedShapes.ts` captures the current
PowerPoint selection through an Office.js-like context and converts each shape
through the shared snapshot mapper. `packages/addin/taskpane/taskpane.js`
provides the user-facing capture, approval, copy, selection-context, and
override-candidate controls without weakening the no-agent deterministic
boundary.
