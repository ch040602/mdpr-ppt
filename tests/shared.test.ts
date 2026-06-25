import assert from "node:assert/strict";
import { existsSync, mkdtempSync, rmSync, writeFileSync, readFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import test from "node:test";
import { runCli } from "../packages/cli/src/main";
import { captureSelectedShapeSnapshots } from "../packages/addin/src/office/getSelectedShapes";
import {
  applyObjectMapToSelection,
  createSelectionContext,
  createUserOverrideCandidate,
  loadPptxObjectMap,
  mapShapeToMdprObject,
  validatePptSelection,
  writeSelectionContext,
  writeUserOverrideCandidate,
  runValidateSchemaSync,
  defaultSharedSchemaNames,
  type MdprPptSelection,
  type PptxObjectMapEntry,
} from "../packages/shared/src/index";

const sourceSha256 = "a".repeat(64);

test("captureSelectedShapeSnapshots loads selected PowerPoint shapes through an Office.js-like context", async () => {
  const loadedProperties: Array<string | string[]> = [];
  let syncCalled = false;
  const snapshots = await captureSelectedShapeSnapshots({
    presentation: {
      getSelectedShapes: () => ({
        load: (properties) => loadedProperties.push(properties),
        items: [{
          id: "shape-1",
          name: "mdpr:slide-4:region-main:b12",
          type: "geometricShape",
          left: 184,
          top: 126,
          width: 420,
          height: 180,
          rotation: 0,
          fill: { color: "#F5F7FB", transparency: 0 },
          line: { color: "#C9CED8", weight: 1 },
          textFrame: { hasText: true },
        }],
      }),
    },
    sync: async () => {
      syncCalled = true;
    },
  });

  assert.equal(syncCalled, true);
  assert.equal(loadedProperties.length, 1);
  assert.equal(snapshots[0].pptShapeId, "shape-1");
  assert.equal(snapshots[0].name, "mdpr:slide-4:region-main:b12");
  assert.deepEqual(snapshots[0].bboxPt, { left: 184, top: 126, width: 420, height: 180, rotation: 0 });
});

test("PowerPoint add-in manifest and taskpane expose the approved selection rail UI", () => {
  const manifest = readFileSync("packages/addin/manifest.xml", "utf-8");
  const taskpaneHtml = readFileSync("packages/addin/taskpane/index.html", "utf-8");
  const taskpaneJs = readFileSync("packages/addin/taskpane/taskpane.js", "utf-8");

  assert.match(manifest, /<Hosts>\s*<Host Name="Presentation"\/>\s*<\/Hosts>/);
  assert.match(manifest, /PowerPointApi/);
  assert.match(manifest, /taskpane\/index\.html/);
  assert.match(manifest, /https:\/\/localhost:3000/);
  assert.match(manifest, /assets\/icon\.svg/);
  assert.match(taskpaneHtml, /id="captureSelectedShapes"/);
  assert.match(taskpaneHtml, /id="approveSelection"/);
  assert.match(taskpaneHtml, /id="copySelectionJson"/);
  assert.match(taskpaneHtml, /id="copySelectionContext"/);
  assert.match(taskpaneHtml, /id="copyOverrideCandidate"/);
  assert.match(taskpaneHtml, /approved selection rail/i);
  assert.match(taskpaneJs, /getSelectedShapes/);
  assert.match(taskpaneJs, /parseMdprShapeName/);
  assert.match(taskpaneJs, /agent-hint-final-decision/);
  assert.match(taskpaneJs, /delete selectionContext\.shapes/);
  assert.doesNotMatch(taskpaneJs, /fetch\(/);
  assert.equal(existsSync("packages/addin/assets/icon.svg"), true);
  assert.equal(existsSync("scripts/serve-addin.mjs"), true);
});

test("validatePptSelection requires user approval and allowed use rails", () => {
  const selection = fixtureSelection();
  assert.deepEqual(validatePptSelection(selection), []);

  const withoutApproval = {
    ...selection,
    selection: { ...selection.selection, userApproved: false },
  };
  assert.match(validatePptSelection(withoutApproval).join("\n"), /userApproved/);
});

test("mapShapeToMdprObject resolves MDPR block mapping from shape metadata", () => {
  const objectMap: PptxObjectMapEntry[] = [{
    slideId: "slide-4",
    regionId: "region-main",
    blockIds: ["b12", "b13"],
    shapeName: "mdpr:slide-4:region-main:b12",
    objectKind: "native-text",
    role: "body",
    editable: true,
  }];

  const mapped = mapShapeToMdprObject(fixtureSelection().shapes[0], objectMap);
  assert.equal(mapped?.slideId, "slide-4");
  assert.deepEqual(mapped?.blockIds, ["b12", "b13"]);
});

test("loadPptxObjectMap accepts MDPR manifests and object-map schema files", () => {
  const entry = fixtureObjectMap()[0];
  assert.deepEqual(loadPptxObjectMap({ schemaVersion: "mdpr-pptx-object-map-v1", objects: [entry] }), [entry]);
  assert.deepEqual(loadPptxObjectMap({ pptxObjects: [entry] }), [entry]);
});

test("applyObjectMapToSelection refreshes approved selection mapping from shape names", () => {
  const selection = {
    ...fixtureSelection(),
    mdprMapping: {
      sourceSha256,
      slideId: "stale-slide",
      regionId: "stale-region",
      blockIds: ["stale-block"],
      mappingConfidence: 0.1,
    },
  };
  const mapped = applyObjectMapToSelection(selection, fixtureObjectMap());
  assert.equal(mapped.mdprMapping.slideId, "slide-4");
  assert.equal(mapped.mdprMapping.regionId, "region-main");
  assert.deepEqual(mapped.mdprMapping.blockIds, ["b12", "b13"]);
  assert.equal(mapped.mdprMapping.mappingConfidence, 1);
});

test("createSelectionContext exports weak review input without style coordinates", () => {
  const context = createSelectionContext(fixtureSelection(), {
    screenshotPath: ".mdpresent/review/slide-4-selection.png",
    userInstruction: "Keep the table and caption together.",
  });

  assert.equal(context.schemaVersion, "mdpr-selection-context-v1");
  assert.deepEqual(context.overlappedBlocks, ["b12", "b13"]);
  assert.equal(JSON.stringify(context).includes("bboxPt"), false);
  assert.equal(JSON.stringify(context).includes("#F5F7FB"), false);
});

test("writeSelectionContext writes deterministic JSON", () => {
  const workDir = mkdtempSync(join(tmpdir(), "mdpr-ppt-"));
  try {
    const outPath = join(workDir, "selection-context.json");
    writeSelectionContext(outPath, createSelectionContext(fixtureSelection()));
    const parsed = JSON.parse(readFileSync(outPath, "utf-8"));
    assert.equal(parsed.schemaVersion, "mdpr-selection-context-v1");
  } finally {
    rmSync(workDir, { recursive: true, force: true });
  }
});

test("createUserOverrideCandidate exports semantic pinBlock without raw geometry", () => {
  const candidate = createUserOverrideCandidate(fixtureSelection(), {
    selectionRef: ".mdpresent/ppt/selection.mapped.json",
    op: "pinBlock",
  });
  assert.equal(candidate.schemaVersion, "mdpr-user-override-candidate-v1");
  assert.equal(candidate.source.userApproved, true);
  assert.deepEqual(candidate.operations, [{
    op: "pinBlock",
    target: {
      slideId: "slide-4",
      blockId: "b12",
    },
    value: {
      slot: "region-main",
    },
  }]);
  assert.equal(JSON.stringify(candidate).includes("bboxPt"), false);
  assert.equal(JSON.stringify(candidate).includes("left"), false);
});

test("writeUserOverrideCandidate writes deterministic approved-rail JSON", () => {
  const workDir = mkdtempSync(join(tmpdir(), "mdpr-ppt-override-"));
  try {
    const outPath = join(workDir, "override.candidate.json");
    writeUserOverrideCandidate(outPath, createUserOverrideCandidate(fixtureSelection(), {
      selectionRef: ".mdpresent/ppt/selection.mapped.json",
      op: "pinBlock",
    }));
    const parsed = JSON.parse(readFileSync(outPath, "utf-8"));
    assert.equal(parsed.schemaVersion, "mdpr-user-override-candidate-v1");
  } finally {
    rmSync(workDir, { recursive: true, force: true });
  }
});

test("CLI map-selection writes a selection remapped from an MDPR manifest object map", () => {
  const workDir = mkdtempSync(join(tmpdir(), "mdpr-ppt-map-selection-"));
  try {
    const selectionPath = join(workDir, "selection.json");
    const manifestPath = join(workDir, "manifest.json");
    const outPath = join(workDir, "mapped", "selection.json");
    writeFileSync(selectionPath, JSON.stringify({
      ...fixtureSelection(),
      mdprMapping: {
        sourceSha256,
        slideId: "stale-slide",
        regionId: "stale-region",
        blockIds: ["stale-block"],
        mappingConfidence: 0.1,
      },
    }), "utf-8");
    writeFileSync(manifestPath, JSON.stringify({ pptxObjects: fixtureObjectMap() }), "utf-8");
    const stdout: string[] = [];
    const exitCode = runCli(["map-selection", selectionPath, "--object-map", manifestPath, "--out", outPath], {
      stdout: (value) => stdout.push(value),
      stderr: (value) => stdout.push(value),
    });
    assert.equal(exitCode, 0);
    assert.match(stdout.join("\n"), /"mappingConfidence": 1/);
    const mapped = JSON.parse(readFileSync(outPath, "utf-8")) as MdprPptSelection;
    assert.equal(mapped.mdprMapping.slideId, "slide-4");
    assert.deepEqual(mapped.mdprMapping.blockIds, ["b12", "b13"]);
  } finally {
    rmSync(workDir, { recursive: true, force: true });
  }
});

test("CLI map-selection fails when no selected shape maps to MDPR output objects", () => {
  const workDir = mkdtempSync(join(tmpdir(), "mdpr-ppt-map-selection-miss-"));
  try {
    const selectionPath = join(workDir, "selection.json");
    const manifestPath = join(workDir, "manifest.json");
    const outPath = join(workDir, "mapped", "selection.json");
    writeFileSync(selectionPath, JSON.stringify({
      ...fixtureSelection(),
      shapes: [{ ...fixtureSelection().shapes[0], name: "manual-shape-without-mdpr-tag" }],
    }), "utf-8");
    writeFileSync(manifestPath, JSON.stringify({ pptxObjects: fixtureObjectMap() }), "utf-8");
    const exitCode = runCli(["map-selection", selectionPath, "--object-map", manifestPath, "--out", outPath], {
      stdout: () => undefined,
      stderr: () => undefined,
    });
    assert.equal(exitCode, 1);
    const mapped = JSON.parse(readFileSync(outPath, "utf-8")) as MdprPptSelection;
    assert.equal(mapped.mdprMapping.mappingConfidence, 0);
  } finally {
    rmSync(workDir, { recursive: true, force: true });
  }
});

test("CLI override-candidate writes a semantic approved-rail override candidate", () => {
  const workDir = mkdtempSync(join(tmpdir(), "mdpr-ppt-override-cli-"));
  try {
    const selectionPath = join(workDir, "selection.json");
    const outPath = join(workDir, "override.candidate.json");
    writeFileSync(selectionPath, JSON.stringify(fixtureSelection()), "utf-8");
    const exitCode = runCli(["override-candidate", selectionPath, "--op", "pinBlock", "--out", outPath], {
      stdout: () => undefined,
      stderr: () => undefined,
    });
    assert.equal(exitCode, 0);
    const candidate = JSON.parse(readFileSync(outPath, "utf-8"));
    assert.equal(candidate.operations[0].op, "pinBlock");
    assert.equal(candidate.operations[0].target.blockId, "b12");
    assert.equal(candidate.operations[0].value.slot, "region-main");
  } finally {
    rmSync(workDir, { recursive: true, force: true });
  }
});

test("validate schema sync checks all MDPR-owned bridge schemas", () => {
  assert.deepEqual([...defaultSharedSchemaNames].sort(), [
    "mdpr-ppt-selection.schema.json",
    "mdpr-ppt-pack-candidate.schema.json",
    "mdpr-user-override-candidate.schema.json",
    "mdpr-selection-context.schema.json",
    "mdpr-pptx-object-map.schema.json",
  ].sort());

  const mdprPath = join("..", "mdpr-skill", ".cache", "mdpr");
  if (!existsSync(join(mdprPath, "schemas", "mdpr-ppt-selection.schema.json"))) return;

  const result = runValidateSchemaSync({ mdprPath });
  assert.equal(result.status, "pass", result.findings.join("\n"));
});

test("CLI validate-schema-sync passes against the local MDPR source-of-truth schemas", () => {
  const mdprPath = join("..", "mdpr-skill", ".cache", "mdpr");
  if (!existsSync(join(mdprPath, "schemas", "mdpr-ppt-selection.schema.json"))) return;

  const stdout: string[] = [];
  const exitCode = runCli(["validate-schema-sync", "--mdpr-path", mdprPath], {
    stdout: (value) => stdout.push(value),
    stderr: (value) => stdout.push(value),
  });
  assert.equal(exitCode, 0, stdout.join("\n"));
  assert.match(stdout.join("\n"), /"status": "pass"/);
});

function fixtureSelection(): MdprPptSelection {
  return {
    schemaVersion: "mdpr-ppt-selection-v1",
    source: {
      tool: "mdpr-ppt",
      host: "powerpoint",
      presentationId: "ppt-pres-id",
      pptxSha256: "b".repeat(64),
      capturedAt: "2026-06-25T00:00:00.000Z",
    },
    selection: {
      kind: "shape-selection",
      userApproved: true,
      slideId: "ppt-slide-id",
      shapeCount: 1,
    },
    mdprMapping: {
      sourceSha256,
      slideId: "slide-4",
      regionId: "region-main",
      blockIds: ["b12", "b13"],
      mappingConfidence: 0.91,
    },
    shapes: [{
      pptShapeId: "shape-id",
      name: "mdpr:slide-4:region-main:b12",
      type: "geometricShape",
      bboxPt: {
        left: 184,
        top: 126,
        width: 420,
        height: 180,
        rotation: 0,
      },
      styleSnapshot: {
        fill: { color: "#F5F7FB", transparency: 0 },
        line: { color: "#C9CED8", weight: 1 },
        text: { hasText: true },
      },
    }],
    allowedUses: ["selection-context", "pack-candidate", "user-override-candidate"],
    disallowedUses: ["agent-hint-final-decision"],
  };
}

function fixtureObjectMap(): PptxObjectMapEntry[] {
  return [{
    slideId: "slide-4",
    regionId: "region-main",
    blockIds: ["b12", "b13"],
    shapeName: "mdpr:slide-4:region-main:b12",
    objectKind: "native-text",
    role: "body",
    editable: true,
  }];
}
