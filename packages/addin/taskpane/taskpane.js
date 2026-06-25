/* global Office, PowerPoint */

const state = {
  snapshots: [],
  approved: false,
  lastSelection: undefined,
};

const ids = {
  statusBadge: "statusBadge",
  sourceSha256: "sourceSha256",
  pptxSha256: "pptxSha256",
  presentationId: "presentationId",
  slideId: "slideId",
  userInstruction: "userInstruction",
  transformKind: "transformKind",
  outputJson: "outputJson",
  objectInfo: "objectInfo",
  captureSelectedShapes: "captureSelectedShapes",
  approveSelection: "approveSelection",
  copySelectionJson: "copySelectionJson",
  copyObjectInfo: "copyObjectInfo",
  copySelectionContext: "copySelectionContext",
  copyOverrideCandidate: "copyOverrideCandidate",
  copyTransformCandidate: "copyTransformCandidate",
};

function element(id) {
  const value = document.getElementById(id);
  if (!value) throw new Error(`Missing UI element: ${id}`);
  return value;
}

function inputValue(id) {
  return element(id).value.trim();
}

function setStatus(label) {
  element(ids.statusBadge).textContent = label;
}

function writeOutput(value) {
  const text = typeof value === "string" ? value : JSON.stringify(value, null, 2);
  element(ids.outputJson).value = text;
  return text;
}

function writeObjectInfo(value) {
  const text = typeof value === "string" ? value : JSON.stringify(value, null, 2);
  element(ids.objectInfo).value = text;
  return text;
}

function assertSha256(value, label) {
  if (!/^[a-f0-9]{64}$/.test(value)) {
    throw new Error(`${label} must be a 64 character lowercase hex digest.`);
  }
}

function shapeToSnapshot(shape) {
  return {
    pptShapeId: shape.id || shape.name || "unknown-shape",
    name: shape.name || shape.id || "unknown-shape",
    type: shape.type || "unknown",
    bboxPt: {
      left: Number(shape.left || 0),
      top: Number(shape.top || 0),
      width: Number(shape.width || 0),
      height: Number(shape.height || 0),
      ...(typeof shape.rotation === "number" ? { rotation: shape.rotation } : {}),
    },
    styleSnapshot: {
      ...(shape.fill ? { fill: {
        color: shape.fill.color,
        transparency: shape.fill.transparency,
      } } : {}),
      ...(shape.line ? { line: {
        color: shape.line.color,
        weight: shape.line.weight,
      } } : {}),
      ...(shape.textFrame ? { text: {
        hasText: shape.textFrame.hasText,
      } } : {}),
    },
  };
}

function parseMdprShapeName(name) {
  const match = /^mdpr:([^:]+):([^:]+):(.+)$/.exec(name || "");
  if (!match) return undefined;
  return {
    slideId: match[1],
    regionId: match[2],
    blockIds: match[3].split(",").map((blockId) => blockId.trim()).filter(Boolean),
  };
}

function inferMdprMapping(snapshots) {
  const parsed = snapshots.map((snapshot) => parseMdprShapeName(snapshot.name)).filter(Boolean);
  if (parsed.length === 0) {
    return {
      slideId: "unmapped-slide",
      blockIds: [],
      mappingConfidence: 0,
    };
  }
  const slideIds = [...new Set(parsed.map((item) => item.slideId))];
  const regionIds = [...new Set(parsed.map((item) => item.regionId))];
  const blockIds = [...new Set(parsed.flatMap((item) => item.blockIds))];
  return {
    slideId: slideIds.length === 1 ? slideIds[0] : "mixed-slides",
    ...(regionIds.length === 1 ? { regionId: regionIds[0] } : {}),
    blockIds,
    mappingConfidence: Number((parsed.length / Math.max(snapshots.length, 1)).toFixed(3)),
  };
}

function buildObjectInfo() {
  const mapping = inferMdprMapping(state.snapshots);
  return {
    schemaVersion: "mdpr-ppt-object-info-v1",
    capturedShapes: state.snapshots.length,
    inferredMdprMapping: mapping,
    objects: state.snapshots.map((snapshot, index) => ({
      index,
      pptShapeId: snapshot.pptShapeId,
      name: snapshot.name,
      type: snapshot.type,
      bboxPt: snapshot.bboxPt,
      styleSnapshot: snapshot.styleSnapshot || {},
      mdprTag: parseMdprShapeName(snapshot.name) || null,
    })),
  };
}

async function captureSelectedShapesFromPowerPoint() {
  if (typeof PowerPoint === "undefined" || !PowerPoint.run) {
    throw new Error("PowerPoint runtime is not available. Sideload this taskpane in PowerPoint.");
  }
  return PowerPoint.run(async (context) => {
    const selectedShapes = context.presentation.getSelectedShapes();
    selectedShapes.load([
      "items/id",
      "items/name",
      "items/type",
      "items/left",
      "items/top",
      "items/width",
      "items/height",
      "items/rotation",
      "items/fill",
      "items/line",
      "items/textFrame",
    ]);
    await context.sync();
    return (selectedShapes.items || []).map(shapeToSnapshot);
  });
}

function buildSelection() {
  const sourceSha256 = inputValue(ids.sourceSha256);
  const pptxSha256 = inputValue(ids.pptxSha256);
  assertSha256(sourceSha256, "MDPR source SHA-256");
  assertSha256(pptxSha256, "PPTX SHA-256");
  if (state.snapshots.length === 0) throw new Error("Capture at least one selected shape first.");
  const inferredMapping = inferMdprMapping(state.snapshots);
  const selection = {
    schemaVersion: "mdpr-ppt-selection-v1",
    source: {
      tool: "mdpr-ppt",
      host: "powerpoint",
      presentationId: inputValue(ids.presentationId) || "powerpoint-presentation",
      pptxSha256,
      capturedAt: new Date().toISOString(),
    },
    selection: {
      kind: "shape-selection",
      userApproved: state.approved,
      slideId: inputValue(ids.slideId) || "ppt-slide-current",
      shapeCount: state.snapshots.length,
    },
    mdprMapping: {
      sourceSha256,
      ...inferredMapping,
    },
    shapes: state.snapshots,
    allowedUses: ["selection-context", "pack-candidate", "user-override-candidate"],
    disallowedUses: ["agent-hint-final-decision"],
  };
  state.lastSelection = selection;
  return selection;
}

function requireApprovedSelection() {
  const selection = state.lastSelection || buildSelection();
  if (!selection.selection.userApproved) {
    throw new Error("Approve the selection before exporting JSON.");
  }
  return selection;
}

function buildSelectionContext(selection) {
  const selectionContext = {
    schemaVersion: "mdpr-selection-context-v1",
    source: {
      kind: "mdpr-ppt",
      sourceSha256: selection.mdprMapping.sourceSha256,
    },
    slideId: selection.mdprMapping.slideId,
    overlappedBlocks: selection.mdprMapping.blockIds,
    overlappedRegions: selection.mdprMapping.regionId ? [selection.mdprMapping.regionId] : [],
    userInstruction: inputValue(ids.userInstruction) || undefined,
  };
  delete selectionContext.shapes;
  delete selectionContext.bboxPt;
  delete selectionContext.styleSnapshot;
  return selectionContext;
}

function buildOverrideCandidate(selection) {
  return {
    schemaVersion: "mdpr-user-override-candidate-v1",
    source: {
      kind: "mdpr-ppt",
      selectionRef: "taskpane-selection.json",
      userApproved: true,
    },
    operations: [{
      op: "pinBlock",
      target: {
        slideId: selection.mdprMapping.slideId,
        blockId: selection.mdprMapping.blockIds[0] || "unmapped-block",
      },
      value: selection.mdprMapping.regionId ? { slot: selection.mdprMapping.regionId } : undefined,
    }],
    requiresApproval: true,
    constraints: {
      preferSemanticOverride: true,
      rawCoordinatesAreLastResort: true,
    },
  };
}

function buildTransformCandidate(selection) {
  const transformKind = inputValue(ids.transformKind) || "callout";
  const blockIds = selection.mdprMapping.blockIds || [];
  const firstShape = selection.shapes[0] || {};
  return {
    schemaVersion: "mdpr-ppt-pack-candidate-v1",
    kind: "component-pack",
    source: {
      tool: "mdpr-ppt",
      selectionRef: "taskpane-selection.json",
      userApproved: true,
      pptxSha256: selection.source.pptxSha256,
    },
    tokens: {
      extractedFromSelection: true,
      styleSnapshot: firstShape.styleSnapshot || {},
    },
    components: [{
      id: `selected-${transformKind}`,
      kind: transformKind,
      sourceSlideId: selection.mdprMapping.slideId,
      sourceRegionId: selection.mdprMapping.regionId,
      sourceBlockIds: blockIds,
      slots: {
        content: blockIds.length ? blockIds : ["unmapped-block"],
      },
      editable: true,
      transformIntent: inputValue(ids.userInstruction) || `Convert selected object to MDPR ${transformKind}.`,
    }],
    requiresApproval: true,
    constraints: {
      noRawUseInAgentHints: true,
      requiresDesignLockUpdate: true,
    },
  };
}

async function copyText(value) {
  if (navigator.clipboard?.writeText) {
    await navigator.clipboard.writeText(value);
    return;
  }
  element(ids.outputJson).select();
  document.execCommand("copy");
}

async function runAction(action) {
  try {
    await action();
  } catch (error) {
    setStatus("Error");
    writeOutput(error instanceof Error ? error.message : String(error));
  }
}

function bindUi() {
  element(ids.captureSelectedShapes).addEventListener("click", () => runAction(async () => {
    state.snapshots = await captureSelectedShapesFromPowerPoint();
    state.approved = false;
    const objectInfo = buildObjectInfo();
    writeObjectInfo(objectInfo);
    writeOutput(objectInfo);
    setStatus(state.snapshots.length > 0 ? "Captured" : "Empty");
  }));

  element(ids.approveSelection).addEventListener("click", () => runAction(async () => {
    state.approved = true;
    const selection = buildSelection();
    writeOutput(selection);
    setStatus("Approved");
  }));

  element(ids.copySelectionJson).addEventListener("click", () => runAction(async () => {
    const text = writeOutput(requireApprovedSelection());
    await copyText(text);
    setStatus("Copied");
  }));

  element(ids.copyObjectInfo).addEventListener("click", () => runAction(async () => {
    if (state.snapshots.length === 0) throw new Error("Capture at least one selected shape first.");
    const text = writeObjectInfo(buildObjectInfo());
    writeOutput(JSON.parse(text));
    await copyText(text);
    setStatus("Copied");
  }));

  element(ids.copySelectionContext).addEventListener("click", () => runAction(async () => {
    const text = writeOutput(buildSelectionContext(requireApprovedSelection()));
    await copyText(text);
    setStatus("Copied");
  }));

  element(ids.copyOverrideCandidate).addEventListener("click", () => runAction(async () => {
    const text = writeOutput(buildOverrideCandidate(requireApprovedSelection()));
    await copyText(text);
    setStatus("Copied");
  }));

  element(ids.copyTransformCandidate).addEventListener("click", () => runAction(async () => {
    const text = writeOutput(buildTransformCandidate(requireApprovedSelection()));
    await copyText(text);
    setStatus("Copied");
  }));
}

if (typeof Office !== "undefined" && Office.onReady) {
  Office.onReady(() => {
    bindUi();
    setStatus("Ready");
  });
} else {
  window.addEventListener("DOMContentLoaded", () => {
    bindUi();
    setStatus("Preview");
  });
}
