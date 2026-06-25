import { existsSync, mkdirSync, readFileSync, writeFileSync } from "node:fs";
import { dirname, resolve } from "node:path";

export type PptBoundingBoxPt = {
  left: number;
  top: number;
  width: number;
  height: number;
  rotation?: number;
};

export type PptShapeSnapshot = {
  pptShapeId: string;
  name: string;
  type: string;
  bboxPt: PptBoundingBoxPt;
  styleSnapshot?: {
    fill?: { color?: string; transparency?: number };
    line?: { color?: string; weight?: number };
    text?: { hasText?: boolean };
  };
};

export type MdprPptSelection = {
  schemaVersion: "mdpr-ppt-selection-v1";
  source: {
    tool: "mdpr-ppt";
    host: "powerpoint";
    presentationId: string;
    pptxSha256: string;
    capturedAt: string;
  };
  selection: {
    kind: "shape-selection" | "anchor-point" | "style-sample";
    userApproved: boolean;
    slideId: string;
    shapeCount: number;
  };
  mdprMapping: {
    sourceSha256: string;
    slideId: string;
    regionId?: string;
    blockIds: string[];
    mappingConfidence: number;
  };
  shapes: PptShapeSnapshot[];
  allowedUses: Array<"selection-context" | "pack-candidate" | "user-override-candidate">;
  disallowedUses: string[];
};

export type PptxObjectMapEntry = {
  slideId: string;
  layoutSlideId?: string;
  regionId: string;
  blockIds: string[];
  shapeName: string;
  objectKind: string;
  role?: string;
  editable: boolean;
};

export type MdprSelectionContext = {
  schemaVersion: "mdpr-selection-context-v1";
  source: {
    kind: "mdpr-ppt";
    sourceSha256: string;
  };
  slideId: string;
  overlappedBlocks: string[];
  overlappedRegions: string[];
  screenshotPath?: string;
  selectionPath?: string;
  userInstruction?: string;
};

export type MdprUserOverrideCandidate = {
  schemaVersion: "mdpr-user-override-candidate-v1";
  source: {
    kind: "mdpr-ppt";
    selectionRef: string;
    userApproved: true;
  };
  operations: Array<{
    op: "pinBlock" | "moveBlock" | "hideBlock";
    target: {
      slideId: string;
      blockId: string;
    };
    value?: {
      slot?: string;
    };
  }>;
  requiresApproval: true;
  constraints: {
    preferSemanticOverride: true;
    rawCoordinatesAreLastResort: true;
  };
};

export const defaultSharedSchemaNames = [
  "mdpr-ppt-selection.schema.json",
  "mdpr-ppt-pack-candidate.schema.json",
  "mdpr-user-override-candidate.schema.json",
  "mdpr-selection-context.schema.json",
  "mdpr-pptx-object-map.schema.json",
] as const;

export type ValidateSchemaSyncInput = {
  mdprPath?: string;
  localSchemaDir?: string;
  sharedSchemaNames?: readonly string[];
};

export type ValidateSchemaSyncResult = {
  status: "pass" | "fail";
  findings: string[];
  localSchemaDir: string;
  mdprSchemaDir: string;
};

export function runValidateSchemaSync(input: ValidateSchemaSyncInput = {}): ValidateSchemaSyncResult {
  const localSchemaDir = resolve(input.localSchemaDir ?? "schemas");
  const mdprSchemaDir = resolve(input.mdprPath ?? "../mdpr-skill/.cache/mdpr", "schemas");
  const findings: string[] = [];
  for (const schemaName of input.sharedSchemaNames ?? defaultSharedSchemaNames) {
    const localPath = resolve(localSchemaDir, schemaName);
    const mdprPath = resolve(mdprSchemaDir, schemaName);
    if (!existsSync(localPath)) {
      findings.push(`local shared schema is missing: ${localPath}`);
      continue;
    }
    if (!existsSync(mdprPath)) {
      findings.push(`MDPR shared schema is missing: ${mdprPath}`);
      continue;
    }
    const localSchema = readJsonObject(localPath, `local shared schema ${schemaName}`);
    const mdprSchema = readJsonObject(mdprPath, `MDPR shared schema ${schemaName}`);
    const localParseError = typeof localSchema.__parseError === "string" ? localSchema.__parseError : undefined;
    const mdprParseError = typeof mdprSchema.__parseError === "string" ? mdprSchema.__parseError : undefined;
    if (localParseError || mdprParseError) {
      if (localParseError) findings.push(localParseError);
      if (mdprParseError) findings.push(mdprParseError);
      continue;
    }
    if (stableStringify(localSchema) !== stableStringify(mdprSchema)) {
      findings.push(`shared schema drift: ${schemaName}`);
    }
  }
  return {
    status: findings.length === 0 ? "pass" : "fail",
    findings,
    localSchemaDir,
    mdprSchemaDir,
  };
}

export function validatePptSelection(selection: MdprPptSelection): string[] {
  const findings: string[] = [];
  if (selection.schemaVersion !== "mdpr-ppt-selection-v1") findings.push("schemaVersion must be mdpr-ppt-selection-v1");
  if (selection.source.tool !== "mdpr-ppt") findings.push("source.tool must be mdpr-ppt");
  if (selection.source.host !== "powerpoint") findings.push("source.host must be powerpoint");
  if (!/^[a-f0-9]{64}$/.test(selection.source.pptxSha256)) findings.push("source.pptxSha256 must be a 64 character lowercase hex digest");
  if (!/^[a-f0-9]{64}$/.test(selection.mdprMapping.sourceSha256)) findings.push("mdprMapping.sourceSha256 must be a 64 character lowercase hex digest");
  if (!selection.selection.userApproved) findings.push("selection.userApproved must be true before export");
  if (selection.selection.shapeCount !== selection.shapes.length) findings.push("selection.shapeCount must match shapes.length");
  if (!selection.allowedUses.includes("selection-context")) findings.push("allowedUses must include selection-context");
  if (!selection.disallowedUses.includes("agent-hint-final-decision")) findings.push("disallowedUses must include agent-hint-final-decision");
  if (selection.mdprMapping.mappingConfidence < 0 || selection.mdprMapping.mappingConfidence > 1) findings.push("mdprMapping.mappingConfidence must be between 0 and 1");
  return findings;
}

export function mapShapeToMdprObject(
  shape: PptShapeSnapshot,
  objectMap: PptxObjectMapEntry[],
): PptxObjectMapEntry | undefined {
  return objectMap.find((entry) => entry.shapeName === shape.name)
    ?? objectMap.find((entry) => shape.name.includes(entry.shapeName))
    ?? objectMap.find((entry) => entry.blockIds.some((blockId) => shape.name.includes(blockId)));
}

export function loadPptxObjectMap(input: unknown): PptxObjectMapEntry[] {
  if (Array.isArray(input)) return input.filter(isPptxObjectMapEntry);
  if (!input || typeof input !== "object") return [];
  const record = input as Record<string, unknown>;
  if (Array.isArray(record.objects)) return record.objects.filter(isPptxObjectMapEntry);
  if (Array.isArray(record.pptxObjects)) return record.pptxObjects.filter(isPptxObjectMapEntry);
  return [];
}

export function applyObjectMapToSelection(
  selection: MdprPptSelection,
  objectMap: PptxObjectMapEntry[],
): MdprPptSelection {
  const matches = selection.shapes
    .map((shape) => mapShapeToMdprObject(shape, objectMap))
    .filter((entry): entry is PptxObjectMapEntry => Boolean(entry));
  if (matches.length === 0) {
    return {
      ...selection,
      mdprMapping: {
        ...selection.mdprMapping,
        mappingConfidence: 0,
      },
    };
  }
  const slideIds = unique(matches.map((entry) => entry.slideId));
  const regionIds = unique(matches.map((entry) => entry.regionId));
  const blockIds = unique(matches.flatMap((entry) => entry.blockIds));
  return {
    ...selection,
    mdprMapping: {
      ...selection.mdprMapping,
      slideId: slideIds.length === 1 ? slideIds[0] : selection.mdprMapping.slideId,
      regionId: regionIds.length === 1 ? regionIds[0] : selection.mdprMapping.regionId,
      blockIds: blockIds.length > 0 ? blockIds : selection.mdprMapping.blockIds,
      mappingConfidence: Number((matches.length / Math.max(selection.shapes.length, 1)).toFixed(3)),
    },
  };
}

export function createSelectionContext(
  selection: MdprPptSelection,
  options: {
    screenshotPath?: string;
    selectionPath?: string;
    userInstruction?: string;
  } = {},
): MdprSelectionContext {
  const findings = validatePptSelection(selection);
  if (findings.length > 0) {
    throw new Error(`Invalid mdpr-ppt selection: ${findings.join("; ")}`);
  }
  return {
    schemaVersion: "mdpr-selection-context-v1",
    source: {
      kind: "mdpr-ppt",
      sourceSha256: selection.mdprMapping.sourceSha256,
    },
    slideId: selection.mdprMapping.slideId,
    overlappedBlocks: [...selection.mdprMapping.blockIds],
    overlappedRegions: selection.mdprMapping.regionId ? [selection.mdprMapping.regionId] : [],
    ...(options.screenshotPath ? { screenshotPath: options.screenshotPath } : {}),
    ...(options.selectionPath ? { selectionPath: options.selectionPath } : {}),
    ...(options.userInstruction ? { userInstruction: options.userInstruction } : {}),
  };
}

export function writeSelectionContext(path: string, context: MdprSelectionContext): void {
  mkdirSync(dirname(path), { recursive: true });
  writeFileSync(path, JSON.stringify(context, null, 2) + "\n", "utf-8");
}

export function createUserOverrideCandidate(
  selection: MdprPptSelection,
  options: {
    selectionRef: string;
    op: "pinBlock" | "moveBlock" | "hideBlock";
    blockId?: string;
    slot?: string;
  },
): MdprUserOverrideCandidate {
  const findings = validatePptSelection(selection);
  if (findings.length > 0) {
    throw new Error(`Invalid mdpr-ppt selection: ${findings.join("; ")}`);
  }
  const blockId = options.blockId ?? selection.mdprMapping.blockIds[0];
  if (!blockId) throw new Error("selection mapping must include at least one blockId");
  const slot = options.slot ?? selection.mdprMapping.regionId;
  if (options.op !== "hideBlock" && !slot) throw new Error(`${options.op} requires a mapped regionId or --slot`);
  return {
    schemaVersion: "mdpr-user-override-candidate-v1",
    source: {
      kind: "mdpr-ppt",
      selectionRef: options.selectionRef,
      userApproved: true,
    },
    operations: [{
      op: options.op,
      target: {
        slideId: selection.mdprMapping.slideId,
        blockId,
      },
      ...(options.op === "hideBlock" ? {} : { value: { slot } }),
    }],
    requiresApproval: true,
    constraints: {
      preferSemanticOverride: true,
      rawCoordinatesAreLastResort: true,
    },
  };
}

export function writeUserOverrideCandidate(path: string, candidate: MdprUserOverrideCandidate): void {
  mkdirSync(dirname(path), { recursive: true });
  writeFileSync(path, JSON.stringify(candidate, null, 2) + "\n", "utf-8");
}

function isPptxObjectMapEntry(value: unknown): value is PptxObjectMapEntry {
  if (!value || typeof value !== "object" || Array.isArray(value)) return false;
  const entry = value as Record<string, unknown>;
  return typeof entry.slideId === "string"
    && typeof entry.regionId === "string"
    && Array.isArray(entry.blockIds)
    && entry.blockIds.every((blockId) => typeof blockId === "string")
    && typeof entry.shapeName === "string"
    && entry.shapeName.startsWith("mdpr:")
    && typeof entry.objectKind === "string"
    && typeof entry.editable === "boolean";
}

function unique(values: string[]): string[] {
  return [...new Set(values.filter(Boolean))];
}

function readJsonObject(path: string, label: string): Record<string, unknown> {
  try {
    return JSON.parse(readFileSync(path, "utf-8")) as Record<string, unknown>;
  } catch (error) {
    return {
      __parseError: `${label} is invalid JSON: ${error instanceof Error ? error.message : String(error)}`,
    };
  }
}

function stableStringify(value: unknown): string {
  if (Array.isArray(value)) return `[${value.map(stableStringify).join(",")}]`;
  if (value && typeof value === "object") {
    const entries = Object.entries(value as Record<string, unknown>)
      .sort(([left], [right]) => left.localeCompare(right))
      .map(([key, item]) => `${JSON.stringify(key)}:${stableStringify(item)}`);
    return `{${entries.join(",")}}`;
  }
  return JSON.stringify(value);
}
