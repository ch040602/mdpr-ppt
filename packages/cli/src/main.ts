import { mkdirSync, readFileSync, writeFileSync } from "node:fs";
import { dirname } from "node:path";
import { pathToFileURL } from "node:url";
import {
  applyObjectMapToSelection,
  createSelectionContext,
  createUserOverrideCandidate,
  loadPptxObjectMap,
  runValidateSchemaSync,
  validatePptSelection,
  writeSelectionContext,
  writeUserOverrideCandidate,
  type MdprPptSelection,
} from "../../shared/src/index";

export type CliIo = {
  stdout: (value: string) => void;
  stderr: (value: string) => void;
};

const defaultIo: CliIo = {
  stdout: (value) => process.stdout.write(`${value}\n`),
  stderr: (value) => process.stderr.write(`${value}\n`),
};

export function runCli(argv: string[], io: CliIo = defaultIo): number {
  try {
    const args = [...argv];
    const command = args.shift();
    if (!command || command === "--help" || command === "help") {
      io.stdout(helpText());
      return 0;
    }
    if (command === "validate-selection") return runValidateSelection(args, io);
    if (command === "map-selection") return runMapSelection(args, io);
    if (command === "selection-context") return runSelectionContext(args, io);
    if (command === "override-candidate") return runOverrideCandidate(args, io);
    if (command === "validate-schema-sync") return runValidateSchemaSyncCommand(args, io);
    io.stderr(`Unknown command: ${command}`);
    io.stderr(helpText());
    return 2;
  } catch (error) {
    io.stderr(error instanceof Error ? error.message : String(error));
    return 1;
  }
}

function helpText(): string {
  return [
    "mdpr-ppt",
    "",
    "Commands:",
    "  validate-selection <mdpr-ppt-selection.json>",
    "  map-selection <mdpr-ppt-selection.json> --object-map <mdpr-manifest-or-object-map.json> --out <mdpr-ppt-selection.json>",
    "  selection-context <mdpr-ppt-selection.json> --out <mdpr-selection-context.json>",
    "  override-candidate <mdpr-ppt-selection.json> --op pinBlock|moveBlock|hideBlock --out <mdpr-user-override-candidate.json>",
    "  validate-schema-sync --mdpr-path <MdPr>",
  ].join("\n");
}

function runValidateSelection(args: string[], io: CliIo): number {
  const path = args.shift();
  if (!path || path.startsWith("--")) throw new Error("validate-selection requires a selection JSON path");
  const selection = readSelection(path);
  const findings = validatePptSelection(selection);
  io.stdout(JSON.stringify({ status: findings.length === 0 ? "pass" : "fail", findings }, null, 2));
  return findings.length === 0 ? 0 : 1;
}

function runMapSelection(args: string[], io: CliIo): number {
  const path = args.shift();
  if (!path || path.startsWith("--")) throw new Error("map-selection requires a selection JSON path");
  const options = parseOptions(args);
  const objectMapPath = requireOption(options, "object-map");
  const outPath = requireOption(options, "out");
  const mapped = applyObjectMapToSelection(readSelection(path), loadPptxObjectMap(readJson(objectMapPath)));
  writeJson(outPath, mapped);
  io.stdout(JSON.stringify({
    status: mapped.mdprMapping.mappingConfidence > 0 ? "pass" : "fail",
    out: outPath,
    mappingConfidence: mapped.mdprMapping.mappingConfidence,
    blockIds: mapped.mdprMapping.blockIds,
  }, null, 2));
  return mapped.mdprMapping.mappingConfidence > 0 ? 0 : 1;
}

function runSelectionContext(args: string[], io: CliIo): number {
  const path = args.shift();
  if (!path || path.startsWith("--")) throw new Error("selection-context requires a selection JSON path");
  const options = parseOptions(args);
  const outPath = requireOption(options, "out");
  const context = createSelectionContext(readSelection(path), {
    screenshotPath: options.screenshot,
    selectionPath: path,
    userInstruction: options.instruction,
  });
  writeSelectionContext(outPath, context);
  io.stdout(JSON.stringify({ status: "pass", out: outPath }, null, 2));
  return 0;
}

function runOverrideCandidate(args: string[], io: CliIo): number {
  const path = args.shift();
  if (!path || path.startsWith("--")) throw new Error("override-candidate requires a selection JSON path");
  const options = parseOptions(args);
  const outPath = requireOption(options, "out");
  const op = readOverrideOp(requireOption(options, "op"));
  const candidate = createUserOverrideCandidate(readSelection(path), {
    selectionRef: path,
    op,
    blockId: options["block-id"],
    slot: options.slot,
  });
  writeUserOverrideCandidate(outPath, candidate);
  io.stdout(JSON.stringify({ status: "pass", out: outPath, op }, null, 2));
  return 0;
}

function runValidateSchemaSyncCommand(args: string[], io: CliIo): number {
  const options = parseOptions(args);
  const result = runValidateSchemaSync({
    mdprPath: options["mdpr-path"],
    localSchemaDir: options["local-schema-dir"],
  });
  io.stdout(JSON.stringify(result, null, 2));
  return result.status === "pass" ? 0 : 1;
}

function readSelection(path: string): MdprPptSelection {
  return JSON.parse(readFileSync(path, "utf-8")) as MdprPptSelection;
}

function readJson(path: string): unknown {
  return JSON.parse(readFileSync(path, "utf-8")) as unknown;
}

function writeJson(path: string, value: unknown): void {
  mkdirSync(dirname(path), { recursive: true });
  writeFileSync(path, JSON.stringify(value, null, 2) + "\n", "utf-8");
}

function parseOptions(args: string[]): Record<string, string> {
  const options: Record<string, string> = {};
  for (let index = 0; index < args.length; index += 1) {
    const arg = args[index];
    if (!arg.startsWith("--")) throw new Error(`Unexpected positional argument: ${arg}`);
    const key = arg.slice(2);
    const next = args[index + 1];
    if (!next || next.startsWith("--")) {
      options[key] = "true";
    } else {
      options[key] = next;
      index += 1;
    }
  }
  return options;
}

function requireOption(options: Record<string, string>, key: string): string {
  const value = options[key];
  if (!value) throw new Error(`Missing required option --${key}`);
  return value;
}

function readOverrideOp(value: string): "pinBlock" | "moveBlock" | "hideBlock" {
  if (value === "pinBlock" || value === "moveBlock" || value === "hideBlock") return value;
  throw new Error(`Unsupported override candidate op: ${value}`);
}

if (process.argv[1] && import.meta.url === pathToFileURL(process.argv[1]).href) {
  process.exitCode = runCli(process.argv.slice(2));
}
