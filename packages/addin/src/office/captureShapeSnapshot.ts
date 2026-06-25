import type { PptBoundingBoxPt, PptShapeSnapshot } from "../../../shared/src/index";

export type OfficeShapeLike = {
  id?: string;
  name?: string;
  type?: string;
  left: number;
  top: number;
  width: number;
  height: number;
  rotation?: number;
  fill?: { color?: string; transparency?: number };
  line?: { color?: string; weight?: number };
  textFrame?: { hasText?: boolean };
};

export function captureShapeSnapshot(shape: OfficeShapeLike): PptShapeSnapshot {
  const bboxPt: PptBoundingBoxPt = {
    left: shape.left,
    top: shape.top,
    width: shape.width,
    height: shape.height,
    ...(typeof shape.rotation === "number" ? { rotation: shape.rotation } : {}),
  };
  return {
    pptShapeId: shape.id ?? shape.name ?? "unknown-shape",
    name: shape.name ?? shape.id ?? "unknown-shape",
    type: shape.type ?? "unknown",
    bboxPt,
    styleSnapshot: {
      ...(shape.fill ? { fill: shape.fill } : {}),
      ...(shape.line ? { line: shape.line } : {}),
      ...(shape.textFrame ? { text: { hasText: shape.textFrame.hasText } } : {}),
    },
  };
}
