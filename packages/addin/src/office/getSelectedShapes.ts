import { captureShapeSnapshot, type OfficeShapeLike } from "./captureShapeSnapshot";
import type { PptShapeSnapshot } from "../../../shared/src/index";

export type OfficeShapeCollectionLike = {
  items?: OfficeShapeLike[];
  load: (properties: string | string[]) => void;
};

export type OfficePresentationContextLike = {
  presentation: {
    getSelectedShapes: () => OfficeShapeCollectionLike;
  };
  sync: () => Promise<void> | void;
};

const selectedShapeProperties = [
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
];

export async function captureSelectedShapeSnapshots(
  context: OfficePresentationContextLike,
): Promise<PptShapeSnapshot[]> {
  const selectedShapes = context.presentation.getSelectedShapes();
  selectedShapes.load(selectedShapeProperties);
  await context.sync();
  return (selectedShapes.items ?? []).map(captureShapeSnapshot);
}
