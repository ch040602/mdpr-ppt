export type OoxmlExtractionWarning = {
  code: "OOXML_EXTRACTOR_NOT_IMPLEMENTED";
  message: string;
};

export function inspectPptxFallback(): { warnings: OoxmlExtractionWarning[] } {
  return {
    warnings: [{
      code: "OOXML_EXTRACTOR_NOT_IMPLEMENTED",
      message: "Deep grouped-shape and table fallback extraction is reserved for the next mdpr-ppt milestone.",
    }],
  };
}
