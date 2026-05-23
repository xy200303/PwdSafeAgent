import { describe, expect, it } from "vitest";
import { dataUrlToBlob, dataUrlToUint8Array } from "../../src/renderer/src/ui/previewDataUrl";

describe("artifact preview renderers", () => {
  it("converts data urls to bytes for PDF.js", () => {
    const bytes = dataUrlToUint8Array("data:application/pdf;base64,JVBERi0x");

    expect(Array.from(bytes)).toEqual([37, 80, 68, 70, 45, 49]);
  });

  it("converts data urls to blobs for docx-preview", () => {
    const blob = dataUrlToBlob(
      "data:application/vnd.openxmlformats-officedocument.wordprocessingml.document;base64,UEsDBA=="
    );

    expect(blob.type).toBe("application/vnd.openxmlformats-officedocument.wordprocessingml.document");
    expect(blob.size).toBe(4);
  });
});
