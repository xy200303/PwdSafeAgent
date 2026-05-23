export function dataUrlToUint8Array(dataUrl: string): Uint8Array<ArrayBuffer> {
  const binary = decodeBase64(getDataUrlPayload(dataUrl));
  const buffer = new ArrayBuffer(binary.length);
  const bytes = new Uint8Array(buffer);
  for (let index = 0; index < binary.length; index += 1) {
    bytes[index] = binary.charCodeAt(index);
  }
  return bytes;
}

export function dataUrlToBlob(dataUrl: string): Blob {
  const mimeType = dataUrl.match(/^data:([^;,]+)[;,]/)?.[1] || "application/octet-stream";
  return new Blob([dataUrlToUint8Array(dataUrl).buffer], { type: mimeType });
}

function getDataUrlPayload(dataUrl: string): string {
  const commaIndex = dataUrl.indexOf(",");
  if (!dataUrl.startsWith("data:") || commaIndex < 0) {
    throw new Error("Invalid data URL");
  }
  return dataUrl.slice(commaIndex + 1);
}

function decodeBase64(payload: string): string {
  if (typeof globalThis.atob === "function") {
    return globalThis.atob(payload);
  }
  if (typeof Buffer !== "undefined") {
    return Buffer.from(payload, "base64").toString("binary");
  }
  throw new Error("No base64 decoder is available");
}
