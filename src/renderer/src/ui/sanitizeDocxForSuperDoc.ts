import PizZip from "pizzip";

const WORD_XML_ENTRY_PATTERN =
  /^word\/(document|header\d+|footer\d+|footnotes|endnotes|comments)\.xml$/i;

export function sanitizeDocxForSuperDoc(documentBuffer: ArrayBuffer): ArrayBuffer {
  const zip = new PizZip(documentBuffer);
  let changed = false;

  for (const entryName of Object.keys(zip.files)) {
    if (!WORD_XML_ENTRY_PATTERN.test(entryName)) continue;
    const file = zip.file(entryName);
    if (!file) continue;

    const xml = file.asText();
    const sanitizedXml = unwrapStructuredDocumentTags(xml);
    if (sanitizedXml === xml) continue;

    zip.file(entryName, sanitizedXml);
    changed = true;
  }

  if (!changed) return documentBuffer;

  const output = zip.generate({ type: "uint8array", compression: "DEFLATE" });
  const buffer = new Uint8Array(output.byteLength);
  buffer.set(output);
  return buffer.buffer;
}

function unwrapStructuredDocumentTags(xml: string): string {
  const parser = new DOMParser();
  const document = parser.parseFromString(xml, "application/xml");
  if (document.getElementsByTagName("parsererror").length > 0) return xml;

  let changed = false;

  while (true) {
    const controls = Array.from(document.getElementsByTagNameNS("*", "sdt"));
    if (!controls.length) break;

    let unwrappedInPass = false;

    for (let index = controls.length - 1; index >= 0; index -= 1) {
      const control = controls[index];
      const parent = control.parentNode;
      if (!parent) continue;

      const content = findChildByLocalName(control, "sdtContent");
      if (!content) {
        parent.removeChild(control);
        changed = true;
        unwrappedInPass = true;
        continue;
      }

      const contentChildren = Array.from(content.childNodes);
      for (const child of contentChildren) {
        parent.insertBefore(child, control);
      }

      parent.removeChild(control);
      changed = true;
      unwrappedInPass = true;
    }

    if (!unwrappedInPass) break;
  }

  if (!changed) return xml;
  return new XMLSerializer().serializeToString(document);
}

function findChildByLocalName(node: Element, localName: string): Element | null {
  for (const child of Array.from(node.childNodes)) {
    if (child.nodeType !== Node.ELEMENT_NODE) continue;
    const element = child as Element;
    if (element.localName === localName) return element;
  }

  return null;
}
