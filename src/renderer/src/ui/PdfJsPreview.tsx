import { useEffect, useRef, useState, type JSX } from "react";
import { getDocument, GlobalWorkerOptions } from "pdfjs-dist";
import type { PDFDocumentProxy } from "pdfjs-dist";
import { dataUrlToUint8Array } from "./previewDataUrl";

GlobalWorkerOptions.workerSrc = new URL("pdfjs-dist/build/pdf.worker.min.mjs", import.meta.url).toString();

export default function PdfJsPreview({ dataUrl, name }: { dataUrl: string; name: string }): JSX.Element {
  const containerRef = useRef<HTMLDivElement>(null);
  const [error, setError] = useState("");
  const [pageCount, setPageCount] = useState(0);

  useEffect(() => {
    let cancelled = false;
    let documentTask: PDFDocumentProxy | undefined;
    const container = containerRef.current;
    if (!container) return;

    container.replaceChildren();
    setError("");
    setPageCount(0);

    void (async () => {
      try {
        const loadingTask = getDocument({ data: dataUrlToUint8Array(dataUrl) });
        documentTask = await loadingTask.promise;
        if (cancelled) return;

        setPageCount(documentTask.numPages);
        for (let pageNumber = 1; pageNumber <= documentTask.numPages; pageNumber += 1) {
          if (cancelled) return;
          const page = await documentTask.getPage(pageNumber);
          const viewport = page.getViewport({ scale: 1.35 });
          const ratio = window.devicePixelRatio || 1;
          const canvas = document.createElement("canvas");
          const context = canvas.getContext("2d");
          if (!context) continue;

          canvas.width = Math.floor(viewport.width * ratio);
          canvas.height = Math.floor(viewport.height * ratio);
          canvas.style.width = `${viewport.width}px`;
          canvas.style.height = `${viewport.height}px`;
          context.setTransform(ratio, 0, 0, ratio, 0, 0);

          const pageShell = document.createElement("section");
          pageShell.className = "pdf-page";
          pageShell.setAttribute("aria-label", `${name} 第 ${pageNumber} 页`);
          pageShell.appendChild(canvas);
          container.appendChild(pageShell);

          await page.render({ canvas, canvasContext: context, viewport }).promise;
        }
      } catch (renderError) {
        if (!cancelled) setError(renderError instanceof Error ? renderError.message : String(renderError));
      }
    })();

    return () => {
      cancelled = true;
      void documentTask?.destroy();
      container.replaceChildren();
    };
  }, [dataUrl, name]);

  return (
    <div className="pdfjs-preview">
      <div className="preview-library-label">PDF.js 开源预览{pageCount ? ` · ${pageCount} 页` : ""}</div>
      {error ? <p className="preview-empty">PDF 渲染失败：{error}</p> : null}
      <div ref={containerRef} className="pdf-pages" />
    </div>
  );
}
