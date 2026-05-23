import { useEffect, useRef, useState, type CSSProperties, type JSX } from "react";
import { renderAsync } from "docx-preview";
import { dataUrlToBlob } from "./previewDataUrl";

const MIN_DOCX_SCALE = 0.42;
const FIT_WIDTH_GUTTER = 8;

export default function DocxPreview({ dataUrl }: { dataUrl: string }): JSX.Element {
  const viewportRef = useRef<HTMLDivElement>(null);
  const containerRef = useRef<HTMLDivElement>(null);
  const renderIdRef = useRef(0);
  const naturalWidthRef = useRef(0);
  const [error, setError] = useState("");
  const [isRendering, setIsRendering] = useState(false);
  const [scale, setScale] = useState(1);

  function updateScale(): void {
    const viewport = viewportRef.current;
    const naturalWidth = naturalWidthRef.current;
    if (!viewport || !naturalWidth) return;

    const availableWidth = Math.max(0, viewport.clientWidth - FIT_WIDTH_GUTTER);
    const nextScale = Math.min(1, Math.max(MIN_DOCX_SCALE, availableWidth / naturalWidth));
    const roundedScale = Math.round(nextScale * 1000) / 1000;
    setScale((currentScale) => (Math.abs(currentScale - roundedScale) < 0.005 ? currentScale : roundedScale));
  }

  useEffect(() => {
    const renderId = renderIdRef.current + 1;
    renderIdRef.current = renderId;
    const container = containerRef.current;
    if (!container) return;

    container.replaceChildren();
    naturalWidthRef.current = 0;
    setError("");
    setIsRendering(true);
    setScale(1);

    void (async () => {
      const stagingContainer = document.createElement("div");

      try {
        await renderAsync(dataUrlToBlob(dataUrl), stagingContainer, undefined, {
          breakPages: true,
          className: "docx-preview-document",
          inWrapper: true,
          ignoreFonts: false,
          renderHeaders: true,
          renderFooters: true
        });

        if (renderIdRef.current !== renderId || !containerRef.current) return;
        containerRef.current.replaceChildren(...Array.from(stagingContainer.childNodes));
        naturalWidthRef.current = measureDocxNaturalWidth(containerRef.current);
        requestAnimationFrame(() => {
          if (renderIdRef.current === renderId) updateScale();
        });
        setIsRendering(false);
      } catch (renderError) {
        if (renderIdRef.current !== renderId) return;
        setError(renderError instanceof Error ? renderError.message : String(renderError));
        setIsRendering(false);
      }
    })();

    return () => {
      if (renderIdRef.current === renderId) {
        renderIdRef.current += 1;
      }
    };
  }, [dataUrl]);

  useEffect(() => {
    const viewport = viewportRef.current;
    if (!viewport) return;

    const observer = new ResizeObserver(() => updateScale());
    observer.observe(viewport);
    return () => observer.disconnect();
  }, []);

  return (
    <div className="docx-preview-host">
      <div className="preview-library-label">
        docx-preview 开源预览{scale < 0.995 ? ` · 适配 ${Math.round(scale * 100)}%` : ""}
      </div>
      {isRendering ? <p className="preview-empty">正在渲染 Word 预览...</p> : null}
      {error ? <p className="preview-empty">Word 渲染失败：{error}</p> : null}
      <div ref={viewportRef} className="docx-viewport">
        <div
          ref={containerRef}
          className="docx-pages"
          style={{ "--docx-scale": scale.toString() } as CSSProperties}
        />
      </div>
    </div>
  );
}

function measureDocxNaturalWidth(container: HTMLElement): number {
  const pages = Array.from(
    container.querySelectorAll<HTMLElement>("section.docx-preview-document, section.docx")
  );
  const widths = pages.map((page) => parseCssPixels(page.style.width) || page.scrollWidth || page.offsetWidth);
  return Math.max(...widths, 0);
}

function parseCssPixels(value: string): number {
  const parsed = Number.parseFloat(value);
  return Number.isFinite(parsed) ? parsed : 0;
}
