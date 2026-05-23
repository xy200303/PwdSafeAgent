import { useEffect, useMemo, useRef, useState, type CSSProperties, type JSX } from "react";
import { renderAsync } from "docx-preview";
import {
  FilePreviewCore,
  createDownloadPlugin,
  createImagePreviewPlugin,
  createPdfPreviewPlugin,
  withPlugins,
  type FileInfo,
  type FilePreviewPlugin
} from "@dev_xiaoyun/vane-file-preview";
import type { ArtifactPreview } from "../../../shared/types";
import { dataUrlToBlob } from "./previewDataUrl";

const DOCX_MIME_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
const DEFAULT_DOCX_SCALE_FACTOR = 0.78;
const DOCX_ZOOM_STEP = 0.08;
const MIN_DOCX_SCALE = 0.3;
const MAX_DOCX_SCALE = 1;
const FIT_WIDTH_GUTTER = 96;

const VanePreview = withPlugins(FilePreviewCore, [
  createPdfPreviewPlugin({ useIframe: true, enableToolbar: false }),
  createFitDocxPreviewPlugin(),
  createImagePreviewPlugin({ enableZoom: true, enableRotate: true }),
  createDownloadPlugin({ showFileInfo: true })
]);

export default function VaneFilePreview({ preview }: { preview: ArtifactPreview }): JSX.Element {
  const [error, setError] = useState("");
  const previewUrl = useObjectUrl(preview.dataUrl);
  const file = useMemo(() => buildVaneFileInfo(preview, previewUrl), [preview, previewUrl]);

  if (preview.dataUrl && !previewUrl) {
    return <p className="preview-empty">正在准备文件预览...</p>;
  }

  if (!previewUrl || !file) {
    return <p className="preview-empty">该文件暂不支持 Vane 内置预览。</p>;
  }

  return (
    <div className="vane-preview-host">
      <div className="preview-library-label">Vane File Preview 开源预览</div>
      {error ? <p className="preview-empty">文件预览失败：{error}</p> : null}
      <div className="vane-preview-frame">
        <VanePreview
          key={`${preview.artifactId}-${preview.mode}`}
          file={file}
          enableDefaultToolbar={false}
          onLoadError={(loadError) => setError(loadError.message)}
          containerStyle={{
            width: "100%",
            height: "100%",
            minHeight: 0,
            background: "#fff",
            borderRadius: 0
          }}
          contentStyle={{
            width: "100%",
            height: "100%",
            minHeight: 0,
            background: "#fff"
          }}
        />
      </div>
    </div>
  );
}

export function buildVaneFileInfo(preview: ArtifactPreview, url: string): FileInfo {
  const extension = getPreviewExtension(preview);
  return {
    name: preview.name,
    url,
    size: preview.size ?? estimateDataUrlSize(preview.dataUrl ?? ""),
    type: preview.mimeType || getMimeType(preview),
    extension,
    previewMode: "offline",
    docxMode: "docx-preview"
  };
}

function useObjectUrl(dataUrl?: string): string {
  const [objectUrl, setObjectUrl] = useState("");

  useEffect(() => {
    if (!dataUrl) {
      setObjectUrl("");
      return undefined;
    }

    const nextObjectUrl = URL.createObjectURL(dataUrlToBlob(dataUrl));
    setObjectUrl(nextObjectUrl);
    return () => URL.revokeObjectURL(nextObjectUrl);
  }, [dataUrl]);

  return objectUrl;
}

function getPreviewExtension(preview: ArtifactPreview): string {
  const nameExtension = preview.name.match(/\.[^.]+$/)?.[0]?.toLowerCase();
  if (nameExtension) return nameExtension;
  return preview.kind === "other" ? "" : `.${preview.kind}`;
}

function getMimeType(preview: ArtifactPreview): string {
  if (preview.mode === "pdf") return "application/pdf";
  if (preview.mode === "docx") return DOCX_MIME_TYPE;
  if (preview.mode === "image") return `image/${preview.kind === "jpg" ? "jpeg" : preview.kind}`;
  return "application/octet-stream";
}

function estimateDataUrlSize(dataUrl: string): number {
  if (!dataUrl) return 0;
  const payload = dataUrl.slice(dataUrl.indexOf(",") + 1);
  const padding = payload.endsWith("==") ? 2 : payload.endsWith("=") ? 1 : 0;
  return Math.max(0, Math.floor((payload.length * 3) / 4) - padding);
}

function createFitDocxPreviewPlugin(): FilePreviewPlugin {
  return {
    name: "FitDocxPreviewPlugin",
    version: "1.0.0",
    description: "DOCX preview rendered with docx-preview and fit-width scaling",
    supportedTypes: [DOCX_MIME_TYPE],
    supportedExtensions: [".docx"],
    hooks: {
      canHandle: (file) => file.extension.toLowerCase() === ".docx" || file.type === DOCX_MIME_TYPE,
      getPriority: () => 20,
      render: (context) => <FitDocxPreview file={context.file} />,
      getActions: (context) => ({
        download: () => downloadPreviewFile(context.file),
        save: () => downloadPreviewFile(context.file)
      })
    }
  };
}

function FitDocxPreview({ file }: { file: FileInfo }): JSX.Element {
  const viewportRef = useRef<HTMLDivElement>(null);
  const pagesRef = useRef<HTMLDivElement>(null);
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
    setScale((currentScale) => {
      const nextScale = clampDocxScale((availableWidth / naturalWidth) * DEFAULT_DOCX_SCALE_FACTOR);
      return Math.abs(currentScale - nextScale) < 0.005 ? currentScale : nextScale;
    });
  }

  useEffect(() => {
    const renderId = renderIdRef.current + 1;
    renderIdRef.current = renderId;
    const pages = pagesRef.current;
    if (!pages) return;

    pages.replaceChildren();
    naturalWidthRef.current = 0;
    setError("");
    setIsRendering(true);
    setScale(1);

    void (async () => {
      const stagingContainer = document.createElement("div");

      try {
        const response = await fetch(file.url);
        if (!response.ok) throw new Error(`HTTP ${response.status}: ${response.statusText}`);

        await renderAsync(await response.arrayBuffer(), stagingContainer, undefined, {
          breakPages: true,
          className: "docx-preview-document",
          ignoreFonts: false,
          inWrapper: true,
          renderFooters: true,
          renderHeaders: true,
          useBase64URL: true
        });

        if (renderIdRef.current !== renderId || !pagesRef.current) return;
        pagesRef.current.replaceChildren(...Array.from(stagingContainer.childNodes));
        naturalWidthRef.current = measureDocxNaturalWidth(pagesRef.current);
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
  }, [file.url]);

  useEffect(() => {
    const viewport = viewportRef.current;
    if (!viewport) return;

    const observer = new ResizeObserver(() => updateScale());
    observer.observe(viewport);
    return () => observer.disconnect();
  }, []);

  return (
    <div ref={viewportRef} className="vane-docx-preview">
      <div className="vane-docx-toolbar">
        <span>Word 标准版式 · {Math.round(scale * 100)}%</span>
        <button type="button" onClick={() => setScale((currentScale) => clampDocxScale(currentScale - DOCX_ZOOM_STEP))}>
          缩小
        </button>
        <button type="button" onClick={updateScale}>
          适配
        </button>
        <button type="button" onClick={() => setScale((currentScale) => clampDocxScale(currentScale + DOCX_ZOOM_STEP))}>
          放大
        </button>
      </div>
      {isRendering ? <p className="preview-empty vane-docx-status">正在渲染 Word 标准版式...</p> : null}
      {error ? <p className="preview-empty vane-docx-status error">Word 渲染失败：{error}</p> : null}
      <div
        ref={pagesRef}
        className="vane-docx-pages"
        style={{ "--docx-scale": scale.toString() } as CSSProperties}
      />
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

function clampDocxScale(scale: number): number {
  const safeScale = Number.isFinite(scale) ? scale : DEFAULT_DOCX_SCALE_FACTOR;
  return Math.round(Math.min(MAX_DOCX_SCALE, Math.max(MIN_DOCX_SCALE, safeScale)) * 1000) / 1000;
}

function downloadPreviewFile(file: FileInfo): void {
  const link = document.createElement("a");
  link.href = file.url;
  link.download = file.name;
  link.click();
}
