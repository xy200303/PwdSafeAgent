import { useEffect, useMemo, useState, type JSX } from "react";
import {
  FilePreviewCore,
  createDownloadPlugin,
  createImagePreviewPlugin,
  createPdfPreviewPlugin,
  withPlugins,
  type FileInfo
} from "@dev_xiaoyun/vane-file-preview";
import type { ArtifactPreview } from "../../../shared/types";
import { dataUrlToBlob } from "./previewDataUrl";

const VanePreview = withPlugins(FilePreviewCore, [
  createPdfPreviewPlugin({ useIframe: true, enableToolbar: false }),
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
    previewMode: "offline"
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
  if (preview.mode === "image") return `image/${preview.kind === "jpg" ? "jpeg" : preview.kind}`;
  return "application/octet-stream";
}

function estimateDataUrlSize(dataUrl: string): number {
  if (!dataUrl) return 0;
  const payload = dataUrl.slice(dataUrl.indexOf(",") + 1);
  const padding = payload.endsWith("==") ? 2 : payload.endsWith("=") ? 1 : 0;
  return Math.max(0, Math.floor((payload.length * 3) / 4) - padding);
}
function downloadPreviewFile(file: FileInfo): void {
  const link = document.createElement("a");
  link.href = file.url;
  link.download = file.name;
  link.click();
}
