import { useEffect, useMemo, useState, type JSX } from "react";
import {
  SuperDocEditor,
  type SuperDocContentErrorEvent,
  type SuperDocExceptionEvent
} from "@superdoc-dev/react";
import "@superdoc-dev/react/style.css";
import type { ArtifactPreview } from "../../../shared/types";
import { dataUrlToArrayBuffer } from "./previewDataUrl";
import { sanitizeDocxForSuperDoc } from "./sanitizeDocxForSuperDoc";

const DOCX_MIME_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

export default function SuperDocPreview({ preview }: { preview: ArtifactPreview }): JSX.Element {
  const [error, setError] = useState("");
  const [isLoadingDocument, setIsLoadingDocument] = useState(true);
  const documentFile = useMemo(() => buildDocumentFile(preview), [preview]);

  useEffect(() => {
    setError("");
    setIsLoadingDocument(true);
  }, [preview.artifactId]);

  if (!documentFile) {
    return <p className="preview-empty">当前 DOCX 预览缺少文档数据。</p>;
  }

  return (
    <div className="superdoc-preview-host">
      <div className="superdoc-preview-status" aria-live="polite">
        {isLoadingDocument ? "正在加载 Word 文档内容..." : "SuperDoc 原生 DOCX 预览"}
        {error ? ` · Word 渲染失败：${error}` : ""}
      </div>
      <div className="superdoc-preview-frame">
        <SuperDocEditor
          key={preview.artifactId}
          document={documentFile}
          documentMode="viewing"
          role="viewer"
          contained
          hideToolbar
          className="sd-theme-word superdoc-preview-editor"
          style={{ height: "100%", width: "100%" }}
          renderLoading={() => <p className="preview-empty">正在初始化 Word 工作区...</p>}
          onReady={() => {
            setError("");
            setIsLoadingDocument(false);
          }}
          onContentError={(event) => {
            setError(describeContentError(event));
            setIsLoadingDocument(false);
          }}
          onException={(event) => {
            const message = describeException(event);
            if (message) setError(message);
            setIsLoadingDocument(false);
          }}
        />
      </div>
    </div>
  );
}

function buildDocumentFile(preview: ArtifactPreview): File | null {
  if (!preview.dataUrl) return null;
  const buffer = sanitizeDocxForSuperDoc(dataUrlToArrayBuffer(preview.dataUrl));
  const fileName = preview.name.endsWith(".docx") ? preview.name : `${preview.name || "document"}.docx`;
  return new File([buffer], fileName, { type: preview.mimeType || DOCX_MIME_TYPE });
}

function describeContentError(event: SuperDocContentErrorEvent): string {
  return extractKnownErrorMessage(event) || "内容解析失败";
}

function describeException(event: SuperDocExceptionEvent): string | null {
  return extractKnownErrorMessage(event);
}

function extractKnownErrorMessage(error: unknown): string | null {
  if (error && typeof error === "object") {
    if ("error" in error) {
      const nestedMessage = getErrorMessage((error as { error?: unknown }).error);
      if (nestedMessage) return nestedMessage;
    }
    if ("message" in error) {
      const directMessage = getErrorMessage((error as { message?: unknown }).message);
      if (directMessage) return directMessage;
    }
  }

  return getErrorMessage(error);
}

function getErrorMessage(error: unknown): string | null {
  if (error instanceof Error && error.message) return error.message;
  if (typeof error === "string" && error.trim()) return error.trim();
  return null;
}
