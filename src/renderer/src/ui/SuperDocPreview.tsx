import { useEffect, useMemo, useState, type JSX } from "react";
import {
  SuperDocEditor,
  type SuperDocContentErrorEvent,
  type SuperDocExceptionEvent,
  type SuperDocReadyEvent
} from "@superdoc-dev/react";
import "@superdoc-dev/react/style.css";
import type { ArtifactPreview } from "../../../shared/types";
import { dataUrlToArrayBuffer } from "./previewDataUrl";
import { sanitizeDocxForSuperDoc } from "./sanitizeDocxForSuperDoc";

export default function SuperDocPreview({ preview }: { preview: ArtifactPreview }): JSX.Element {
  const [error, setError] = useState("");
  const [isLoadingDocument, setIsLoadingDocument] = useState(true);
  const documentBuffer = useMemo(() => buildDocumentBuffer(preview), [preview]);

  useEffect(() => {
    setError("");
    setIsLoadingDocument(true);
  }, [preview.artifactId]);

  if (!documentBuffer) {
    return <p className="preview-empty">当前 DOCX 预览缺少文档数据。</p>;
  }

  return (
    <div className="superdoc-preview-host">
      <div className="preview-library-label">SuperDoc 原生 DOCX 预览</div>
      {isLoadingDocument ? <p className="preview-empty">正在加载 Word 文档内容...</p> : null}
      {error ? <p className="preview-empty">Word 渲染失败：{error}</p> : null}
      <div className="superdoc-preview-frame">
        <SuperDocEditor
          key={preview.artifactId}
          documentMode="viewing"
          role="viewer"
          contained
          className="sd-theme-word superdoc-preview-editor"
          style={{ height: "100%" }}
          renderLoading={() => <p className="preview-empty">正在初始化 Word 工作区...</p>}
          onReady={(event) => void loadDocumentIntoEditor(event, documentBuffer, setError, setIsLoadingDocument)}
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

function buildDocumentBuffer(preview: ArtifactPreview): ArrayBuffer | null {
  if (!preview.dataUrl) return null;
  return sanitizeDocxForSuperDoc(dataUrlToArrayBuffer(preview.dataUrl));
}

async function loadDocumentIntoEditor(
  event: SuperDocReadyEvent,
  documentBuffer: ArrayBuffer,
  setError: (message: string) => void,
  setIsLoadingDocument: (value: boolean) => void
): Promise<void> {
  setError("");
  setIsLoadingDocument(true);

  try {
    const host = event.superdoc as {
      ui?: { document?: { replaceFile?: (file: ArrayBuffer) => Promise<void> } };
      activeEditor?: { replaceFile?: (file: ArrayBuffer) => Promise<void> } | null;
    };

    if (host.ui?.document?.replaceFile) {
      await host.ui.document.replaceFile(documentBuffer);
      setIsLoadingDocument(false);
      return;
    }

    if (!host.activeEditor?.replaceFile) {
      throw new Error("SuperDoc 编辑器尚未就绪");
    }

    await host.activeEditor.replaceFile(documentBuffer);
    setIsLoadingDocument(false);
  } catch (error) {
    setError(getErrorMessage(error) || "文档加载失败");
    setIsLoadingDocument(false);
  }
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
