import { lazy, Suspense, type JSX } from "react";
import * as Dialog from "@radix-ui/react-dialog";
import * as ScrollArea from "@radix-ui/react-scroll-area";
import { IncremarkContent } from "@incremark/react";
import { X } from "lucide-react";
import type { ArtifactPreview } from "../../../shared/types";

const LazyVaneFilePreview = lazy(() => import("./VaneFilePreview"));
const LazySuperDocPreview = lazy(() => import("./SuperDocPreview"));

export default function ArtifactPreviewPanel({
  preview,
  error,
  onClose
}: {
  preview?: ArtifactPreview;
  error: string;
  onClose: () => void;
}): JSX.Element {
  const isDocxPreview = preview?.mode === "docx";

  return (
    <Dialog.Root
      open
      onOpenChange={(open) => {
        if (!open) onClose();
      }}
    >
      <Dialog.Portal>
        <Dialog.Overlay className="preview-backdrop" />
        <Dialog.Content className={isDocxPreview ? "preview-panel preview-panel-docx" : "preview-panel"}>
          <header>
            <div>
              <Dialog.Title asChild>
                <strong>{preview?.name || "预览失败"}</strong>
              </Dialog.Title>
              <span>{preview ? preview.kind.toUpperCase() : error}</span>
            </div>
            <Dialog.Close asChild>
              <button type="button" aria-label="关闭预览">
                <X size={16} />
              </button>
            </Dialog.Close>
          </header>
          {preview?.summary ? (
            <Dialog.Description className="preview-summary">{preview.summary}</Dialog.Description>
          ) : (
            <Dialog.Description className="sr-only">文件预览内容</Dialog.Description>
          )}
          {isDocxPreview ? (
            <div className="preview-docx-shell">
              <div className="preview-docx-body">
                {error ? <p className="preview-empty">{error}</p> : null}
                <Suspense fallback={<PreviewRendererLoading label="正在加载 SuperDoc..." />}>
                  {preview ? <LazySuperDocPreview preview={preview} /> : null}
                </Suspense>
              </div>
            </div>
          ) : (
            <ScrollArea.Root className="preview-scroll">
              <ScrollArea.Viewport className="preview-body">
                {error ? <p className="preview-empty">{error}</p> : null}
              {preview && isVanePreviewMode(preview) ? (
                <Suspense fallback={<PreviewRendererLoading label="正在加载 Vane File Preview..." />}>
                  <LazyVaneFilePreview preview={preview} />
                </Suspense>
              ) : null}
              {preview?.mode === "text" && preview.kind === "md" && preview.text ? (
                <div className="preview-markdown">
                  <IncremarkContent content={preview.text} isFinished />
                </div>
              ) : null}
              {preview?.mode === "text" && preview.kind !== "md" && preview.text ? <pre>{preview.text}</pre> : null}
              {preview?.mode === "unsupported" ? <p className="preview-empty">{preview.summary}</p> : null}
              </ScrollArea.Viewport>
              <ScrollArea.Scrollbar className="scrollbar" orientation="vertical">
                <ScrollArea.Thumb className="scrollbar-thumb" />
              </ScrollArea.Scrollbar>
            </ScrollArea.Root>
          )}
        </Dialog.Content>
      </Dialog.Portal>
    </Dialog.Root>
  );
}

function isVanePreviewMode(preview: ArtifactPreview): boolean {
  return Boolean(preview.dataUrl && ["image", "pdf"].includes(preview.mode));
}

function PreviewRendererLoading({ label }: { label: string }): JSX.Element {
  return <p className="preview-empty">{label}</p>;
}
