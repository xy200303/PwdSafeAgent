import { execFile } from "node:child_process";
import { existsSync } from "node:fs";
import { mkdir } from "node:fs/promises";
import { basename, dirname, extname, join } from "node:path";
import { promisify } from "node:util";

const execFileAsync = promisify(execFile);

type ExecFileImpl = (
  file: string,
  args: readonly string[],
  options: { cwd?: string; timeout?: number; windowsHide?: boolean }
) => Promise<{ stdout: string; stderr: string }>;

export interface PdfExportOptions {
  libreOfficePath?: string;
  candidateCommands?: string[];
  execFileImpl?: ExecFileImpl;
  timeoutMs?: number;
}

export type PdfExportResult =
  | {
      status: "success";
      outputPath: string;
      fileName: string;
      converter: string;
      summary: string;
    }
  | {
      status: "unavailable";
      outputPath: string;
      fileName: string;
      converter: "";
      summary: string;
    }
  | {
      status: "failed";
      outputPath: string;
      fileName: string;
      converter: string;
      summary: string;
    };

export async function exportDocxToPdf(
  docxPath: string,
  outputDir: string,
  options: PdfExportOptions = {}
): Promise<PdfExportResult> {
  const fileName = `${basename(docxPath, extname(docxPath))}.pdf`;
  const outputPath = join(outputDir, fileName);

  if (!existsSync(docxPath)) {
    return {
      status: "failed",
      outputPath,
      fileName,
      converter: "",
      summary: `Word 文件不存在：${docxPath}`
    };
  }

  const execImpl = options.execFileImpl ?? defaultExecFile;
  const converter = await findLibreOfficeExecutable({
    ...options,
    execFileImpl: execImpl
  });
  if (!converter) {
    return {
      status: "unavailable",
      outputPath,
      fileName,
      converter: "",
      summary: "未检测到 LibreOffice/soffice，已跳过 PDF 导出"
    };
  }

  await mkdir(outputDir, { recursive: true });
  try {
    await execImpl(converter, ["--headless", "--convert-to", "pdf", "--outdir", outputDir, docxPath], {
      cwd: dirname(docxPath),
      timeout: options.timeoutMs ?? 120000,
      windowsHide: true
    });
  } catch (error) {
    return {
      status: "failed",
      outputPath,
      fileName,
      converter,
      summary: `PDF 导出失败：${error instanceof Error ? error.message : String(error)}`
    };
  }

  if (!existsSync(outputPath)) {
    return {
      status: "failed",
      outputPath,
      fileName,
      converter,
      summary: `PDF 导出未生成目标文件：${fileName}`
    };
  }

  return {
    status: "success",
    outputPath,
    fileName,
    converter,
    summary: `已导出 ${fileName}`
  };
}

export async function findLibreOfficeExecutable(options: PdfExportOptions = {}): Promise<string | null> {
  const execImpl = options.execFileImpl ?? defaultExecFile;
  const candidates = buildLibreOfficeCandidates(options.libreOfficePath, options.candidateCommands);

  for (const candidate of candidates) {
    if (!candidate) continue;
    if (hasPathSeparator(candidate) && !existsSync(candidate)) continue;

    try {
      await execImpl(candidate, ["--version"], {
        timeout: 8000,
        windowsHide: true
      });
      return candidate;
    } catch {
      continue;
    }
  }

  return null;
}

export function buildLibreOfficeCandidates(explicitPath?: string, extraCandidates: string[] = []): string[] {
  const candidates = [
    explicitPath?.trim(),
    ...extraCandidates,
    process.env.LIBREOFFICE_PATH?.trim(),
    "soffice",
    "libreoffice",
    "C:\\Program Files\\LibreOffice\\program\\soffice.exe",
    "C:\\Program Files (x86)\\LibreOffice\\program\\soffice.exe",
    "/Applications/LibreOffice.app/Contents/MacOS/soffice",
    "/usr/bin/libreoffice",
    "/usr/local/bin/libreoffice",
    "/usr/bin/soffice",
    "/usr/local/bin/soffice"
  ].filter((candidate): candidate is string => Boolean(candidate));

  return Array.from(new Set(candidates));
}

function hasPathSeparator(value: string): boolean {
  return /[\\/]/.test(value) || /^[a-zA-Z]:/.test(value);
}

function defaultExecFile(
  file: string,
  args: readonly string[],
  options: { cwd?: string; timeout?: number; windowsHide?: boolean }
): Promise<{ stdout: string; stderr: string }> {
  return execFileAsync(file, [...args], options);
}
