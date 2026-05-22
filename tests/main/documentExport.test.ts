import { mkdtemp, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { basename, join } from "node:path";
import { describe, expect, it } from "vitest";
import { buildLibreOfficeCandidates, exportDocxToPdf, findLibreOfficeExecutable } from "../../src/main/documentExport";

describe("documentExport", () => {
  it("builds unique LibreOffice candidates with explicit path first", () => {
    const candidates = buildLibreOfficeCandidates("fake-office", ["fake-office", "soffice"]);

    expect(candidates[0]).toBe("fake-office");
    expect(new Set(candidates).size).toBe(candidates.length);
    expect(candidates).toContain("soffice");
  });

  it("finds the first executable candidate that responds to --version", async () => {
    const converter = await findLibreOfficeExecutable({
      libreOfficePath: "fake-office",
      candidateCommands: [],
      execFileImpl: async (file, args) => {
        if (file === "fake-office" && args.includes("--version")) {
          return { stdout: "LibreOffice 24", stderr: "" };
        }
        throw new Error("not found");
      }
    });

    expect(converter).toBe("fake-office");
  });

  it("exports docx to pdf with an injected converter", async () => {
    const dir = await mkdtemp(join(tmpdir(), "pwd-safe-agent-pdf-"));
    const docxPath = join(dir, "密码应用方案.docx");
    await writeFile(docxPath, "fake docx");

    try {
      const result = await exportDocxToPdf(docxPath, dir, {
        libreOfficePath: "fake-office",
        execFileImpl: async (file, args) => {
          if (file !== "fake-office") throw new Error("unexpected converter");
          if (args.includes("--version")) return { stdout: "LibreOffice 24", stderr: "" };
          const sourcePath = args.at(-1) ?? "";
          await writeFile(join(dir, `${basename(sourcePath, ".docx")}.pdf`), "%PDF-1.7");
          return { stdout: "convert ok", stderr: "" };
        }
      });

      expect(result.status).toBe("success");
      expect(result.fileName).toBe("密码应用方案.pdf");
      expect(result.outputPath).toBe(join(dir, "密码应用方案.pdf"));
    } finally {
      await rm(dir, { recursive: true, force: true });
    }
  });

  it("returns unavailable when no converter can be executed", async () => {
    const dir = await mkdtemp(join(tmpdir(), "pwd-safe-agent-pdf-"));
    const docxPath = join(dir, "密码应用方案.docx");
    await writeFile(docxPath, "fake docx");

    try {
      const result = await exportDocxToPdf(docxPath, dir, {
        candidateCommands: ["missing-office"],
        execFileImpl: async () => {
          throw new Error("missing");
        }
      });

      expect(result.status).toBe("unavailable");
      expect(result.summary).toContain("未检测到 LibreOffice");
    } finally {
      await rm(dir, { recursive: true, force: true });
    }
  });
});
