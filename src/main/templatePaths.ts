import { join, resolve } from "node:path";

export const PROJECT_BUNDLED_DOCS_RELATIVE_PATH = "resources/docs";
export const RUNTIME_BUNDLED_DOCS_RELATIVE_PATH = "docs";
export const BUILT_IN_TEMPLATE_DIRNAME = "templates";
export const BUILT_IN_DESIGN_DIRNAME = "design";
export const BUILT_IN_EXAMPLES_DIRNAME = "examples";
export const BUILT_IN_REFERENCES_DIRNAME = "references";

export const BUILT_IN_TEMPLATE_DOCX_BASENAME = "密码应用方案.docx";
export const BUILT_IN_TEMPLATE_JSON_BASENAME = "密码应用方案.template.json";
export const BUILT_IN_TEMPLATE_MARKDOWN_BASENAME = "密码应用方案.md";

export const BUILT_IN_TEMPLATE_DOCX_RELATIVE_PATH = `docs/${BUILT_IN_TEMPLATE_DIRNAME}/${BUILT_IN_TEMPLATE_DOCX_BASENAME}`;
export const BUILT_IN_TEMPLATE_JSON_RELATIVE_PATH = `docs/${BUILT_IN_TEMPLATE_DIRNAME}/${BUILT_IN_TEMPLATE_JSON_BASENAME}`;
export const BUILT_IN_TEMPLATE_MARKDOWN_RELATIVE_PATH = `docs/${BUILT_IN_TEMPLATE_DIRNAME}/${BUILT_IN_TEMPLATE_MARKDOWN_BASENAME}`;

export function getProjectBundledDocsDir(rootDir: string): string {
  return join(rootDir, "resources", "docs");
}

function getBuiltInTemplateAssetPath(docsDir: string, basename: string): string {
  return join(docsDir, BUILT_IN_TEMPLATE_DIRNAME, basename);
}

export function getBuiltInTemplateDocxPath(docsDir: string): string {
  return getBuiltInTemplateAssetPath(docsDir, BUILT_IN_TEMPLATE_DOCX_BASENAME);
}

export function getBuiltInTemplateJsonPath(docsDir: string): string {
  return getBuiltInTemplateAssetPath(docsDir, BUILT_IN_TEMPLATE_JSON_BASENAME);
}

export function getBuiltInTemplateMarkdownPath(docsDir: string): string {
  return getBuiltInTemplateAssetPath(docsDir, BUILT_IN_TEMPLATE_MARKDOWN_BASENAME);
}

export function isBuiltInTemplateJsonPath(filePath: string, docsDir: string): boolean {
  return resolve(filePath).toLowerCase() === resolve(getBuiltInTemplateJsonPath(docsDir)).toLowerCase();
}
