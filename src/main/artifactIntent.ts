const ARTIFACT_ACTION_PATTERN = "(?:生成|编制|编写|撰写|输出|导出|创建|制作|起草|整理|形成|完善|补全|写)";
const ARTIFACT_TARGET_PATTERN = "(?:密码应用方案|商用密码应用方案|方案|报告|文档|初稿|草稿|word|docx|pdf)";
const DIAGRAM_TARGET_PATTERN = "(?:流程图|架构图|拓扑图|部署图|技术架构|业务流程|图示|配图)";

const artifactActionBeforeTarget = new RegExp(`${ARTIFACT_ACTION_PATTERN}.{0,18}${ARTIFACT_TARGET_PATTERN}`, "i");
const artifactTargetBeforeAction = new RegExp(`${ARTIFACT_TARGET_PATTERN}.{0,18}${ARTIFACT_ACTION_PATTERN}`, "i");
const diagramIntent = new RegExp(DIAGRAM_TARGET_PATTERN, "i");
const templateDrivenIntent = /(?:按(?:照)?|基于|根据).{0,12}(?:模板|附件|资料|文档|word)/i;
const passwordSchemeIntent = /密码应用方案/i;
const deliveryContent = /方案|报告|文档|密码|密钥|加密|签名|验签|证书|身份鉴别|合规|等保/;

export function hasSchemeArtifactIntent(prompt: string): boolean {
  const text = prompt.trim();
  if (!text) return false;
  if (artifactActionBeforeTarget.test(text) || artifactTargetBeforeAction.test(text)) return true;

  if (templateDrivenIntent.test(text) && /(方案|报告|文档|输出|导出|生成|编制|撰写|整理)/i.test(text)) {
    return true;
  }

  return passwordSchemeIntent.test(text) && /(?:开始|继续|完善|补全|形成|帮我|给我|出一份|写一份)/.test(text);
}

export function hasDiagramArtifactIntent(prompt: string): boolean {
  const text = prompt.trim();
  if (!text) return false;
  return diagramIntent.test(text) || hasSchemeArtifactIntent(text);
}

export function shouldCreateDraftArtifact(prompt: string, content: string): boolean {
  if (content.trim().length < 120) return false;
  if (!hasSchemeArtifactIntent(prompt)) return false;
  return deliveryContent.test(content);
}
