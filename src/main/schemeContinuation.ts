import type { SchemeProgressItem, SchemeProgressSection } from "../shared/types";

const SCHEME_COMPLETION_NOTICE_RE =
  /^> 章节生成未完成：当前[^\n]*\n> 当前 Word 只能视为阶段性文件，不是完整方案。\n> (?:下一步应继续处理[^\n]*|下一步应继续补齐未完成章节。)\s*(?:\n\n)?/;

const GENERIC_CONTINUE_RE =
  /^(?:继续(?:生成|写|做)?|继续(?:啊|呀|吧|一下)?|接着(?:继续|来)?|往下继续|继续下去|continue|go on)\s*['"“”‘’`]*\s*[!！。.?？~～]*$/i;

export function stripSyntheticSchemeCompletionNotice(content: string): string {
  return content.replace(SCHEME_COMPLETION_NOTICE_RE, "").trim();
}

export function rewriteContinuationPrompt(message: string, progress?: SchemeProgressItem): string {
  const trimmed = message.trim();
  if (!trimmed || !progress || !hasIncompleteSchemeProgress(progress) || !GENERIC_CONTINUE_RE.test(trimmed)) {
    return trimmed;
  }

  const nextSection = getNextActionableSection(progress);
  const progressText = `${progress.completed}/${progress.total}${progress.failed ? `，失败 ${progress.failed}` : ""}`;
  const nextText = nextSection ? `${nextSection.number} ${nextSection.title}` : "下一未完成章节";
  const draftedCount = progress.drafted ?? progress.sections.filter((section) => section.status === "drafted").length;
  const priorityText = draftedCount > 0 ? "先把已起草章节写入 Word，再继续后续章节。" : "直接继续下一批未完成章节。";

  return [
    trimmed,
    "请继续当前未完成的密码应用方案生成任务，不要重复总结当前进度，也不要只回复“未完成”。",
    `当前章节进度：${progressText}。`,
    `当前优先处理：${nextText}。`,
    priorityText,
    "请直接调用必要工具推进下一步。"
  ].join("\n");
}

function hasIncompleteSchemeProgress(progress: SchemeProgressItem): boolean {
  if (progress.total <= 0) return false;
  return progress.completed < progress.total || progress.failed > 0 || (progress.drafted ?? 0) > 0;
}

function getNextActionableSection(progress: SchemeProgressItem): SchemeProgressSection | undefined {
  return (
    progress.sections.find((section) => section.status === "failed") ??
    progress.sections.find((section) => section.status === "drafted") ??
    progress.sections.find((section) => section.status === "pending")
  );
}
