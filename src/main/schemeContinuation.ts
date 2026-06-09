const SCHEME_COMPLETION_NOTICE_RE =
  /^> (?:章节|模块)生成未完成：当前[^\n]*\n> 当前(?: Word)?\s*(?:文件)?只能视为阶段性文件，不是完整方案。\n> (?:下一步应继续处理[^\n]*|下一步应继续补齐未完成(?:章节|模块)。)\s*(?:\n\n)?/;

export function stripSyntheticSchemeCompletionNotice(content: string): string {
  return content.replace(SCHEME_COMPLETION_NOTICE_RE, "").trim();
}
