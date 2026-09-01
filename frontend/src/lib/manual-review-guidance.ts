type GuidanceIssue = {
  target_field?: string;
  candidate_values?: Array<string | number | boolean | null>;
  source_files?: string[];
  source_sheets?: string[];
};

const DATE_FIELDS = new Set(["入职日期", "离职日期", "转正日期", "生效日期"]);

export function getManualReviewGuidance(issue: GuidanceIssue): string {
  const field = issue.target_field || "这个字段";
  const candidates = (issue.candidate_values || []).filter((value) => value !== null && value !== "");
  const sources = [...(issue.source_files || []), ...(issue.source_sheets || [])].filter(Boolean);
  const parts = [`请先核对${field}的来源，再决定写入值。`];
  if (candidates.length) parts.push(`系统找到的候选值：${candidates.join("、")}。`);
  if (sources.length) parts.push(`参考来源：${sources.join(" · ")}。`);
  if (DATE_FIELDS.has(field)) parts.push("日期请填写 YYYY-MM-DD，例如 2026-07-01。");
  else parts.push("金额、天数或小时请填写数字；确认无误后点击“确定更新”。");
  return parts.join(" ");
}
