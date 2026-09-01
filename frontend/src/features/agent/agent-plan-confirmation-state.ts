export type PlanQuestionOption = { id: string; label: string; instruction: string };
export type PlanQuestion = { id: string; title: string; question: string; options: PlanQuestionOption[] };
export type PlanConfirmationAction = "auto_process" | "answer_questions" | "confirm_plan";

function questionId(index: number) {
  return `question-${index + 1}`;
}

export function presentPlanQuestions(questions: string[], projectMonth = "当前项目月份"): PlanQuestion[] {
  return questions.map((question, index) => {
    const lower = question.toLowerCase();
    if (lower.includes("salary_month") || question.includes("处理月份") || question.includes("项目月份")) {
      return { id: questionId(index), title: "处理月份", question, options: [
        { id: "use_project_month", label: `按项目月份 ${projectMonth} 处理`, instruction: `按项目月份 ${projectMonth} 处理，不按文件名月份覆盖。` },
        { id: "use_file_month", label: "按文件月份处理", instruction: "按总表与来源文件对应的文件月份处理。" },
      ] };
    }
    if (question.includes("补发") || question.includes("补扣")) {
      return { id: questionId(index), title: "补发补扣", question, options: [
        { id: "leave_blank", label: "本月留空，列入待确认", instruction: "本月没有来源依据的补发补扣留空，并列入待确认。" },
        { id: "use_provided_detail", label: "使用我补充的明细", instruction: "使用我在对话中补充的本月补发补扣明细。" },
      ] };
    }
    if (question.includes("奖金") || question.includes("工资核算")) {
      return { id: questionId(index), title: "奖金来源", question, options: [
        { id: "use_latest_source", label: "继续使用现有奖金来源", instruction: "继续使用当前已提供的奖金来源表处理本月奖金。" },
        { id: "wait_for_new_source", label: "等待本月新来源", instruction: "没有本月新奖金来源时不写入，保留并列入待确认。" },
      ] };
    }
    return { id: questionId(index), title: `待确认事项 ${index + 1}`, question, options: [
      { id: "follow_recommendation", label: "按 Agent 建议处理", instruction: `按 Agent 对“${question}”给出的建议处理。` },
      { id: "leave_unresolved", label: "保留并列入待确认", instruction: `“${question}”暂不写入，保留并列入待确认。` },
    ] };
  });
}

export function buildPlanResponse(questions: PlanQuestion[], answers: Record<string, string>, customInstruction = "") {
  const answerText = questions.map((question) => question.options.find((option) => option.id === answers[question.id])?.instruction).filter(Boolean).join("\n");
  return [answerText, customInstruction.trim()].filter(Boolean).join("\n").slice(0, 4000);
}

export function planConfirmationAction({ hasQuestions, hasUnansweredQuestions = hasQuestions }: { hasQuestions: boolean; hasUnansweredQuestions?: boolean }): PlanConfirmationAction {
  if (!hasQuestions) return "auto_process";
  return hasUnansweredQuestions ? "answer_questions" : "confirm_plan";
}
