"use client";

import { useEffect, useMemo, useRef, useState } from "react";
import { AgentRun } from "@/lib/api";
import { buildPlanResponse, presentPlanQuestions } from "./agent-plan-confirmation-state";

export function AgentPlanConfirmation({ run, busy, onCustomize }: {
  run: AgentRun;
  busy: boolean;
  onCustomize: (instruction: string) => void;
}) {
  const plan = run.model_plan;
  const [answers, setAnswers] = useState<Record<string, string>>({});
  const submittedAnswerSet = useRef("");
  const effectiveQuestions = useMemo(() => {
    const existing = plan?.questions || [];
    if (existing.length || !run.month_confirmation?.required || run.plan_confirmation?.confirmed) return existing;
    return [
      `当前项目月份为 ${run.salary_month || "当前项目月份"}，但上传文件属于 ${run.month_confirmation.filename_month || "其他月份"} 所属批次。请确认本次应按项目月份处理，还是按上传文件所属批次处理；若按项目月份处理，请补充对应月份的数据源。`,
    ];
  }, [plan?.questions, run.month_confirmation, run.plan_confirmation?.confirmed, run.salary_month]);
  const allQuestions = useMemo(() => {
    if (run.plan_confirmation?.confirmed) return effectiveQuestions;
    const issues = run.basic_processor?.issues || [];
    const issueQuestions = issues
      .map((issue) => [issue.item, issue.detail].filter(Boolean).join("："))
      .filter(Boolean);
    return [...effectiveQuestions, ...issueQuestions];
  }, [effectiveQuestions, run.basic_processor?.issues, run.plan_confirmation?.confirmed]);
  const questions = useMemo(
    () => presentPlanQuestions(allQuestions, run.salary_month || "当前项目月份")
      // Month selection is an automatic project-level default. Filter it on
      // the client too so older runs cannot resurrect the retired prompt.
      // Keep it when the question explicitly concerns a conflicting data
      // source; that is a real business decision, not a filename detail.
      .filter((question) => question.title !== "处理月份" || /数据源|source|补充|薪资数据/i.test(question.question)),
    [allQuestions, run.salary_month],
  );
  const unansweredCount = questions.filter((question) => !answers[question.id]).length;
  const selectAllRecommended = () => {
    setAnswers(Object.fromEntries(questions.map((question) => [question.id, "follow_recommendation"])));
  };

  useEffect(() => {
    setAnswers({});
    submittedAnswerSet.current = "";
  }, [plan?.questions, run.basic_processor?.issue_count]);

  useEffect(() => {
    if (busy || !questions.length || unansweredCount) return;
    const instruction = buildPlanResponse(questions, answers);
    if (!instruction || submittedAnswerSet.current === instruction) return;
    submittedAnswerSet.current = instruction;
    onCustomize(instruction);
  }, [answers, busy, onCustomize, questions, unansweredCount]);

  // Complete plans start automatically. This panel only collects genuine
  // business choices, then submits them as soon as every choice is present.
  if (!run.plan_confirmation?.required || run.plan_confirmation.confirmed || !questions.length) return null;

  return (
    <section aria-labelledby="agent-plan-heading" className="my-6 rounded-xl border border-blue-200 bg-white p-5">
      <h2 id="agent-plan-heading" className="text-base font-semibold text-slate-900">确认本次处理</h2>
      <p className="mt-1 text-xs leading-5 text-slate-500">原件不会修改；确认后只生成独立草稿。</p>
      {plan && questions.length ? <>
        <div className="mt-4 space-y-3" aria-label="需要确认的处理选择">
          <div className="flex flex-wrap items-center justify-between gap-3">
            <p className="text-sm leading-6 text-slate-700">以下事项已按问题类型归类。默认可一次性交给 Agent 统一处理，也可以单独调整某一类。</p>
            <button
              type="button"
              disabled={busy}
              onClick={selectAllRecommended}
              className="min-h-10 rounded-md border border-blue-600 bg-blue-600 px-3 py-2 text-xs font-medium text-white transition-colors hover:bg-blue-700 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-blue-500 disabled:cursor-not-allowed disabled:opacity-50"
            >
              一键按 Agent 建议处理全部
            </button>
          </div>
          {questions.map((question) => <section key={question.id} className="rounded-lg border border-slate-200 bg-slate-50/70 p-3.5" aria-labelledby={`${question.id}-title`}>
            <div className="flex flex-wrap items-center justify-between gap-2"><h3 id={`${question.id}-title`} className="text-sm font-medium text-slate-900">{question.title}</h3>{answers[question.id] ? <span className="text-xs text-emerald-700">已选择</span> : <span className="text-xs text-amber-700">待选择</span>}</div>
            <div className="mt-3 grid gap-2 sm:grid-cols-2">{question.options.map((option) => <button key={option.id} type="button" disabled={busy} onClick={() => setAnswers((current) => ({ ...current, [question.id]: option.id }))} className={`min-h-10 rounded-md border px-3 py-2 text-left text-xs font-medium transition-colors focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-blue-500 disabled:cursor-not-allowed disabled:opacity-50 ${answers[question.id] === option.id ? "border-blue-600 bg-blue-50 text-blue-900" : "border-slate-300 bg-white text-slate-700 hover:border-blue-400 hover:bg-blue-50"}`} aria-pressed={answers[question.id] === option.id}>{option.label}</button>)}</div>
            <details className="mt-2 text-[11px] leading-5 text-slate-500"><summary className="cursor-pointer select-none">查看问题详情</summary><p className="mt-1">{question.question}</p></details>
          </section>)}
        </div>
        <details className="mt-4 rounded-lg border border-slate-200 bg-white px-3.5 py-3"><summary className="cursor-pointer text-sm font-medium text-slate-700">查看本次处理范围（{plan.steps.length} 项）</summary><p className="mt-2 whitespace-pre-wrap text-xs leading-5 text-slate-600">{plan.summary}</p><ol className="mt-3 space-y-2 text-xs leading-5 text-slate-700">{plan.steps.map((step, index) => <li key={index}>{step}</li>)}</ol></details>
      </> : null}
    </section>
  );
}
