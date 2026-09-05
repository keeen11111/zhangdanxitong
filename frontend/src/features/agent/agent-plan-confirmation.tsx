"use client";

import { useEffect, useMemo, useState } from "react";
import { AgentRun } from "@/lib/api";
import { buildPlanResponse, presentPlanQuestions, retainPlanAnswers, shouldShowPlanConfirmation } from "./agent-plan-confirmation-state";

export function AgentPlanConfirmation({ run, busy, onCustomize, onAnswersChanged }: {
  run: AgentRun;
  busy: boolean;
  onCustomize: (instruction: string) => boolean | void | Promise<boolean | void>;
  onAnswersChanged?: (instruction: string) => void;
}) {
  const plan = run.model_plan;
  const [answers, setAnswers] = useState<Record<string, string>>({});
  const [customAnswers, setCustomAnswers] = useState<Record<string, string>>({});
  const [submitted, setSubmitted] = useState(false);
  const effectiveQuestions = useMemo(() => {
    const existing = plan?.questions || [];
    // 月份差异由后端自动按文件月份处理（confirmed=true），只有未被
    // 自动确认时才在前端合成确认问题。
    const monthPending = Boolean(
      run.month_confirmation?.required && !run.month_confirmation?.confirmed,
    );
    if (existing.length || !monthPending || run.plan_confirmation?.confirmed) return existing;
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
  const questionKey = questions.map((question) => `${question.id}:${question.question}`).join("\u0000");
  useEffect(() => {
    onAnswersChanged?.(buildPlanResponse(questions, answers, "", customAnswers));
  }, [answers, customAnswers, onAnswersChanged, questionKey, questions]);
  const unansweredCount = questions.filter((question) => {
    const answer = answers[question.id];
    if (!answer) return true;
    // 自定义答案必须填了文字才算答完。
    return answer === "custom_answer" && !customAnswers[question.id]?.trim();
  }).length;
  const answeredCount = questions.length - unansweredCount;
  const selectAllRecommended = () => {
    setAnswers(Object.fromEntries(questions.map((question) => [question.id, "follow_recommendation"])));
  };

  useEffect(() => {
    setAnswers((current) => retainPlanAnswers(questions, current));
    setCustomAnswers((current) => retainPlanAnswers(questions, current));
    setSubmitted(false);
  }, [questionKey, questions, run.run_id]);

  const submitAnswers = async () => {
    if (busy || unansweredCount) return;
    const instruction = buildPlanResponse(questions, answers, "", customAnswers);
    if (instruction) {
      setSubmitted(true);
      try {
        const result = await onCustomize(instruction);
        if (result === false) setSubmitted(false);
      } catch {
        setSubmitted(false);
      }
    }
  };
  if (!shouldShowPlanConfirmation({
    required: Boolean(run.plan_confirmation?.required),
    confirmed: Boolean(run.plan_confirmation?.confirmed),
    hasQuestions: Boolean(questions.length),
    submitting: submitted,
  })) return null;

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
            <div className="flex flex-wrap items-center justify-between gap-2"><h3 id={`${question.id}-title`} className="text-sm font-medium text-slate-900">{question.title}</h3>{answers[question.id] === "custom_answer" ? <span className="text-xs text-amber-700">{customAnswers[question.id]?.trim() ? "已填写" : "填写中"}</span> : answers[question.id] ? <span className="text-xs text-emerald-700">已选择</span> : <span className="text-xs text-amber-700">待选择</span>}</div>
            <p className="mt-1.5 text-xs leading-5 text-slate-600">{question.question}</p>
            <div className="mt-3 grid gap-2 sm:grid-cols-2">{question.options.map((option) => <button key={option.id} type="button" disabled={busy} onClick={() => setAnswers((current) => ({ ...current, [question.id]: option.id }))} className={`min-h-10 rounded-md border px-3 py-2 text-left text-xs font-medium transition-colors focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-blue-500 disabled:cursor-not-allowed disabled:opacity-50 ${answers[question.id] === option.id ? "border-blue-600 bg-blue-50 text-blue-900" : "border-slate-300 bg-white text-slate-700 hover:border-blue-400 hover:bg-blue-50"}`} aria-pressed={answers[question.id] === option.id}>{option.label}</button>)}<button key="custom-answer" type="button" disabled={busy} onClick={() => setAnswers((current) => ({ ...current, [question.id]: "custom_answer" }))} className={`min-h-10 rounded-md border px-3 py-2 text-left text-xs font-medium transition-colors focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-blue-500 disabled:cursor-not-allowed disabled:opacity-50 ${answers[question.id] === "custom_answer" ? "border-blue-600 bg-blue-50 text-blue-900" : "border-dashed border-slate-400 bg-white text-slate-600 hover:border-blue-400 hover:bg-blue-50"}`} aria-pressed={answers[question.id] === "custom_answer"}>自定义答案…</button></div>
            {answers[question.id] === "custom_answer" ? (
              <div className="mt-2">
                <label htmlFor={`${question.id}-custom-input`} className="sr-only">自定义答案</label>
                <textarea id={`${question.id}-custom-input`} value={customAnswers[question.id] || ""} onChange={(event) => setCustomAnswers((current) => ({ ...current, [question.id]: event.target.value }))} disabled={busy} rows={2} maxLength={1000} className="w-full rounded-md border border-slate-300 bg-white px-3 py-2 text-xs leading-5 text-slate-800 outline-none transition-shadow placeholder:text-slate-400 focus:border-blue-500 focus:ring-2 focus:ring-blue-500/20 disabled:cursor-not-allowed disabled:bg-slate-50" placeholder="直接输入你的口径或处理方式，例如：只处理派遣人员，其他人全部删除，人数按来源23人对齐" />
                <p className="mt-1 text-[10px] text-slate-400">你的回答会作为该问题的处理指令提交给 Agent</p>
              </div>
            ) : null}
          </section>)}
        </div>
        <div className="mt-4 flex flex-wrap items-center justify-between gap-3 border-t border-slate-200 pt-4">
          <p className="text-xs leading-5 text-slate-500">{unansweredCount ? `已答 ${answeredCount}/${questions.length}，全部答完后可开始处理` : `已答 ${answeredCount}/${questions.length}，确认后开始处理`}</p>
          <button
            type="button"
            disabled={busy || unansweredCount > 0}
            onClick={submitAnswers}
            className="min-h-10 rounded-md border border-blue-600 bg-blue-600 px-4 py-2 text-xs font-medium text-white transition-colors hover:bg-blue-700 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-blue-500 disabled:cursor-not-allowed disabled:opacity-50"
          >
            确认并开始处理
          </button>
        </div>
        <details className="mt-4 rounded-lg border border-slate-200 bg-white px-3.5 py-3"><summary className="cursor-pointer text-sm font-medium text-slate-700">查看本次处理范围（{plan.steps.length} 项）</summary><p className="mt-2 whitespace-pre-wrap text-xs leading-5 text-slate-600">{plan.summary}</p><ol className="mt-3 space-y-2 text-xs leading-5 text-slate-700">{plan.steps.map((step, index) => <li key={index}>{step}</li>)}</ol></details>
      </> : null}
    </section>
  );
}
