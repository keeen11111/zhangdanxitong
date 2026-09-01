export const FINANCIAL_INTEGRATION_TIMEOUT_MS = 15 * 60 * 1000;

export interface FinancialProgressStage {
  key: string;
  label: string;
  status: "pending" | "running" | "completed";
}

export function financialIntegrationProgressPercent(stages: FinancialProgressStage[]) {
  if (!stages.length) return 4;
  const completedCount = stages.filter((stage) => stage.status === "completed").length;
  if (completedCount === stages.length) return 100;
  const hasRunningStage = stages.some((stage) => stage.status === "running");
  return Math.min(96, Math.round((completedCount / stages.length) * 100) + (hasRunningStage ? 12 : 4));
}
