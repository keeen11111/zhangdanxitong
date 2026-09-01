type CompletedIntegration = { completed_at: string };

type IntegrationFile = { file_type: string };
type IntegrationProgressState = { status: string };

export function hasFinancialWorkbookFiles(files: IntegrationFile[]): boolean {
  return files.some(
    (file) => file.file_type === "financial_master" || file.file_type === "financial_source",
  );
}

export function shouldLoadFinancialIntegrationResult(
  files: IntegrationFile[],
  progress: IntegrationProgressState | null,
): boolean {
  return hasFinancialWorkbookFiles(files) && progress !== null && progress.status !== "idle";
}

export function selectLatestIntegrationResult(
  financial: CompletedIntegration | null,
  pipeline: CompletedIntegration | null,
): "financial" | "pipeline" | null {
  if (!financial) return pipeline ? "pipeline" : null;
  if (!pipeline) return "financial";

  const financialTime = Date.parse(financial.completed_at);
  const pipelineTime = Date.parse(pipeline.completed_at);
  if (Number.isNaN(financialTime) || Number.isNaN(pipelineTime)) return "pipeline";
  return pipelineTime > financialTime ? "pipeline" : "financial";
}
