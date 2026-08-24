export interface IntegrationEstimateInput {
  masterRows: number;
  masterColumns: number;
  masterSheets: number;
  formulaCount: number;
  changeCells: number;
  changeFiles: number;
}

export function estimateIntegrationSeconds(input: IntegrationEstimateInput) {
  const masterCells = Math.max(0, input.masterRows * input.masterColumns);
  const seconds =
    10 +
    input.masterSheets * 0.25 +
    input.changeFiles * 0.75 +
    masterCells / 18_000 +
    Math.max(0, input.changeCells) / 30_000 +
    Math.max(0, input.formulaCount) / 5_000;
  return Math.max(15, Math.min(600, Math.round(seconds)));
}

export function formatEstimateRange(seconds: number) {
  const safeSeconds = Math.max(1, Math.round(seconds));
  const lower = Math.max(1, safeSeconds - Math.max(5, Math.round(safeSeconds * 0.15)));
  const upper = safeSeconds + Math.max(10, Math.round(safeSeconds * 0.25));
  if (upper < 60) return `约 ${lower}–${upper} 秒`;
  const lowerMinutes = Math.max(1, Math.floor(lower / 60));
  const upperMinutes = Math.max(lowerMinutes + 1, Math.ceil(upper / 60));
  return `约 ${lowerMinutes}–${upperMinutes} 分钟`;
}
