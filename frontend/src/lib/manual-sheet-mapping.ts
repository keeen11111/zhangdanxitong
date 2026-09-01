export interface ManualSheetMappingSelection {
  issue_id: string;
  target_sheet: string;
}

export function validManualSheetMappings(
  selections: ManualSheetMappingSelection[],
): ManualSheetMappingSelection[] {
  return selections.filter((selection) => selection.issue_id.trim() && selection.target_sheet.trim());
}
