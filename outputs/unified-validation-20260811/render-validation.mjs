import fs from "node:fs/promises";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const input = await FileBlob.load("unified-payroll.xlsx");
const workbook = await SpreadsheetFile.importXlsx(input);

for (const [sheetName, range, filename] of [
  ["工资核算", "A1:Z12", "payroll-preview.png"],
  ["待人工处理", "A1:H12", "issues-preview.png"],
]) {
  const preview = await workbook.render({ sheetName, range, scale: 1, format: "png" });
  await fs.writeFile(filename, new Uint8Array(await preview.arrayBuffer()));
}
