import fs from "node:fs/promises";
import path from "node:path";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";


const [inputPath, outputDir] = process.argv.slice(2);
if (!inputPath || !outputDir) {
  throw new Error("usage: render_keyuan_workbook.mjs <input.xlsx> <output-directory>");
}

await fs.mkdir(outputDir, { recursive: true });
const workbook = await SpreadsheetFile.importXlsx(await FileBlob.load(inputPath));
const sheetInspection = await workbook.inspect({ kind: "sheet", include: "id,name", maxChars: 20000 });
const sheetRecords = sheetInspection.ndjson.split("\n").filter(Boolean).map((line) => JSON.parse(line));
const sheetNames = sheetRecords.flatMap((record) => {
  if (record.kind !== "sheet") return [];
  if (Array.isArray(record.data)) return record.data.map((item) => item.name).filter(Boolean);
  if (record.data?.name) return [record.data.name];
  if (record.name) return [record.name];
  return [];
});

const formulaErrors = await workbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!|#NULL!",
  options: { useRegex: true, maxResults: 500 },
  summary: "final formula error scan",
  maxChars: 50000,
});

const renders = [];
for (let index = 0; index < sheetNames.length; index += 1) {
  const sheetName = sheetNames[index];
  const safeName = sheetName.replace(/[<>:"/\\|?*]/g, "_");
  try {
    const preview = await workbook.render({ sheetName, autoCrop: "all", scale: 0.65, format: "png" });
    const target = path.join(outputDir, `${String(index + 1).padStart(2, "0")}-${safeName}.png`);
    await fs.writeFile(target, new Uint8Array(await preview.arrayBuffer()));
    renders.push({ sheetName, status: "rendered", target });
  } catch (error) {
    renders.push({ sheetName, status: "failed", error: String(error) });
  }
}

const report = {
  inputPath,
  sheetNames,
  formulaErrors: formulaErrors.ndjson,
  renders,
};
await fs.writeFile(path.join(outputDir, "render-report.json"), JSON.stringify(report, null, 2), "utf8");
console.log(JSON.stringify(report, null, 2));
