import { Loader2, TableProperties } from "lucide-react";

import { WorkbookAnalysis } from "@/lib/api";

export function MasterWorkbookAnalysis({
  analysis,
  loading,
}: {
  analysis: WorkbookAnalysis | null;
  loading: boolean;
}) {
  if (loading) {
    return (
      <div className="mt-4 flex items-center gap-2 rounded-lg bg-sky-50 px-4 py-3 text-sm text-sky-900" role="status">
        <Loader2 className="h-4 w-4 animate-spin" />
        正在梳理总表 Sheet、数据与公式类型…
      </div>
    );
  }
  if (!analysis) return null;

  return (
    <section className="mt-4 overflow-hidden rounded-lg border border-sky-200" aria-labelledby="master-analysis-title">
      <div className="flex flex-col gap-2 bg-sky-50 px-4 py-3 sm:flex-row sm:items-center sm:justify-between">
        <div className="flex items-center gap-2">
          <TableProperties className="h-4 w-4 text-sky-800" />
          <h3 id="master-analysis-title" className="text-sm font-semibold text-sky-950">总表内容分类</h3>
        </div>
        <p className="text-xs tabular-nums text-sky-800">
          {analysis.sheet_count} 个 Sheet · {analysis.formula_count} 个公式单元格
        </p>
      </div>
      <div className="flex flex-wrap gap-2 border-b border-sky-100 px-4 py-3">
        {analysis.categories.map((category) => (
          <span key={category} className="rounded-full bg-slate-100 px-2.5 py-1 text-xs text-slate-700">
            {category}
          </span>
        ))}
      </div>
      <div className="max-h-64 overflow-auto">
        <table className="w-full text-left text-xs">
          <thead className="sticky top-0 bg-white text-slate-500">
            <tr className="border-b border-slate-200">
              <th className="px-4 py-2 font-medium">Sheet</th>
              <th className="px-4 py-2 font-medium">分类</th>
              <th className="hidden px-4 py-2 font-medium sm:table-cell">内容</th>
              <th className="px-4 py-2 text-right font-medium">规模</th>
            </tr>
          </thead>
          <tbody className="divide-y divide-slate-100">
            {analysis.sheets.map((sheet) => (
              <tr key={sheet.name}>
                <td className="max-w-48 truncate px-4 py-2.5 font-medium text-slate-800" title={sheet.name}>{sheet.name}</td>
                <td className="px-4 py-2.5 text-slate-600">{sheet.category}</td>
                <td className="hidden px-4 py-2.5 text-slate-500 sm:table-cell">{sheet.content_types.join("、")}</td>
                <td className="whitespace-nowrap px-4 py-2.5 text-right tabular-nums text-slate-500">
                  {sheet.row_count} × {sheet.col_count}
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </section>
  );
}
