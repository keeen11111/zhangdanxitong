"use client";

import Link from "next/link";
import { useParams, useSearchParams } from "next/navigation";
import {
  Database,
  FileSpreadsheet,
  GitCompare,
  ListChecks,
  CheckCircle2,
  FileDown,
} from "lucide-react";

import { cn } from "@/lib/utils";

const WORKSPACE_ITEMS = [
  { key: "base", label: "原始库", icon: Database },
  { key: "sources", label: "源头数据", icon: FileSpreadsheet },
  { key: "mapping", label: "匹配引擎", icon: GitCompare },
  { key: "audit", label: "差异审核", icon: ListChecks },
  { key: "final", label: "最终库", icon: CheckCircle2 },
  { key: "export", label: "导出中心", icon: FileDown },
] as const;

export function ProjectWorkspaceNav() {
  const params = useParams<{ projectId: string }>();
  const searchParams = useSearchParams();
  const activeKey = searchParams.get("step") || "base";

  return (
    <nav
      aria-label="项目处理阶段"
      className="mb-6 overflow-x-auto rounded-lg border border-slate-200/90 bg-white p-1.5 shadow-[0_1px_2px_rgba(15,23,42,0.03)]"
    >
      <div className="flex min-w-max items-center gap-1">
        {WORKSPACE_ITEMS.map((item) => {
          const Icon = item.icon;
          const active = item.key === activeKey;
          const href =
            item.key === "base"
              ? `/projects/${params.projectId}`
              : `/projects/${params.projectId}?step=${item.key}`;

          return (
            <Link
              key={item.key}
              href={href}
              aria-current={active ? "page" : undefined}
              className={cn(
                "flex items-center gap-2 rounded-md px-3 py-2 text-sm font-medium transition-colors duration-200",
                active
                  ? "bg-blue-50 text-blue-700 shadow-[inset_0_-2px_0_#2563eb]"
                  : "text-slate-600 hover:bg-slate-50 hover:text-slate-950"
              )}
            >
              <Icon className="h-4 w-4" aria-hidden="true" />
              {item.label}
            </Link>
          );
        })}
      </div>
    </nav>
  );
}
