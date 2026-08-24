"use client";

import { useEffect, useState } from "react";
import {
  BarChart,
  Bar,
  PieChart,
  Pie,
  Cell,
  XAxis,
  YAxis,
  Tooltip,
  ResponsiveContainer,
  CartesianGrid,
} from "recharts";
import {
  Database,
  Columns3,
  CheckCircle2,
  AlertCircle,
  Loader2,
  Hash,
  Type,
  Calendar,
  TrendingUp,
} from "lucide-react";

import { api, FileProfile, ColumnProfile } from "@/lib/api";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { Badge } from "@/components/ui/badge";
import { Progress } from "@/components/ui/progress";
import { cn } from "@/lib/utils";

const PIE_COLORS = [
  "#3b82f6", "#10b981", "#f59e0b", "#ef4444", "#8b5cf6",
  "#ec4899", "#14b8a6", "#f97316", "#6366f1", "#84cc16",
  "#06b6d4", "#a855f7", "#eab308", "#64748b", "#fb7185",
];

const DTYPE_ICON = {
  number: Hash,
  text: Type,
  datetime: Calendar,
  empty: AlertCircle,
} as const;

const DTYPE_LABEL = {
  number: "数值",
  text: "文本",
  datetime: "日期",
  empty: "空列",
};

/** 缺失率对应的颜色 */
function missingColor(rate: number): string {
  if (rate === 0) return "text-green-600";
  if (rate < 10) return "text-amber-600";
  if (rate < 30) return "text-orange-600";
  return "text-red-600";
}

/** 列质量卡片 */
function ColumnCard({ col }: { col: ColumnProfile }) {
  const Icon = DTYPE_ICON[col.dtype] || Type;
  return (
    <div className="rounded-lg border bg-white p-3 shadow-sm">
      <div className="flex items-center justify-between">
        <div className="flex items-center gap-1.5 min-w-0">
          <Icon className="h-3.5 w-3.5 shrink-0 text-muted-foreground" />
          <span className="truncate text-sm font-medium" title={col.name}>
            {col.name}
          </span>
        </div>
        <Badge variant="outline" className="shrink-0 text-[10px]">
          {DTYPE_LABEL[col.dtype]}
        </Badge>
      </div>

      {/* 缺失率 */}
      <div className="mt-2">
        <div className="flex items-center justify-between text-[11px]">
          <span className="text-muted-foreground">缺失率</span>
          <span className={cn("font-semibold", missingColor(col.missing_rate))}>
            {col.missing_rate}%
          </span>
        </div>
        <Progress
          value={100 - col.missing_rate}
          className="mt-1 h-1.5"
        />
      </div>

      {/* 统计信息 */}
      <div className="mt-2 flex flex-wrap gap-x-3 gap-y-0.5 text-[10px] text-muted-foreground">
        <span>唯一值 {col.unique_count}</span>
        <span>缺失 {col.missing_count}</span>
        {col.dtype === "number" && col.min !== null && (
          <span>min {col.min}</span>
        )}
        {col.dtype === "number" && col.max !== null && (
          <span>max {col.max}</span>
        )}
        {col.dtype === "number" && col.mean !== null && (
          <span>均值 {col.mean}</span>
        )}
      </div>

      {/* 示例值 */}
      {col.sample_values.length > 0 && (
        <div className="mt-1.5 truncate text-[10px] text-slate-400" title={col.sample_values.join(", ")}>
          示例: {col.sample_values.join(" | ")}
        </div>
      )}
    </div>
  );
}

/** 数据可视化看板 */
export function DataProfileBoard({
  projectId,
  fileId,
  compact = false,
}: {
  projectId: string;
  fileId: string;
  compact?: boolean;
}) {
  const [profile, setProfile] = useState<FileProfile | null>(null);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    let cancelled = false;
    setLoading(true);
    api
      .profileFile(projectId, fileId)
      .then((p) => {
        if (!cancelled) {
          setProfile(p);
          setError(null);
        }
      })
      .catch((e) => {
        if (!cancelled) setError(e instanceof Error ? e.message : "画像加载失败");
      })
      .finally(() => {
        if (!cancelled) setLoading(false);
      });
    return () => {
      cancelled = true;
    };
  }, [projectId, fileId]);

  if (loading) {
    return (
      <div className="flex h-32 items-center justify-center rounded-lg border border-dashed text-muted-foreground">
        <Loader2 className="mr-2 h-4 w-4 animate-spin" /> 生成数据画像中...
      </div>
    );
  }

  if (error || !profile) {
    return (
      <div className="rounded-lg border border-red-200 bg-red-50 p-3 text-sm text-red-700">
        数据画像加载失败：{error || "未知错误"}
      </div>
    );
  }

  const { overview } = profile;

  return (
    <div className="space-y-4">
      {/* 1. 总览 metric 卡片 */}
      <div className={cn("grid gap-3", compact ? "grid-cols-2 sm:grid-cols-4" : "grid-cols-2 md:grid-cols-4")}>
        <Card className="border-teal-100">
          <CardContent className="flex items-center gap-3 p-4">
            <div className="flex h-9 w-9 items-center justify-center rounded-md bg-teal-50">
              <Database className="h-4 w-4 text-teal-700" />
            </div>
            <div>
              <div className="text-xs text-muted-foreground">总行数</div>
              <div className="text-lg font-bold">{overview.total_rows}</div>
            </div>
          </CardContent>
        </Card>
        <Card className="border-slate-200">
          <CardContent className="flex items-center gap-3 p-4">
            <div className="flex h-9 w-9 items-center justify-center rounded-md bg-slate-100">
              <Columns3 className="h-4 w-4 text-slate-700" />
            </div>
            <div>
              <div className="text-xs text-muted-foreground">总列数</div>
              <div className="text-lg font-bold">{overview.total_cols}</div>
            </div>
          </CardContent>
        </Card>
        <Card className="border-green-100">
          <CardContent className="flex items-center gap-3 p-4">
            <div className="flex h-9 w-9 items-center justify-center rounded-lg bg-green-50">
              <CheckCircle2 className="h-4 w-4 text-green-600" />
            </div>
            <div>
              <div className="text-xs text-muted-foreground">完整度</div>
              <div className="text-lg font-bold">{overview.completeness_pct}%</div>
            </div>
          </CardContent>
        </Card>
        <Card className="border-amber-100">
          <CardContent className="flex items-center gap-3 p-4">
            <div className="flex h-9 w-9 items-center justify-center rounded-lg bg-amber-50">
              <AlertCircle className="h-4 w-4 text-amber-600" />
            </div>
            <div>
              <div className="text-xs text-muted-foreground">缺失单元格</div>
              <div className="text-lg font-bold">{overview.missing_cells}</div>
            </div>
          </CardContent>
        </Card>
      </div>

      {/* 2. 列质量诊断网格 */}
      <Card>
        <CardHeader className="pb-3">
          <CardTitle className="flex items-center gap-2 text-sm">
            <TrendingUp className="h-4 w-4 text-teal-700" />
            列质量诊断
            <Badge variant="secondary" className="text-[10px]">
              {profile.columns_profile.length} 列
            </Badge>
          </CardTitle>
        </CardHeader>
        <CardContent>
          <div className="grid gap-2.5 sm:grid-cols-2 lg:grid-cols-3 xl:grid-cols-4">
            {profile.columns_profile.map((col) => (
              <ColumnCard key={col.name} col={col} />
            ))}
          </div>
        </CardContent>
      </Card>

      {/* 3. 关键分布图 */}
      {profile.distributions.length > 0 && (
        <Card>
          <CardHeader className="pb-3">
            <CardTitle className="text-sm">关键字段分布</CardTitle>
          </CardHeader>
          <CardContent>
            <div className="grid gap-6 md:grid-cols-2">
              {profile.distributions.slice(0, 6).map((dist) => (
                <div key={dist.column} className="space-y-2">
                  <div className="flex items-center justify-between">
                    <h4 className="text-sm font-medium">{dist.column}</h4>
                    <Badge variant="outline" className="text-[10px]">
                      {dist.type === "category" ? "分类分布" : "数值直方图"}
                    </Badge>
                  </div>
                  {dist.type === "category" && dist.items.length <= 8 ? (
                    // 分类少 → 饼图
                    <ResponsiveContainer width="100%" height={200}>
                      <PieChart>
                        <Pie
                          data={dist.items}
                          dataKey="count"
                          nameKey="label"
                          cx="50%"
                          cy="50%"
                          outerRadius={70}
                          label={(entry) => `${entry.label}: ${entry.count}`}
                          labelLine={false}
                          fontSize={10}
                        >
                          {dist.items.map((_, i) => (
                            <Cell key={i} fill={PIE_COLORS[i % PIE_COLORS.length]} />
                          ))}
                        </Pie>
                        <Tooltip />
                      </PieChart>
                    </ResponsiveContainer>
                  ) : (
                    // 直方图或多分类 → 柱状图
                    <ResponsiveContainer width="100%" height={200}>
                      <BarChart
                        data={dist.items}
                        margin={{ top: 5, right: 10, left: -20, bottom: 40 }}
                      >
                        <CartesianGrid strokeDasharray="3 3" stroke="#f0f0f0" />
                        <XAxis
                          dataKey="label"
                          tick={{ fontSize: 9 }}
                          angle={-30}
                          textAnchor="end"
                          interval={0}
                        />
                        <YAxis tick={{ fontSize: 10 }} />
                        <Tooltip />
                        <Bar dataKey="count" fill="#3b82f6" radius={[4, 4, 0, 0]} />
                      </BarChart>
                    </ResponsiveContainer>
                  )}
                </div>
              ))}
            </div>
          </CardContent>
        </Card>
      )}
    </div>
  );
}
