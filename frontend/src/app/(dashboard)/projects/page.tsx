"use client";

import { useEffect, useState } from "react";
import Link from "next/link";
import { useRouter } from "next/navigation";
import { toast } from "sonner";
import {
  Plus,
  FolderKanban,
  Calendar,
  FileText,
  Trash2,
  ArrowRight,
  Loader2,
} from "lucide-react";

import { api, Project } from "@/lib/api";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { Badge } from "@/components/ui/badge";
import { Card, CardContent } from "@/components/ui/card";
import {
  Dialog,
  DialogContent,
  DialogHeader,
  DialogTitle,
  DialogDescription,
  DialogFooter,
  DialogTrigger,
} from "@/components/ui/dialog";

const STATUS_LABEL: Record<string, { text: string; cls: string }> = {
  import: { text: "数据导入", cls: "bg-slate-100 text-slate-700" },
  config: { text: "字段映射", cls: "bg-teal-50 text-teal-700" },
  work: { text: "诊断补齐", cls: "bg-amber-100 text-amber-700" },
  export: { text: "可导出", cls: "bg-green-100 text-green-700" },
};

export default function ProjectsPage() {
  const router = useRouter();
  const [projects, setProjects] = useState<Project[]>([]);
  const [loading, setLoading] = useState(true);
  const [open, setOpen] = useState(false);
  const [creating, setCreating] = useState(false);
  const [name, setName] = useState("");
  const [month, setMonth] = useState("2026.06");

  async function load() {
    try {
      const list = await api.listProjects();
      setProjects(list);
    } catch (e) {
      toast.error(e instanceof Error ? e.message : "加载失败");
    } finally {
      setLoading(false);
    }
  }

  useEffect(() => {
    load();
  }, []);

  async function onCreate(e: React.FormEvent) {
    e.preventDefault();
    setCreating(true);
    try {
      const p = await api.createProject({ name, salary_month: month });
      toast.success("项目已创建");
      setOpen(false);
      router.push(`/projects/${p.id}`);
    } catch (err) {
      toast.error(err instanceof Error ? err.message : "创建失败");
    } finally {
      setCreating(false);
    }
  }

  async function onDelete(p: Project) {
    if (!confirm(`确认删除项目「${p.name}」？此操作不可撤销。`)) return;
    try {
      await api.deleteProject(p.id);
      toast.success("已删除");
      load();
    } catch (e) {
      toast.error(e instanceof Error ? e.message : "删除失败");
    }
  }

  return (
    <div className="mx-auto max-w-7xl space-y-6">
      {/* 页头 */}
      <div className="flex flex-col gap-4 border-b border-slate-200 pb-5 sm:flex-row sm:items-end sm:justify-between">
        <div>
          <p className="mb-1 text-xs font-semibold uppercase tracking-[0.16em] text-teal-700">Payroll workspace</p>
          <h1 className="text-2xl font-semibold tracking-tight text-slate-950">薪资项目</h1>
          <p className="mt-1.5 text-sm text-muted-foreground">集中管理每个月度薪资批次及其处理进度</p>
        </div>
        <Dialog open={open} onOpenChange={setOpen}>
          <DialogTrigger asChild>
            <Button className="shadow-sm">
              <Plus className="mr-2 h-4 w-4" />
              新建项目
            </Button>
          </DialogTrigger>
          <DialogContent>
            <DialogHeader>
              <DialogTitle>新建薪资项目</DialogTitle>
              <DialogDescription>为本月薪资批次创建一个处理空间</DialogDescription>
            </DialogHeader>
            <form onSubmit={onCreate} className="space-y-4 py-2">
              <div className="space-y-2">
                <Label htmlFor="pname">项目名称</Label>
                <Input
                  id="pname"
                  required
                  value={name}
                  onChange={(e) => setName(e.target.value)}
                  placeholder="如：2026年6月薪资批次"
                />
              </div>
              <div className="space-y-2">
                <Label htmlFor="pmonth">薪资月份</Label>
                <Input
                  id="pmonth"
                  required
                  value={month}
                  onChange={(e) => setMonth(e.target.value)}
                  placeholder="2026.06"
                />
              </div>
              <DialogFooter>
                <Button type="button" variant="outline" onClick={() => setOpen(false)}>
                  取消
                </Button>
                <Button type="submit" disabled={creating}>
                  {creating && <Loader2 className="mr-2 h-4 w-4 animate-spin" />}
                  创建并开始
                </Button>
              </DialogFooter>
            </form>
          </DialogContent>
        </Dialog>
      </div>

      {!loading && projects.length > 0 && (
        <div className="grid overflow-hidden rounded-lg border bg-white sm:grid-cols-3 sm:divide-x">
          <div className="border-b px-5 py-4 sm:border-b-0">
            <div className="text-xs font-medium text-slate-500">项目总数</div>
            <div className="mt-1 text-2xl font-semibold text-slate-900">{projects.length}</div>
          </div>
          <div className="border-b px-5 py-4 sm:border-b-0">
            <div className="text-xs font-medium text-slate-500">已导入文件</div>
            <div className="mt-1 text-2xl font-semibold text-slate-900">{projects.reduce((sum, project) => sum + project.file_count, 0)}</div>
          </div>
          <div className="px-5 py-4">
            <div className="text-xs font-medium text-slate-500">最近薪资月份</div>
            <div className="mt-1 text-2xl font-semibold text-slate-900">{projects[0]?.salary_month || "—"}</div>
          </div>
        </div>
      )}

      {/* 项目列表 */}
      {loading ? (
        <div className="flex h-52 items-center justify-center rounded-lg border bg-white text-muted-foreground" aria-busy="true">
          <Loader2 className="mr-2 h-5 w-5 animate-spin" /> 加载中...
        </div>
      ) : projects.length === 0 ? (
        <Card className="border-dashed bg-white shadow-none">
          <CardContent className="flex flex-col items-center justify-center py-16 text-center">
            <div className="mb-4 flex h-12 w-12 items-center justify-center rounded-lg border bg-slate-50">
              <FolderKanban className="h-5 w-5 text-slate-500" />
            </div>
            <h2 className="text-sm font-semibold text-slate-900">还没有薪资项目</h2>
            <p className="mt-1 text-sm text-muted-foreground">创建第一个月度批次，开始导入和处理数据</p>
          </CardContent>
        </Card>
      ) : (
        <section className="overflow-hidden rounded-lg border bg-white" aria-labelledby="project-list-title">
          <div className="flex items-center justify-between border-b px-4 py-3 sm:px-5">
            <div>
              <h2 id="project-list-title" className="text-sm font-semibold text-slate-900">项目列表</h2>
              <p className="mt-0.5 text-xs text-slate-500">选择一个项目继续处理</p>
            </div>
            <Badge variant="outline" className="font-normal text-slate-500">{projects.length} 个项目</Badge>
          </div>
          <div className="hidden grid-cols-[minmax(0,1fr)_120px_120px_120px_48px] border-b bg-slate-50 px-5 py-2 text-[11px] font-medium uppercase tracking-wider text-slate-500 md:grid">
            <span>项目</span><span>薪资月份</span><span>文件</span><span>状态</span><span />
          </div>
          <div className="divide-y">
            {projects.map((p) => {
              const st = STATUS_LABEL[p.status] || STATUS_LABEL.import;
              return (
                <div key={p.id} className="group flex items-center transition-colors hover:bg-slate-50/80">
                  <Link href={`/projects/${p.id}`} className="grid min-w-0 flex-1 grid-cols-[minmax(0,1fr)_auto] items-center gap-3 px-4 py-4 sm:px-5 md:grid-cols-[minmax(0,1fr)_120px_120px_120px]">
                    <span className="flex min-w-0 items-center gap-3">
                      <span className="flex h-9 w-9 shrink-0 items-center justify-center rounded-md border bg-white text-teal-700">
                        <FileText className="h-4 w-4" />
                      </span>
                      <span className="min-w-0">
                        <span className="block truncate text-sm font-semibold text-slate-900">{p.name}</span>
                        <span className="mt-0.5 block text-xs text-slate-500 md:hidden">{p.salary_month} · {p.file_count} 个文件</span>
                      </span>
                    </span>
                    <ArrowRight className="h-4 w-4 text-slate-300 transition-transform group-hover:translate-x-0.5 group-hover:text-teal-700 md:hidden" />
                    <span className="hidden items-center gap-1.5 text-xs text-slate-600 md:flex"><Calendar className="h-3.5 w-3.5" />{p.salary_month}</span>
                    <span className="hidden text-xs text-slate-600 md:block">{p.file_count} 个文件</span>
                    <span className="hidden md:block"><Badge variant="secondary" className={st.cls}>{st.text}</Badge></span>
                  </Link>
                  <button
                    onClick={() => onDelete(p)}
                    className="mr-3 flex h-8 w-8 shrink-0 items-center justify-center rounded-md text-slate-400 transition-colors hover:bg-red-50 hover:text-red-600 sm:mr-4"
                    aria-label={`删除项目 ${p.name}`}
                    title="删除项目"
                  >
                    <Trash2 className="h-4 w-4" />
                  </button>
                </div>
              );
            })}
          </div>
        </section>
      )}
    </div>
  );
}
