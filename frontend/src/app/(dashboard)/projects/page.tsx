"use client";

import { useEffect, useMemo, useState } from "react";
import Link from "next/link";
import { useRouter, useSearchParams } from "next/navigation";
import { ArrowRight, CalendarDays, CheckCircle2, Clock3, FileSpreadsheet, FolderKanban, Loader2, Plus, Search, X } from "lucide-react";
import { toast } from "sonner";

import { api, getStoredUser, Project } from "@/lib/api";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";

function dateLabel(value?: string | null) {
  if (!value) return "暂无记录";
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return value.replace("T", " ").slice(0, 16);
  return new Intl.DateTimeFormat("zh-CN", { month: "2-digit", day: "2-digit", hour: "2-digit", minute: "2-digit", hour12: false }).format(date);
}

function statusFor(project: Project) {
  if (project.pending_issue_count > 0) return { label: `${project.pending_issue_count} 项待确认`, tone: "text-amber-700", dot: "bg-amber-500" };
  if (project.has_result) return { label: "已完成", tone: "text-emerald-700", dot: "bg-emerald-500" };
  if (!project.file_count) return { label: "等待文件", tone: "text-slate-500", dot: "bg-slate-300" };
  return { label: "可以开始分析", tone: "text-blue-700", dot: "bg-blue-500" };
}

export default function ProjectsPage() {
  const router = useRouter();
  const searchParams = useSearchParams();
  const user = getStoredUser();
  const [projects, setProjects] = useState<Project[]>([]);
  const [query, setQuery] = useState("");
  const [loading, setLoading] = useState(true);
  const [creating, setCreating] = useState(false);
  const [showCreate, setShowCreate] = useState(false);
  const [name, setName] = useState("");
  const [month, setMonth] = useState(() => {
    const now = new Date();
    return `${now.getFullYear()}.${String(now.getMonth() + 1).padStart(2, "0")}`;
  });

  async function load() {
    setLoading(true);
    try { setProjects(await api.listProjects()); }
    catch (error) { toast.error(error instanceof Error ? error.message : "项目列表加载失败"); }
    finally { setLoading(false); }
  }

  useEffect(() => { void load(); }, []);

  useEffect(() => {
    if (searchParams.get("create") !== "1") return;
    setShowCreate(true);
    router.replace("/projects");
  }, [router, searchParams]);

  async function createProject(event: React.FormEvent) {
    event.preventDefault();
    const trimmed = name.trim();
    if (!trimmed) return;
    setCreating(true);
    try {
      const project = await api.createProject({ name: trimmed, salary_month: month.trim() });
      toast.success("已创建月度处理");
      router.push(`/projects/${project.id}/agent`);
    } catch (error) { toast.error(error instanceof Error ? error.message : "创建失败"); }
    finally { setCreating(false); }
  }

  const visible = useMemo(() => {
    const needle = query.trim().toLowerCase();
    return projects.filter((project) => !needle || `${project.name} ${project.salary_month}`.toLowerCase().includes(needle));
  }, [projects, query]);
  const pending = projects.reduce((sum, item) => sum + item.pending_issue_count, 0);
  const completed = projects.filter((item) => item.has_result && item.pending_issue_count === 0).length;

  return <div className="payroll-page max-w-[1180px]">
    <header className="flex flex-col gap-5 border-b border-slate-200/90 pb-7 sm:flex-row sm:items-end sm:justify-between">
      <div><p className="payroll-kicker">Financial workspace</p><h1 className="page-heading mt-2">项目</h1><p className="mt-2 max-w-xl text-sm leading-6 text-slate-600">选择项目，继续上次的月度处理，或从一批新文件开始。</p>{user?.tenant_name ? <p className="mt-3 text-xs text-slate-400">当前工作区：{user.tenant_name}</p> : null}</div>
      <Button type="button" className="h-10 bg-blue-700 px-4 text-sm text-white shadow-sm hover:bg-blue-800" onClick={() => setShowCreate(true)}><Plus className="mr-2 h-4 w-4" />新建月度处理</Button>
    </header>

    <section className="metric-strip grid sm:grid-cols-3 sm:divide-x sm:divide-slate-200/90" aria-label="工作区摘要"><div className="border-b border-slate-200/90 px-5 py-4 sm:border-b-0"><p className="text-xs text-slate-500">项目处理</p><p className="mt-1 text-2xl font-semibold tabular-nums text-slate-950">{projects.length}</p></div><div className="border-b border-slate-200/90 px-5 py-4 sm:border-b-0"><p className="text-xs text-slate-500">待回答问题</p><p className={pending ? "mt-1 text-2xl font-semibold tabular-nums text-amber-700" : "mt-1 text-2xl font-semibold tabular-nums text-slate-950"}>{pending}</p></div><div className="px-5 py-4"><p className="text-xs text-slate-500">已完成批次</p><p className="mt-1 text-2xl font-semibold tabular-nums text-emerald-700">{completed}</p></div></section>

    <div className="mt-9 flex flex-col gap-3 sm:flex-row sm:items-center sm:justify-between"><div><h2 className="text-base font-semibold text-slate-950">最近项目</h2><p className="mt-1 text-xs text-slate-500">每个项目拥有独立的对话、上下文和处理经验。</p></div><label className="relative block sm:w-72"><span className="sr-only">搜索项目</span><Search className="pointer-events-none absolute left-3 top-1/2 h-4 w-4 -translate-y-1/2 text-slate-400" /><Input value={query} onChange={(event) => setQuery(event.target.value)} placeholder="搜索项目或月份" className="h-10 border-slate-300 bg-white pl-9 shadow-sm" /></label></div>

    {loading ? <div className="mt-4 flex h-48 items-center justify-center rounded-lg border border-slate-200/90 bg-white text-sm text-slate-500 shadow-[0_1px_2px_rgba(15,23,42,0.03)]" aria-busy="true"><Loader2 className="mr-2 h-4 w-4 animate-spin text-blue-700" />正在加载项目...</div> : null}
    {!loading && !visible.length ? <section className="mt-4 rounded-lg border border-dashed border-slate-300 bg-white px-6 py-16 text-center shadow-[0_1px_2px_rgba(15,23,42,0.03)]"><FolderKanban className="mx-auto h-8 w-8 text-slate-300" /><h2 className="mt-4 text-sm font-semibold text-slate-900">{projects.length ? "没有匹配的项目" : "还没有项目处理记录"}</h2><p className="mt-1 text-sm text-slate-500">{projects.length ? "换个搜索条件试试。" : "创建一次月度处理，Agent 会在对话中完成剩余步骤。"}</p>{!projects.length ? <Button type="button" className="mt-5 bg-blue-700 hover:bg-blue-800" onClick={() => setShowCreate(true)}><Plus className="mr-2 h-4 w-4" />创建第一个项目</Button> : null}</section> : null}
    {!loading && visible.length ? <section className="mt-4 overflow-hidden rounded-lg border border-slate-200/90 bg-white shadow-[0_1px_2px_rgba(15,23,42,0.03)]" aria-label="项目列表"><div className="divide-y divide-slate-200/90">{visible.map((project) => { const state = statusFor(project); return <Link key={project.id} href={`/projects/${project.id}/agent`} className="group flex min-w-0 items-center gap-4 px-4 py-4 transition-colors hover:bg-blue-50/40 focus-visible:relative focus-visible:z-10 sm:px-5"><span className="flex h-10 w-10 shrink-0 items-center justify-center rounded-lg bg-blue-50 text-blue-700"><FolderKanban className="h-5 w-5" /></span><span className="min-w-0 flex-1"><span className="flex flex-wrap items-center gap-x-3 gap-y-1"><span className="truncate text-sm font-semibold text-slate-950">{project.name}</span><span className="inline-flex items-center gap-1.5 text-xs text-slate-500"><CalendarDays className="h-3.5 w-3.5" />所属月 {project.salary_month}</span></span><span className="mt-1.5 flex flex-wrap items-center gap-x-4 gap-y-1 text-xs text-slate-500"><span className={`inline-flex items-center gap-1.5 font-medium ${state.tone}`}><span className={`h-1.5 w-1.5 rounded-full ${state.dot}`} />{state.label}</span><span className="inline-flex items-center gap-1.5"><FileSpreadsheet className="h-3.5 w-3.5" />{project.file_count} 个文件</span><span className="inline-flex items-center gap-1.5"><Clock3 className="h-3.5 w-3.5" />最近 {dateLabel(project.updated_at || project.created_at)}</span></span></span><span className="hidden text-xs font-medium text-blue-700 sm:inline-flex sm:items-center sm:gap-1 opacity-0 transition-opacity group-hover:opacity-100">进入对话<ArrowRight className="h-3.5 w-3.5" /></span><ArrowRight className="h-4 w-4 shrink-0 text-slate-300 sm:hidden" /></Link>; })}</div></section> : null}

    {showCreate ? <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-950/45 p-4 backdrop-blur-[2px]" role="dialog" aria-modal="true" aria-labelledby="new-run-title"><div className="w-full max-w-md rounded-xl border border-slate-200 bg-white shadow-2xl"><div className="flex items-start justify-between border-b border-slate-200 px-5 py-4"><div><h2 id="new-run-title" className="text-base font-semibold text-slate-950">新建月度处理</h2><p className="mt-1 text-xs text-slate-500">先建立项目上下文，进入对话后再上传文件。</p></div><button type="button" className="inline-flex h-9 w-9 cursor-pointer items-center justify-center rounded-lg text-slate-400 hover:bg-slate-100 hover:text-slate-700" onClick={() => setShowCreate(false)} aria-label="关闭"><X className="h-4 w-4" /></button></div><form onSubmit={createProject} className="space-y-4 px-5 py-5"><div><label htmlFor="project-name" className="mb-1.5 block text-sm font-medium text-slate-700">项目名称</label><Input id="project-name" value={name} onChange={(event) => setName(event.target.value)} placeholder="例如：北京科园" required autoFocus /></div><div><label htmlFor="salary-month" className="mb-1.5 block text-sm font-medium text-slate-700">所属月</label><Input id="salary-month" value={month} onChange={(event) => setMonth(event.target.value)} placeholder="2026.07" required /></div><div className="flex justify-end gap-2 border-t border-slate-100 pt-4"><Button type="button" variant="outline" onClick={() => setShowCreate(false)}>取消</Button><Button type="submit" className="bg-blue-700 hover:bg-blue-800" disabled={creating}>{creating ? <Loader2 className="mr-2 h-4 w-4 animate-spin" /> : <CheckCircle2 className="mr-2 h-4 w-4" />}创建并进入对话</Button></div></form></div></div> : null}
  </div>;
}
