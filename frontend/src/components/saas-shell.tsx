"use client";

import { useEffect, useState } from "react";
import Link from "next/link";
import { usePathname, useRouter } from "next/navigation";
import { ChevronLeft, ChevronRight, CircleHelp, FolderKanban, LogOut, Menu, Plus, Settings2, X } from "lucide-react";
import { toast } from "sonner";

import { api, clearAuth, getStoredUser, Project } from "@/lib/api";
import { cn } from "@/lib/utils";
import { mergeRecentProject, PROJECT_CREATED_EVENT } from "@/lib/project-list";
import { AgentMascot } from "@/components/agent-mascot";

const navItems = [
  { href: "/projects", label: "项目", icon: FolderKanban },
] as const;

export function SaaSShell({ children }: { children: React.ReactNode }) {
  const pathname = usePathname();
  const router = useRouter();
  const user = getStoredUser();
  const [projects, setProjects] = useState<Project[]>([]);
  const [collapsed, setCollapsed] = useState(false);
  const [mobileOpen, setMobileOpen] = useState(false);
  const [online, setOnline] = useState<boolean | null>(null);

  useEffect(() => {
    let active = true;
    void Promise.allSettled([api.health(), api.listProjects()]).then(([health, companyList]) => {
      if (!active) return;
      setOnline(health.status === "fulfilled" && health.value.status === "ok");
      if (companyList.status === "fulfilled") setProjects(companyList.value.slice(0, 8));
    });
    return () => { active = false; };
  }, [pathname]);

  useEffect(() => {
    const handleProjectCreated = (event: Event) => {
      const project = (event as CustomEvent<Project>).detail;
      if (!project?.id) return;
      setProjects((current) => mergeRecentProject(current, project));
    };
    window.addEventListener(PROJECT_CREATED_EVENT, handleProjectCreated);
    return () => window.removeEventListener(PROJECT_CREATED_EVENT, handleProjectCreated);
  }, []);

  useEffect(() => setMobileOpen(false), [pathname]);

  function logout() {
    clearAuth();
    toast.success("已退出登录");
    router.replace("/login");
  }

  const initial = user?.name?.slice(0, 1) || user?.email?.slice(0, 1).toUpperCase() || "U";
  const currentProjectId = pathname.match(/\/projects\/([^/]+)/)?.[1];
  const isAgentRoute = pathname.includes("/agent");

  return (
    <div className="app-shell flex min-h-screen text-slate-900">
      <a href="#main-content" className="sr-only z-[70] rounded-md bg-white px-3 py-2 text-sm font-medium text-slate-900 focus:not-sr-only focus:fixed focus:left-3 focus:top-3">跳到主要内容</a>
      <button type="button" aria-label="打开导航" className="fixed left-3 top-3 z-40 inline-flex h-10 w-10 cursor-pointer items-center justify-center rounded-lg border border-slate-200 bg-white text-slate-700 shadow-sm lg:hidden" onClick={() => setMobileOpen(true)}><Menu className="h-4 w-4" /></button>
      {mobileOpen ? <button type="button" aria-label="关闭导航" className="fixed inset-0 z-40 bg-slate-950/20 lg:hidden" onClick={() => setMobileOpen(false)} /> : null}
      <aside className={cn("app-sidebar fixed inset-y-0 left-0 z-50 flex w-[240px] flex-col border-r border-slate-200 text-slate-700 transition-transform duration-200 lg:sticky lg:top-0 lg:z-20 lg:h-screen lg:translate-x-0", mobileOpen ? "translate-x-0" : "-translate-x-full", collapsed && "lg:w-[68px]")} aria-label="主导航">
        <div className="flex h-[68px] items-center justify-between border-b border-slate-200 px-4">
          <Link href="/projects" className={cn("flex min-w-0 items-center gap-3", collapsed && "lg:mx-auto")}><AgentMascot className="h-8 w-6" priority /><span className={cn("truncate text-[15px] font-semibold tracking-tight text-slate-950", collapsed && "lg:hidden")}>财务 Agent</span></Link>
          <button type="button" className="inline-flex h-9 w-9 cursor-pointer items-center justify-center rounded-lg text-slate-500 hover:bg-slate-100 hover:text-slate-900 lg:hidden" onClick={() => setMobileOpen(false)} aria-label="关闭导航"><X className="h-4 w-4" /></button>
        </div>
        <div className="px-3 py-4"><Link href="/projects?create=1" className={cn("flex h-10 items-center gap-2 rounded-lg bg-blue-600 px-3 text-sm font-semibold text-white shadow-[0_6px_16px_rgba(37,99,235,0.2)] transition-colors hover:bg-blue-500", collapsed && "lg:justify-center lg:px-0")}><Plus className="h-4 w-4" /><span className={collapsed ? "lg:hidden" : undefined}>新建月度处理</span></Link></div>
        <nav className="space-y-1 px-3" aria-label="工作区"><p className={cn("mb-2 px-3 text-[10px] font-semibold uppercase tracking-[0.16em] text-slate-500", collapsed && "lg:hidden")}>工作台</p>{navItems.map((item) => { const Icon = item.icon; const active = item.href === "/projects" ? pathname.startsWith("/projects") : pathname.startsWith(item.href); return <Link key={item.href} href={item.href} aria-current={active ? "page" : undefined} title={collapsed ? item.label : undefined} className={cn("flex h-10 items-center gap-3 rounded-lg px-3 text-sm font-medium transition-colors", active ? "bg-blue-50 text-blue-900 shadow-[inset_3px_0_0_#60a5fa]" : "text-slate-600 hover:bg-slate-100 hover:text-slate-950", collapsed && "lg:justify-center lg:px-0")}><Icon className={cn("h-[17px] w-[17px]", active ? "text-blue-600" : "text-slate-500")} /><span className={collapsed ? "lg:hidden" : undefined}>{item.label}</span></Link>; })}</nav>
        <div className={cn("mt-6 min-h-0 flex-1 overflow-y-auto px-3", collapsed && "lg:hidden")}><div className="mb-2 flex items-center justify-between px-3 text-[10px] font-semibold uppercase tracking-[0.16em] text-slate-500"><span>最近项目</span><Link href="/projects" className="cursor-pointer text-slate-500 hover:text-blue-700" aria-label="查看全部项目"><ChevronRight className="h-3.5 w-3.5" /></Link></div><div className="space-y-1">{projects.map((project) => <Link key={project.id} href={`/projects/${project.id}/agent`} className={cn("flex items-center gap-2 rounded-lg px-3 py-2 text-sm text-slate-600 transition-colors hover:bg-slate-100 hover:text-slate-950", currentProjectId === project.id && "bg-blue-50 text-blue-900")}><span className="flex h-6 w-6 shrink-0 items-center justify-center rounded-md bg-slate-100 text-[11px] font-semibold text-slate-600">{project.name.slice(0, 1)}</span><span className="min-w-0 flex-1 truncate">{project.name}</span><span className="font-mono text-[10px] text-slate-500">{project.salary_month}</span></Link>)}{!projects.length ? <p className="px-3 py-3 text-xs leading-5 text-slate-500">还没有项目，先创建一次月度处理。</p> : null}</div></div>
        <div className="border-t border-slate-200 p-3"><div className={cn("flex items-center gap-2 px-2 py-2", collapsed && "lg:justify-center lg:px-0")}><span className="flex h-8 w-8 shrink-0 items-center justify-center rounded-full bg-blue-50 text-xs font-semibold text-blue-700">{initial}</span><div className={cn("min-w-0 flex-1", collapsed && "lg:hidden")}><p className="truncate text-sm font-medium text-slate-900">{user?.name || "当前用户"}</p><p className="truncate text-xs text-slate-500">{user?.tenant_name || "财务工作区"}</p></div><button type="button" onClick={logout} className={cn("inline-flex h-9 w-9 cursor-pointer items-center justify-center rounded-lg text-slate-500 hover:bg-red-50 hover:text-red-700", collapsed && "lg:hidden")} aria-label="退出登录" title="退出登录"><LogOut className="h-4 w-4" /></button></div><div className={cn("mt-1 flex items-center justify-between px-2 text-[11px] text-slate-500", collapsed && "lg:hidden")}><span className="inline-flex items-center gap-1.5"><span className={cn("h-1.5 w-1.5 rounded-full", online === true ? "bg-emerald-500" : online === false ? "bg-red-500" : "bg-slate-400")} />{online === true ? "服务正常" : online === false ? "服务未连接" : "检查服务"}</span><span className="inline-flex items-center gap-2"><Settings2 className="h-3.5 w-3.5" /><CircleHelp className="h-3.5 w-3.5" /></span></div><button type="button" className="mt-2 hidden h-8 w-full cursor-pointer items-center justify-center border-t border-slate-200 pt-2 text-slate-500 hover:text-slate-900 lg:flex" onClick={() => setCollapsed((value) => !value)} aria-label={collapsed ? "展开侧栏" : "收起侧栏"}>{collapsed ? <ChevronRight className="h-4 w-4" /> : <ChevronLeft className="h-4 w-4" />}</button></div>
      </aside>
      <main id="main-content" className={cn("min-w-0 flex-1", isAgentRoute && "overflow-hidden bg-white")}><div className={cn("mx-auto w-full max-w-[1500px]", isAgentRoute ? "h-screen" : "min-h-screen px-4 pb-8 pt-16 sm:px-6 lg:px-10 lg:pt-8")}>{children}</div></main>
    </div>
  );
}
