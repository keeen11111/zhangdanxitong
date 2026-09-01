"use client";

import { useState } from "react";
import { useRouter } from "next/navigation";
import { Bot, Building2, CheckCircle2, Loader2, LockKeyhole, Mail, UserRound } from "lucide-react";
import { toast } from "sonner";

import { api, setAuth } from "@/lib/api";
import { AgentMascot } from "@/components/agent-mascot";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";

export default function LoginPage() {
  const router = useRouter();
  const [mode, setMode] = useState<"login" | "register">("login");
  const [loading, setLoading] = useState(false);
  const [email, setEmail] = useState("demo@autopayroll.com");
  const [password, setPassword] = useState("demo123");
  const [name, setName] = useState("");
  const [tenant, setTenant] = useState("");

  async function submit(event: React.FormEvent) {
    event.preventDefault();
    setLoading(true);
    try {
      const result = mode === "login"
        ? await api.login({ email, password })
        : await api.register({ email, password, name, tenant_name: tenant || undefined });
      setAuth(result.access_token, result.user);
      toast.success(mode === "login" ? "欢迎回来" : "注册成功");
      router.replace("/projects");
    } catch (error) { toast.error(error instanceof Error ? error.message : "操作失败"); }
    finally { setLoading(false); }
  }

  return (
    <main className="login-canvas flex min-h-screen items-center justify-center px-4 py-8 sm:px-6 sm:py-12">
      <div className="grid w-full max-w-[980px] overflow-hidden rounded-xl border border-slate-200 bg-white shadow-[0_18px_60px_rgba(15,23,42,0.10)] lg:grid-cols-[1fr_420px]">
        <section className="hidden border-r border-blue-100 bg-blue-50/70 p-10 lg:flex lg:flex-col lg:justify-between xl:p-12">
          <div>
            <div className="flex items-center gap-3"><AgentMascot className="h-8 w-6" priority /><span className="text-base font-semibold tracking-tight text-slate-950">财务 Agent</span></div>
            <p className="mt-20 text-[11px] font-semibold uppercase tracking-[0.18em] text-blue-700">Payroll operations workspace</p>
            <h1 className="mt-4 max-w-sm text-[32px] font-semibold leading-[1.18] tracking-[-0.03em] text-slate-950">让每一次工资更新，都有依据可追溯。</h1>
            <p className="mt-5 max-w-sm text-sm leading-7 text-slate-600">把总表、来源文件和操作手册交给 Agent。系统先完成低风险处理，只在需要判断时把清晰的选择交给你。</p>
          </div>
          <div className="space-y-3 text-sm text-slate-600"><p className="flex items-center gap-2.5"><CheckCircle2 className="h-4 w-4 text-emerald-600" />公司记忆隔离，处理经验持续沉淀</p><p className="flex items-center gap-2.5"><CheckCircle2 className="h-4 w-4 text-emerald-600" />每次写入都有文件和单元格依据</p><p className="flex items-center gap-2.5"><CheckCircle2 className="h-4 w-4 text-emerald-600" />正式发布前保留风险检查</p></div>
        </section>
        <section className="p-6 sm:p-10">
          <div className="mb-8 flex items-center gap-3 lg:hidden"><AgentMascot className="h-8 w-6" priority /><span className="text-base font-semibold tracking-tight text-slate-950">财务 Agent</span></div>
          <div><p className="payroll-kicker">Secure workspace</p><h2 className="mt-2 text-2xl font-semibold tracking-tight text-slate-950">{mode === "login" ? "登录工作区" : "创建工作区"}</h2><p className="mt-2 text-sm leading-6 text-slate-500">{mode === "login" ? "继续处理公司的月度文件。" : "创建账号后开始管理公司处理。"}</p></div>
          <div className="mt-7 grid grid-cols-2 border-b border-slate-200"><button type="button" onClick={() => setMode("login")} className={`cursor-pointer border-b-2 py-3 text-sm font-medium transition-colors ${mode === "login" ? "border-blue-700 text-blue-700" : "border-transparent text-slate-500 hover:text-slate-800"}`}>登录</button><button type="button" onClick={() => setMode("register")} className={`cursor-pointer border-b-2 py-3 text-sm font-medium transition-colors ${mode === "register" ? "border-blue-700 text-blue-700" : "border-transparent text-slate-500 hover:text-slate-800"}`}>注册</button></div>
          <form onSubmit={submit} className="mt-6 space-y-4">
            {mode === "register" ? <><div><label htmlFor="name" className="mb-1.5 block text-sm font-medium text-slate-700">姓名</label><div className="relative"><UserRound className="pointer-events-none absolute left-3 top-1/2 h-4 w-4 -translate-y-1/2 text-slate-400" /><Input id="name" value={name} onChange={(event) => setName(event.target.value)} className="h-11 pl-9" placeholder="张三" required /></div></div><div><label htmlFor="tenant" className="mb-1.5 block text-sm font-medium text-slate-700">工作区名称</label><div className="relative"><Building2 className="pointer-events-none absolute left-3 top-1/2 h-4 w-4 -translate-y-1/2 text-slate-400" /><Input id="tenant" value={tenant} onChange={(event) => setTenant(event.target.value)} className="h-11 pl-9" placeholder="例如：财务共享中心" /></div></div></> : null}
            <div><label htmlFor="email" className="mb-1.5 block text-sm font-medium text-slate-700">邮箱</label><div className="relative"><Mail className="pointer-events-none absolute left-3 top-1/2 h-4 w-4 -translate-y-1/2 text-slate-400" /><Input id="email" type="email" autoComplete="email" value={email} onChange={(event) => setEmail(event.target.value)} className="h-11 pl-9" placeholder="you@company.com" required /></div></div><div><label htmlFor="password" className="mb-1.5 block text-sm font-medium text-slate-700">密码</label><div className="relative"><LockKeyhole className="pointer-events-none absolute left-3 top-1/2 h-4 w-4 -translate-y-1/2 text-slate-400" /><Input id="password" type="password" autoComplete={mode === "login" ? "current-password" : "new-password"} value={password} onChange={(event) => setPassword(event.target.value)} className="h-11 pl-9" placeholder="请输入密码" minLength={mode === "register" ? 6 : undefined} required /></div></div><Button type="submit" className="mt-2 h-11 w-full bg-blue-700 text-white shadow-sm hover:bg-blue-800" disabled={loading}>{loading ? <Loader2 className="mr-2 h-4 w-4 animate-spin" /> : <Bot className="mr-2 h-4 w-4" />}{mode === "login" ? "进入工作区" : "创建并进入"}</Button>
          </form>
          <p className="mt-6 text-center text-xs leading-5 text-slate-400">文件只在你的工作区内处理。正式结果发布前会保留完整审计记录。</p>
        </section>
      </div>
    </main>
  );
}
