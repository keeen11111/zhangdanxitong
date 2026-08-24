"use client";

import { useState } from "react";
import { useRouter } from "next/navigation";
import { toast } from "sonner";
import { Loader2, Wallet, Building2, Mail, Lock, User as UserIcon, ShieldCheck, Database, FileCheck2 } from "lucide-react";

import { api, setAuth } from "@/lib/api";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { Card, CardContent } from "@/components/ui/card";
import { Tabs, TabsContent, TabsList, TabsTrigger } from "@/components/ui/tabs";

export default function LoginPage() {
  const router = useRouter();
  const [loading, setLoading] = useState(false);

  // 登录表单
  const [loginEmail, setLoginEmail] = useState("demo@autopayroll.com");
  const [loginPwd, setLoginPwd] = useState("demo123");

  // 注册表单
  const [regName, setRegName] = useState("");
  const [regEmail, setRegEmail] = useState("");
  const [regPwd, setRegPwd] = useState("");
  const [regTenant, setRegTenant] = useState("");

  async function onLogin(e: React.FormEvent) {
    e.preventDefault();
    setLoading(true);
    try {
      const res = await api.login({ email: loginEmail, password: loginPwd });
      setAuth(res.access_token, res.user);
      toast.success(`欢迎回来，${res.user.name}`);
      router.replace("/projects");
    } catch (err) {
      toast.error(err instanceof Error ? err.message : "登录失败");
    } finally {
      setLoading(false);
    }
  }

  async function onRegister(e: React.FormEvent) {
    e.preventDefault();
    setLoading(true);
    try {
      const res = await api.register({
        email: regEmail,
        password: regPwd,
        name: regName,
        tenant_name: regTenant || undefined,
      });
      setAuth(res.access_token, res.user);
      toast.success(`注册成功，欢迎 ${res.user.name}`);
      router.replace("/projects");
    } catch (err) {
      toast.error(err instanceof Error ? err.message : "注册失败");
    } finally {
      setLoading(false);
    }
  }

  return (
    <main className="min-h-screen bg-slate-100 p-4 lg:p-8">
      <div className="mx-auto grid min-h-[calc(100vh-2rem)] max-w-6xl overflow-hidden rounded-lg border border-slate-200 bg-white shadow-sm lg:min-h-[calc(100vh-4rem)] lg:grid-cols-[1.05fr_0.95fr]">
        <section className="relative hidden flex-col justify-between overflow-hidden bg-[#101c2c] p-10 text-white lg:flex">
          <div className="absolute inset-x-0 top-0 h-1 bg-teal-500" />
          <div>
            <div className="flex items-center gap-3">
              <div className="flex h-10 w-10 items-center justify-center rounded-md bg-teal-600">
                <Wallet className="h-5 w-5" />
              </div>
              <div>
                <div className="text-base font-semibold">薪资数据中心</div>
                <div className="text-[10px] uppercase tracking-[0.18em] text-slate-400">Payroll Operations</div>
              </div>
            </div>

            <div className="mt-24 max-w-md">
              <p className="text-xs font-semibold uppercase tracking-[0.18em] text-teal-300">企业薪资数据工作台</p>
              <h1 className="mt-4 text-4xl font-semibold leading-tight tracking-tight">
                从原始数据到标准账单，流程清晰且结果可核验。
              </h1>
              <p className="mt-5 text-sm leading-7 text-slate-300">
                统一处理人员档案、薪资明细与考勤数据；关键缺失项在导出前强制拦截，避免不完整结果进入交付环节。
              </p>
            </div>

            <div className="mt-12 grid gap-3 text-sm text-slate-200">
              <div className="flex items-center gap-3"><Database className="h-4 w-4 text-teal-300" />多文件归集与结构化数据库</div>
              <div className="flex items-center gap-3"><ShieldCheck className="h-4 w-4 text-teal-300" />必填字段与高风险数据校验</div>
              <div className="flex items-center gap-3"><FileCheck2 className="h-4 w-4 text-teal-300" />按企业模板输出可追溯结果</div>
            </div>
          </div>
          <p className="text-xs text-slate-500">仅供获授权的企业用户访问</p>
        </section>

        <section className="flex items-center justify-center p-6 sm:p-10 lg:p-14">
          <div className="w-full max-w-md">
            <div className="mb-8 lg:hidden">
              <div className="flex items-center gap-3">
                <div className="flex h-10 w-10 items-center justify-center rounded-md bg-teal-600 text-white">
                  <Wallet className="h-5 w-5" />
                </div>
                <div>
                  <div className="font-semibold text-slate-900">薪资数据中心</div>
                  <div className="text-xs text-slate-500">企业薪资数据工作台</div>
                </div>
              </div>
            </div>

            <div className="mb-7">
              <h2 className="text-2xl font-semibold tracking-tight text-slate-900">访问工作空间</h2>
              <p className="mt-2 text-sm text-slate-500">登录现有账号，或为企业创建新的独立空间。</p>
            </div>

            <Card className="border-slate-200 shadow-none">
              <CardContent className="p-5 sm:p-6">
            <Tabs defaultValue="login">
              <TabsList className="grid w-full grid-cols-2 bg-slate-100">
                <TabsTrigger value="login">登录</TabsTrigger>
                <TabsTrigger value="register">注册</TabsTrigger>
              </TabsList>

              {/* 登录 */}
              <TabsContent value="login">
                <form onSubmit={onLogin} className="space-y-4 pt-2">
                  <div className="space-y-2">
                    <Label htmlFor="email">邮箱</Label>
                    <div className="relative">
                      <Mail className="pointer-events-none absolute left-3 top-1/2 h-4 w-4 -translate-y-1/2 text-muted-foreground" />
                      <Input
                        id="email"
                        type="email"
                        required
                        className="pl-9"
                        value={loginEmail}
                        onChange={(e) => setLoginEmail(e.target.value)}
                        placeholder="you@company.com"
                      />
                    </div>
                  </div>
                  <div className="space-y-2">
                    <Label htmlFor="pwd">密码</Label>
                    <div className="relative">
                      <Lock className="pointer-events-none absolute left-3 top-1/2 h-4 w-4 -translate-y-1/2 text-muted-foreground" />
                      <Input
                        id="pwd"
                        type="password"
                        required
                        className="pl-9"
                        value={loginPwd}
                        onChange={(e) => setLoginPwd(e.target.value)}
                        placeholder="••••••"
                      />
                    </div>
                  </div>
                  <Button type="submit" className="w-full" disabled={loading}>
                    {loading && <Loader2 className="mr-2 h-4 w-4 animate-spin" />}
                    登录
                  </Button>
                  <p className="text-center text-xs text-muted-foreground">
                    演示账号：demo@autopayroll.com / demo123
                  </p>
                </form>
              </TabsContent>

              {/* 注册 */}
              <TabsContent value="register">
                <form onSubmit={onRegister} className="space-y-4 pt-2">
                  <div className="space-y-2">
                    <Label htmlFor="rname">姓名</Label>
                    <div className="relative">
                      <UserIcon className="pointer-events-none absolute left-3 top-1/2 h-4 w-4 -translate-y-1/2 text-muted-foreground" />
                      <Input
                        id="rname"
                        required
                        className="pl-9"
                        value={regName}
                        onChange={(e) => setRegName(e.target.value)}
                        placeholder="张三"
                      />
                    </div>
                  </div>
                  <div className="space-y-2">
                    <Label htmlFor="remail">邮箱</Label>
                    <div className="relative">
                      <Mail className="pointer-events-none absolute left-3 top-1/2 h-4 w-4 -translate-y-1/2 text-muted-foreground" />
                      <Input
                        id="remail"
                        type="email"
                        required
                        className="pl-9"
                        value={regEmail}
                        onChange={(e) => setRegEmail(e.target.value)}
                        placeholder="you@company.com"
                      />
                    </div>
                  </div>
                  <div className="space-y-2">
                    <Label htmlFor="rpwd">密码（至少 6 位）</Label>
                    <div className="relative">
                      <Lock className="pointer-events-none absolute left-3 top-1/2 h-4 w-4 -translate-y-1/2 text-muted-foreground" />
                      <Input
                        id="rpwd"
                        type="password"
                        required
                        minLength={6}
                        className="pl-9"
                        value={regPwd}
                        onChange={(e) => setRegPwd(e.target.value)}
                        placeholder="••••••"
                      />
                    </div>
                  </div>
                  <div className="space-y-2">
                    <Label htmlFor="rtenant">企业/租户名（可选）</Label>
                    <div className="relative">
                      <Building2 className="pointer-events-none absolute left-3 top-1/2 h-4 w-4 -translate-y-1/2 text-muted-foreground" />
                      <Input
                        id="rtenant"
                        className="pl-9"
                        value={regTenant}
                        onChange={(e) => setRegTenant(e.target.value)}
                        placeholder="留空则用邮箱域名"
                      />
                    </div>
                  </div>
                  <Button type="submit" className="w-full" disabled={loading}>
                    {loading && <Loader2 className="mr-2 h-4 w-4 animate-spin" />}
                    注册并登录
                  </Button>
                </form>
              </TabsContent>
            </Tabs>
              </CardContent>
            </Card>

            <p className="mt-6 text-center text-xs text-slate-400">
              登录即表示你已获得当前企业空间的访问授权
            </p>
          </div>
        </section>
      </div>
    </main>
  );
}
