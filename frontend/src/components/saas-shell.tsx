"use client";

import { useEffect, useState } from "react";
import Link from "next/link";
import { useRouter } from "next/navigation";
import { ChevronDown, FileSpreadsheet, LogOut } from "lucide-react";
import { toast } from "sonner";

import { api, clearAuth, getStoredUser } from "@/lib/api";
import { Avatar, AvatarFallback } from "@/components/ui/avatar";
import { Button } from "@/components/ui/button";
import {
  DropdownMenu,
  DropdownMenuContent,
  DropdownMenuItem,
  DropdownMenuLabel,
  DropdownMenuSeparator,
  DropdownMenuTrigger,
} from "@/components/ui/dropdown-menu";

export function SaaSShell({ children }: { children: React.ReactNode }) {
  const router = useRouter();
  const user = getStoredUser();
  const [serviceOnline, setServiceOnline] = useState<boolean | null>(null);

  useEffect(() => {
    let active = true;
    async function checkService() {
      try {
        const result = await api.health();
        if (active) setServiceOnline(result.status === "ok");
      } catch {
        if (active) setServiceOnline(false);
      }
    }
    void checkService();
    const timer = window.setInterval(checkService, 30_000);
    return () => {
      active = false;
      window.clearInterval(timer);
    };
  }, []);

  function logout() {
    clearAuth();
    toast.success("已退出登录");
    router.replace("/login");
  }

  const initial = user?.name?.[0] || user?.email?.[0]?.toUpperCase() || "U";

  return (
    <div className="min-h-screen bg-slate-50">
      <header className="sticky top-0 z-30 border-b border-slate-200 bg-white/95 backdrop-blur">
        <div className="mx-auto flex h-14 max-w-7xl items-center justify-between px-4 sm:px-6">
          <div className="flex items-center gap-7">
            <Link href="/projects" className="flex items-center gap-2 text-sm font-semibold text-slate-950">
              <span className="flex h-8 w-8 items-center justify-center rounded-md bg-sky-700 text-white">
                <FileSpreadsheet className="h-4 w-4" />
              </span>
              外服账单系统
            </Link>
            <nav className="hidden sm:block" aria-label="主导航">
              <Link href="/projects" className="text-sm text-slate-600 hover:text-slate-950">
                整合任务
              </Link>
            </nav>
          </div>

          <div className="flex items-center gap-3">
            <span className="hidden items-center gap-1.5 text-xs text-slate-500 md:flex">
              <span
                className={`h-2 w-2 rounded-full ${serviceOnline === true ? "bg-emerald-500" : serviceOnline === false ? "bg-red-500" : "bg-slate-300"}`}
              />
              {serviceOnline === true ? "服务正常" : serviceOnline === false ? "服务未连接" : "检查服务"}
            </span>
            <DropdownMenu>
              <DropdownMenuTrigger asChild>
                <Button variant="ghost" className="h-10 gap-2 px-2">
                  <Avatar className="h-7 w-7">
                    <AvatarFallback className="bg-slate-100 text-xs text-slate-700">{initial}</AvatarFallback>
                  </Avatar>
                  <span className="hidden text-sm text-slate-700 sm:inline">{user?.name}</span>
                  <ChevronDown className="h-3.5 w-3.5 text-slate-400" />
                </Button>
              </DropdownMenuTrigger>
              <DropdownMenuContent align="end" className="w-56">
                <DropdownMenuLabel>
                  <p className="text-sm font-medium">{user?.name}</p>
                  <p className="mt-0.5 text-xs font-normal text-slate-500">{user?.tenant_name}</p>
                </DropdownMenuLabel>
                <DropdownMenuSeparator />
                <DropdownMenuItem onClick={logout} className="text-red-700 focus:text-red-700">
                  <LogOut className="mr-2 h-4 w-4" />退出登录
                </DropdownMenuItem>
              </DropdownMenuContent>
            </DropdownMenu>
          </div>
        </div>
      </header>
      <main className="px-4 py-8 sm:px-6 lg:py-10">{children}</main>
    </div>
  );
}
