"use client";

import { useEffect, useState } from "react";
import { useRouter } from "next/navigation";

import { api, getStoredUser, clearAuth, User } from "@/lib/api";
import { Loader2 } from "lucide-react";

/** 登录态守卫：未登录跳转 /login，登录后渲染子节点并提供 user 上下文 */
export function AuthGuard({ children }: { children: React.ReactNode }) {
  const router = useRouter();
  const [user, setUser] = useState<User | null>(null);
  const [checking, setChecking] = useState(true);

  useEffect(() => {
    const stored = getStoredUser();
    if (!stored) {
      router.replace("/login");
      return;
    }
    // 校验 token 有效性
    api
      .me()
      .then((u) => {
        setUser(u);
        setChecking(false);
      })
      .catch(() => {
        clearAuth();
        router.replace("/login");
      });
  }, [router]);

  if (checking || !user) {
    return (
      <div className="flex min-h-screen items-center justify-center text-muted-foreground">
        <Loader2 className="mr-2 h-5 w-5 animate-spin" /> 验证登录态...
      </div>
    );
  }

  return <>{children}</>;
}
