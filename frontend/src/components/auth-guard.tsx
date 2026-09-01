"use client";

import { useEffect, useState } from "react";

import { api, getStoredUser, clearAuth, User } from "@/lib/api";
import { Loader2 } from "lucide-react";

/** 登录态守卫：未登录跳转 /login，登录后渲染子节点并提供 user 上下文 */
export function AuthGuard({ children }: { children: React.ReactNode }) {
  const [user, setUser] = useState<User | null>(null);
  const [checking, setChecking] = useState(true);

  useEffect(() => {
    let active = true;
    let timeoutId: ReturnType<typeof setTimeout> | undefined;

    const redirectToLogin = () => {
      clearAuth();
      if (!active) return;

      // Use a full navigation here.  This guard itself lives below the
      // dashboard layout, so a client-side route transition that fails must
      // not leave the user indefinitely on this loading screen.
      setChecking(false);
      window.location.replace("/login");
    };

    const stored = getStoredUser();
    if (!stored) {
      redirectToLogin();
    } else {
      // A local backend should answer immediately.  If it is unavailable or
      // a network request hangs, recover to a clean login instead of showing
      // an endless \"验证登录态\" spinner.
      timeoutId = setTimeout(redirectToLogin, 8_000);

      api
        .me()
        .then((u) => {
          if (!active) return;
          if (timeoutId) clearTimeout(timeoutId);
          setUser(u);
          setChecking(false);
        })
        .catch(redirectToLogin);
    }

    return () => {
      active = false;
      if (timeoutId) clearTimeout(timeoutId);
    };
  }, []);

  if (checking) {
    return (
      <div className="flex min-h-screen items-center justify-center text-muted-foreground">
        <Loader2 className="mr-2 h-5 w-5 animate-spin" /> 验证登录态...
      </div>
    );
  }

  return user ? <>{children}</> : null;
}
