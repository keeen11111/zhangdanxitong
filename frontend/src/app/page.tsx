"use client";

import { useEffect } from "react";
import { useRouter } from "next/navigation";
import { api, getStoredUser } from "@/lib/api";

/** 首页：根据登录态跳转 */
export default function Home() {
  const router = useRouter();
  useEffect(() => {
    let active = true;
    const fallbackTimer = window.setTimeout(() => {
      if (active) router.replace("/login");
    }, 8000);
    const user = getStoredUser();
    if (user) {
      // 验证 token 是否仍有效
      api
        .me()
        .then(() => {
          if (active) router.replace("/projects");
        })
        .catch(() => {
          if (active) router.replace("/login");
        });
    } else {
      router.replace("/login");
    }
    return () => {
      active = false;
      window.clearTimeout(fallbackTimer);
    };
  }, [router]);

  return (
    <div className="flex min-h-screen items-center justify-center text-muted-foreground">
      加载中...
    </div>
  );
}
