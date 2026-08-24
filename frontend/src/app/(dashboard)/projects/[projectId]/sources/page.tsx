"use client";
import { useParams, useRouter } from "next/navigation";
import { useEffect } from "react";

export default function RedirectSources() {
  const params = useParams<{ projectId: string }>();
  const router = useRouter();
  useEffect(() => {
    router.replace(`/projects/${params.projectId}?step=sources`);
  }, [params.projectId, router]);
  return null;
}
