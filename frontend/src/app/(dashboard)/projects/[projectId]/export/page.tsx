"use client";
import { useParams, useRouter } from "next/navigation";
import { useEffect } from "react";

export default function RedirectExport() {
  const params = useParams<{ projectId: string }>();
  const router = useRouter();
  useEffect(() => {
    router.replace(`/projects/${params.projectId}?step=export`);
  }, [params.projectId, router]);
  return null;
}
