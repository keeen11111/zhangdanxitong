"use client";
import { useParams, useRouter } from "next/navigation";
import { useEffect } from "react";

export default function RedirectAudit() {
  const params = useParams<{ projectId: string }>();
  const router = useRouter();
  useEffect(() => {
    router.replace(`/projects/${params.projectId}?step=audit`);
  }, [params.projectId, router]);
  return null;
}
