"use client";
import { useParams, useRouter } from "next/navigation";
import { useEffect } from "react";

export default function RedirectMapping() {
  const params = useParams<{ projectId: string }>();
  const router = useRouter();
  useEffect(() => {
    router.replace(`/projects/${params.projectId}?step=mapping`);
  }, [params.projectId, router]);
  return null;
}
