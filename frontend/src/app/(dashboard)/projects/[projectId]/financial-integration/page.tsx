"use client";

import { useParams, useRouter } from "next/navigation";
import { useEffect } from "react";

/** Keep historical bookmarks working while exposing one project workspace. */
export default function FinancialIntegrationRoute() {
  const { projectId } = useParams<{ projectId: string }>();
  const router = useRouter();

  useEffect(() => {
    router.replace(`/projects/${projectId}`);
  }, [projectId, router]);

  return null;
}
