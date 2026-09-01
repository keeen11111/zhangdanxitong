"use client";

import { useParams } from "next/navigation";

import { AgentWorkbenchPage } from "@/features/agent/agent-workbench-page";

export default function ProjectIntegrationPage() {
  const { projectId } = useParams<{ projectId: string }>();
  return <AgentWorkbenchPage projectId={projectId} />;
}
