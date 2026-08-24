"use client";

import { useParams } from "next/navigation";

import { MasterUpdatePage } from "@/features/integration/master-update-page";

export default function ProjectIntegrationPage() {
  const { projectId } = useParams<{ projectId: string }>();
  return <MasterUpdatePage projectId={projectId} />;
}
