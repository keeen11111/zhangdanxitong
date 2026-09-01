"use client";

import { useParams } from "next/navigation";

import { MasterUpdatePage } from "@/features/integration/master-update-page";

/** Historical payroll projects remain accessible without mixing their files into the common financial flow. */
export default function LegacyPayrollProjectPage() {
  const { projectId } = useParams<{ projectId: string }>();
  return <MasterUpdatePage projectId={projectId} />;
}
