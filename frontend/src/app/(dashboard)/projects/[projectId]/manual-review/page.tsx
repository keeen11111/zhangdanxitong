import { ManualReviewPage } from "@/features/integration/manual-review-page";

export default function ProjectManualReviewPage({
  params,
}: {
  params: { projectId: string };
}) {
  return <ManualReviewPage projectId={params.projectId} />;
}
