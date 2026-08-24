import { ExportResultPage } from "@/features/integration/export-result-page";

export default function ProjectResultPage({ params }: { params: { projectId: string } }) {
  return <ExportResultPage projectId={params.projectId} />;
}
