export type WorkQueueStatus = "review_required" | "ready_for_release" | "blocked";

export function getWorkQueueStatusDisplay(status: WorkQueueStatus) {
  switch (status) {
    case "blocked":
      return { label: "未完成", step: "需重新整合", className: "bg-red-100 text-red-800" };
    case "review_required":
      return { label: "未完成", step: "需修正来源或映射", className: "bg-amber-100 text-amber-800" };
    case "ready_for_release":
      return { label: "未完成", step: "待确认发布", className: "bg-teal-100 text-teal-800" };
  }
}
