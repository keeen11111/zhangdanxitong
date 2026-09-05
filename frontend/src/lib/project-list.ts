import type { Project } from "./api";

/** Browser event used to keep the persistent shell in sync with project CRUD. */
export const PROJECT_CREATED_EVENT = "payroll:project-created";

export function mergeRecentProject(
  current: Project[],
  created: Project,
  limit = 8,
): Project[] {
  const withoutDuplicate = current.filter((project) => project.id !== created.id);
  return [created, ...withoutDuplicate].slice(0, limit);
}
