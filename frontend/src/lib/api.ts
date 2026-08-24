// 后端 API 客户端 + 类型定义

import { API_BASE } from "./utils";

// ---------- 类型 ----------
export interface User {
  id: string;
  email: string;
  name: string;
  tenant_id: string;
  tenant_name: string;
  is_active: boolean;
  created_at: string;
}

export interface TokenOut {
  access_token: string;
  token_type: string;
  user: User;
}

export interface HealthStatus {
  status: string;
  service: string;
  environment: string;
}

export interface Project {
  id: string;
  name: string;
  salary_month: string;
  status: string;
  created_at: string;
  updated_at: string;
  file_count: number;
}

export interface FileMeta {
  id: string;
  file_type: string;
  original_name: string;
  sheet_name: string | null;
  row_count: number;
  col_count: number;
  created_at: string;
  processing_seconds?: number | null;
  normalization_note?: string | null;
}

export interface UploadFailure {
  filename: string;
  message: string;
}

export interface ResilientUploadResult {
  uploaded: FileMeta[];
  failed: UploadFailure[];
}

export interface FilePreview {
  file_id: string;
  columns: string[];
  rows: Record<string, unknown>[];
  row_count: number;
  col_count: number;
}

export interface WorkbookSheetAnalysis {
  name: string;
  category: string;
  row_count: number;
  col_count: number;
  non_empty_cells: number;
  formula_count: number;
  content_types: string[];
}

export interface WorkbookAnalysis {
  file_id: string;
  filename: string;
  sheet_count: number;
  formula_count: number;
  categories: string[];
  sheets: WorkbookSheetAnalysis[];
}

// ---------- 数据画像 ----------
export interface ColumnProfile {
  name: string;
  dtype: "number" | "text" | "datetime" | "empty";
  missing_count: number;
  missing_rate: number;
  unique_count: number;
  sample_values: string[];
  min?: number | null;
  max?: number | null;
  mean?: number | null;
}

export interface DistributionItem {
  label: string;
  count: number;
}

export interface Distribution {
  column: string;
  type: "category" | "histogram";
  items: DistributionItem[];
}

export interface FileProfile {
  file_id: string;
  filename: string;
  file_type: string;
  overview: {
    total_rows: number;
    total_cols: number;
    missing_cells: number;
    completeness_pct: number;
  };
  columns_profile: ColumnProfile[];
  distributions: Distribution[];
}

export interface DiffResult {
  new_hires: Record<string, unknown>[];
  departed: Record<string, unknown>[];
  transfers: Record<string, unknown>[];
}

export interface AuditMetrics {
  total: number;
  qualified: number;
  completeness_pct: number;
  high_risk_count: number;
  warning_count: number;
  missing_field_breakdown: Record<string, number>;
}

export interface AuditResult {
  columns: string[];
  rows: Record<string, unknown>[];
  status_col: string[];
  missing_col: string[][];
  metrics: AuditMetrics;
}

export interface ExportResult {
  file_id: string;
  filename: string;
  log: string[];
}

// ---------- Modules（板块化工作台） ----------
export interface ModuleInfo {
  id: string;
  name: string;
  source_file: string;
  source_sheet: string;
  row_count: number;
  col_count: number;
}

export interface ModuleData {
  id: string;
  name: string;
  source_file: string;
  source_sheet: string;
  columns: string[];
  rows: Record<string, unknown>[];
}

export interface ModuleSaveIn {
  columns: string[];
  rows: Record<string, unknown>[];
}

// ---------- Token 管理 ----------
const TOKEN_KEY = "autopayroll_token";
const USER_KEY = "autopayroll_user";

export function getToken(): string | null {
  if (typeof window === "undefined") return null;
  return localStorage.getItem(TOKEN_KEY);
}

export function setAuth(token: string, user: User) {
  localStorage.setItem(TOKEN_KEY, token);
  localStorage.setItem(USER_KEY, JSON.stringify(user));
}

export function clearAuth() {
  localStorage.removeItem(TOKEN_KEY);
  localStorage.removeItem(USER_KEY);
}

export function getStoredUser(): User | null {
  if (typeof window === "undefined") return null;
  const s = localStorage.getItem(USER_KEY);
  if (!s) return null;
  try {
    return JSON.parse(s) as User;
  } catch {
    // A partially written/stale browser cache must not strand the app on
    // the root loading screen.  The next visit can start a clean login.
    clearAuth();
    return null;
  }
}

// ---------- fetch 封装 ----------
async function request<T>(
  path: string,
  options: RequestInit = {}
): Promise<T> {
  const token = getToken();
  const headers: Record<string, string> = {
    ...(options.headers as Record<string, string>),
  };
  if (token) headers["Authorization"] = `Bearer ${token}`;
  // FormData 时不要手动设 Content-Type，浏览器自动加 boundary
  if (!(options.body instanceof FormData) && options.body) {
    headers["Content-Type"] = "application/json";
  }
  const res = await fetch(`${API_BASE}${path}`, { ...options, headers });
  const text = await res.text();
  let data: unknown = null;
  if (text) {
    try {
      data = JSON.parse(text);
    } catch {
      data = text;
    }
  }
  if (!res.ok) {
    const msg =
      (data && typeof data === "object" && "detail" in data
        ? String((data as { detail: unknown }).detail)
        : typeof data === "string"
        ? data
        : `请求失败 (${res.status})`) || `请求失败 (${res.status})`;
    throw new Error(msg);
  }
  return data as T;
}

async function uploadFilesIndividually(
  projectId: string,
  files: File[],
  fileType = "source"
): Promise<ResilientUploadResult> {
  const uploaded: FileMeta[] = [];
  const failed: UploadFailure[] = [];

  // Keep failures isolated while allowing a small amount of parallelism.  A
  // bounded worker pool avoids both one oversized multipart request and the
  // previous one-file-at-a-time upload delay.
  let nextIndex = 0;
  async function worker() {
    while (nextIndex < files.length) {
      const file = files[nextIndex++];
      const formData = new FormData();
      formData.append("file", file);
      formData.append("file_type", fileType);

      try {
        const record = await request<FileMeta>(
          `/api/projects/${projectId}/files`,
          { method: "POST", body: formData }
        );
        if (record) {
          uploaded.push(record);
        } else {
          failed.push({ filename: file.name, message: "服务未返回文件信息" });
        }
      } catch (error) {
        failed.push({
          filename: file.name,
          message: error instanceof Error ? error.message : "上传失败",
        });
      }
    }
  }
  await Promise.all(
    Array.from({ length: Math.min(3, files.length) }, () => worker())
  );

  return { uploaded, failed };
}

// ---------- Auth API ----------
export const api = {
  health: () => request<HealthStatus>("/api/health"),

  register: (payload: {
    email: string;
    password: string;
    name: string;
    tenant_name?: string;
  }) =>
    request<TokenOut>("/api/auth/register", {
      method: "POST",
      body: JSON.stringify(payload),
    }),

  login: (payload: { email: string; password: string }) =>
    request<TokenOut>("/api/auth/login", {
      method: "POST",
      body: JSON.stringify(payload),
    }),

  me: () => request<User>("/api/auth/me"),

  // ---------- Projects ----------
  listProjects: () => request<Project[]>("/api/projects"),
  createProject: (payload: { name: string; salary_month: string }) =>
    request<Project>("/api/projects", {
      method: "POST",
      body: JSON.stringify(payload),
    }),
  getProject: (id: string) => request<Project>(`/api/projects/${id}`),
  deleteProject: (id: string) =>
    request<void>(`/api/projects/${id}`, { method: "DELETE" }),
  updateProjectStatus: (id: string, status: string) =>
    request<Project>(`/api/projects/${id}/status?status=${status}`, {
      method: "PATCH",
    }),

  // ---------- Files ----------
  uploadFile: (
    projectId: string,
    file: File,
    fileType: string,
    sheetName?: string
  ) => {
    const fd = new FormData();
    fd.append("file", file);
    fd.append("file_type", fileType);
    if (sheetName) fd.append("sheet_name", sheetName);
    return request<FileMeta>(`/api/projects/${projectId}/files`, {
      method: "POST",
      body: fd,
    });
  },
  uploadFilesAuto: (projectId: string, files: File[]) => {
    const fd = new FormData();
    files.forEach((file) => fd.append("files", file));
    return request<FileMeta[]>(`/api/projects/${projectId}/files/batch-auto`, {
      method: "POST",
      body: fd,
    });
  },
  uploadFilesIndividually,
  listFiles: (projectId: string) =>
    request<FileMeta[]>(`/api/projects/${projectId}/files`),
  deleteFile: (projectId: string, fileId: string) =>
    request<void>(`/api/projects/${projectId}/files/${fileId}`, {
      method: "DELETE",
    }),
  updateFileType: (projectId: string, fileId: string, fileType: string) =>
    request<FileMeta>(`/api/projects/${projectId}/files/${fileId}`, {
      method: "PATCH",
      body: JSON.stringify({ file_type: fileType }),
    }),
  previewFile: (projectId: string, fileId: string) =>
    request<FilePreview>(`/api/projects/${projectId}/files/${fileId}/preview`),
  analyzeFile: (projectId: string, fileId: string) =>
    request<WorkbookAnalysis>(`/api/projects/${projectId}/files/${fileId}/analysis`),
  profileFile: (projectId: string, fileId: string) =>
    request<FileProfile>(`/api/projects/${projectId}/files/${fileId}/profile`),

  // ---------- Payroll ----------
  computePayrollDiff: (
    projectId: string,
    payload: {
      last_month_file_id: string;
      current_file_id: string;
      id_col_lm: string;
      id_col_cs: string;
      compare_cols: string[];
      social_file_ids: string[];
    }
  ) =>
    request<DiffResult>(`/api/payroll/${projectId}/diff`, {
      method: "POST",
      body: JSON.stringify(payload),
    }),

  audit: (projectId: string) =>
    request<AuditResult>(`/api/payroll/${projectId}/audit`),

  saveAudit: (projectId: string, rows: Record<string, unknown>[]) =>
    request<AuditResult>(`/api/payroll/${projectId}/audit/save`, {
      method: "POST",
      body: JSON.stringify({ rows }),
    }),

  export: (projectId: string) =>
    request<ExportResult>(`/api/payroll/${projectId}/export`, {
      method: "POST",
    }),

  exportDownloadUrl: (projectId: string) =>
    `${API_BASE}/api/payroll/${projectId}/export/download`,

  // ---------- Modules（板块化工作台） ----------
  loadModules: (projectId: string) =>
    request<ModuleInfo[]>(`/api/modules/${projectId}/load`, { method: "POST" }),
  listModules: (projectId: string) =>
    request<ModuleInfo[]>(`/api/modules/${projectId}`),
  getModule: (projectId: string, moduleId: string) =>
    request<ModuleData>(`/api/modules/${projectId}/${moduleId}`),
  saveModule: (projectId: string, moduleId: string, payload: ModuleSaveIn) =>
    request<ModuleData>(`/api/modules/${projectId}/${moduleId}`, {
      method: "PUT",
      body: JSON.stringify(payload),
    }),
  addRow: (projectId: string, moduleId: string, row: Record<string, unknown>) =>
    request<ModuleData>(`/api/modules/${projectId}/${moduleId}/rows`, {
      method: "POST",
      body: JSON.stringify({ row }),
    }),
  deleteRow: (projectId: string, moduleId: string, rowIdx: number) =>
    request<ModuleData>(`/api/modules/${projectId}/${moduleId}/rows/${rowIdx}`, {
      method: "DELETE",
    }),
  renameModule: (projectId: string, moduleId: string, name: string) =>
    request<ModuleInfo>(`/api/modules/${projectId}/${moduleId}/rename`, {
      method: "PATCH",
      body: JSON.stringify({ name }),
    }),
  exportModules: (projectId: string) =>
    request<{ filename: string; log: string[] }>(`/api/modules/${projectId}/export`, {
      method: "POST",
    }),
  moduleExportDownloadUrl: (projectId: string) =>
    `${API_BASE}/api/modules/${projectId}/export/download`,

  // ---------- Entities（实体化工作台） ----------
  uploadTemplate: (projectId: string, file: File) => {
    const fd = new FormData();
    fd.append("file", file);
    return request<{ id: string; filename: string; message: string }>(
      `/api/entities/${projectId}/template`,
      { method: "POST", body: fd }
    );
  },
  importEntities: (projectId: string) =>
    request<EntityOverview[]>(`/api/entities/${projectId}/import`, { method: "POST" }),
  listEntities: (projectId: string) =>
    request<EntityOverview[]>(`/api/entities/${projectId}`),
  getEntity: (projectId: string, entityId: string) =>
    request<EntityData>(`/api/entities/${projectId}/${entityId}`),
  saveEntity: (projectId: string, entityId: string, payload: { columns: string[]; rows: Record<string, unknown>[] }) =>
    request<EntityData>(`/api/entities/${projectId}/${entityId}`, {
      method: "PUT",
      body: JSON.stringify(payload),
    }),
  addEntityRow: (projectId: string, entityId: string, row: Record<string, unknown>) =>
    request<EntityData>(`/api/entities/${projectId}/${entityId}/rows`, {
      method: "POST",
      body: JSON.stringify({ row }),
    }),
  deleteEntityRow: (projectId: string, entityId: string, rowIdx: number) =>
    request<{ ok: boolean; remaining: number }>(
      `/api/entities/${projectId}/${entityId}/rows/${rowIdx}`,
      { method: "DELETE" }
    ),
  getEntityLogs: (projectId: string) =>
    request<{ logs: string[]; imported_at: string | null }>(`/api/entities/${projectId}/logs`),
  exportEntities: (projectId: string) =>
    request<{ filename: string; log: string[] }>(`/api/entities/${projectId}/export`, {
      method: "POST",
    }),
  entityExportDownloadUrl: (projectId: string) =>
    `${API_BASE}/api/entities/${projectId}/export/download`,

  // ---------- Snapshots（月度快照） ----------
  listSnapshots: (projectId: string) =>
    request<SnapshotInfo[]>(`/api/entities/${projectId}/snapshots`),
  saveSnapshot: (projectId: string, label?: string) =>
    request<SnapshotInfo>(`/api/entities/${projectId}/snapshots`, {
      method: "POST",
      body: JSON.stringify({ label: label || null }),
    }),
  loadSnapshot: (projectId: string, snapshotId: string) =>
    request<EntityOverview[]>(`/api/entities/${projectId}/snapshots/${snapshotId}/load`, {
      method: "POST",
    }),
  cloneSnapshot: (projectId: string, snapshotId: string) =>
    request<EntityOverview[]>(`/api/entities/${projectId}/snapshots/${snapshotId}/clone`, {
      method: "POST",
    }),
  deleteSnapshot: (projectId: string, snapshotId: string) =>
    request<{ ok: boolean }>(`/api/entities/${projectId}/snapshots/${snapshotId}`, {
      method: "DELETE",
    }),

  // ---------- Pipeline（6模块流水线） ----------
  loadBaseDb: (projectId: string, refresh = false) =>
    request<PipelineDbResponse>(`/api/pipeline/${projectId}/base-db/load?refresh=${refresh}`, { method: "POST" }),
  getBaseDb: (projectId: string) =>
    request<PipelineDbResponse>(`/api/pipeline/${projectId}/base-db`),
  extractSourceData: (projectId: string) =>
    request<PipelineDbResponse>(`/api/pipeline/${projectId}/source-data/extract`, { method: "POST" }),
  getSourceData: (projectId: string) =>
    request<PipelineDbResponse>(`/api/pipeline/${projectId}/source-data`),
  computeDiff: (projectId: string) =>
    request<DiffResponse>(`/api/pipeline/${projectId}/diff/compute`, { method: "POST" }),
  getDiff: (projectId: string) =>
    request<DiffResponse>(`/api/pipeline/${projectId}/diff`),
  updateDiffItem: (projectId: string, payload: DiffUpdateIn) =>
    request<DiffResponse>(`/api/pipeline/${projectId}/diff/update`, {
      method: "POST",
      body: JSON.stringify(payload),
    }),
  batchUpdateDiff: (projectId: string, payload: BatchUpdateIn) =>
    request<DiffResponse>(`/api/pipeline/${projectId}/diff/batch-update`, {
      method: "POST",
      body: JSON.stringify(payload),
    }),
  buildFinalDb: (projectId: string) =>
    request<PipelineDbResponse>(`/api/pipeline/${projectId}/final-db/build`, { method: "POST" }),
  getFinalDb: (projectId: string) =>
    request<PipelineDbResponse>(`/api/pipeline/${projectId}/final-db`),
  exportPipeline: (projectId: string) =>
    request<PipelineExportResult>(`/api/pipeline/${projectId}/export`, { method: "POST" }),
  integratePipeline: (projectId: string) =>
    request<IntegrationResult>(`/api/pipeline/${projectId}/integrate`, { method: "POST" }),
  getLatestPipelineExport: (projectId: string) =>
    request<PipelineExportResult>(`/api/pipeline/${projectId}/export/latest`),
  renamePipelineExport: (projectId: string, filename: string) =>
    request<PipelineExportResult>(`/api/pipeline/${projectId}/export/filename`, {
      method: "PATCH",
      body: JSON.stringify({ filename }),
    }),
  previewPipelineExport: (projectId: string) =>
    request<PipelineExportPreview>(`/api/pipeline/${projectId}/export/preview`),
  previewManualIssuesExport: (projectId: string) =>
    request<PipelineExportPreview>(`/api/pipeline/${projectId}/export/issues/preview`),
  previewReviewExport: (projectId: string) =>
    request<PipelineExportPreview>(`/api/pipeline/${projectId}/export/review/preview`),
  confirmDepartmentTransfers: (projectId: string, issueIds: string[]) =>
    request<PipelineExportResult>(
      `/api/pipeline/${projectId}/manual-review/department-transfers/confirm`,
      { method: "POST", body: JSON.stringify({ issue_ids: issueIds }) },
    ),
  resolveManualIssue: (projectId: string, issueId: string, action: "apply_proposed" | "keep_current") =>
    request<PipelineExportResult>(
      `/api/pipeline/${projectId}/manual-review/issue/resolve`,
      { method: "POST", body: JSON.stringify({ issue_id: issueId, action }) },
    ),
  pipelineExportDownloadUrl: (projectId: string) =>
    `${API_BASE}/api/pipeline/${projectId}/export/download`,
  manualIssuesExportDownloadUrl: (projectId: string) =>
    `${API_BASE}/api/pipeline/${projectId}/export/issues/download`,
  reviewExportDownloadUrl: (projectId: string) =>
    `${API_BASE}/api/pipeline/${projectId}/export/review/download`,
};

// ---------- Pipeline Types ----------
export interface PipelineDbResponse {
  entities: Record<string, Record<string, unknown>[]>;
  overview: EntityOverview[];
  source?: string;
  logs?: string[];
  validation?: PipelineValidation;
  created_at: string;
}

export interface PipelineValidationItem {
  row_index: number;
  employee_id: string;
  employee_name: string;
  status: string;
  missing_fields: string[];
}

export interface PipelineValidation {
  metrics: AuditMetrics;
  blockers: PipelineValidationItem[];
  warnings: PipelineValidationItem[];
  preparation_issue_count?: number;
}

export interface IntegrationStage {
  key: string;
  label: string;
  status: "completed";
}

export interface ManualIssue {
  issue_id?: string;
  issue_type: string;
  employee_id?: string;
  person_name: string;
  target_field: string;
  problem_type?: string;
  difficulty?: "简单" | "一般" | "复杂";
  suggestion?: string;
  recommended_action?: "apply_proposed" | "keep_current" | "copy_only";
  action_options?: Array<{ value: string; label: string }>;
  copy_text?: string;
  current_value?: string | null;
  proposed_value?: string | null;
  status?: "pending" | "confirmed";
  message: string;
  source_files: string[];
  source_sheets: string[];
}

export interface IntegrationResult {
  status: "completed" | "completed_with_issues" | "failed";
  filename: string | null;
  stages: IntegrationStage[];
  overview: EntityOverview[];
  validation: PipelineValidation;
  log: string[];
  completed_at: string;
  issue_count: number;
  source_person_count: number;
  output_person_count: number;
  input_person_count: number;
  manual_only_person_count: number;
  coverage_gap_count: number;
  source_file_count: number;
  issues: ManualIssue[];
  issues_filename?: string | null;
  review_filename?: string | null;
  duration_seconds: number;
  stage_durations: Record<string, number>;
}

export interface PipelineExportResult {
  filename: string;
  log: string[];
  issue_count: number;
  source_person_count: number;
  output_person_count: number;
  input_person_count: number;
  manual_only_person_count: number;
  coverage_gap_count: number;
  source_file_count: number;
  issues: ManualIssue[];
  validation: PipelineValidation;
  completed_at: string;
  issues_filename?: string | null;
  review_filename?: string | null;
  duration_seconds: number;
  stage_durations: Record<string, number>;
}

export interface PipelineExportPreview {
  filename: string;
  sheet_name: string;
  rows: unknown[][];
}

export interface DiffResponse {
  diff: Record<string, unknown[] | Record<string, unknown>>;
  summary: Record<string, unknown>;
  created_at: string;
}

export interface DiffUpdateIn {
  diff_type: string;
  item_index: number;
  status: string;
  updates?: Record<string, unknown> | null;
}

export interface BatchUpdateIn {
  diff_type: string;
  indices: number[];
  status: string;
}

// ---------- Entity Types ----------
export interface EntityOverview {
  id: string;
  name: string;
  icon: string;
  primary_key: string;
  row_count: number;
  field_count: number;
  completeness_pct: number;
}

export interface EntityFieldDef {
  name: string;
  dtype: string;
  required: boolean;
}

export interface EntityData {
  id: string;
  name: string;
  icon: string;
  primary_key: string;
  fields: EntityFieldDef[];
  columns: string[];
  rows: Record<string, unknown>[];
}

export interface SnapshotInfo {
  id: string;
  salary_month: string;
  label: string | null;
  row_count: number;
  created_at: string;
}
