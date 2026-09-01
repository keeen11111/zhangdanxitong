// 后端 API 客户端 + 类型定义

import { API_BASE } from "./utils";
import { fetchWithTimeout } from "./request-timeout";
import { FINANCIAL_INTEGRATION_TIMEOUT_MS } from "./financial-workbook-flow";
import { recoverFromUnauthorized } from "./auth-recovery";
import { AgentStreamEvent, parseSseFrames } from "./agent-stream";

const UPLOAD_REQUEST_TIMEOUT_MS = 75_000;
const UPLOAD_TIMEOUT_MESSAGE = "上传处理超过 75 秒，已停止等待。请确认文件未加密，或另存为普通 .xlsx 后重试";
const DEFAULT_REQUEST_TIMEOUT_MS = 20_000;
const DEFAULT_REQUEST_TIMEOUT_MESSAGE = "请求处理超过 20 秒，已停止等待。请刷新页面后重试";

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
  has_result: boolean;
  result_completed_at: string | null;
  pending_issue_count: number;
}

export interface MatchingPolicy {
  project_id: string;
  company: string;
  salary_month: string;
  auto_match_threshold: number;
  field_weights: Record<string, number>;
  source_rules: Array<{
    company?: string | null;
    month?: string | null;
    source_sheet?: string | null;
    target_sheet?: string | null;
    target_field?: string | null;
    action: "allow" | "review" | "ignore";
  }>;
  updated_at: string;
}

// ---------- Document-driven Excel Agent ----------
// These contracts intentionally describe the approval boundary. The server
// may add fields over time, but the client only renders the safe, typed subset.
export type AgentRunStatus =
  | "draft"
  | "planning"
  | "ready"
  | "processing"
  | "execution_incomplete"
  | "review"
  | "completed"
  | "published"
  | "blocked"
  | "failed";

export type AgentWorkItemStatus = "pending" | "needs_conversation" | "applied" | "skipped" | "failed";

export type AgentDecisionAction = "apply" | "keep_current" | "skip";

export interface AgentRun {
  execution_result?: { status?: string; code?: string | null; content?: string } | null;
  validation?: { status?: string; detail?: string } | null;
  draft_filename?: string | null;
  execution_mode?: "demo" | "keyuan_reference";
  demo_result?: { filename: string; sha256: string; provenance: string };
  reference_result?: { filename: string; sha256: string; provenance: string };
  plan_confirmation?: { required: boolean; confirmed: boolean };
  file_manifest?: Array<{ filename: string; role: string; sha256: string }>;
  model_plan?: { summary: string; steps: string[]; questions: string[]; model: string };
  basic_processor?: { status?: string; change_count?: number; issue_count?: number; issues?: Array<{ item?: string; detail?: string }> } | null;
  run_id: string;
  project_id: string;
  company_id?: string | null;
  company_name?: string | null;
  salary_month?: string | null;
  rule_version?: string | null;
  status: AgentRunStatus;
  total_items?: number;
  pending_items?: number;
  high_risk_count?: number;
  master_file?: string | null;
  source_files?: string[];
  created_at?: string;
  updated_at?: string;
  detail?: string | null;
  month_confirmation?: { required: boolean; confirmed: boolean; filename_month?: string; configured_month?: string };
  conversation?: AgentMessage[];
}

export interface AgentResultChange {
  target_sheet?: string;
  target_cell?: string;
  sheet?: string;
  old_value?: unknown;
  new_value?: unknown;
  source_file?: string;
  source_sheet?: string;
  source_cells?: string[];
  inserted_row?: number;
}

export interface AgentRunResult {
  run_id: string;
  available: boolean;
  can_download: boolean;
  filename: string | null;
  sha256?: string | null;
  completed_at?: string | null;
  validation: { status?: string; detail?: string };
  change_count: number;
  changes: AgentResultChange[];
  pagination: { page: number; page_size: number; total_items: number; total_pages: number };
}

export interface AgentWorkItem {
  item_id: string;
  run_id?: string;
  sequence?: number;
  person_name?: string | null;
  employee_ref?: string | null;
  category?: string | null;
  field?: string | null;
  source_values?: Record<string, unknown>;
  current_value?: string | number | boolean | null;
  proposed_value?: string | number | boolean | null;
  candidate_values?: Array<string | number | boolean | null>;
  recommendation?: string | null;
  reason?: string | null;
  risk_level?: "low" | "medium" | "high";
  status: AgentWorkItemStatus;
  confidence?: number | null;
  conversation?: AgentMessage[];
  updated_at?: string;
  revision?: number;
  decision?: string | null;
}

export interface AgentMessage {
  id?: string;
  role: "agent" | "user" | "system";
  content: string;
  created_at?: string;
}

export interface AgentRunEvent {
  event_id: string;
  run_id: string;
  revision: number;
  type: string;
  payload: Record<string, unknown>;
  item_id?: string;
  at?: string;
}

export interface AgentMemoryRule {
  rule_id: string;
  company_id?: string | null;
  title: string;
  description: string;
  status: "candidate" | "active" | "suspended";
  confidence?: number;
  evidence_count?: number;
  last_used_at?: string | null;
  source?: string | null;
}

export interface AgentMemoryResponse {
  rules: AgentMemoryRule[];
  suggestions?: AgentMemoryRule[];
}

export interface AgentRunBundle {
  run: AgentRun;
  items: AgentWorkItem[];
  memory?: AgentMemoryResponse;
}

export type AgentMaterialKind = "manual" | "transcript" | "recording" | "rule_package" | "unknown";

export interface AgentMaterial {
  material_id: string;
  project_id: string;
  filename: string;
  kind: AgentMaterialKind;
  status: "ready" | "transcription_required" | "unsupported" | "error";
  mime_type: string;
  size_bytes: number;
  created_at: string;
  text_preview?: string;
}

export interface AgentDecisionInput {
  action: AgentDecisionAction;
  value?: string | number | boolean | null;
  note?: string;
  remember?: boolean;
  source?: string | null;
  expected_revision?: number;
}

export interface AgentModelStatus {
  configured: boolean;
  provider: string | null;
  model: string | null;
  api_style: "chat_completions" | "responses" | null;
  reasoning_effort: string | null;
  enable_thinking: boolean;
  fallback_configured?: boolean;
  fallback_provider?: string | null;
  fallback_model?: string | null;
}

export interface AgentRulePackageResult {
  project_id: string;
  package: {
    package_id: string;
    company_id: string;
    version: string;
    effective_period: string;
    status: "candidate" | "active" | "retired";
    rules: Record<string, unknown>[];
    sources: Record<string, unknown>[];
  };
  model: AgentModelStatus;
}

export interface AgentRulePackage {
  package_id: string;
  project_id: string;
  filename: string;
  version: string;
  effective_period: string;
  status: "candidate" | "active" | "retired";
  rules: Array<{
    rule_id?: string;
    condition?: Record<string, unknown>;
    action?: Record<string, unknown>;
    exceptions?: Array<Record<string, unknown>>;
    source_refs?: string[];
  }>;
  sources: Array<{
    kind?: string;
    ref?: string;
    excerpt?: string;
    revision?: number;
  }>;
  created_at: string;
}

function mapAgentStatus(status: unknown): AgentRunStatus {
  const value = String(status || "");
  if (value === "awaiting_review") return "review";
  if (value === "ready_to_publish") return "ready";
  if (value === "needs_review") return "review";
  if (value === "execution_incomplete") return "execution_incomplete";
  if (value === "processing") return "processing";
  if (value === "published") return "published";
  if (value === "failed") return "failed";
  if (value === "blocked") return "blocked";
  return (value as AgentRunStatus) || "draft";
}

function mapAgentRun(raw: Record<string, unknown>): AgentRun {
  const summary = (raw.summary || {}) as Record<string, unknown>;
  const execution = (raw.execution_result || {}) as Record<string, unknown>;
  return {
    ...(raw as unknown as AgentRun),
    // Older persisted runs recorded a turn-limit outcome as awaiting_review.
    // It is resumable work, not a user review state.
    status: execution.code === "MAX_TURNS_EXCEEDED" ? "execution_incomplete" : mapAgentStatus(raw.status),
    total_items: Number(summary.total ?? raw.total_items ?? 0),
    pending_items: Number(summary.needs_review ?? raw.pending_items ?? 0),
    high_risk_count: Number(raw.high_risk_count ?? 0),
  };
}

function normaliseEventCursor(value: number) {
  return Number.isFinite(value) ? Math.max(0, Math.floor(value)) : 0;
}

function mapAgentItem(raw: Record<string, unknown>): AgentWorkItem {
  const status = String(raw.status || "pending");
  const mappedStatus: AgentWorkItemStatus = status === "needs_review"
    ? "needs_conversation"
    : status === "resolved" || status === "auto_applied"
    ? "applied"
    : status === "skipped"
    ? "skipped"
    : status === "failed"
    ? "failed"
    : "pending";
  const messages = Array.isArray(raw.messages) ? raw.messages : [];
  return {
    item_id: String(raw.id || raw.item_id || ""),
    run_id: raw.run_id as string | undefined,
    sequence: Number(raw.sequence || 0),
    person_name: (raw.person_name as string | null) ?? null,
    employee_ref: (raw.person_key as string | null) ?? null,
    category: (raw.issue_type as string | null) ?? null,
    field: (raw.target_cell as string | null) ?? null,
    source_values: (raw.source_values as Record<string, unknown>) || {
      来源值: raw.proposed_value,
    },
    current_value: raw.current_value as AgentWorkItem["current_value"],
    proposed_value: raw.proposed_value as AgentWorkItem["proposed_value"],
    candidate_values: Array.isArray(raw.candidate_values) ? raw.candidate_values as AgentWorkItem["candidate_values"] : [],
    recommendation: (raw.suggested_action as string | null) || (mappedStatus === "needs_conversation" ? "需要确认" : "按规则处理"),
    reason: (raw.message as string | null) || null,
    risk_level: raw.issue_type === "conflicting_value" ? "high" : "low",
    status: mappedStatus,
    confidence: typeof raw.confidence === "number" ? raw.confidence : null,
    conversation: messages.map((message) => {
      const entry = message as Record<string, unknown>;
      return {
        role: entry.role === "user" ? "user" : entry.role === "system" ? "system" : "agent",
        content: String(entry.content || ""),
        created_at: entry.at as string | undefined,
      } as AgentMessage;
    }),
    updated_at: raw.updated_at as string | undefined,
    revision: typeof raw.revision === "number" ? raw.revision : 1,
    decision: typeof raw.decision === "string" ? raw.decision : null,
  };
}

function mapMemoryRule(raw: Record<string, unknown>): AgentMemoryRule {
  const conditions = raw.conditions && typeof raw.conditions === "object"
    ? Object.entries(raw.conditions as Record<string, unknown>).map(([key, value]) => `${key}=${String(value)}`).join("；")
    : "公司确认的可复用处理规则";
  return {
    rule_id: String(raw.id || raw.rule_id || ""),
    title: String(raw.diff_type || "Agent 处理规则"),
    description: String(raw.note || conditions),
    status: raw.status === "suspended" ? "suspended" : raw.status === "active" ? "active" : "candidate",
    confidence: typeof raw.confidence === "number" ? raw.confidence : undefined,
    evidence_count: typeof raw.use_count === "number" ? raw.use_count : undefined,
    last_used_at: (raw.last_used_at as string | null) ?? null,
    source: (raw.rule_version as string | null) ?? null,
  };
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

export interface FinancialWorkbookIntegration {
  filename: string;
  review_filename?: string | null;
  auto_update_count: number;
  source_file_count: number;
  issues: Record<string, unknown>[];
  matches: Record<string, unknown>[];
  updates: Record<string, unknown>[];
  status: "review_required" | "ready_for_release" | "published" | "blocked";
  release_checks: {
    can_release: boolean;
    issue_count: number;
    matched_sheet_count: number;
    auto_update_count: number;
  };
  completed_at: string;
  published_at?: string | null;
  published_version?: FinancialWorkbookPublishedVersion | null;
}

export interface FinancialWorkbookPublishedVersion {
  version_id: string;
  filename: string;
  original_draft_filename: string;
  auto_update_count: number;
  source_file_count: number;
  release_checks: {
    can_release: boolean;
    issue_count: number;
    matched_sheet_count: number;
    auto_update_count: number;
  };
  published_at: string;
  published_by: {
    id: string;
    name: string;
  };
  sha256: string;
}

export interface FinancialWorkQueueItem {
  project_id: string;
  project_name: string;
  salary_month: string;
  status: "review_required" | "ready_for_release" | "blocked";
  issue_count: number;
  action_label: string;
  completed_at: string;
}

export interface FinancialWorkQueue {
  items: FinancialWorkQueueItem[];
  total: number;
  page: number;
  page_size: number;
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

function recoverExpiredSession(status: number): boolean {
  if (typeof window === "undefined") return false;

  return recoverFromUnauthorized(status, window.location.pathname, {
    clear: clearAuth,
    redirect: (target) => window.location.replace(target),
  });
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
type RequestTimeoutOptions = {
  timeoutMs?: number;
  timeoutMessage?: string;
};

async function request<T>(
  path: string,
  options: RequestInit = {},
  timeoutOptions: RequestTimeoutOptions = {},
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
  const res = await fetchWithTimeout(
    `${API_BASE}${path}`,
    { ...options, headers },
    timeoutOptions.timeoutMs ?? DEFAULT_REQUEST_TIMEOUT_MS,
    timeoutOptions.timeoutMessage ?? DEFAULT_REQUEST_TIMEOUT_MESSAGE,
  );
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
    recoverExpiredSession(res.status);
    const msg =
      (data && typeof data === "object" && "detail" in data
        ? String((data as { detail: unknown }).detail)
        : typeof data === "string"
        ? data
        : `请求失败 (${res.status})`) || `请求失败 (${res.status})`;
    const error = new Error(msg) as Error & { status?: number };
    error.status = res.status;
    throw error;
  }
  return data as T;
}

async function uploadFilesIndividually(
  projectId: string,
  files: File[],
  fileType = "source",
  password?: string,
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
      try {
        const record = await uploadProjectFile(projectId, file, fileType, undefined, password);
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

function uploadProjectFile(
  projectId: string,
  file: File,
  fileType: string,
  sheetName?: string,
  password?: string,
) {
  const formData = new FormData();
  formData.append("file", file);
  formData.append("file_type", fileType);
  if (sheetName) formData.append("sheet_name", sheetName);
  if (password) formData.append("password", password);

  return request<FileMeta>(
    `/api/projects/${projectId}/files`,
    {
      method: "POST",
      body: formData,
    },
    {
      timeoutMs: UPLOAD_REQUEST_TIMEOUT_MS,
      timeoutMessage: UPLOAD_TIMEOUT_MESSAGE,
    },
  );
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

  // ---------- Agent（文档驱动 Excel 处理） ----------
  startAgentRun: (projectId: string, instruction = "按公司生效规则更新总表", autoPublish = false, demo = false) =>
    request<Record<string, unknown>>("/api/agent/runs", {
      method: "POST",
      body: JSON.stringify({ project_id: projectId, instruction, auto_publish: autoPublish, demo }),
    }, {
      timeoutMs: FINANCIAL_INTEGRATION_TIMEOUT_MS,
      timeoutMessage: "Agent 正在处理文件超过 15 分钟，请刷新页面查看运行状态",
    }).then(mapAgentRun),
  getAgentRun: (runId: string) => request<Record<string, unknown>>(`/api/agent/runs/${runId}`).then(mapAgentRun),
  getAgentDemo: (projectId: string) => request<{ available: boolean; filename: string | null }>(`/api/agent/projects/${projectId}/demo`),
  downloadAgentDemo: async (runId: string) => {
    const response = await fetchWithTimeout(`${API_BASE}/api/agent/runs/${runId}/demo/download`, {
      headers: { Authorization: `Bearer ${getToken() || ""}` },
    }, 60_000, "下载超时，请重试");
    if (!response.ok) throw new Error("演示文件下载失败，请重新运行演示");
    return response.blob();
  },
  downloadKeyuanReferenceResult: async (runId: string) => {
    const response = await fetchWithTimeout(`${API_BASE}/api/agent/runs/${runId}/result/download`, {
      headers: { Authorization: `Bearer ${getToken() || ""}` },
    }, 60_000, "下载超时，请重试");
    if (!response.ok) throw new Error("结果表格下载失败，请重新开始该案例");
    return response.blob();
  },
  downloadAcceptedAgentFinal: async (runId: string) => {
    const response = await fetchWithTimeout(`${API_BASE}/api/agent/runs/${runId}/acceptance/download`, {
      headers: { Authorization: `Bearer ${getToken() || ""}` },
    }, 60_000, "下载超时，请重试");
    if (!response.ok) throw new Error("正式稿下载失败，请确认验收状态后重试");
    return response.blob();
  },
  listAgentRuns: (projectId: string) =>
    request<{ project_id: string; items: Record<string, unknown>[]; total: number }>(
      `/api/agent/runs?project_id=${encodeURIComponent(projectId)}`,
    ).then((raw) => ({ ...raw, items: raw.items.map(mapAgentRun) })),
  listAgentEvents: (runId: string, afterRevision = 0) =>
    request<{ run_id: string; revision: number; events: AgentRunEvent[] }>(
      `/api/agent/runs/${encodeURIComponent(runId)}/events?after_revision=${normaliseEventCursor(afterRevision)}`,
    ),
  streamAgentEvents: async (
    runId: string,
    afterRevision: number,
    onEvent: (event: AgentRunEvent) => void,
    signal?: AbortSignal,
  ) => {
    const token = getToken();
    const response = await fetch(
      `${API_BASE}/api/agent/runs/${encodeURIComponent(runId)}/events/stream?after_revision=${normaliseEventCursor(afterRevision)}`,
      {
        headers: {
          Accept: "text/event-stream",
          ...(token ? { Authorization: `Bearer ${token}` } : {}),
        },
        signal,
      },
    );
    if (!response.ok) {
      recoverExpiredSession(response.status);
      throw new Error(`处理事件流连接失败 (${response.status})`);
    }
    if (!response.body) throw new Error("处理事件流没有响应体");
    const reader = response.body.getReader();
    const decoder = new TextDecoder();
    let buffer = "";
    while (true) {
      const { value, done } = await reader.read();
      buffer += decoder.decode(value || new Uint8Array(), { stream: !done });
      const parsed = parseSseFrames(buffer);
      buffer = parsed.remaining;
      for (const event of parsed.events) {
        const payload = event.payload as unknown as AgentRunEvent;
        if (typeof payload.revision === "number" && typeof payload.type === "string") onEvent(payload);
      }
      if (done) break;
    }
  },
  planAgentRun: (runId: string, instruction?: string) =>
    request<Record<string, unknown>>(`/api/agent/runs/${runId}/plan`, { method: "POST", body: JSON.stringify({ confirm_month: false, instruction }) }, { timeoutMs: 330_000, timeoutMessage: "计划分析耗时较长，请刷新查看状态；尚未执行工作簿修改" }).then(mapAgentRun),
  confirmAgentPlan: (runId: string, selectedSteps?: string[], instruction?: string) =>
    request<Record<string, unknown>>(`/api/agent/runs/${runId}/plan`, { method: "POST", body: JSON.stringify({ confirm_month: true, confirm_plan: true, selected_steps: selectedSteps, instruction: instruction || undefined }) }).then(mapAgentRun),
  executeAgentRun: (runId: string) =>
    request<Record<string, unknown>>(`/api/agent/runs/${runId}/execute`, { method: "POST" }, { timeoutMs: FINANCIAL_INTEGRATION_TIMEOUT_MS, timeoutMessage: "Agent 仍在处理，请刷新查看运行记录，不要重复提交" }).then(mapAgentRun),
  processAgentRun: (runId: string, instruction?: string) =>
    request<Record<string, unknown>>(`/api/agent/runs/${runId}/process`, {
      method: "POST",
      body: JSON.stringify(instruction ? { instruction } : {}),
    }).then(mapAgentRun),
  getAgentResult: (runId: string, page = 1, pageSize = 50) =>
    request<AgentRunResult>(`/api/agent/runs/${encodeURIComponent(runId)}/output?page=${page}&page_size=${pageSize}`),
  downloadAgentOutput: async (runId: string) => {
    const response = await fetchWithTimeout(`${API_BASE}/api/agent/runs/${encodeURIComponent(runId)}/output/download`, {
      headers: { Authorization: `Bearer ${getToken() || ""}` },
    }, 60_000, "下载超时，请重试");
    if (!response.ok) throw new Error("更新后表格下载失败，请确认校验状态");
    return response.blob();
  },
  listAgentItems: (runId: string) =>
    request<{ items?: Record<string, unknown>[] } | Record<string, unknown>[]>(`/api/agent/runs/${runId}/items`).then((raw) =>
      (Array.isArray(raw) ? raw : raw.items || []).map(mapAgentItem)),
  sendAgentItemMessage: (runId: string, itemId: string, message: string) =>
    request<{ item?: Record<string, unknown> } | Record<string, unknown>>(`/api/agent/runs/${runId}/items/${itemId}/message`, {
      method: "POST",
      body: JSON.stringify({ message }),
    }).then((raw) => mapAgentItem((raw as { item?: Record<string, unknown> }).item || raw as Record<string, unknown>)),
  sendAgentRunMessage: (runId: string, message: string) =>
    request<{ run_id: string; message: AgentMessage }>(`/api/agent/runs/${runId}/message`, {
      method: "POST",
      body: JSON.stringify({ message }),
    }),
  streamAgentRunMessage: async (
    runId: string,
    message: string,
    onEvent: (event: AgentStreamEvent) => void,
  ) => {
    const token = getToken();
    const response = await fetchWithTimeout(
      `${API_BASE}/api/agent/runs/${encodeURIComponent(runId)}/message/stream`,
      {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          Accept: "text/event-stream",
          ...(token ? { Authorization: `Bearer ${token}` } : {}),
        },
        body: JSON.stringify({ message }),
      },
      12_000,
      "未能连接到 Agent 服务，请检查服务状态后重试",
    );
    if (!response.ok) {
      recoverExpiredSession(response.status);
      const body = await response.text();
      throw new Error(body || `消息发送失败 (${response.status})`);
    }
    if (!response.body) throw new Error("Agent 服务未返回流式响应");

    const reader = response.body.getReader();
    const decoder = new TextDecoder();
    let buffer = "";
    while (true) {
      const { value, done } = await reader.read();
      buffer += decoder.decode(value || new Uint8Array(), { stream: !done });
      const parsed = parseSseFrames(buffer);
      buffer = parsed.remaining;
      parsed.events.forEach(onEvent);
      if (done) break;
    }
  },
  applyAgentItem: (runId: string, itemId: string, decision: AgentDecisionInput) =>
    request<Record<string, unknown>>(`/api/agent/runs/${runId}/items/${itemId}/apply`, {
      method: "POST",
      body: JSON.stringify({ ...decision, action: decision.action === "apply" ? "apply_proposed" : decision.action }),
    }).then((raw) => mapAgentItem(((raw.items as Record<string, unknown>[] | undefined) || []).find((item) => item.id === itemId) || raw)),
  retryAgentItem: (runId: string, itemId: string) =>
    request<Record<string, unknown>>(`/api/agent/runs/${runId}/items/${itemId}/retry`, {
      method: "POST",
    }).then((raw) => mapAgentItem(((raw.items as Record<string, unknown>[] | undefined) || []).find((item) => item.id === itemId) || raw)),
  publishAgentRun: (runId: string) =>
    request<Record<string, unknown>>(`/api/agent/runs/${runId}/publish`, { method: "POST", body: JSON.stringify({ confirm_month: true }) }).then(mapAgentRun),
  listAgentMemory: (projectId: string) =>
    request<{ items?: Record<string, unknown>[] }>(`/api/agent/memory?project_id=${encodeURIComponent(projectId)}`).then((raw) => ({
      rules: (raw.items || []).filter((item) => item.status === "active").map(mapMemoryRule),
      suggestions: (raw.items || []).filter((item) => item.status === "candidate").map(mapMemoryRule),
    })),
  suspendAgentMemory: (ruleId: string) =>
    request<AgentMemoryRule>(`/api/agent/memory/${ruleId}/suspend`, { method: "POST" }),
  activateAgentMemory: (ruleId: string) =>
    request<AgentMemoryRule>(`/api/agent/memory/${ruleId}/activate`, { method: "POST" }),
  listAgentMaterials: (projectId: string) =>
    request<{ project_id: string; items: AgentMaterial[]; total: number }>(
      `/api/agent/projects/${encodeURIComponent(projectId)}/materials`,
    ),
  uploadAgentMaterial: (projectId: string, file: File) => {
    const formData = new FormData();
    formData.append("file", file);
    return request<AgentMaterial>(
      `/api/agent/projects/${encodeURIComponent(projectId)}/materials`,
      { method: "POST", body: formData },
      { timeoutMs: UPLOAD_REQUEST_TIMEOUT_MS, timeoutMessage: UPLOAD_TIMEOUT_MESSAGE },
    );
  },
  compileAgentMaterials: (projectId: string, version?: string, effectivePeriod?: string) =>
    request<AgentRulePackageResult>(`/api/agent/projects/${encodeURIComponent(projectId)}/materials/compile`, {
      method: "POST",
      body: JSON.stringify({ version, effective_period: effectivePeriod }),
    }, { timeoutMs: 90_000, timeoutMessage: "规则编译超过 90 秒，请稍后查看材料状态" }),
  listAgentRulePackages: (projectId: string) =>
    request<{ project_id: string; items: AgentRulePackage[]; total: number }>(
      `/api/agent/projects/${encodeURIComponent(projectId)}/rule-packages`,
    ),
  activateAgentRulePackage: (projectId: string, packageId: string) =>
    request<{ project_id: string; package: AgentRulePackage }>(
      `/api/agent/projects/${encodeURIComponent(projectId)}/rule-packages/${encodeURIComponent(packageId)}/activate`,
      { method: "POST", body: JSON.stringify({ confirm: true }) },
    ),
  getAgentModelStatus: () => request<AgentModelStatus>("/api/agent/model/status"),

  // ---------- Files ----------
  uploadFile: (
    projectId: string,
    file: File,
    fileType: string,
    sheetName?: string,
    password?: string,
  ) => uploadProjectFile(projectId, file, fileType, sheetName, password),
  uploadFilesAuto: (projectId: string, files: File[], password?: string) => {
    const fd = new FormData();
    files.forEach((file) => fd.append("files", file));
    if (password) fd.append("password", password);
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

  // ---------- 通用财务工作簿 ----------
  createFinancialWorkbookIntegration: (projectId: string) =>
    request<FinancialWorkbookIntegration>(
      `/api/financial-workbooks/${projectId}/integrations`,
      { method: "POST" },
      {
        timeoutMs: FINANCIAL_INTEGRATION_TIMEOUT_MS,
        timeoutMessage: "财务整合处理超过 15 分钟，请检查文件大小或格式后重试",
      },
    ),
  getFinancialWorkbookIntegrationProgress: (projectId: string) =>
    request<IntegrationProgress>(
      `/api/financial-workbooks/${projectId}/integrations/progress`,
    ),
  getFinancialWorkQueue: (page = 1, pageSize = 50) =>
    request<FinancialWorkQueue>(
      `/api/financial-workbooks/work-queue?page=${page}&page_size=${pageSize}`,
    ),
  getLatestFinancialWorkbookIntegration: (projectId: string) =>
    request<FinancialWorkbookIntegration>(
      `/api/financial-workbooks/${projectId}/integrations/latest`,
    ),
  releaseFinancialWorkbookIntegration: (projectId: string) =>
    request<FinancialWorkbookIntegration>(
      `/api/financial-workbooks/${projectId}/integrations/latest/release`,
      { method: "POST" },
    ),
  listPublishedFinancialWorkbookVersions: (projectId: string) =>
    request<FinancialWorkbookPublishedVersion[]>(
      `/api/financial-workbooks/${projectId}/integrations/published`,
    ),
  financialWorkbookDraftDownloadUrl: (projectId: string) =>
    `${API_BASE}/api/financial-workbooks/${projectId}/integrations/latest/draft/download`,
  financialWorkbookIntegrationDownloadUrl: (projectId: string) =>
    `${API_BASE}/api/financial-workbooks/${projectId}/integrations/latest/download`,
  financialWorkbookReviewDownloadUrl: (projectId: string) =>
    `${API_BASE}/api/financial-workbooks/${projectId}/integrations/latest/review/download`,
  publishedFinancialWorkbookDownloadUrl: (projectId: string, versionId: string) =>
    `${API_BASE}/api/financial-workbooks/${projectId}/integrations/published/${versionId}/download`,

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
    request<IntegrationResult>(
      `/api/pipeline/${projectId}/integrate`,
      { method: "POST" },
      {
        timeoutMs: FINANCIAL_INTEGRATION_TIMEOUT_MS,
        timeoutMessage: "整合处理超过 15 分钟，已停止等待。请刷新页面后查看处理结果。",
      },
    ),
  getIntegrationProgress: (projectId: string) =>
    request<IntegrationProgress>(`/api/pipeline/${projectId}/integrate/progress`),
  getLatestPipelineExport: (projectId: string) =>
    request<PipelineExportResult>(`/api/pipeline/${projectId}/export/latest`),
  getMatchingPolicy: (projectId: string) =>
    request<MatchingPolicy>(`/api/pipeline/${projectId}/matching-policy`),
  updateMatchingPolicy: (projectId: string, policy: Pick<MatchingPolicy, "auto_match_threshold" | "field_weights" | "source_rules">) =>
    request<MatchingPolicy>(`/api/pipeline/${projectId}/matching-policy`, {
      method: "PUT",
      body: JSON.stringify(policy),
    }),
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
  applyOnlineManualReview: (projectId: string, decisions: ManualReviewDecision[]) =>
    request<PipelineExportResult>(
      `/api/pipeline/${projectId}/manual-review/apply`,
      { method: "POST", body: JSON.stringify({ decisions }) },
    ),
  applyManualSheetMapping: (projectId: string, issueId: string, targetSheet: string) =>
    request<PipelineExportResult>(
      `/api/pipeline/${projectId}/manual-review/sheet-mappings`,
      { method: "POST", body: JSON.stringify({ issue_id: issueId, target_sheet: targetSheet }) },
      {
        timeoutMs: FINANCIAL_INTEGRATION_TIMEOUT_MS,
        timeoutMessage: "正在根据确认结果更新总表，请稍后刷新查看最新结果。",
      },
    ),
  applyManualSheetMappings: (projectId: string, mappings: Array<{ issue_id: string; target_sheet: string }>) =>
    request<PipelineExportResult>(
      `/api/pipeline/${projectId}/manual-review/sheet-mappings/batch`,
      { method: "POST", body: JSON.stringify({ mappings }) },
      {
        timeoutMs: FINANCIAL_INTEGRATION_TIMEOUT_MS,
        timeoutMessage: "正在根据确认结果批量更新总表，请稍后刷新查看最新结果。",
      },
    ),
  ignoreSemanticManualIssue: (projectId: string, issueId: string) =>
    request<PipelineExportResult>(
      `/api/pipeline/${projectId}/manual-review/semantic-issues/ignore`,
      { method: "POST", body: JSON.stringify({ issue_id: issueId }) },
      {
        timeoutMs: FINANCIAL_INTEGRATION_TIMEOUT_MS,
        timeoutMessage: "正在保存本次不更新的选择，请稍后刷新查看最新结果。",
      },
    ),
  importManualReviewWorkbook: (projectId: string, file: File) => {
    const formData = new FormData();
    formData.append("file", file);
    return request<PipelineExportResult>(
      `/api/pipeline/${projectId}/manual-review/import`,
      { method: "POST", body: formData },
    );
  },
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
  status: "pending" | "running" | "completed";
}

export interface IntegrationProgress {
  status: "idle" | "processing" | "completed" | "completed_with_issues" | "blocked" | "failed";
  stages: IntegrationStage[];
  detail?: string | null;
  updated_at: string;
}

export interface ManualIssue {
  issue_id?: string;
  issue_type: string;
  employee_id?: string;
  person_name?: string;
  target_field?: string;
  problem_type?: string;
  difficulty?: "简单" | "一般" | "复杂";
  suggestion?: string;
  recommended_action?: "apply_proposed" | "keep_current" | "copy_only";
  action_options?: Array<{ value: string; label: string }>;
  copy_text?: string;
  current_value?: string | number | boolean | null;
  proposed_value?: string | number | boolean | null;
  candidate_values?: Array<string | number | boolean | null>;
  target_sheet?: string;
  target_cell?: string;
  target_location?: string;
  can_update_online?: boolean;
  status?: "pending" | "confirmed";
  message: string;
  source_files: string[];
  source_sheets: string[];
  candidate_target_sheets?: string[];
  confidence?: number;
}

export interface ManualReviewDecision {
  issue_id: string;
  outcome: "更新到总表" | "保留总表";
  value: string | number | boolean | null;
  note: string;
  remember?: boolean;
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
  target_sheets?: string[];
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
