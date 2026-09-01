"use client";

import { useRef, useState } from "react";
import { FileSpreadsheet, Loader2, Trash2, Upload } from "lucide-react";

import { FileMeta } from "@/lib/api";
import { cn } from "@/lib/utils";

interface WorkbookUploadPanelProps {
  title: string;
  description: string;
  helper: string;
  files: FileMeta[];
  multiple?: boolean;
  uploading: boolean;
  disabled?: boolean;
  tone: "master" | "changes";
  onSelect: (files: File[]) => void;
  onRemove: (file: FileMeta) => void;
}

function fileSummary(file: FileMeta) {
  const size = !file.row_count && !file.col_count
    ? "图片 / 图表附件"
    : `${file.row_count} 行 · ${file.col_count} 列`;
  const details = [
    size,
    file.normalization_note,
    file.processing_seconds ? `处理 ${file.processing_seconds.toFixed(1)} 秒` : null,
  ].filter(Boolean);
  return details.join(" · ");
}

export function WorkbookUploadPanel({
  title,
  description,
  helper,
  files,
  multiple = false,
  uploading,
  disabled = false,
  tone,
  onSelect,
  onRemove,
}: WorkbookUploadPanelProps) {
  const inputRef = useRef<HTMLInputElement>(null);
  const [dragging, setDragging] = useState(false);
  const isMaster = tone === "master";
  const unavailable = disabled || uploading;

  function acceptFiles(selected: FileList | File[]) {
    if (unavailable) return;
    onSelect(Array.from(selected));
  }

  return (
    <section
      aria-busy={uploading}
      className={cn(
        "rounded-md border bg-white p-4 sm:p-5",
        isMaster ? "border-teal-200" : "border-slate-200"
      )}
    >
      <div className="flex items-start justify-between gap-4">
        <div>
          <div className="flex items-center gap-2">
            <span
              className={cn(
                "inline-flex h-6 min-w-6 items-center justify-center rounded-full px-2 text-xs font-semibold",
                isMaster ? "bg-teal-700 text-white" : "bg-slate-100 text-slate-700"
              )}
            >
              {isMaster ? "1" : "2"}
            </span>
            <h2 className="text-sm font-semibold text-slate-950">{title}</h2>
          </div>
          <p className="mt-2 text-sm leading-6 text-slate-600">{description}</p>
        </div>
        <span className="shrink-0 text-xs tabular-nums text-slate-500">{files.length} 个文件</span>
      </div>

      <input
        ref={inputRef}
        className="hidden"
        type="file"
        accept=".xlsx,.xls"
        multiple={multiple}
        onChange={(event) => {
          if (event.target.files) acceptFiles(event.target.files);
          event.target.value = "";
        }}
      />
      <button
        type="button"
        disabled={unavailable}
        onClick={() => inputRef.current?.click()}
        onDragEnter={(event) => {
          event.preventDefault();
          if (!unavailable) setDragging(true);
        }}
        onDragOver={(event) => event.preventDefault()}
        onDragLeave={() => setDragging(false)}
        onDrop={(event) => {
          event.preventDefault();
          setDragging(false);
          acceptFiles(event.dataTransfer.files);
        }}
        className={cn(
          "mt-4 flex min-h-28 w-full items-center justify-center gap-3 rounded-md border border-dashed px-5 text-left transition-colors focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-teal-700 focus-visible:ring-offset-2 disabled:cursor-not-allowed disabled:opacity-60",
          dragging
            ? "border-teal-600 bg-teal-50"
            : isMaster
            ? "border-teal-300 bg-teal-50/50 hover:border-teal-500"
            : "border-slate-300 bg-slate-50 hover:border-teal-500 hover:bg-teal-50/50"
        )}
      >
        {uploading ? (
          <Loader2 className="h-5 w-5 shrink-0 animate-spin text-teal-700" />
        ) : (
          <Upload className="h-5 w-5 shrink-0 text-teal-700" />
        )}
        <span aria-live="polite">
          <span className="block text-sm font-medium text-slate-900">
            {uploading ? "正在上传并检查文件…" : disabled ? "请先移除当前总表" : helper}
          </span>
          <span className="mt-1 block text-xs text-slate-500">
            {uploading
              ? "普通 .xlsx 通常几秒完成；旧版 .xls 最长约 1 分钟"
              : "支持普通 .xlsx / .xls；加密文件会在上传时弹出密码输入"}
          </span>
        </span>
      </button>


      {files.length ? (
        <ul className="mt-3 divide-y divide-slate-100 rounded-lg border border-slate-200 px-3">
          {files.map((file) => (
            <li key={file.id} className="flex items-center gap-3 py-3">
              <FileSpreadsheet className="h-4 w-4 shrink-0 text-slate-400" />
              <div className="min-w-0 flex-1">
                <p className="truncate text-sm font-medium text-slate-800" title={file.original_name}>
                  {file.original_name}
                </p>
                <p className="mt-0.5 text-xs text-slate-500">{fileSummary(file)}</p>
              </div>
              <button
                type="button"
                onClick={() => onRemove(file)}
                className="inline-flex h-8 w-8 items-center justify-center rounded-md text-slate-400 hover:bg-red-50 hover:text-red-700 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-red-600"
                aria-label={`移除 ${file.original_name}`}
              >
                <Trash2 className="h-4 w-4" />
              </button>
            </li>
          ))}
        </ul>
      ) : (
        <p className="mt-3 text-xs text-slate-500">尚未上传。</p>
      )}
    </section>
  );
}
