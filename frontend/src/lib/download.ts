import { getToken } from "@/lib/api";
import { triggerBrowserDownload } from "@/lib/browser-download";

type SaveFilePicker = (options: {
  suggestedName: string;
  types: Array<{
    description: string;
    accept: Record<string, string[]>;
  }>;
}) => Promise<{
  createWritable: () => Promise<{
    write: (data: Blob) => Promise<void>;
    close: () => Promise<void>;
  }>;
}>;

type SavePickerWindow = Window & { showSaveFilePicker?: SaveFilePicker };

async function fetchWorkbookBlob(url: string): Promise<Blob> {
  const response = await fetch(url, { headers: { Authorization: `Bearer ${getToken()}` } });
  if (!response.ok) throw new Error("下载失败");
  return response.blob();
}

export async function downloadWorkbookDirect(url: string, filename: string): Promise<void> {
  const blob = await fetchWorkbookBlob(url);
  triggerBrowserDownload(blob, filename);
}

export async function downloadWorkbookWithSaveAs(url: string, filename: string): Promise<void> {
  const savePicker = (window as SavePickerWindow).showSaveFilePicker;
  let fileHandle: Awaited<ReturnType<SaveFilePicker>> | undefined;

  if (savePicker) {
    try {
      fileHandle = await savePicker({
        suggestedName: filename,
        types: [{
          description: "Excel 工作簿",
          accept: { "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": [".xlsx"] },
        }],
      });
    } catch (error) {
      if (error instanceof DOMException && error.name === "AbortError") return;
      throw error;
    }
  }

  const blob = await fetchWorkbookBlob(url);

  if (fileHandle) {
    const writable = await fileHandle.createWritable();
    await writable.write(blob);
    await writable.close();
    return;
  }

  triggerBrowserDownload(blob, filename);
}
