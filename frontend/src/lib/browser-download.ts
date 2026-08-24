type DownloadAnchor = {
  href: string;
  download: string;
  click: () => void;
};

type BrowserDownloadEnvironment = {
  createAnchor: () => DownloadAnchor;
  attachAnchor: (anchor: DownloadAnchor) => void;
  removeAnchor: (anchor: DownloadAnchor) => void;
  createObjectUrl: (blob: Blob) => string;
  revokeObjectUrl: (url: string) => void;
  scheduleCleanup: (cleanup: () => void) => void;
};

function defaultEnvironment(): BrowserDownloadEnvironment {
  return {
    createAnchor: () => document.createElement("a"),
    attachAnchor: (anchor) => document.body.append(anchor as HTMLAnchorElement),
    removeAnchor: (anchor) => (anchor as HTMLAnchorElement).remove(),
    createObjectUrl: (blob) => URL.createObjectURL(blob),
    revokeObjectUrl: (url) => URL.revokeObjectURL(url),
    scheduleCleanup: (cleanup) => window.setTimeout(cleanup, 1000),
  };
}

export function triggerBrowserDownload(
  blob: Blob,
  filename: string,
  environment: BrowserDownloadEnvironment = defaultEnvironment(),
): void {
  const objectUrl = environment.createObjectUrl(blob);
  const anchor = environment.createAnchor();
  anchor.href = objectUrl;
  anchor.download = filename;
  environment.attachAnchor(anchor);
  try {
    anchor.click();
  } finally {
    environment.scheduleCleanup(() => {
      environment.removeAnchor(anchor);
      environment.revokeObjectUrl(objectUrl);
    });
  }
}
