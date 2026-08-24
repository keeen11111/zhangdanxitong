const { app, BrowserWindow, dialog } = require("electron");
const { spawn } = require("node:child_process");
const http = require("node:http");
const path = require("node:path");

const HOST = "127.0.0.1";
const API_PORT = 18000;
const WEB_PORT = 13000;
let backendProcess;
let webProcess;

function waitForServer(port, label) {
  return new Promise((resolve, reject) => {
    const deadline = Date.now() + 30000;
    const check = () => {
      const request = http.get({ host: HOST, port, path: "/" }, (response) => {
        response.resume();
        if (response.statusCode && response.statusCode < 500) return resolve();
        retry();
      });
      request.on("error", retry);
      request.setTimeout(1000, () => request.destroy());
    };
    const retry = () => {
      if (Date.now() >= deadline) return reject(new Error(`${label} 启动超时`));
      setTimeout(check, 250);
    };
    check();
  });
}

function stopProcess(child) {
  if (child && !child.killed) child.kill();
}

async function startApplication() {
  const userDataDir = path.join(app.getPath("userData"), "data");
  const runtimeEnv = {
    ...process.env,
    PAYROLL_DATA_DIR: userDataDir,
    PAYROLL_ENV: "desktop",
    PAYROLL_CORS_ORIGINS: `http://${HOST}:${WEB_PORT}`,
  };
  backendProcess = spawn(
    path.join(process.resourcesPath, "backend", "外服账单系统后端.exe"),
    [],
    { env: runtimeEnv, windowsHide: true }
  );
  backendProcess.on("error", (error) => dialog.showErrorBox("后端启动失败", error.message));
  await waitForServer(API_PORT, "数据服务");

  webProcess = spawn(
    process.execPath,
    [path.join(process.resourcesPath, "frontend", "server.js")],
    {
      env: { ...runtimeEnv, ELECTRON_RUN_AS_NODE: "1", HOSTNAME: HOST, PORT: String(WEB_PORT) },
      windowsHide: true,
    }
  );
  webProcess.on("error", (error) => dialog.showErrorBox("界面启动失败", error.message));
  await waitForServer(WEB_PORT, "界面服务");

  const window = new BrowserWindow({
    width: 1360,
    height: 900,
    minWidth: 1024,
    minHeight: 720,
    title: "外服账单系统",
    autoHideMenuBar: true,
    webPreferences: { contextIsolation: true, nodeIntegration: false },
  });
  await window.loadURL(`http://${HOST}:${WEB_PORT}`);
}

app.whenReady().then(startApplication).catch((error) => {
  dialog.showErrorBox("外服账单系统启动失败", error.message);
  app.quit();
});

app.on("window-all-closed", () => app.quit());
app.on("before-quit", () => {
  stopProcess(webProcess);
  stopProcess(backendProcess);
});
