const { cpSync, existsSync, mkdirSync, rmSync } = require("node:fs");
const { execFileSync } = require("node:child_process");
const path = require("node:path");

const frontendDir = path.resolve(__dirname, "..");
const projectDir = path.resolve(frontendDir, "..");
const standaloneDir = path.join(frontendDir, ".next", "standalone");
const runtimeStaticDir = path.join(standaloneDir, ".next", "static");
const sourceStaticDir = path.join(frontendDir, ".next", "static");
const distDir = path.join(projectDir, "dist", "外服账单系统后端");

function run(command, args, options = {}) {
  if (process.platform === "win32") {
    const commandLine = [command, ...args]
      .map((argument) => (/[\s"]/u.test(argument) ? `"${argument.replace(/"/gu, '\\"')}"` : argument))
      .join(" ");
    execFileSync("cmd.exe", ["/d", "/s", "/c", commandLine], { stdio: "inherit", ...options });
    return;
  }
  execFileSync(command, args, { stdio: "inherit", ...options });
}

run("npm.cmd", ["run", "build"], {
  cwd: frontendDir,
  env: { ...process.env, NEXT_PUBLIC_API_BASE: "http://127.0.0.1:18000" },
});

mkdirSync(path.dirname(runtimeStaticDir), { recursive: true });
rmSync(runtimeStaticDir, { recursive: true, force: true });
cpSync(sourceStaticDir, runtimeStaticDir, { recursive: true });

rmSync(distDir, { recursive: true, force: true });
run("python", [
  "-m", "PyInstaller", "--noconfirm", "--clean", "--onedir",
  "--name", "外服账单系统后端",
  "--add-data", "backend/data/templates;templates",
  "desktop_backend.py",
], { cwd: projectDir });

run("npx.cmd", ["electron-builder", "--win", "nsis", "--x64"], { cwd: frontendDir });
