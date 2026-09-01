"""Portable launcher for the bundled payroll application."""

from __future__ import annotations

import os
import subprocess
import sys
import time
import urllib.request
import webbrowser
from pathlib import Path


HOST = "127.0.0.1"
API_PORT = 18000
WEB_PORT = 13000


def resource_root() -> Path:
    return Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))


def wait_for_url(port: int, timeout: float = 40.0) -> None:
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        try:
            with urllib.request.urlopen(f"http://{HOST}:{port}/", timeout=1):
                return
        except Exception:
            time.sleep(0.25)
    raise RuntimeError(f"服务启动超时（端口 {port}）")


def main() -> None:
    root = resource_root()
    data_dir = Path(os.environ.get("LOCALAPPDATA", root)) / "外服账单系统" / "data"
    data_dir.mkdir(parents=True, exist_ok=True)
    env = {
        **os.environ,
        "PAYROLL_DATA_DIR": str(data_dir),
        "PAYROLL_ENV": "desktop",
        "PAYROLL_CORS_ORIGINS": f"http://{HOST}:{WEB_PORT}",
    }
    backend = subprocess.Popen(
        [str(root / "backend" / "外服账单系统后端.exe")],
        cwd=str(root / "backend"),
        env=env,
        creationflags=subprocess.CREATE_NO_WINDOW,
    )
    frontend = None
    try:
        wait_for_url(API_PORT)
        frontend = subprocess.Popen(
            [str(root / "node.exe"), str(root / "frontend" / "server.js")],
            cwd=str(root / "frontend"),
            env={**env, "ELECTRON_RUN_AS_NODE": "1", "HOSTNAME": HOST, "PORT": str(WEB_PORT)},
            creationflags=subprocess.CREATE_NO_WINDOW,
        )
        wait_for_url(WEB_PORT)
        webbrowser.open(f"http://{HOST}:{WEB_PORT}")
        print(f"外服账单系统已启动：http://{HOST}:{WEB_PORT}")
        print("关闭此窗口将停止服务。")
        while True:
            if backend.poll() is not None:
                raise RuntimeError("后端服务意外退出")
            if frontend.poll() is not None:
                raise RuntimeError("前端服务意外退出")
            time.sleep(1)
    except KeyboardInterrupt:
        pass
    except Exception as exc:
        print(f"启动失败：{exc}")
        input("按回车键退出……")
    finally:
        for process in (frontend, backend):
            if process and process.poll() is None:
                process.terminate()
                try:
                    process.wait(timeout=5)
                except subprocess.TimeoutExpired:
                    process.kill()


if __name__ == "__main__":
    main()
