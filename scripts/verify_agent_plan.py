"""Exercise the real upload/create/plan HTTP routes in an isolated test database.

Uses the configured model. Does not confirm or execute the plan, and never
uploads a final reference workbook. Password is read only from the process env.
"""
from __future__ import annotations

import argparse
import hashlib
import json
import os
from pathlib import Path
import sys
import tempfile


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--master", type=Path, required=True)
    parser.add_argument("--source", type=Path, action="append", required=True)
    parser.add_argument("--material", type=Path, action="append", default=[])
    parser.add_argument("--month", required=True)
    parser.add_argument("--instruction", required=True)
    args = parser.parse_args()
    paths = [args.master, *args.source, *args.material]
    hashes = {str(path): hashlib.sha256(path.read_bytes()).hexdigest() for path in paths}
    output = Path(tempfile.mkdtemp(prefix="agent-plan-check-"))
    os.environ["PAYROLL_DATA_DIR"] = str(output)
    os.environ["PAYROLL_DATABASE_URL"] = "sqlite:///" + (output / "check.db").as_posix()
    sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
    from fastapi.testclient import TestClient
    from backend.auth import get_current_user
    from backend.database import SessionLocal, init_db
    from backend.main import app
    from backend.models import Project, Tenant, User

    init_db()
    with SessionLocal() as db:
        tenant = Tenant(name="Agent 验证隔离空间")
        user = User(name="验证用户", email="agent-check@example.com", hashed_password="disabled", tenant=tenant)
        project = Project(name="科园样本 Agent 只读验证", salary_month=args.month, owner=user)
        db.add(project)
        db.commit()
        project_id = project.id
        app.dependency_overrides[get_current_user] = lambda: user
        with TestClient(app) as client:
            for path in [args.master, *args.source]:
                role = "financial_master" if path == args.master else "financial_source"
                with path.open("rb") as stream:
                    response = client.post(f"/api/projects/{project_id}/files", data={
                        "file_type": role, "password": os.getenv("SAMPLE_WORKBOOK_PASSWORD", ""),
                    }, files={"file": (path.name, stream)})
                response.raise_for_status()
                print(f"Uploaded {role}: {path.name}", flush=True)
            for path in args.material:
                with path.open("rb") as stream:
                    response = client.post(f"/api/agent/projects/{project_id}/materials", files={"file": (path.name, stream)})
                response.raise_for_status()
            response = client.post("/api/agent/runs", json={"project_id": project_id, "instruction": args.instruction})
            response.raise_for_status()
            run_id = response.json()["run_id"]
            print(f"Created read-only run: {run_id}; output: {output}", flush=True)
            response = client.post(f"/api/agent/runs/{run_id}/plan", json={"confirm_plan": False})
            if response.status_code != 200:
                print(response.text, flush=True)
            response.raise_for_status()
            result = response.json()
            (output / "plan.json").write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
            print(json.dumps(result["model_plan"], ensure_ascii=False, indent=2), flush=True)
            assert not result.get("draft_filename")
            assert not list((output / "exports").rglob("*.xlsx"))
            assert hashes == {str(path): hashlib.sha256(path.read_bytes()).hexdigest() for path in paths}
            print("PASS: real Agent plan generated; zero drafts; all original file hashes unchanged", flush=True)


if __name__ == "__main__":
    main()
