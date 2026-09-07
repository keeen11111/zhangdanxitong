"""Exercise the same authenticated HTTP upload/integrate/download flow as the UI."""
from io import BytesIO
from pathlib import Path
from uuid import uuid4

from fastapi.testclient import TestClient
from openpyxl import load_workbook
from sqlalchemy import create_engine
from sqlalchemy.orm import sessionmaker

from backend.database import Base, get_db
from backend.main import app
from backend.routers import pipeline, projects
from backend.tests.core.test_structured_hire import structured_files


def test_personnel_notice_http_workflow_and_tenant_isolation(tmp_path: Path, monkeypatch) -> None:
    master, source, _ = structured_files(tmp_path)
    engine = create_engine(f'sqlite:///{tmp_path / "test.db"}', connect_args={'check_same_thread': False})
    Base.metadata.create_all(engine)
    sessions = sessionmaker(bind=engine)
    def database():
        with sessions() as session:
            yield session
    app.dependency_overrides[get_db] = database
    for module in (projects, pipeline):
        monkeypatch.setattr(module, 'UPLOAD_DIR', str(tmp_path / 'uploads'))
    monkeypatch.setattr(pipeline, 'EXPORT_DIR', str(tmp_path / 'exports'))
    monkeypatch.setattr(pipeline, 'SESSION_DIR', str(tmp_path / 'sessions'))
    (tmp_path / 'sessions').mkdir()
    try:
        client = TestClient(app)
        registration = client.post('/api/auth/register', json={
            'email': 'merge@example.com', 'name': '系统验证', 'tenant_name': '合并测试', 'password': uuid4().hex,
        })
        assert registration.status_code == 200, registration.text
        headers = {'Authorization': 'Bearer ' + registration.json()['access_token']}
        project = client.post('/api/projects', json={'name': '入职合并验证', 'salary_month': '2026.05'}, headers=headers)
        assert project.status_code == 201, project.text
        project_id = project.json()['id']
        response = client.post(f'/api/projects/{project_id}/files/batch-auto', files=[
            ('files', ('工资总表.xlsx', master.read_bytes())), ('files', ('客户通知.xlsx', source.read_bytes())),
        ], headers=headers)
        assert response.status_code == 201, response.text
        assert {f['file_type'] for f in response.json()} == {'template', 'source'}
        response = client.post(f'/api/pipeline/{project_id}/integrate', headers=headers)
        assert response.status_code == 200, response.text
        result = response.json()
        assert result['status'] == 'completed'
        assert result['output_person_count'] == 2
        assert result['validation']['personnel_coverage']['matched_field_count'] == 17
        for suffix in ('download', 'review/download'):
            response = client.get(f'/api/pipeline/{project_id}/export/{suffix}', headers=headers)
            assert response.status_code == 200, response.text
            book = load_workbook(BytesIO(response.content), data_only=True)
            assert book['工资核算']['N4'].value == 200
            assert book['工资核算']['U4'].value is None
        other = client.post('/api/auth/register', json={'email': 'other@example.com', 'name': '其他租户', 'password': uuid4().hex})
        forbidden = client.get(f'/api/pipeline/{project_id}/export/download', headers={
            'Authorization': 'Bearer ' + other.json()['access_token'],
        })
        assert forbidden.status_code in (403, 404)
        again = client.post(f'/api/pipeline/{project_id}/integrate', headers=headers)
        assert again.status_code == 200
        assert again.json()['output_person_count'] == 2
    finally:
        app.dependency_overrides.clear()
        engine.dispose()
