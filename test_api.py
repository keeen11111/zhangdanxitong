import requests

BASE = "http://127.0.0.1:8000"

# 1. 登录
r = requests.post(f"{BASE}/api/auth/login", json={"email": "demo@autopayroll.com", "password": "demo123"})
tok = r.json()["access_token"]
H = {"Authorization": f"Bearer {tok}"}
print("1. 登录:", r.status_code)

# 2. 列项目
r = requests.get(f"{BASE}/api/projects", headers=H)
projects = r.json()
print("2. 项目数:", len(projects))
if not projects:
    r = requests.post(f"{BASE}/api/projects", json={"name": "API测试", "salary_month": "2026.06"}, headers=H)
    pid = r.json()["id"]
else:
    pid = projects[0]["id"]
print("   项目ID:", pid)

# 3. 上传文件
with open("test_sample.xlsx", "rb") as f:
    r = requests.post(
        f"{BASE}/api/projects/{pid}/files",
        headers=H,
        files={"file": ("test_sample.xlsx", f, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")},
        data={"file_type": "last_month"},
    )
print("3. 上传文件:", r.status_code, r.json().get("original_name", "ERR:" + r.text))

# 4. 加载板块
r = requests.post(f"{BASE}/api/modules/{pid}/load", headers=H)
print("4. 加载板块:", r.status_code)
if r.status_code == 200:
    mods = r.json()
    print("   板块数:", len(mods))
    for m in mods:
        print("   板块:", m["name"], f"({m['row_count']}行 x {m['col_count']}列)")

    # 5. 获取第一个板块数据
    mid = mods[0]["id"]
    r = requests.get(f"{BASE}/api/modules/{pid}/{mid}", headers=H)
    print("5. 获取板块数据:", r.status_code)
    if r.status_code == 200:
        data = r.json()
        print("   列:", data["columns"])
        print("   首行:", data["rows"][0])

        # 6. 保存（修改第一行姓名）
        data["rows"][0]["姓名"] = "张三丰"
        r = requests.put(
            f"{BASE}/api/modules/{pid}/{mid}",
            headers=H,
            json={"columns": data["columns"], "rows": data["rows"]},
        )
        print("6. 保存板块:", r.status_code)
        if r.status_code == 200:
            print("   首行姓名=", r.json()["rows"][0]["姓名"])

    # 7. 导出
    r = requests.post(f"{BASE}/api/modules/{pid}/export", headers=H)
    print("7. 导出:", r.status_code)
    if r.status_code == 200:
        print("   文件名:", r.json()["filename"])
        print("   日志:", r.json()["log"])
    else:
        print("   ERR:", r.text)
else:
    print("   ERR:", r.text)
