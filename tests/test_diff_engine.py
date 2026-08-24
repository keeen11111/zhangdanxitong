"""流水线差异整合的回归测试。"""

from core.diff_engine import apply_confirmed_diff, compute_diff


def test_confirmed_new_hire_keeps_salary_and_unique_social_identity() -> None:
    base = {"employee_profile": [], "salary_detail": []}
    source = {
        "employee_profile": [
            {
                "工号": "E001",
                "姓名": "新增员工",
                "本月状态": "入职",
                "公司": "测试公司",
                "转正日期": "2026-12-01",
            }
        ],
        "salary_detail": [
            {
                "工号": "E001",
                "姓名": "新增员工",
                "月基本薪资": 2500,
                "月绩效标准": 1500,
                "交通通讯补贴标准": 400,
            }
        ],
        "social_security": [
            {
                "姓名": "新增员工",
                "身份证号": "110101199001011234",
            }
        ],
    }

    diff = compute_diff(base, source)
    diff["hire"][0]["status"] = "confirmed"
    final = apply_confirmed_diff(base, diff)

    profile = final["employee_profile"][0]
    assert profile["身份证号"] == "110101199001011234"
    assert profile["转正日期"] == "2026-12-01"

    salary = final["salary_detail"][0]
    assert salary["工号"] == "E001"
    assert salary["月基本薪资"] == 2500
    assert salary["月绩效标准"] == 1500
    assert salary["交通通讯补贴标准"] == 400
