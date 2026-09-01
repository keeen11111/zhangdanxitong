import pandas as pd

from core.entity_engine import _parse_bonus_sheet


def test_keyuan_bonus_sums_monthly_and_quarterly_columns() -> None:
    source = pd.DataFrame([
        ["姓名", "工号", "绩效奖金-月度", "绩效奖金-季度", "项目奖金-非DTP"],
        ["仅月度", "E001", 2000, 0, 100],
        ["仅季度", "E002", 0, 3000, 200],
        ["两列都有", "E003", 500, 600, 300],
        ["均为零", "E004", 0, 0, 400],
    ])

    _, salary_rows, _ = _parse_bonus_sheet(source, "奖金.xlsx", "奖金-7月")

    assert [row["绩效奖金"] for row in salary_rows] == [2000, 3000, 1100, 0]
    assert [row["业绩奖金"] for row in salary_rows] == [100, 200, 300, 400]
