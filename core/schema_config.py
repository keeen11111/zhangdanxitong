"""业务实体 Schema 配置：实体定义 + 导入映射 + 导出映射。

核心设计：
    源数据列 --(导入映射)--> 实体字段 --(导出映射)--> 模板列位

三层解耦：源格式变化只调导入映射，模板变化只调导出映射。
"""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Optional


# ============================================================
# 实体定义
# ============================================================
@dataclass
class EntityField:
    """实体字段定义。"""
    name: str  # 标准字段名（中文）
    dtype: str  # text / amount / date / int
    required: bool = False  # 是否核心项
    aliases: list[str] = field(default_factory=list)  # 源数据中可能的列名别名


@dataclass
class EntityDef:
    """业务实体定义。"""
    id: str  # 实体标识，如 employee_profile
    name: str  # 显示名，如 员工档案
    primary_key: str  # 主键字段名
    foreign_key: Optional[str]  # 外键（关联到员工档案），None 表示独立
    fields: list[EntityField]
    icon: str = "📋"  # 前端图标


# ============================================================
# 5 个业务实体
# ============================================================

EMPLOYEE_PROFILE = EntityDef(
    id="employee_profile",
    name="员工档案",
    primary_key="工号",
    foreign_key=None,
    icon="👤",
    fields=[
        EntityField("工号", "text", required=True, aliases=["新工号", "员工工号", "员工编号", "工 号", "员工ID", "ID", "工牌号", "人员编码"]),
        EntityField("姓名", "text", required=True, aliases=["人员姓名", "员工姓名", "名字"]),
        EntityField("工作地点", "text", aliases=["工作地", "地点"]),
        EntityField("公司", "text", aliases=["新公司", "所属公司", "缴费公司", "单位名称", "公司名称"]),
        EntityField("部门", "text", aliases=["新成本中心/部门", "成本中心/部门", "部门", "一级部门", "成本中心"]),
        EntityField("岗位", "text", aliases=["新岗位", "岗位", "职位"]),
        EntityField("人员类别", "text", aliases=["人员类型", "类别"]),
        EntityField("合同类别", "text", aliases=["合同类型"]),
        EntityField("本月状态", "text", aliases=["状态"]),
        EntityField("入职日期", "date", aliases=["入职时间", "入职日"]),
        EntityField("离职日期", "date", aliases=["离职时间", "离职日"]),
        EntityField("是否试用期", "text", aliases=["试用期状态"]),
        EntityField("转正日期", "date", aliases=["转正时间", "试用期结束时间"]),
        EntityField("试用期标准", "amount", aliases=["试用期薪资"]),
        EntityField("纳税类别", "text", aliases=["纳税类型"]),
        EntityField("身份证号", "text", required=True, aliases=["身份证号码", "证件号码", "身份证", "身份证号"]),
        EntityField("银行卡号", "text", aliases=["工资银行卡账号", "银行卡号", "银行账号"]),
        EntityField("开户银行", "text", aliases=["银行", "开户行"]),
    ],
)

SALARY_DETAIL = EntityDef(
    id="salary_detail",
    name="薪资明细",
    primary_key="工号",
    foreign_key="工号",
    icon="💰",
    fields=[
        EntityField("工号", "text", required=True, aliases=["新工号", "员工工号", "员工编号"]),
        EntityField("姓名", "text", aliases=["人员姓名"]),
        EntityField("月基本薪资", "amount", aliases=["基本工资", "月基本工资", "基本薪资"]),
        EntityField("月绩效标准", "amount", aliases=["绩效工资", "绩效标准", "月绩效"]),
        EntityField("DTP绩效标准", "amount", aliases=["额外DTP绩效标准", "DTP标准"]),
        EntityField("考核成绩", "amount", aliases=["绩效考核", "考核分数"]),
        EntityField("绩效奖金", "amount", aliases=["绩效奖金(月)", "绩效奖金_月度", "绩效奖金_月", "绩效"]),
        EntityField("业绩奖金", "amount", aliases=["销售佣金(月)", "提成小计", "提成", "业绩", "佣金", "项目奖金_非DTP"]),
        EntityField("DTP绩效奖金", "amount", aliases=["计件奖金(月)", "计件激励小计", "计件激励", "项目奖金_DTP", "DTP奖金"]),
        EntityField("年终奖", "amount", aliases=["年终奖"]),
        EntityField("其它工资调差", "amount", aliases=["税前补发", "工资调差", "其他调差"]),
        EntityField("交通通讯补贴标准", "amount", aliases=["交通通讯补贴", "车贴", "交通补贴"]),
        # 模板公式会区分“50元/天”“500元/月”，必须保留单位，不能清洗成 0。
        EntityField("餐补标准", "text", aliases=["餐补", "饭贴", "伙食津贴"]),
        EntityField("药师津贴", "amount", aliases=["药师津贴(月)"]),
        EntityField("店长津贴", "amount", aliases=["岗位津贴(月)", "岗位津贴"]),
        EntityField("其他补贴", "amount", aliases=["其他补贴特殊约定", "综合津贴(月)", "其他津贴"]),
        EntityField("过节费", "amount", aliases=["节日津贴标准"]),
        EntityField("防暑降温费标准", "amount", aliases=["高温费", "高温费标准"]),
        EntityField("天津取暖补贴", "amount", aliases=["取暖补贴"]),
        EntityField("出差补贴", "amount", aliases=["出差津贴"]),
        EntityField("独生子女补贴", "amount", aliases=["独生子女费", "独生子女津贴"]),
        EntityField("补偿金", "amount", aliases=["经济补偿金"]),
        # 本月实际发放金额（部分来自考勤计算）
        EntityField("本月交通通讯补贴", "amount", aliases=["本月交通补贴"]),
        EntityField("本月餐补", "amount", aliases=["本月伙食补贴"]),
        EntityField("本月防暑降温费", "amount", aliases=["防暑降温费", "本月高温费"]),
        EntityField("本月节日津贴", "amount", aliases=["节日津贴", "本月过节费"]),
    ],
)

ATTENDANCE_RECORD = EntityDef(
    id="attendance_record",
    name="考勤记录",
    primary_key="工号",
    foreign_key="工号",
    icon="📅",
    fields=[
        EntityField("工号", "text", required=True, aliases=["新工号", "员工工号", "员工编号"]),
        EntityField("姓名", "text", aliases=["人员姓名"]),
        EntityField("实际出勤天数", "amount", aliases=["出勤天数", "本月实际出勤天数", "出勤"]),
        EntityField("普通加班时长", "amount", aliases=["平时加班（小时）", "加班时长"]),
        EntityField("工作日加班时长", "amount", aliases=["工作日加班时长（加班费）(求和)", "工作日加班（求和）"]),
        EntityField("公休日加班时长", "amount", aliases=["公休日加班时长（加班费）(求和)", "公休日加班（求和）"]),
        EntityField("法定节日加班时长", "amount", aliases=["节假日加班（小时）", "节假日加班时长", "法定加班时长"]),
        EntityField("迟到早退次数", "int", aliases=["迟到早退、缺卡扣款次数", "迟到次数"]),
        EntityField("病假时长", "amount", aliases=["病假（天）", "病假"]),
        EntityField("事假时长", "amount", aliases=["事假（天）", "事假"]),
        EntityField("产假时长", "amount", aliases=["产假（天）", "产假"]),
        EntityField("迟到早退扣款", "amount", aliases=["迟到早退、缺卡扣款", "迟到扣款"]),
        EntityField("病假扣款", "amount", aliases=["病假扣款"]),
        EntityField("事假扣款", "amount", aliases=["事假扣款"]),
        EntityField("产假扣款", "amount", aliases=["产假扣款"]),
        EntityField("加班费", "amount", aliases=["加班工资", "加班费"]),
        EntityField("班制", "text", aliases=["排班", "班制"]),
        EntityField("出差天数", "amount", aliases=["出差", "出差（天）"]),
    ],
)

SOCIAL_SECURITY = EntityDef(
    id="social_security",
    name="社保公积金",
    primary_key="身份证号",
    foreign_key="身份证号",
    icon="🏥",
    fields=[
        EntityField("姓名", "text", required=True, aliases=["姓名"]),
        EntityField("身份证号", "text", required=True, aliases=["身份证号码", "证件号码", "身份证"]),
        EntityField("缴费公司", "text", aliases=["缴费公司", "单位名称"]),
        EntityField("单位申报基数", "amount", aliases=["申报基数", "社保基数"]),
        EntityField("养老基数", "amount", aliases=["养老保险基数"]),
        EntityField("失业基数", "amount", aliases=["失业保险基数"]),
        EntityField("工伤基数", "amount", aliases=["工伤保险基数"]),
        EntityField("医疗基数", "amount", aliases=["医疗保险基数"]),
        # 个人部分
        EntityField("个人养老", "amount", aliases=["个人养老", "养老保险（个人）", "个人养老8%"]),
        EntityField("个人医疗", "amount", aliases=["个人基本医疗", "医疗保险（个人）", "个人医疗2%+3"]),
        EntityField("个人公积金", "amount", aliases=["个人公积金", "公积金（个人）", "个人住房12%"]),
        EntityField("个人失业", "amount", aliases=["个人失业", "失业保险（个人）", "个人失业0.5%"]),
        EntityField("个人合计", "amount", aliases=["保险公积金合计（个人）", "个人部分合计"]),
        # 公司部分
        EntityField("公司养老", "amount", aliases=["公司养老", "养老保险（公司）", "单位养老16%"]),
        EntityField("公司医疗", "amount", aliases=["公司基本医疗", "医疗保险（公司）", "单位医疗9.8%"]),
        EntityField("公司公积金", "amount", aliases=["公司公积金", "公积金（公司）", "单位住房12%"]),
        EntityField("公司失业", "amount", aliases=["公司失业保险", "失业保险（公司）", "单位失业0.5%"]),
        EntityField("公司生育", "amount", aliases=["公司生育保险", "生育保险（公司）", "单位生育"]),
        EntityField("公司工伤", "amount", aliases=["公司工伤保险", "工伤保险（公司）", "单位工伤0.3%"]),
        EntityField("公司补充医疗", "amount", aliases=["公司补充医疗保险", "商业保险", "补充医疗"]),
        EntityField("公司合计", "amount", aliases=["保险公积金合计（公司）", "公司部分合计"]),
    ],
)

TAX_RECORD = EntityDef(
    id="tax_record",
    name="个税记录",
    primary_key="工号",
    foreign_key="工号",
    icon="🧾",
    fields=[
        EntityField("工号", "text", required=True, aliases=["新工号", "员工工号", "工号"]),
        EntityField("姓名", "text", aliases=["姓名"]),
        EntityField("本次扣税基数", "amount", aliases=["扣税基数"]),
        EntityField("本次应扣税额", "amount", aliases=["应扣税额", "本次扣税"]),
        EntityField("个税调整项", "amount", aliases=["个税调整", "补扣个人所得税"]),
        EntityField("子女及配偶补充医疗代扣", "amount", aliases=["子女配偶补充医疗", "税后扣款"]),
        EntityField("上月累计扣税基数", "amount", aliases=["上月累计"]),
        EntityField("累计本月扣税基数", "amount", aliases=["累计扣税基数", "累计应纳税所得额"]),
        EntityField("累计子女教育", "amount", aliases=["累计子女教育支出扣除"]),
        EntityField("累计继续教育", "amount", aliases=["累计继续教育支出扣除"]),
        EntityField("累计住房贷款利息", "amount", aliases=["累计住房贷款利息支出扣除"]),
        EntityField("累计住房租金", "amount", aliases=["累计住房租金支出扣除"]),
        EntityField("累计赡养老人", "amount", aliases=["累计赡养老人支出扣除"]),
        EntityField("累计婴幼儿照护", "amount", aliases=["累计婴幼儿照护费用"]),
        EntityField("累计应扣缴税额", "amount", aliases=["累计应扣税"]),
        EntityField("上月累计扣缴税额", "amount", aliases=["累计已扣税", "累计已预缴税额"]),
    ],
)

ALL_ENTITIES: list[EntityDef] = [
    EMPLOYEE_PROFILE,
    SALARY_DETAIL,
    ATTENDANCE_RECORD,
    SOCIAL_SECURITY,
    TAX_RECORD,
]

ENTITY_BY_ID: dict[str, EntityDef] = {e.id: e for e in ALL_ENTITIES}


# ============================================================
# 导出映射：实体字段 → 模板「工资核算」Sheet 列位
# ============================================================
# 格式: (列字母, 实体ID, 字段名, 是否公式列)
# 公式列为 True 表示该列在模板中已有公式，导出时不覆盖

EXPORT_MAPPING: list[tuple[str, str, str, bool]] = [
    # === A-B: 员工基本信息 ===
    ("A", "employee_profile", "工号", False),
    ("B", "employee_profile", "姓名", False),
    ("C", "employee_profile", "工作地点", False),
    ("D", "employee_profile", "公司", False),
    ("E", "employee_profile", "部门", False),
    ("F", "employee_profile", "岗位", False),
    ("G", "employee_profile", "人员类别", False),
    ("H", "employee_profile", "合同类别", False),
    ("I", "employee_profile", "本月状态", False),
    ("J", "employee_profile", "入职日期", False),
    ("K", "employee_profile", "离职日期", False),
    ("L", "employee_profile", "是否试用期", False),
    ("M", "employee_profile", "转正日期", False),
    ("N", "employee_profile", "试用期标准", False),
    ("O", "employee_profile", "纳税类别", False),
    # === P-W: 薪资标准 ===
    ("P", "salary_detail", "交通通讯补贴标准", False),
    ("Q", "salary_detail", "过节费", False),
    ("R", "salary_detail", "防暑降温费标准", False),
    ("S", "salary_detail", "天津取暖补贴", False),
    ("T", "salary_detail", "餐补标准", False),
    ("U", "salary_detail", "药师津贴", False),
    ("V", "salary_detail", "店长津贴", False),
    ("W", "salary_detail", "其他补贴", False),
    # === X-AB: 薪资结构与绩效 ===
    ("X", "salary_detail", None, True),  # 月度合计(12月薪) - 公式
    ("Y", "salary_detail", None, True),  # 固浮比 - 公式
    ("Z", "salary_detail", "月基本薪资", False),
    ("AA", "salary_detail", "月绩效标准", False),
    ("AB", "salary_detail", "DTP绩效标准", False),
    # === AC-AH: 奖金与调差 ===
    ("AC", "salary_detail", "考核成绩", False),
    ("AD", "salary_detail", "绩效奖金", False),
    ("AE", "salary_detail", "业绩奖金", False),
    ("AF", "salary_detail", "DTP绩效奖金", False),
    ("AG", "salary_detail", "年终奖", False),
    ("AH", "salary_detail", "其它工资调差", False),
    # === AI-AT: 考勤数据 ===
    ("AI", "attendance_record", "实际出勤天数", False),
    ("AJ", "attendance_record", "普通加班时长", False),
    ("AK", "attendance_record", "法定节日加班时长", False),
    ("AL", "attendance_record", "迟到早退次数", False),
    ("AM", "attendance_record", "病假时长", False),
    ("AN", "attendance_record", "事假时长", False),
    ("AO", "attendance_record", "产假时长", False),
    ("AP", "attendance_record", "迟到早退扣款", False),
    ("AQ", "attendance_record", "病假扣款", False),
    ("AR", "attendance_record", "事假扣款", False),
    ("AS", "attendance_record", "产假扣款", False),
    ("AT", "attendance_record", "加班费", False),
    # === AU-AZ: 本月各类补贴（实发） ===
    ("AU", "salary_detail", "本月交通通讯补贴", False),
    ("AV", "salary_detail", "本月餐补", False),
    ("AW", "salary_detail", "本月防暑降温费", False),
    ("AX", "salary_detail", "本月节日津贴", False),
    ("AY", "salary_detail", "独生子女补贴", False),
    ("AZ", "salary_detail", "出差补贴", False),
    # === BA: 应发工资合计 - 公式列 ===
    ("BA", None, None, True),
    # === BB-BO: 社保公积金 ===
    ("BB", "social_security", "缴费公司", False),
    ("BC", "social_security", "个人合计", False),
    ("BD", "social_security", "个人养老", False),
    ("BE", "social_security", "个人医疗", False),
    ("BF", "social_security", "个人公积金", False),
    ("BG", "social_security", "个人失业", False),
    ("BH", "social_security", "公司合计", False),
    ("BI", "social_security", "公司养老", False),
    ("BJ", "social_security", "公司医疗", False),
    ("BK", "social_security", "公司公积金", False),
    ("BL", "social_security", "公司失业", False),
    ("BM", "social_security", "公司生育", False),
    ("BN", "social_security", "公司工伤", False),
    ("BO", "social_security", "公司补充医疗", False),
    # === BP-BW: 个税与实发 ===
    ("BP", "tax_record", "本次扣税基数", False),
    ("BQ", "tax_record", "本次应扣税额", False),
    ("BR", "tax_record", "个税调整项", False),
    ("BS", "tax_record", "子女及配偶补充医疗代扣", False),
    ("BT", None, None, True),  # 工资实发金额 - 公式
    ("BU", None, None, True),  # 应收、应付工资 - 公式
    ("BV", "salary_detail", "补偿金", False),
    ("BW", None, None, True),  # 本次实发金额 - 公式
    # === BX-BZ: 人员身份信息 ===
    ("BX", "employee_profile", "身份证号", False),
    ("BY", "employee_profile", "银行卡号", False),
    ("BZ", "employee_profile", "开户银行", False),
    # === CA-CN: 累计税务数据 ===
    ("CA", "tax_record", "上月累计扣税基数", False),
    ("CB", "tax_record", "累计本月扣税基数", False),
    ("CC", "tax_record", "累计子女教育", False),
    ("CD", "tax_record", "累计继续教育", False),
    ("CE", "tax_record", "累计住房贷款利息", False),
    ("CF", "tax_record", "累计住房租金", False),
    ("CG", "tax_record", "累计赡养老人", False),
    ("CH", "tax_record", "累计婴幼儿照护", False),
    ("CI", None, None, True),  # 累计专项扣除 - 公式
    ("CJ", None, None, True),  # 累计减除费用 - 公式
    ("CK", None, None, True),  # 减专项扣除后累计扣税基数 - 公式
    ("CL", "tax_record", "累计应扣缴税额", False),
    ("CM", "tax_record", "上月累计扣缴税额", False),
    ("CN", None, None, True),  # 本次应扣税额 - 公式
]


# ============================================================
# 源文件类型 → 实体映射规则
# ============================================================
# 定义每种源文件应该映射到哪些实体，以及从哪个 Sheet 读取

SOURCE_FILE_RULES: dict[str, dict] = {
    "salary": {
        "description": "薪资数据（含考勤、奖金、入离职等）",
        "entities": ["employee_profile", "salary_detail", "attendance_record"],
        "sheet_rules": {
            # Sheet名匹配（模糊匹配）→ (实体ID, 表头所在行号)
            "入离职": ("employee_profile", 1),  # 第1行是表头
            "考勤": ("attendance_record", None),  # 需要嗅探表头
            "奖金": ("salary_detail", None),
        },
    },
    "social_hean": {
        "description": "社保-鹤安长泰",
        "entities": ["social_security"],
        "sheet_rules": {
            "Sheet1": ("social_security", 1),  # 0-based: 第2行是表头（第1行是标题）
        },
    },
    "social_yiyao": {
        "description": "社保-益药科园",
        "entities": ["social_security"],
        "sheet_rules": {
            "Sheet1": ("social_security", 1),
        },
    },
    "tax": {
        "description": "个税数据",
        "entities": ["tax_record"],
        "sheet_rules": {},
    },
}


def detect_file_type(filename: str) -> str:
    """根据文件名嗅探文件类型。"""
    name = filename.lower()
    if "鹤安" in filename or "hean" in name:
        return "social_hean"
    if "益药" in filename or "yiyao" in name:
        return "social_yiyao"
    if "薪资" in filename or "工资" in filename or "salary" in name:
        return "salary"
    if "个税" in filename or "tax" in name:
        return "tax"
    if "社保" in filename or "social" in name:
        return "social_hean"  # 默认归为社保
    return "salary"  # 默认


def get_column_letter(col_idx: int) -> str:
    """将列序号（1-based）转换为 Excel 列字母。"""
    result = ""
    while col_idx > 0:
        col_idx, remainder = divmod(col_idx - 1, 26)
        result = chr(65 + remainder) + result
    return result


def column_letter_to_index(letter: str) -> int:
    """将 Excel 列字母转换为序号（1-based）。"""
    result = 0
    for char in letter:
        result = result * 26 + (ord(char) - 64)
    return result


# 构建导出映射的快速查找表：列字母 → (实体ID, 字段名, 是否公式)
EXPORT_BY_COLUMN: dict[str, tuple[str | None, str | None, bool]] = {
    col: (eid, fname, is_formula) for col, eid, fname, is_formula in EXPORT_MAPPING
}
