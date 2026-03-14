"""
应收凭证处理器
每张销售发票 → 3 行分录：
  借：应收账款_客户名    (1131xxx)  价税合计
  贷：主营业务收入       (5101)     不含税金额
  贷：应交税金_销项税额  (2171001005) 税额
"""
import pandas as pd
from utils.subject import SubjectMatcher
from utils.date_utils import extract_voucher_date, detect_date_col, format_date_ymd, last_day_of_month

# 列名候选（容错不同导出格式）
DATE_COLS = ['开票日期', '日期', '发票日期']
CUSTOMER_COLS = ['购方名称', '客户名称', '购买方名称', '购方单位', '购买方单位名称']
INVOICE_NO_COLS = ['数电发票号码', '发票号码', '号码', '电子发票号码', '发票编号']
AMT_EX_TAX_COLS = ['不含税金额', '金额', '价款', '不含税价款']
TAX_COLS = ['税额', '税金']
AMT_INC_TAX_COLS = ['价税合计', '含税金额', '合计', '价税合计金额']
IS_POSITIVE_COLS = ['是否正数发票', '正负标志', '发票方向', '正负']

FIXED_INCOME_CODE = '5101'
FIXED_TAX_CODE = '2171001005'


def _find_col(df: pd.DataFrame, candidates: list[str]) -> str | None:
    for c in candidates:
        if c in df.columns:
            return c
    return None


def process_ar(
    df_raw: pd.DataFrame,
    matcher: SubjectMatcher,
    voucher_no: int,
) -> tuple[list[dict], list[str]]:
    """
    处理应收发票，返回 (rows, warnings)
    - rows: 凭证行 dict 列表
    - warnings: 文字警告列表（黄色行汇总）
    """
    df = df_raw.copy()
    warnings: list[str] = []
    rows: list[dict] = []

    # 过滤合计行（序号列为"合计"的行）
    seq_col = '序号' if '序号' in df.columns else df.columns[0]
    df = df[df[seq_col].astype(str).str.strip() != '合计'].copy()
    df = df.reset_index(drop=True)

    # 检测必要列
    date_col = _find_col(df, DATE_COLS)
    customer_col = _find_col(df, CUSTOMER_COLS)
    invoice_no_col = _find_col(df, INVOICE_NO_COLS)
    amt_ex_col = _find_col(df, AMT_EX_TAX_COLS)
    tax_col = _find_col(df, TAX_COLS)
    amt_inc_col = _find_col(df, AMT_INC_TAX_COLS)
    is_pos_col = _find_col(df, IS_POSITIVE_COLS)

    missing = [name for name, col in [
        ('开票日期', date_col), ('购方名称', customer_col),
        ('不含税金额', amt_ex_col), ('税额', tax_col), ('价税合计', amt_inc_col)
    ] if col is None]
    if missing:
        raise ValueError(f"应收发票文件缺少以下列（或列名不匹配）：{missing}\n实际列名：{list(df.columns)}")

    # 自动提取凭证日期
    voucher_date_obj = extract_voucher_date(df, date_col)
    if voucher_date_obj is None:
        raise ValueError("无法从开票日期列提取凭证日期，请检查日期格式")
    voucher_date = voucher_date_obj.strftime('%Y-%m-%d')

    # 获取科目显示名
    income_name = matcher.get_display_name(FIXED_INCOME_CODE)
    tax_name = matcher.get_display_name(FIXED_TAX_CODE)

    entry_seq = 1  # 全凭证连续编号（所有发票合为一张凭证）

    for row_idx, row in df.iterrows():
        invoice_date = format_date_ymd(row[date_col])
        customer = str(row[customer_col]).strip()
        invoice_no = str(row[invoice_no_col]).strip() if invoice_no_col else ''

        try:
            amt_ex = round(float(row[amt_ex_col]), 2)
            tax = round(float(row[tax_col]), 2)
            amt_inc = round(float(row[amt_inc_col]), 2)
        except (ValueError, TypeError):
            warnings.append(f"第 {row_idx + 2} 行金额解析失败，已跳过：{row.to_dict()}")
            continue

        # 红字发票判断（金额已为负，直接使用）
        # is_positive 为 False/"否" 时为红字，金额本身应已为负
        # 无需额外处理，按实际值填入

        # 摘要
        summary = f"{invoice_date}，{customer}，{invoice_no}，"

        # 应收账款科目匹配
        ar_code, ar_name, is_default = matcher.get_ar_account(customer)
        is_yellow = is_default and customer not in ['', 'nan']
        if is_yellow:
            warnings.append(f"行 {row_idx + 2}：客户「{customer}」未在 1131 子科目中匹配到，已归入 {ar_code}")

        def make_row(debit, credit, code, name, amount) -> dict:
            return {
                '日期': voucher_date,
                '凭证字': '记',
                '凭证号': voucher_no,
                '附件数': 0,
                '分录序号': debit or credit,  # placeholder, reassigned below
                '摘要': summary,
                '科目代码': code,
                '科目名称': name,
                '借方金额': amount if debit else '',
                '贷方金额': amount if credit else '',
                '客户': '', '供应商': '', '职员': '', '项目': '', '部门': '', '存货': '', '自定义辅助核算类别': '', '自定义辅助核算编码': '', '自定义辅助核算类别1': '', '自定义辅助核算编码1': '', '数量': '', '单价': '',
                '原币金额': amount,
                '币别': 'RMB',
                '汇率': 1,
                '_yellow': is_yellow,
            }

        # 行1：借 应收账款
        r1 = {
            '日期': voucher_date, '凭证字': '记', '凭证号': voucher_no, '附件数': 0,
            '分录序号': entry_seq, '摘要': summary,
            '科目代码': ar_code, '科目名称': ar_name,
            '借方金额': amt_inc, '贷方金额': '',
            '客户': '', '供应商': '', '职员': '', '项目': '', '部门': '', '存货': '', '自定义辅助核算类别': '', '自定义辅助核算编码': '', '自定义辅助核算类别1': '', '自定义辅助核算编码1': '', '数量': '', '单价': '',
            '原币金额': amt_inc, '币别': 'RMB', '汇率': 1,
            '_yellow': is_yellow,
        }
        entry_seq += 1

        # 行2：贷 主营业务收入
        r2 = {
            '日期': voucher_date, '凭证字': '记', '凭证号': voucher_no, '附件数': 0,
            '分录序号': entry_seq, '摘要': summary,
            '科目代码': FIXED_INCOME_CODE, '科目名称': income_name,
            '借方金额': '', '贷方金额': amt_ex,
            '客户': '', '供应商': '', '职员': '', '项目': '', '部门': '', '存货': '', '自定义辅助核算类别': '', '自定义辅助核算编码': '', '自定义辅助核算类别1': '', '自定义辅助核算编码1': '', '数量': '', '单价': '',
            '原币金额': amt_ex, '币别': 'RMB', '汇率': 1,
            '_yellow': False,
        }
        entry_seq += 1

        # 行3：贷 应交税金_销项税额
        r3 = {
            '日期': voucher_date, '凭证字': '记', '凭证号': voucher_no, '附件数': 0,
            '分录序号': entry_seq, '摘要': summary,
            '科目代码': FIXED_TAX_CODE, '科目名称': tax_name,
            '借方金额': '', '贷方金额': tax,
            '客户': '', '供应商': '', '职员': '', '项目': '', '部门': '', '存货': '', '自定义辅助核算类别': '', '自定义辅助核算编码': '', '自定义辅助核算类别1': '', '自定义辅助核算编码1': '', '数量': '', '单价': '',
            '原币金额': tax, '币别': 'RMB', '汇率': 1,
            '_yellow': False,
        }
        entry_seq += 1

        rows.extend([r1, r2, r3])

    return rows, warnings
