"""
应付凭证处理器
每张进项发票 → 3 行分录：
  借：费用/资产科目       (根据用途)    不含税金额
  借：应交税金_进项税额   (2171001001)  税额
  贷：应付账款/现金/银行  (根据付款方式) 价税合计
"""
import pandas as pd
from utils.subject import SubjectMatcher
from utils.date_utils import extract_voucher_date, format_date_ymd

DATE_COLS = ['开票日期', '日期', '发票日期']
SUPPLIER_COLS = ['销方名称', '供应商名称', '销售方名称', '销方单位', '销售方单位名称']
INVOICE_NO_COLS = ['数电发票号码', '发票号码', '号码', '电子发票号码', '发票编号']
AMT_EX_TAX_COLS = ['不含税金额', '金额', '价款', '不含税价款']
TAX_COLS = ['税额', '税金']
AMT_INC_TAX_COLS = ['价税合计', '含税金额', '合计', '价税合计金额']
PURPOSE_COLS = ['用途', '发票用途', '备注', '摘要', '用途说明']

FIXED_INPUT_TAX_CODE = '2171001001'

# 费用科目关键词匹配表（按 PURPOSE 列或供应商名称匹配）
EXPENSE_KEYWORD_MAP = [
    (['纸板', '原材料', '瓦楞', '纸箱', '纸品'], '1211001'),
    (['维修', '修理', '设备', '印刷机'], '4105007'),
    (['电费', '供电', '电力'], '4105002'),
    (['汽车', '油费', '加油', '车辆', '运输'], '5501003'),
    (['保险', '财产险'], '5502004'),
    (['办公', '文具', '耗材'], '5502001'),
]

# 付款方式 → 贷方科目
PAYMENT_CREDIT_MAP = {
    'pending': ('2121', '应付账款'),       # 挂账，需匹配子科目
    'cash': ('1001', '现金'),
    'bank_icbc': ('1002001', ''),          # 工商银行
    'bank_abc': ('1002002', ''),           # 农商行
    'prepaid': ('1133001', '其他应收款_预存款'),
    'insurance': ('2181002', ''),          # 其他应付款_保险公司
}


def _find_col(df: pd.DataFrame, candidates: list[str]) -> str | None:
    for c in candidates:
        if c in df.columns:
            return c
    return None


def _match_expense_code(text: str) -> str | None:
    if not text or text == 'nan':
        return None
    for keywords, code in EXPENSE_KEYWORD_MAP:
        for kw in keywords:
            if kw in text:
                return code
    return None


def process_ap(
    df_raw: pd.DataFrame,
    matcher: SubjectMatcher,
    voucher_no: int,
    payment_method: str = 'pending',  # 'pending'|'cash'|'bank_icbc'|'bank_abc'|'prepaid'|'insurance'
) -> tuple[list[dict], list[str]]:
    """
    处理应付发票，返回 (rows, warnings)
    """
    df = df_raw.copy()
    warnings: list[str] = []
    rows: list[dict] = []

    # 过滤合计行
    seq_col = '序号' if '序号' in df.columns else df.columns[0]
    df = df[df[seq_col].astype(str).str.strip() != '合计'].copy()
    df = df.reset_index(drop=True)

    # 检测列
    date_col = _find_col(df, DATE_COLS)
    supplier_col = _find_col(df, SUPPLIER_COLS)
    invoice_no_col = _find_col(df, INVOICE_NO_COLS)
    amt_ex_col = _find_col(df, AMT_EX_TAX_COLS)
    tax_col = _find_col(df, TAX_COLS)
    amt_inc_col = _find_col(df, AMT_INC_TAX_COLS)
    purpose_col = _find_col(df, PURPOSE_COLS)

    missing = [name for name, col in [
        ('开票日期', date_col), ('销方名称', supplier_col),
        ('不含税金额', amt_ex_col), ('税额', tax_col), ('价税合计', amt_inc_col)
    ] if col is None]
    if missing:
        raise ValueError(f"应付发票文件缺少以下列：{missing}\n实际列名：{list(df.columns)}")

    # 凭证日期
    voucher_date_obj = extract_voucher_date(df, date_col)
    if voucher_date_obj is None:
        raise ValueError("无法从开票日期提取凭证日期")
    voucher_date = voucher_date_obj.strftime('%Y-%m-%d')

    # 进项税科目
    input_tax_name = matcher.get_display_name(FIXED_INPUT_TAX_CODE)

    entry_seq = 1

    for row_idx, row in df.iterrows():
        invoice_date = format_date_ymd(row[date_col])
        supplier = str(row[supplier_col]).strip()
        invoice_no = str(row[invoice_no_col]).strip() if invoice_no_col else ''
        purpose = str(row[purpose_col]).strip() if purpose_col else ''

        try:
            amt_ex = round(float(row[amt_ex_col]), 2)
            tax = round(float(row[tax_col]), 2)
            amt_inc = round(float(row[amt_inc_col]), 2)
        except (ValueError, TypeError):
            warnings.append(f"第 {row_idx + 2} 行金额解析失败，已跳过")
            continue

        summary = f"{invoice_date}，{supplier}，{invoice_no}，"

        # ---- 借方1：费用/资产科目 ----
        expense_code = _match_expense_code(purpose) or _match_expense_code(supplier)
        if expense_code:
            expense_name = matcher.get_display_name(expense_code)
            expense_yellow = False
        else:
            expense_code = '待确认'
            expense_name = '待确认'
            expense_yellow = True
            warnings.append(f"行 {row_idx + 2}：供应商「{supplier}」（用途：{purpose}）无法匹配费用科目，请手动确认")

        # ---- 贷方：付款方式科目 ----
        credit_code, credit_name, credit_yellow = _resolve_credit(
            payment_method, supplier, matcher, warnings, row_idx
        )

        row_yellow = expense_yellow or credit_yellow

        # 行1：借 费用
        rows.append({
            '日期': voucher_date, '凭证字': '记', '凭证号': voucher_no, '附件数': 0,
            '分录序号': entry_seq, '摘要': summary,
            '科目代码': expense_code, '科目名称': expense_name,
            '借方金额': amt_ex, '贷方金额': '',
            '客户': '', '供应商': '', '职员': '', '项目': '', '部门': '', '存货': '', '自定义辅助核算类别': '', '自定义辅助核算编码': '', '自定义辅助核算类别1': '', '自定义辅助核算编码1': '', '数量': '', '单价': '',
            '原币金额': amt_ex, '币别': 'RMB', '汇率': 1,
            '_yellow': row_yellow,
        })
        entry_seq += 1

        # 行2：借 进项税
        rows.append({
            '日期': voucher_date, '凭证字': '记', '凭证号': voucher_no, '附件数': 0,
            '分录序号': entry_seq, '摘要': summary,
            '科目代码': FIXED_INPUT_TAX_CODE, '科目名称': input_tax_name,
            '借方金额': tax, '贷方金额': '',
            '客户': '', '供应商': '', '职员': '', '项目': '', '部门': '', '存货': '', '自定义辅助核算类别': '', '自定义辅助核算编码': '', '自定义辅助核算类别1': '', '自定义辅助核算编码1': '', '数量': '', '单价': '',
            '原币金额': tax, '币别': 'RMB', '汇率': 1,
            '_yellow': False,
        })
        entry_seq += 1

        # 行3：贷 付款
        rows.append({
            '日期': voucher_date, '凭证字': '记', '凭证号': voucher_no, '附件数': 0,
            '分录序号': entry_seq, '摘要': summary,
            '科目代码': credit_code, '科目名称': credit_name,
            '借方金额': '', '贷方金额': amt_inc,
            '客户': '', '供应商': '', '职员': '', '项目': '', '部门': '', '存货': '', '自定义辅助核算类别': '', '自定义辅助核算编码': '', '自定义辅助核算类别1': '', '自定义辅助核算编码1': '', '数量': '', '单价': '',
            '原币金额': amt_inc, '币别': 'RMB', '汇率': 1,
            '_yellow': credit_yellow,
        })
        entry_seq += 1

    return rows, warnings


def _resolve_credit(
    payment_method: str,
    supplier: str,
    matcher: SubjectMatcher,
    warnings: list,
    row_idx: int
) -> tuple[str, str, bool]:
    """解析贷方科目，返回 (code, name, is_yellow)"""
    if payment_method == 'pending':
        code, name, not_found = matcher.get_ap_account(supplier)
        if not_found:
            warnings.append(f"行 {row_idx + 2}：供应商「{supplier}」未在 2121 子科目匹配，请手动确认")
            return '待确认', '待确认', True
        return code, name, False

    elif payment_method == 'cash':
        code = '1001'
        return code, matcher.get_display_name(code), False

    elif payment_method == 'bank_icbc':
        code = '1002001'
        return code, matcher.get_display_name(code), False

    elif payment_method == 'bank_abc':
        code = '1002002'
        return code, matcher.get_display_name(code), False

    elif payment_method == 'prepaid':
        code = '1133001'
        return code, matcher.get_display_name(code), False

    elif payment_method == 'insurance':
        code = '2181002'
        # 尝试匹配子科目
        sub_code, sub_name = matcher.find_sub_account('2181', supplier)
        if sub_code:
            return sub_code, sub_name, False
        return code, matcher.get_display_name(code), False

    else:
        return '待确认', '待确认', True
