"""
银行活期交易凭证处理器
支持多种银行导出格式自动识别：
  - 农商行格式：表头在第3-4行，收入金额/支出金额列，对方户名列
  - 工商银行等通用格式：表头在第0行，借方发生额/贷方发生额列，借/贷方向列，对方单位名称列
  - 工商银行单列格式：只有「发生额」一列，搭配「借贷标志」方向列
"""
import pandas as pd
from utils.subject import SubjectMatcher
from utils.date_utils import last_day_of_month

EXPENSE_NOTE_MAP = [
    (['备用金', '转账'], '1001'),
    (['工资'], '2151'),
    (['货款', '往来款', '货款往来'], '__AP__'),
    (['加油费'], '1133001'),
    (['管理费'], '5502013'),
    (['保险费'], '__INSURANCE__'),
]
EXPENSE_SUMMARY_MAP = [
    (['自助交易费', '短信服务费', '手续费', '服务费'], '5503001'),
    (['利息'], '5503002'),
    (['归还本息', '还款'], '2101'),
]
FILTER_ACCOUNT_SUFFIX = '1511'
_EMPTY_AUX = {
    '客户': '', '供应商': '', '职员': '', '项目': '', '部门': '', '存货': '',
    '自定义辅助核算类别': '', '自定义辅助核算编码': '',
    '自定义辅助核算类别1': '', '自定义辅助核算编码1': '',
    '数量': '', '单价': '',
}

# 各格式的列名候选，按优先级排列
_COL_TIME       = ['交易时间', '时间', '日期', '交易日期', '记账日期', '账务日期', '交易日', '入账日期']
_COL_INCOME     = ['收入金额', '贷方发生额', '贷方金额', '收入', '贷方', '收入金额(元)', '收入金额（元）']
_COL_EXPENSE    = ['支出金额', '借方发生额', '借方金额', '支出', '借方', '支出金额(元)', '支出金额（元）']
_COL_AMOUNT     = ['发生额', '金额', '交易金额', '交易金额(元)', '交易金额（元）',
                   '交易金额(人民币)', '记账金额', '业务金额']
_COL_CP_NAME    = ['对方户名', '对方单位名称', '对方名称', '对方姓名', '户名', '收款人名称',
                   '对方账户名称', '对方账号名称', '对方客户名称']
_COL_CP_ACCT    = ['对方账号', '对方账户', '对方卡号']
_COL_SUMMARY    = ['摘要', '交易摘要', '用途', '业务摘要', '附言内容', '用途/摘要', '交易说明']
_COL_NOTE       = ['备注', '附言', '用途说明', '备注说明']
_COL_DEBIT_FLAG = ['借/贷', '借贷标志', '收支类型', '方向', '借贷方向', '交易类型',
                   '业务类型', '借贷', '标志', '账务类型', '收/支', '收支标志', '资金流向', '交易方向']
_COL_OWN_ACCT   = ['本方账号', '本行账号', '账号']

# BUG FIX: 借/贷 方向标志的各种写法 → 归一化
_FLAG_DEBIT_SET  = {'借', '借方', 'D', 'DB', 'DR', 'DR.', 'DEBIT', '出', '支出', '支', '出账', '出款', '-'}
_FLAG_CREDIT_SET = {'贷', '贷方', 'C', 'CR', 'CR.', 'CREDIT', '入', '收入', '收', '进账', '入账', '+'}


def _find_col(df, candidates):
    for c in candidates:
        if c in df.columns:
            return c
    # 模糊匹配：列名包含候选关键词
    for c in candidates:
        for col in df.columns:
            if c in str(col):
                return col
    return None


def _normalize_flag(raw: str) -> str:
    """将各种借贷标志归一化为 '借' 或 '贷'，无法识别返回空串"""
    s = raw.strip().upper()
    if s in {v.upper() for v in _FLAG_DEBIT_SET}:
        return '借'
    if s in {v.upper() for v in _FLAG_CREDIT_SET}:
        return '贷'
    return ''


def read_bank_file(raw_bytes):
    """
    智能识别银行流水格式，返回 (df_data, account_no, company_name)
    """
    from io import BytesIO

    engine = 'calamine'
    if raw_bytes[:4] == b'\xd0\xcf\x11\xe0':
        try:
            import xlrd  # noqa
            engine = 'xlrd'
        except ImportError:
            pass

    # 读前50行扫描表头位置（不同银行会有更多说明行）
    df_head = pd.read_excel(BytesIO(raw_bytes), header=None, nrows=50, engine=engine)

    # 表头检测：同时支持通用/工行/农商行等格式
    header_row = None
    for i in range(len(df_head)):
        row_vals = [str(v).strip() for v in df_head.iloc[i].values]
        has_date   = any(any(k in v for k in _COL_TIME) for v in row_vals)
        has_amount = any(any(k in v for k in (_COL_AMOUNT + _COL_INCOME + _COL_EXPENSE)) for v in row_vals)
        has_flag   = any(any(k in v for k in _COL_DEBIT_FLAG) for v in row_vals)
        has_cp     = any(any(k in v for k in _COL_CP_NAME) for v in row_vals)
        non_empty  = sum(1 for v in row_vals if v and v.lower() not in ('nan', 'none'))
        # 必须同时具备“日期列”和“金额列”才能判定为表头，避免误把说明行当表头
        if (has_date and has_amount) and non_empty >= 3:
            header_row = i
            break

    if header_row is None:
        raise ValueError("找不到含「交易时间/日期/发生额」的表头行，请确认是银行活期交易明细文件。")

    df_data = pd.read_excel(BytesIO(raw_bytes), header=header_row, dtype=str, engine=engine)

    # 提取本方账号
    account_no = ''
    company_name = ''

    own_acct_col = _find_col(df_data, _COL_OWN_ACCT)
    # 避免误把“对方账号”当作本方账号
    if own_acct_col and '对方' in str(own_acct_col):
        own_acct_col = None
    if own_acct_col:
        val = str(df_data[own_acct_col].iloc[0]).strip().rstrip('\\t').strip()
        if val.lower() not in ('nan', ''):
            account_no = val
    elif header_row > 0:
        # 农商行格式：账号在表头前的固定位置
        try:
            val = str(df_head.iloc[1, 1]).strip()
            if val.lower() not in ('nan', ''):
                account_no = val
            val2 = str(df_head.iloc[1, 4]).strip()
            if val2.lower() not in ('nan', ''):
                company_name = val2
        except Exception:
            pass
    if not account_no:
        # 通用兜底：从表头前几行文本里提取“账号/账户/卡号”
        import re
        for i in range(min(6, len(df_head))):
            for v in df_head.iloc[i].values:
                s = str(v)
                m = re.search(r'(账号|账户|卡号)[:：\\s]*([0-9]{6,})', s)
                if m:
                    account_no = m.group(2)
                    break
            if account_no:
                break

    # 清理列名中的 \t
    df_data.columns = [str(c).strip().rstrip('\\t').strip() for c in df_data.columns]

    return df_data, account_no, company_name


def get_bank_account_no(raw_bytes):
    try:
        _, acct, _ = read_bank_file(raw_bytes)
        return acct or '未知账号'
    except Exception:
        return '未知账号'


def process_bank(df_raw, matcher, company_name, income_voucher_start,
                 expense_voucher_start, bank_account_no='', bank_subject_code='',
                 user_rules=None, company_accounts=None):
    """
    返回 (result, warnings, pending_items)
    """
    df = df_raw.copy()
    # 清理列名
    df.columns = [str(c).strip().rstrip('\\t').strip() for c in df.columns]

    warnings, pending_items = [], []
    if user_rules is None:
        user_rules = []
    if company_accounts is None:
        company_accounts = []

    bank_code    = bank_subject_code.strip() if bank_subject_code.strip() else '1002'
    bank_display = matcher.get_display_name(bank_code) or '银行存款'

    time_col       = _find_col(df, _COL_TIME)
    cp_name_col    = _find_col(df, _COL_CP_NAME)
    cp_acct_col    = _find_col(df, _COL_CP_ACCT)
    summary_col    = _find_col(df, _COL_SUMMARY)
    note_col       = _find_col(df, _COL_NOTE)
    debit_flag_col = _find_col(df, _COL_DEBIT_FLAG)

    # ── 金额列检测 ───────────────────────────────────────────
    # BUG FIX: 工行部分版本只有单列「发生额」，搭配借/贷方向列
    single_amount_col = None
    amount_col_any = _find_col(df, _COL_AMOUNT)
    if debit_flag_col:
        income_col  = _find_col(df, ['贷方发生额', '贷方金额'])
        expense_col = _find_col(df, ['借方发生额', '借方金额'])
        if not income_col and not expense_col:
            # 工行单列格式：金额在同一列，方向由标志列决定
            single_amount_col = _find_col(df, _COL_AMOUNT)
            income_col  = single_amount_col
            expense_col = single_amount_col
    else:
        # 农商行格式：有独立收入/支出列
        income_col  = _find_col(df, ['收入金额', '收入'])
        expense_col = _find_col(df, ['支出金额', '支出'])

    if not income_col and not expense_col and not amount_col_any:
        raise ValueError("找不到金额列（收入/支出/发生额），请确认银行流水导出格式。")

    if not time_col:
        raise ValueError("找不到「交易时间」列，请确认是银行活期交易明细文件。")

    # 过滤内部转账行（对方账号以1511结尾）
    if cp_acct_col:
        df = df[~df[cp_acct_col].astype(str).str.strip().str.rstrip('\\t').str.endswith(FILTER_ACCOUNT_SUFFIX)].copy()
    # 内部转账规则（多账号）：只在“对方户名=公司名 且 对方账号=公司其他账号”时生效
    acct_list = [a.account_no.strip() for a in company_accounts if getattr(a, "account_no", "").strip()]
    cur_acct = (bank_account_no or '').strip()
    try:
        cur_idx = acct_list.index(cur_acct)
    except ValueError:
        cur_idx = 0

    if cp_name_col and cp_acct_col and company_name and acct_list:
        def _is_internal(row):
            name = str(row[cp_name_col]).strip().rstrip('\\t')
            acct = str(row[cp_acct_col]).strip().rstrip('\\t')
            if not name or not acct:
                return False
            if name != company_name.strip():
                return False
            if acct not in acct_list:
                return False
            # 第一个账号保留全部内部转账；其余账号过滤掉与“更早账号”的转账
            if cur_idx == 0:
                return False
            try:
                cp_idx = acct_list.index(acct)
            except ValueError:
                return False
            return cp_idx < cur_idx

        df = df[~df.apply(_is_internal, axis=1)].copy()
    else:
        # 仅按户名过滤（单账号或缺少账号列时的兜底）
        if cp_name_col and company_name:
            df = df[df[cp_name_col].astype(str).str.strip().str.rstrip('\\t') != company_name.strip()].copy()

    df['_date']  = pd.to_datetime(df[time_col], errors='coerce')
    df = df.dropna(subset=['_date'])
    df['_month'] = df['_date'].dt.strftime('%Y-%m')

    result = {}
    for month, mdf in df.groupby('_month'):
        last_day     = last_day_of_month(int(month[:4]), int(month[5:7]))
        voucher_date = last_day.strftime('%Y-%m-%d')
        income_rows, expense_rows = [], []
        inc_seq = exp_seq = 1

        for _, row in mdf.iterrows():
            trade_date       = row['_date'].strftime('%Y-%m-%d')
            counterpart      = _s(row, cp_name_col)
            counterpart_acct = _s(row, cp_acct_col) if cp_acct_col else ''
            raw_summary      = _s(row, summary_col)
            note             = _s(row, note_col)
            memo             = f"{trade_date}，{counterpart}" if counterpart else trade_date

            # ── 判断收入/支出金额 ──────────────────────────────
            if debit_flag_col:
                raw_flag = _s(row, debit_flag_col)
                # BUG FIX: 归一化借贷方向标志，兼容"借方"/"贷方"/"D"/"C"等写法
                flag = _normalize_flag(raw_flag)

                if single_amount_col:
                    # 工行单列模式
                    raw_amount = _amt(row.get(single_amount_col)) if single_amount_col else 0.0
                    if flag == '借':
                        inc_val, exp_val = 0.0, raw_amount
                    elif flag == '贷':
                        inc_val, exp_val = raw_amount, 0.0
                    else:
                        # 无法识别方向：用金额符号兜底
                        raw_raw = _amt_raw(row.get(single_amount_col)) if single_amount_col else 0.0
                        if raw_raw < 0:
                            inc_val, exp_val = 0.0, abs(raw_raw)
                        else:
                            inc_val, exp_val = abs(raw_raw), 0.0
                else:
                    raw_income  = _amt(row.get(income_col))  if income_col  else 0.0
                    raw_expense = _amt(row.get(expense_col)) if expense_col else 0.0
                    if flag == '借':
                        inc_val, exp_val = 0.0, raw_expense or raw_income
                    elif flag == '贷':
                        inc_val, exp_val = raw_income or raw_expense, 0.0
                    else:
                        # 标志未识别：直接用原始列值（农商行风格兜底）
                        inc_val  = raw_income
                        exp_val  = raw_expense
            else:
                inc_val = _amt(row.get(income_col))  if income_col  else 0.0
                exp_val = _amt(row.get(expense_col)) if expense_col else 0.0
                if not income_col and not expense_col and amount_col_any:
                    # 只有发生额但没有方向列：用正负号兜底
                    amt_raw = _amt_raw(row.get(amount_col_any))
                    if amt_raw < 0:
                        inc_val, exp_val = 0.0, abs(amt_raw)
                    else:
                        inc_val, exp_val = abs(amt_raw), 0.0

            if inc_val > 0:
                if '利息' in raw_summary or '利息' in note:
                    cr_code  = '5503002'
                    cr_name  = matcher.get_display_name(cr_code)
                    income_rows += [
                        _r(voucher_date, income_voucher_start, inc_seq,     memo, bank_code, bank_display, inc_val, '',      False),
                        _r(voucher_date, income_voucher_start, inc_seq + 1, memo, cr_code,   cr_name,      '',      inc_val, False),
                    ]
                    inc_seq += 2
                else:
                    cr_code, _, is_def = matcher.get_ar_account(counterpart)
                    if is_def:
                        pending_items.append({
                            'trade_date':        trade_date,
                            'month':             str(month),
                            'counterpart':       counterpart,
                            'counterpart_acct':  counterpart_acct,
                            'summary':           raw_summary,
                            'note':              note,
                            'amount':            inc_val,
                            'memo':              memo,
                            'voucher_date':      voucher_date,
                            'bank_code':         bank_code,
                            'bank_name':         bank_display,
                            'direction':         'income',
                        })
                    else:
                        cr_name = matcher.get_display_name(cr_code)
                        income_rows += [
                            _r(voucher_date, income_voucher_start, inc_seq,     memo, bank_code, bank_display, inc_val, '',      False),
                            _r(voucher_date, income_voucher_start, inc_seq + 1, memo, cr_code,   cr_name,      '',      inc_val, False),
                        ]
                        inc_seq += 2

            if exp_val > 0:
                code = _resolve(note, raw_summary, counterpart, matcher, user_rules)
                if code is None:
                    pending_items.append({
                        'trade_date':       trade_date,
                        'month':            str(month),
                        'counterpart':      counterpart,
                        'counterpart_acct': counterpart_acct,
                        'summary':          raw_summary,
                        'note':             note,
                        'amount':           exp_val,
                        'memo':             memo,
                        'voucher_date':     voucher_date,
                        'bank_code':        bank_code,
                        'bank_name':        bank_display,
                        'direction':        'expense',
                    })
                else:
                    dr_name = matcher.get_display_name(code)
                    expense_rows += [
                        _r(voucher_date, expense_voucher_start, exp_seq,     memo, code,      dr_name,      exp_val, '',      False),
                        _r(voucher_date, expense_voucher_start, exp_seq + 1, memo, bank_code, bank_display, '',      exp_val, False),
                    ]
                    exp_seq += 2

        result[str(month)] = (income_rows, expense_rows)

    return result, warnings, pending_items


def generate_pending_rows(pending_items, assignments, matcher,
                          income_voucher_start=1, expense_voucher_start=1):
    from collections import defaultdict
    extra = defaultdict(list)
    inc_seq = defaultdict(lambda: 1)
    exp_seq = defaultdict(lambda: 1)

    for i, item in enumerate(pending_items):
        code = (assignments.get(i) or '').strip()
        if not code:
            continue
        month     = item['month']
        name      = matcher.get_display_name(code)
        val       = item['amount']
        direction = item.get('direction', 'expense')

        if direction == 'income':
            seq  = inc_seq[month]
            v_no = income_voucher_start
            extra[month] += [
                _r(item['voucher_date'], v_no, seq,     item['memo'], item['bank_code'], item['bank_name'], val, '',  False),
                _r(item['voucher_date'], v_no, seq + 1, item['memo'], code,              name,              '',  val, False),
            ]
            inc_seq[month] += 2
        else:
            seq  = exp_seq[month]
            v_no = expense_voucher_start
            extra[month] += [
                _r(item['voucher_date'], v_no, seq,     item['memo'], code,              name,              val, '',  False),
                _r(item['voucher_date'], v_no, seq + 1, item['memo'], item['bank_code'], item['bank_name'], '',  val, False),
            ]
            exp_seq[month] += 2

    return dict(extra)


def _r(date, v_no, seq, memo, code, name, debit, credit, yellow):
    return {
        '日期': date, '凭证字': '记', '凭证号': v_no, '附件数': 0,
        '分录序号': seq, '摘要': memo,
        '科目代码': code, '科目名称': name,
        '借方金额': debit, '贷方金额': credit,
        **_EMPTY_AUX,
        '原币金额': debit if debit else credit,
        '币别': 'RMB', '汇率': 1, '_yellow': yellow,
    }


def _s(row, col):
    if not col or col not in row.index: return ''
    v = str(row[col]).strip().rstrip('\\t').strip()
    return '' if v.lower() == 'nan' else v


def _amt(val):
    try:
        v = float(str(val).replace(',', '').replace(' ', '').strip())
        return round(abs(v), 2) if v != 0 else 0.0
    except (ValueError, TypeError):
        return 0.0


def _amt_raw(val):
    try:
        return float(str(val).replace(',', '').replace(' ', '').strip())
    except (ValueError, TypeError):
        return 0.0


def _resolve(note, summary, counterpart, matcher, user_rules):
    from company_manager import apply_rules
    code = apply_rules(user_rules, note, summary, counterpart)
    if code: return code

    # 检查备注和摘要，优先备注
    for text in [note, summary]:
        if not text:
            continue
        for keywords, mapped_code in EXPENSE_NOTE_MAP:
            for kw in keywords:
                if kw in text:
                    if mapped_code == '__AP__':
                        ap_code, _, not_found = matcher.get_ap_account(counterpart)
                        return None if not_found else ap_code
                    elif mapped_code == '__INSURANCE__':
                        ins_code, _ = matcher.find_sub_account('2181', counterpart)
                        return ins_code or '2181002'
                    else:
                        return mapped_code

    # 摘要二级匹配
    for text in [note, summary]:
        if not text:
            continue
        for keywords, mapped_code in EXPENSE_SUMMARY_MAP:
            for kw in keywords:
                if kw in text:
                    return mapped_code

    return None
