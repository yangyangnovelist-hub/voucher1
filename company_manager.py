"""
公司档案管理器
- 每家公司保存：名称、银行账号列表、科目列表（Excel 字节）
- 数据存储在本地 data/ 目录下（JSON 索引 + 二进制科目文件）
"""
import json
import os
import shutil
from dataclasses import dataclass, field, asdict
from typing import Optional
from io import BytesIO
import pandas as pd
from utils.subject import SubjectMatcher

def _excel_engine(path: str) -> str:
    """Detect engine by file signature to handle mislabeled .xls/.xlsx."""
    try:
        with open(path, "rb") as f:
            sig = f.read(4)
        if sig == b'\xd0\xcf\x11\xe0':
            return "xlrd"
    except Exception:
        pass
    return "calamine"

DATA_DIR = os.path.join(os.path.dirname(__file__), 'data')
INDEX_FILE = os.path.join(DATA_DIR, 'companies.json')


@dataclass
class BankAccount:
    account_no: str          # 账号（必填，用于流水筛选）
    bank_name: str = ''      # 银行名称（选填，仅用于文件命名显示）
    subject_code: str = ''   # 对应银行存款科目代码（如 1002002）


@dataclass
class CompanyProfile:
    name: str                               # 公司名称（户名，用于过滤内部转账）
    short_name: str                         # 显示简称
    bank_accounts: list[BankAccount] = field(default_factory=list)
    subject_file: Optional[str] = None     # 科目文件路径（相对 DATA_DIR）


def _ensure_data_dir():
    os.makedirs(DATA_DIR, exist_ok=True)


def load_index() -> dict[str, dict]:
    """返回 {short_name: profile_dict}"""
    _ensure_data_dir()
    if not os.path.exists(INDEX_FILE):
        return {}
    with open(INDEX_FILE, 'r', encoding='utf-8') as f:
        return json.load(f)


def save_index(index: dict):
    _ensure_data_dir()
    with open(INDEX_FILE, 'w', encoding='utf-8') as f:
        json.dump(index, f, ensure_ascii=False, indent=2)


def list_companies() -> list[str]:
    """返回所有公司简称列表"""
    return list(load_index().keys())


def _load_bank_account(d: dict) -> BankAccount:
    """
    兼容旧版本格式加载（旧版字段顺序：account_no, subject_code, bank_name）。
    若字段 bank_name 看起来像科目代码（纯数字/以1002开头），
    且 subject_code 不像科目代码，则自动对调。
    """
    acct = d.get('account_no', '')
    bank = d.get('bank_name', '')
    code = d.get('subject_code', '')

    # 旧版本可能有 is_default 字段，忽略
    # 检测字段对调：bank_name 存的是科目代码，subject_code 存的是银行名
    def looks_like_code(s):
        return bool(s) and (s.startswith('1002') or (s.isdigit() and len(s) >= 4))

    def looks_like_name(s):
        return bool(s) and not s.isdigit()

    if looks_like_code(bank) and (not code or looks_like_name(code)):
        bank, code = code, bank  # 自动对调

    return BankAccount(account_no=acct, bank_name=bank, subject_code=code)


def get_company(short_name: str) -> Optional[CompanyProfile]:
    index = load_index()
    if short_name not in index:
        return None
    d = index[short_name]
    accounts = [_load_bank_account(a) for a in d.get('bank_accounts', [])]
    return CompanyProfile(
        name=d['name'],
        short_name=d['short_name'],
        bank_accounts=accounts,
        subject_file=d.get('subject_file'),
    )


def save_company(profile: CompanyProfile, subject_bytes: Optional[bytes] = None):
    """保存公司档案。若提供 subject_bytes，写入科目文件。"""
    _ensure_data_dir()
    index = load_index()

    # 保存科目文件
    if subject_bytes is not None:
        fname = f"subjects_{profile.short_name}.xlsx"
        fpath = os.path.join(DATA_DIR, fname)
        with open(fpath, 'wb') as f:
            f.write(subject_bytes)
        profile.subject_file = fname

    index[profile.short_name] = {
        'name': profile.name,
        'short_name': profile.short_name,
        'bank_accounts': [vars(a) for a in profile.bank_accounts],
        'subject_file': profile.subject_file,
    }
    save_index(index)


def delete_company(short_name: str):
    index = load_index()
    if short_name in index:
        subject_file = index[short_name].get('subject_file')
        if subject_file:
            fpath = os.path.join(DATA_DIR, subject_file)
            if os.path.exists(fpath):
                os.remove(fpath)
        del index[short_name]
        save_index(index)


def load_matcher(short_name: str) -> Optional[SubjectMatcher]:
    """加载某公司的科目匹配器"""
    profile = get_company(short_name)
    if not profile or not profile.subject_file:
        return None
    fpath = os.path.join(DATA_DIR, profile.subject_file)
    if not os.path.exists(fpath):
        return None
    df = pd.read_excel(fpath, dtype=str, engine=_excel_engine(fpath))
    return SubjectMatcher(df)


def get_first_bank(short_name: str) -> Optional[BankAccount]:
    """返回排在第一位的银行账号"""
    profile = get_company(short_name)
    if not profile or not profile.bank_accounts:
        return None
    return profile.bank_accounts[0]


def find_bank_by_account(short_name: str, account_no: str) -> Optional[BankAccount]:
    """根据账号查找登记的银行账户"""
    profile = get_company(short_name)
    if not profile:
        return None
    for acct in profile.bank_accounts:
        if acct.account_no.strip() == account_no.strip():
            return acct
    return None


# ── 匹配规则管理 ─────────────────────────────────────────────
# 规则格式：
# {
#   "match_note": "关键词（备注包含）",       # 与 match_summary 二选一或同时
#   "match_summary": "关键词（摘要包含）",
#   "match_counterpart": "对方户名包含",      # 可选，组合条件
#   "subject_code": "5503001",
#   "label": "显示名称（用于 UI）"
# }

def load_rules(short_name: str) -> list[dict]:
    """加载公司自定义匹配规则"""
    _ensure_data_dir()
    path = os.path.join(DATA_DIR, f'rules_{short_name}.json')
    if not os.path.exists(path):
        return []
    with open(path, 'r', encoding='utf-8') as f:
        return json.load(f)


def save_rules(short_name: str, rules: list[dict]):
    """全量保存规则列表"""
    _ensure_data_dir()
    path = os.path.join(DATA_DIR, f'rules_{short_name}.json')
    with open(path, 'w', encoding='utf-8') as f:
        json.dump(rules, f, ensure_ascii=False, indent=2)


def add_rule(short_name: str, rule: dict) -> list[dict]:
    """
    添加一条规则（自动去重：相同匹配条件的旧规则会被替换）。
    返回更新后的规则列表。
    """
    rules = load_rules(short_name)
    # 去掉完全相同匹配条件的旧规则
    rules = [
        r for r in rules
        if not (
            r.get('match_note', '') == rule.get('match_note', '') and
            r.get('match_summary', '') == rule.get('match_summary', '') and
            r.get('match_counterpart', '') == rule.get('match_counterpart', '')
        )
    ]
    rules.insert(0, rule)  # 新规则插到最前，优先级最高
    save_rules(short_name, rules)
    return rules


def apply_rules(rules: list[dict], note: str, summary: str, counterpart: str) -> Optional[str]:
    """
    按顺序匹配规则，返回第一个匹配的 subject_code，无匹配返回 None。
    所有非空条件必须同时满足。
    """
    for rule in rules:
        note_kw = rule.get('match_note', '').strip()
        sum_kw  = rule.get('match_summary', '').strip()
        cp_kw   = rule.get('match_counterpart', '').strip()

        if not note_kw and not sum_kw:
            continue  # 空规则跳过

        ok = True
        if note_kw and note_kw not in (note or ''):
            ok = False
        if sum_kw and sum_kw not in (summary or ''):
            ok = False
        if cp_kw and cp_kw not in (counterpart or ''):
            ok = False

        if ok:
            return rule['subject_code']
    return None
