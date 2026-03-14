"""
每家公司的自定义科目匹配规则。
存储在 data/rules_{company_short}.json

规则格式：
[
  {
    "keyword": "加油费",
    "field": "note" | "summary" | "counterpart" | "any",
    "subject_code": "1133001",
    "subject_name": "其他应收款_中油润德（预存加油费）"
  },
  ...
]

匹配逻辑：
- field="note"        → 检查备注字段
- field="summary"     → 检查摘要字段
- field="counterpart" → 检查对方户名
- field="any"         → 检查以上全部
- 规则按顺序匹配，第一个命中的生效
- 用户自定义规则优先于内置规则
"""
import json
import os

DATA_DIR = os.path.join(os.path.dirname(__file__), 'data')


def _rules_path(company_short: str) -> str:
    os.makedirs(DATA_DIR, exist_ok=True)
    return os.path.join(DATA_DIR, f'rules_{company_short}.json')


def load_rules(company_short: str) -> list[dict]:
    """加载公司自定义规则列表"""
    path = _rules_path(company_short)
    if not os.path.exists(path):
        return []
    try:
        with open(path, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception:
        return []


def save_rules(company_short: str, rules: list[dict]):
    """保存（覆盖）公司规则列表"""
    with open(_rules_path(company_short), 'w', encoding='utf-8') as f:
        json.dump(rules, f, ensure_ascii=False, indent=2)


def add_rule(company_short: str, keyword: str, field: str, subject_code: str, subject_name: str):
    """添加一条规则（如果相同 keyword+field 已存在则更新）"""
    rules = load_rules(company_short)
    # 查找是否已存在相同 keyword+field
    for r in rules:
        if r['keyword'] == keyword and r['field'] == field:
            r['subject_code'] = subject_code
            r['subject_name'] = subject_name
            save_rules(company_short, rules)
            return
    rules.append({
        'keyword': keyword,
        'field': field,
        'subject_code': subject_code,
        'subject_name': subject_name,
    })
    save_rules(company_short, rules)


def delete_rule(company_short: str, index: int):
    """删除指定序号的规则"""
    rules = load_rules(company_short)
    if 0 <= index < len(rules):
        rules.pop(index)
        save_rules(company_short, rules)


def apply_rules(note: str, summary: str, counterpart: str, rules: list[dict]) -> tuple[str, str] | None:
    """
    用自定义规则尝试匹配，返回 (subject_code, subject_name) 或 None。
    """
    for r in rules:
        kw = r.get('keyword', '')
        if not kw:
            continue
        field = r.get('field', 'any')
        matched = False
        if field == 'note':
            matched = kw in note
        elif field == 'summary':
            matched = kw in summary
        elif field == 'counterpart':
            matched = kw in counterpart
        else:  # any
            matched = kw in note or kw in summary or kw in counterpart
        if matched:
            return r['subject_code'], r.get('subject_name', r['subject_code'])
    return None
