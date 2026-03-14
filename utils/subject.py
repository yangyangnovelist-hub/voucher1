"""
科目匹配工具
"""
import pandas as pd
import re


class SubjectMatcher:
    def __init__(self, df: pd.DataFrame):
        df = df.copy()

        # ── 列检测：同时支持关键词匹配和位置兜底 ──
        code_col = self._detect_col(df, ['编码', '代码', '科目编码', '科目代码', 'code'])
        name_col = self._detect_col(df, ['名称', '科目名称', '科目名', 'name'])

        # 兜底：取前两列，哪列更多数字串就是 code 列
        if code_col is None or name_col is None or code_col == name_col:
            code_col, name_col = self._detect_by_content(df)

        df = df[[code_col, name_col]].copy()
        df.columns = ['code', 'name']
        df = df.dropna(subset=['code'])
        df['code'] = df['code'].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)
        df['name'] = df['name'].fillna('').astype(str).str.strip()
        # 过滤空行和纯数字以外的杂行
        df = df[df['code'].str.match(r'^\d+$')]

        self.code_to_name: dict[str, str] = dict(zip(df['code'], df['name']))
        self._df = df

    @staticmethod
    def _detect_col(df, keywords):
        for kw in keywords:
            for col in df.columns:
                if kw in str(col):
                    return col
        return None

    @staticmethod
    def _detect_by_content(df):
        """按内容判断哪列是 code（纯数字比例高），哪列是 name"""
        scores = {}
        for col in df.columns[:5]:  # 只看前5列
            vals = df[col].dropna().astype(str).str.strip()
            numeric_ratio = vals.str.match(r'^\d+$').mean()
            scores[col] = numeric_ratio
        sorted_cols = sorted(scores, key=scores.get, reverse=True)
        code_col = sorted_cols[0]
        # name col: 首选含「名」的，否则取 code_col 右边一列
        remaining = [c for c in df.columns if c != code_col]
        name_col = next((c for c in remaining if '名' in str(c)), remaining[0] if remaining else code_col)
        return code_col, name_col

    def get_display_name(self, code: str) -> str:
        name = self.code_to_name.get(str(code), '')
        if not name:
            return code
        return name.split('_')[-1] if '_' in name else name

    def get_ar_account(self, customer: str) -> tuple[str, str, bool]:
        return self._fuzzy_match(customer, '1131', default_code='1131')

    def get_ap_account(self, supplier: str) -> tuple[str, str, bool]:
        return self._fuzzy_match(supplier, '2121', default_code='2121')

    def find_sub_account(self, parent_prefix: str, name: str) -> tuple[str, str]:
        code, dname, is_def = self._fuzzy_match(name, parent_prefix, default_code='')
        if is_def:
            return '', ''
        return code, dname

    def _fuzzy_match(self, query: str, prefix: str, default_code: str) -> tuple[str, str, bool]:
        if not query or str(query).strip() in ('', 'nan', 'None'):
            return default_code, self.get_display_name(default_code), True

        candidates = [
            (code, name)
            for code, name in self.code_to_name.items()
            if code.startswith(prefix) and code != prefix
        ]
        if not candidates:
            return default_code, self.get_display_name(default_code), True

        query_clean = _clean(query)

        for code, name in candidates:
            if _clean(name) == query_clean:
                return code, self.get_display_name(code), False

        for code, name in candidates:
            nc = _clean(name)
            if query_clean in nc or nc in query_clean:
                return code, self.get_display_name(code), False

        best_code, best_score = '', 0
        for code, name in candidates:
            score = _lcs_len(_clean(name), query_clean)
            if score > best_score and score >= 2:
                best_score = score
                best_code = code

        if best_code:
            return best_code, self.get_display_name(best_code), False

        return default_code, self.get_display_name(default_code), True


def _clean(s: str) -> str:
    return re.sub(r'[\s\W_]+', '', s, flags=re.UNICODE).lower()


def _lcs_len(a: str, b: str) -> int:
    if not a or not b:
        return 0
    m, n = len(a), len(b)
    best = 0
    dp = [[0] * (n + 1) for _ in range(m + 1)]
    for i in range(1, m + 1):
        for j in range(1, n + 1):
            if a[i-1] == b[j-1]:
                dp[i][j] = dp[i-1][j-1] + 1
                best = max(best, dp[i][j])
    return best
