"""
日期工具：从发票 DataFrame 中提取凭证日期（取最后一天）
"""
import pandas as pd
from datetime import date
import calendar


def detect_date_col(df: pd.DataFrame, candidates: list[str]) -> str | None:
    for c in candidates:
        if c in df.columns:
            return c
    return None


def extract_voucher_date(df: pd.DataFrame, date_col: str) -> date | None:
    """
    从 date_col 列提取所有日期，取最大日期所在月份的最后一天作为凭证日期。
    """
    dates = []
    for val in df[date_col].dropna():
        d = _parse_date(val)
        if d:
            dates.append(d)
    if not dates:
        return None
    latest = max(dates)
    last = last_day_of_month(latest.year, latest.month)
    return last


def last_day_of_month(year: int, month: int) -> date:
    day = calendar.monthrange(year, month)[1]
    return date(year, month, day)


def format_date_ymd(val) -> str:
    """将各种格式的日期值转为 YYYY-MM-DD 字符串"""
    d = _parse_date(val)
    return d.strftime('%Y-%m-%d') if d else str(val)


def _parse_date(val) -> date | None:
    if pd.isna(val) if not isinstance(val, str) else val.strip() in ('', 'nan', 'None'):
        return None
    if isinstance(val, (pd.Timestamp, date)):
        if isinstance(val, pd.Timestamp):
            return val.date()
        return val
    s = str(val).strip()
    for fmt in ('%Y-%m-%d', '%Y/%m/%d', '%Y%m%d', '%m/%d/%Y', '%d/%m/%Y'):
        try:
            return pd.to_datetime(s, format=fmt).date()
        except Exception:
            continue
    try:
        return pd.to_datetime(s).date()
    except Exception:
        return None
