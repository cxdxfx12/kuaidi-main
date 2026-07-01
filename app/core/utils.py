"""
共享工具函数
避免在 calculate_service.py 和 rule_service.py 中重复定义相同函数
"""
from datetime import datetime
from typing import Optional


def parse_date(date_str: str) -> Optional[datetime]:
    """解析日期，支持多种格式：2026-06-19 / 2026/6/19 / 2026.6.19 / 20260619 / 2026年6月19日"""
    if not date_str:
        return None
    s = str(date_str).strip()
    if not s:
        return None
    # 统一去掉中文字符和多余符号
    for ch in ["年", "月", "日", ".", "-", "/"]:
        s = s.replace(ch, "-")
    # 去掉连续的"-"
    while "--" in s:
        s = s.replace("--", "-")
    s = s.strip("-")
    parts = s.split("-")
    if len(parts) >= 3:
        try:
            y = int(parts[0])
            m = int(parts[1])
            d = int(parts[2])
            return datetime(y, m, d)
        except (ValueError, TypeError):
            pass
    # 兜底：纯8位数字如 20260619
    if s.isdigit() and len(s) == 8:
        try:
            return datetime(int(s[:4]), int(s[4:6]), int(s[6:8]))
        except (ValueError, TypeError):
            pass
    # 最后再尝试标准格式
    for fmt in ["%Y-%m-%d", "%Y/%m/%d", "%Y%m%d"]:
        try:
            return datetime.strptime(str(date_str).strip(), fmt)
        except Exception:
            continue
    return None
