from __future__ import annotations
import re
from typing import List, Tuple

# 你可以按自己 BOM 特点不断加强这些保护规则
RE_PURE_NUMBER = re.compile(r"^\s*[-+]?\d+(\.\d+)?\s*$")
RE_CODE_LIKE = re.compile(r"^[A-Za-z0-9][A-Za-z0-9\-\._/]*$")  # 类似型号/编码
RE_MIXED_SIZE = re.compile(r"^\s*\d+(\.\d+)?\s*(mm|MM|cm|CM|m|M|inch|INCH|in|IN)\s*$")

# 含中文判断（最重要，用于 QA）
RE_HAS_CN = re.compile(r"[\u4e00-\u9fff]")

def should_skip_cell(text: str) -> bool:
    """这些单元格通常是数量、编号、纯型号、纯代码，不翻译。"""
    if text is None:
        return True
    s = str(text).strip()
    if not s:
        return True

    # 纯数字/尺寸
    if RE_PURE_NUMBER.match(s) or RE_MIXED_SIZE.match(s):
        return True

    # 全英文/数字/符号且像编码
    # 注意：如果你的品名里也会出现纯英文（比如 BOLT），这条会跳过
    # 你可以用“只翻译指定列”策略规避（后面主程序支持）。
    if RE_CODE_LIKE.match(s) and not RE_HAS_CN.search(s):
        return True

    return False

def has_chinese(text: str) -> bool:
    return bool(RE_HAS_CN.search(str(text)))
