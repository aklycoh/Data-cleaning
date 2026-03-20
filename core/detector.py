"""检测表格类型和表头边界。"""

from dataclasses import dataclass, field
from enum import Enum, auto
from typing import Any

from .reader import SheetData, MergeInfo


class TableType(Enum):
    CROSS_TAB = auto()       # 交叉表（表头有水平合并）
    DIMENSION_DATA = auto()  # 维度在上、数据在下


@dataclass
class TableInfo:
    table_type: TableType
    header_end_row: int          # 表头末行索引（0-based，含）
    row_label_cols: list[int]    # 行标签列索引（id_vars）
    value_cols: list[int]        # 数据列索引（value_vars）
    header_labels: list[str]     # 每个数据列的组合列名


def _is_numeric(val: Any) -> bool:
    """判断值是否为数值类型。"""
    if val is None:
        return False
    if isinstance(val, (int, float)):
        return True
    if isinstance(val, str):
        val = val.strip()
        if not val:
            return False
        try:
            float(val.replace(",", ""))
            return True
        except ValueError:
            return False
    return False


def _find_header_end(grid: list[list], merges: list[MergeInfo], n_cols: int) -> int:
    """找到表头末行（0-based）。

    策略：
    1. 含有水平合并的行一定是表头行。
    2. 逐行扫描，当某行数值占比 > 50% 时，该行是数据行，上一行为表头末行。
    3. 如果扫描完没找到，默认第 0 行为表头。
    """
    # 找出含水平合并的行（0-based）
    h_merge_rows: set[int] = set()
    for m in merges:
        if m.is_horizontal:
            for r in range(m.min_row - 1, m.max_row):
                h_merge_rows.add(r)

    header_end = 0

    for i, row in enumerate(grid):
        # 计算数值占比
        non_none = [v for v in row if v is not None]
        if not non_none:
            continue
        num_count = sum(1 for v in non_none if _is_numeric(v))
        ratio = num_count / len(non_none)

        if ratio > 0.5 and i not in h_merge_rows:
            # 这一行是数据行
            header_end = max(i - 1, 0)
            break
    else:
        # 没找到明显数据行，用水平合并的最大行号
        if h_merge_rows:
            header_end = max(h_merge_rows)
        else:
            header_end = 0

    # 确保所有水平合并行都在表头范围内
    if h_merge_rows:
        header_end = max(header_end, max(h_merge_rows))

    # 确保 header_end 不超出 grid 范围，且至少留一行数据
    max_allowed = len(grid) - 2 if len(grid) >= 2 else 0
    header_end = min(header_end, max_allowed)

    return header_end


def _classify_columns(
    grid: list[list],
    merges: list[MergeInfo],
    header_end: int,
    n_cols: int,
) -> tuple[list[int], list[int]]:
    """将列分为行标签列和数据列。

    行标签列：表头区域内未被水平合并覆盖的左侧连续列。
    数据列：其余列。
    """
    # 找出被水平合并覆盖的列
    h_merged_cols: set[int] = set()
    for m in merges:
        if m.is_horizontal and m.min_row - 1 <= header_end:
            for c in range(m.min_col - 1, m.max_col):
                h_merged_cols.add(c)

    # 从左边开始，连续未被水平合并覆盖的列为行标签列
    row_label_cols = []
    for c in range(n_cols):
        if c not in h_merged_cols:
            row_label_cols.append(c)
        else:
            break

    # 如果没有行标签列，至少取第一列
    if not row_label_cols and n_cols > 0:
        row_label_cols = [0]

    value_cols = [c for c in range(n_cols) if c not in row_label_cols]

    return row_label_cols, value_cols


def _build_header_labels(
    grid: list[list],
    header_end: int,
    value_cols: list[int],
) -> list[str]:
    """为每个数据列构建组合列名。

    从上到下读取表头单元格值，用 '_' 连接，去重连续相同值。
    """
    labels = []
    for c in value_cols:
        parts = []
        for r in range(header_end + 1):
            if r >= len(grid) or c >= len(grid[r]):
                continue
            val = grid[r][c]
            s = str(val).strip() if val is not None else ""
            if s and (not parts or parts[-1] != s):
                parts.append(s)
        label = "_".join(parts) if parts else f"列{c + 1}"
        # 去重：如果已有同名列，加后缀
        base, count = label, 1
        while label in labels:
            count += 1
            label = f"{base}_{count}"
        labels.append(label)
    return labels


def detect(sheet_data: SheetData) -> TableInfo:
    """检测表格类型和结构。"""
    grid = sheet_data.grid
    merges = sheet_data.merges
    n_cols = sheet_data.n_cols

    header_end = _find_header_end(grid, merges, n_cols)

    # 判定类型：表头区域是否有水平合并
    has_h_merge = any(
        m.is_horizontal and m.min_row - 1 <= header_end
        for m in merges
    )
    table_type = TableType.CROSS_TAB if has_h_merge else TableType.DIMENSION_DATA

    if table_type == TableType.CROSS_TAB:
        row_label_cols, value_cols = _classify_columns(
            grid, merges, header_end, n_cols
        )
    else:
        # 维度在上表：所有列都是普通列
        row_label_cols = list(range(n_cols))
        value_cols = []

    header_labels = _build_header_labels(grid, header_end, value_cols)

    return TableInfo(
        table_type=table_type,
        header_end_row=header_end,
        row_label_cols=row_label_cols,
        value_cols=value_cols,
        header_labels=header_labels,
    )
