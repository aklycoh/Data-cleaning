"""将检测到的表格结构转换为长表。"""

import pandas as pd

from .reader import SheetData
from .detector import TableInfo, TableType


def _build_row_label_names(
    grid: list[list],
    header_end: int,
    row_label_cols: list[int],
) -> list[str]:
    """为行标签列构建列名。取表头区域最后一行的值。"""
    names = []
    for c in row_label_cols:
        if header_end < len(grid) and c < len(grid[header_end]):
            val = grid[header_end][c]
        else:
            val = None
        name = str(val).strip() if val is not None else f"标签{c + 1}"
        if not name:
            name = f"标签{c + 1}"
        # 去重
        base, count = name, 1
        while name in names:
            count += 1
            name = f"{base}_{count}"
        names.append(name)
    return names


def transform(sheet_data: SheetData, table_info: TableInfo) -> pd.DataFrame:
    """根据表格类型将数据转换为长表格式。"""
    grid = sheet_data.grid
    header_end = table_info.header_end_row
    data_rows = grid[header_end + 1:]

    if table_info.table_type == TableType.CROSS_TAB:
        return _transform_cross_tab(data_rows, grid, table_info)
    else:
        return _transform_dimension_data(data_rows, grid, table_info)


def _transform_cross_tab(
    data_rows: list[list],
    grid: list[list],
    info: TableInfo,
) -> pd.DataFrame:
    """交叉表 → 长表。"""
    row_label_names = _build_row_label_names(
        grid, info.header_end_row, info.row_label_cols
    )

    # 构建宽表
    records = []
    for row in data_rows:
        record = {}
        for i, c in enumerate(info.row_label_cols):
            record[row_label_names[i]] = row[c] if c < len(row) else None
        for i, c in enumerate(info.value_cols):
            record[info.header_labels[i]] = row[c] if c < len(row) else None
        records.append(record)

    df_wide = pd.DataFrame(records)

    if df_wide.empty:
        return pd.DataFrame(columns=row_label_names + ["指标", "值"])

    # 行标签列用 ffill 填充（处理垂直合并）
    for col in row_label_names:
        if col in df_wide.columns:
            df_wide[col] = df_wide[col].ffill()

    # 去除全空的数据行
    value_col_names = [c for c in info.header_labels if c in df_wide.columns]
    df_wide = df_wide.dropna(subset=value_col_names, how="all") if value_col_names else df_wide

    # melt 转长表
    df_long = df_wide.melt(
        id_vars=row_label_names,
        value_vars=value_col_names,
        var_name="指标",
        value_name="值",
    )

    df_long = df_long.reset_index(drop=True)
    return df_long


def _transform_dimension_data(
    data_rows: list[list],
    grid: list[list],
    info: TableInfo,
) -> pd.DataFrame:
    """维度在上表 → 直接构建 DataFrame。"""
    # 用表头最后一行作为列名
    header_end = info.header_end_row
    col_names = []
    for c in range(len(grid[0]) if grid else 0):
        # 从上到下读取表头，去重连续相同值，用 '_' 连接
        parts = []
        for r in range(header_end + 1):
            if r >= len(grid) or c >= len(grid[r]):
                continue
            val = grid[r][c]
            s = str(val).strip() if val is not None else ""
            if s and (not parts or parts[-1] != s):
                parts.append(s)
        name = "_".join(parts) if parts else f"列{c + 1}"
        # 去重
        base, count = name, 1
        while name in col_names:
            count += 1
            name = f"{base}_{count}"
        col_names.append(name)

    n_cols = len(col_names)
    # 规范化数据行长度，确保与列名数量一致
    normalized_rows = []
    for row in data_rows:
        if len(row) < n_cols:
            normalized_rows.append(list(row) + [None] * (n_cols - len(row)))
        elif len(row) > n_cols:
            normalized_rows.append(row[:n_cols])
        else:
            normalized_rows.append(row)

    df = pd.DataFrame(normalized_rows, columns=col_names)

    # 去除全空行
    df = df.dropna(how="all")
    df = df.reset_index(drop=True)
    return df
