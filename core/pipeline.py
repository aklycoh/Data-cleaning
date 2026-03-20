"""编排整个清洗流程。"""

import os
import pandas as pd

from .reader import read_sheet, read_csv, read_xls, get_sheet_names
from .detector import detect
from .transformer import transform


def clean(file_path: str, sheet_name: str | None = None) -> pd.DataFrame:
    """读取 Excel/CSV 文件并转换为长表格式。"""
    ext = os.path.splitext(file_path)[1].lower()
    if ext == ".csv":
        sheet_data = read_csv(file_path)
    elif ext == ".xls":
        sheet_data = read_xls(file_path, sheet_name)
    else:
        sheet_data = read_sheet(file_path, sheet_name)
    table_info = detect(sheet_data)
    df = transform(sheet_data, table_info)
    return df
