"""Streamlit Web UI — 数据清洗工具。"""

import tempfile
import os
from io import BytesIO

import streamlit as st
import pandas as pd

from core.reader import get_sheet_names
from core.pipeline import clean

st.set_page_config(page_title="数据清洗工具", layout="wide")
st.title("数据清洗工具")
st.caption("将含合并单元格的复杂表头表格转换为长表格式，支持 .xlsx / .xls / .csv")

uploaded = st.file_uploader("上传文件", type=["xlsx", "xls", "csv"])

if uploaded is not None:
    suffix = os.path.splitext(uploaded.name)[1].lower()
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
    tmp.write(uploaded.read())
    tmp.close()
    tmp_path = tmp.name

    try:
        sheet_name = None
        if suffix in (".xlsx", ".xls"):
            sheets = get_sheet_names(tmp_path)
            sheet_name = st.selectbox("选择工作表", sheets)

        if st.button("开始清洗"):
            with st.spinner("正在清洗数据..."):
                df = clean(tmp_path, sheet_name)

            st.success(f"清洗完成！共 {len(df)} 行 × {len(df.columns)} 列")
            st.dataframe(df, use_container_width=True)

            buf = BytesIO()
            df.to_excel(buf, index=False, engine="openpyxl")
            buf.seek(0)
            st.download_button(
                label="下载清洗结果 (.xlsx)",
                data=buf,
                file_name="cleaned.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
    finally:
        os.unlink(tmp_path)
