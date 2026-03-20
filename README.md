# Excel 数据清洗工具

将含合并单元格复杂表头的表格（交叉表、维度在上表）自动转换为长表格式，通过本地 Web UI 操作。

## 支持格式

| 格式 | 合并单元格还原 | 说明 |
|------|:-:|------|
| `.xlsx` | ✅ | openpyxl 读取 |
| `.xls` | ✅ | xlrd 读取 |
| `.csv` | — | 无合并单元格 |

## 快速开始

### 1. 安装依赖

```bash
pip install -r requirements.txt
```

### 2. 启动

双击 `start.bat`，或命令行运行：

```bash
streamlit run app.py
```

浏览器打开 http://localhost:8501 即可使用。

## 使用方法

1. 上传文件（.xlsx / .xls / .csv）
2. 选择工作表（Excel 文件）
3. 点击"开始清洗"
4. 预览结果，点击"下载清洗结果"导出 .xlsx

## 项目结构

```
├── app.py              # Streamlit Web UI
├── start.bat           # 一键启动脚本
├── requirements.txt    # 依赖列表
└── core/
    ├── reader.py       # 读取文件 + 还原合并单元格
    ├── detector.py     # 检测表头边界与表格类型
    ├── transformer.py  # 转换为长表
    └── pipeline.py     # 编排清洗流程
```

## 核心逻辑

- **表头检测**：逐行扫描，数值占比 > 50% 的行视为数据起始行，之上为表头；含水平合并的行强制归为表头。
- **类型判定**：表头有水平合并 → 交叉表，否则 → 维度在上表。
- **列名构建**：表头区域自上而下拼接（去重连续相同值），如 `2023年_Q1`。
- **长表转换**：交叉表通过 `pd.melt()` 展开，行标签列垂直合并用 `ffill` 填充。
