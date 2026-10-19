# -*- coding: utf-8 -*-
"""
验收单报表读取与清洗模块

为什么不用 openpyxl：
  SAP 导出的 xlsx（8-10MB / 4-5万行）用 openpyxl(read_only=True) 读出的 max_row 恒为 1，
  普通模式则直接卡死。这里用 zipfile 直接解析 xl/worksheets/sheetN.xml，
  内存占用低且速度稳定（10 万行约 10 秒）。

已知该格式特征（脚本已全部处理）：
  - 表头是 HTML 实体编码，如 &#20844;&#21496; 实为「公司」
  - 单元格 t="inlineStr"，值在 <is><t> 内
  - 「编制验收单时间」是 Excel 日期序列号（如 45658.3629）
  - 无 sharedStrings.xml
"""
import zipfile
import re
import html
import os
import glob
from datetime import datetime, timedelta

import pandas as pd
import numpy as np

CELL_RE = re.compile(r'<c r="([A-Z]+)(\d+)"([^>]*)>(.*?)</c>', re.S)
ROW_RE = re.compile(r'<row[^>]*r="(\d+)"[^>]*>(.*?)</row>', re.S)

# 标准字段（29 列）
REQUIRED_COLS = [
    "公司", "SAP订单号", "物资类别", "供应商名称", "合同/订货函编号", "协议编号",
    "合同创建人", "协议类型", "物料编码", "物料描述", "需求计划编号", "采购数量",
    "单位", "验收数量", "验收金额", "库存地", "验收记录", "是否入库", "使用单位",
    "验收单创建人", "编制验收单时间", "验收申请编号", "验收单编号", "验收状态",
    "验收进度", "验收申请时间", "合同到货时间", "实际到货日期", "逾期天数",
]

NUMERIC_COLS = ["采购数量", "验收数量", "验收金额", "逾期天数"]
TEXT_COLS = [
    "公司", "物资类别", "供应商名称", "合同/订货函编号", "协议编号", "合同创建人",
    "协议类型", "物料编码", "物料描述", "库存地", "使用单位", "验收单创建人",
    "验收状态", "验收进度", "是否入库", "验收单编号",
]


def _col_to_idx(col: str) -> int:
    n = 0
    for ch in col:
        n = n * 26 + (ord(ch) - 64)
    return n - 1


def _excel_serial_to_date(v):
    """Excel 日期序列号 -> datetime.date"""
    try:
        f = float(v)
    except (TypeError, ValueError):
        return None
    if f < 20000 or f > 80000:  # 合理日期区间：1954-2020 之后
        return None
    return (datetime(1899, 12, 30) + timedelta(days=f)).date()


def read_sheet(path: str, sheet_index: int = 1):
    """流式读取单个 sheet，返回 (header, rows)"""
    with zipfile.ZipFile(path) as z:
        names = z.namelist()
        target = f"xl/worksheets/sheet{sheet_index}.xml"
        if target not in names:
            candidates = sorted(n for n in names if re.match(r"xl/worksheets/sheet\d+\.xml$", n))
            if not candidates:
                raise ValueError(f"{path} 中未找到工作表 XML")
            target = candidates[0]
        raw = z.read(target).decode("utf-8", "ignore")

    header = None
    rows = []
    for rm in ROW_RE.finditer(raw):
        rnum = int(rm.group(1))
        body = rm.group(2)
        cells = {}
        maxi = -1
        for cm in CELL_RE.finditer(body):
            col, attrs, inner = cm.group(1), cm.group(3), cm.group(4)
            i = _col_to_idx(col)
            t = re.search(r't="(\w+)"', attrs)
            t = t.group(1) if t else "n"
            if t == "inlineStr":
                val = "".join(re.findall(r"<t[^>]*>(.*?)</t>", inner, re.S))
            else:
                vm = re.search(r"<v>(.*?)</v>", inner, re.S)
                val = vm.group(1) if vm else ""
            cells[i] = html.unescape(val)
            if i > maxi:
                maxi = i
        if rnum == 1:
            header = [cells.get(i, "") for i in range(maxi + 1)]
            continue
        if header is None:
            raise ValueError(f"{path} 第 1 行不是表头，格式不符合预期")
        rows.append([cells.get(i, "") for i in range(len(header))])
    return header, rows


def read_report(path: str) -> pd.DataFrame:
    """读取单份验收单报表 -> 清洗后的 DataFrame"""
    header, rows = read_sheet(path)
    header = [h.strip() for h in header]
    df = pd.DataFrame(rows, columns=header)

    missing = [c for c in REQUIRED_COLS if c not in df.columns]
    if missing:
        raise ValueError(
            f"{os.path.basename(path)} 缺少必需字段: {missing}\n"
            f"实际字段: {list(df.columns)}"
        )

    df = df[REQUIRED_COLS].copy()

    for c in NUMERIC_COLS:
        df[c] = pd.to_numeric(df[c], errors="coerce")
    for c in TEXT_COLS:
        if c in df.columns:
            df[c] = df[c].astype(str).str.strip().replace({"nan": "", "None": "", "NaT": ""})

    # 日期：编制验收单时间为 Excel 序列号；申请时间/到货时间为文本
    df["验收单日期"] = pd.to_datetime(
        df["编制验收单时间"].apply(_excel_serial_to_date), errors="coerce"
    )
    df["验收申请日期"] = pd.to_datetime(df["验收申请时间"], errors="coerce", format="mixed")
    df["实际到货日期"] = pd.to_datetime(df["实际到货日期"], errors="coerce", format="mixed")
    df["合同到货日期"] = pd.to_datetime(df["合同到货时间"], errors="coerce", format="mixed")

    # 业务日期：优先验收单日期，缺失回退申请日期
    df["业务日期"] = df["验收单日期"].fillna(df["验收申请日期"])
    df["源文件"] = os.path.basename(path)
    return df


def detect_year(df: pd.DataFrame) -> int:
    """从业务日期推断该文件对应年份"""
    valid = df["业务日期"].dropna()
    if len(valid) == 0:
        raise ValueError("无法解析任何业务日期，无法判断年份")
    return int(valid.dt.year.mode().iloc[0])


def discover_files(data_dir: str, pattern: str = "验收单报表*.xlsx"):
    """在目录中自动发现验收单报表文件，按年份去重（同年取最新修改的）"""
    files = glob.glob(os.path.join(data_dir, "**", pattern), recursive=True)
    files = [f for f in files if not os.path.basename(f).startswith("~$")]
    if not files:
        raise FileNotFoundError(f"在 {data_dir} 下未找到匹配 {pattern} 的文件")

    by_year = {}
    for f in files:
        m = re.search(r"(20\d{2})", os.path.basename(f))
        if not m:
            continue
        y = int(m.group(1))
        if y not in by_year or os.path.getmtime(f) > os.path.getmtime(by_year[y]):
            by_year[y] = f
    if not by_year:
        raise FileNotFoundError(f"文件名中未找到年份（如 2025/2026）: {files}")
    return dict(sorted(by_year.items()))


def load_dataset(paths):
    """读取多份报表并按业务年份合并"""
    frames = []
    for p in paths:
        df = read_report(p)
        df["数据年份"] = detect_year(df)
        frames.append(df)
    out = pd.concat(frames, ignore_index=True)
    out["年月"] = out["业务日期"].dt.to_period("M").astype(str)
    out["月序"] = out["业务日期"].dt.month
    return out


def quality_report(df: pd.DataFrame) -> dict:
    """数据质量体检，供报告「数据说明」章节使用"""
    n = len(df)
    missing = {}
    for c in ["验收金额", "验收数量", "供应商名称", "物资类别", "使用单位"]:
        if c in df.columns:
            if df[c].dtype == object:
                miss = int((df[c] == "").sum())
            else:
                miss = int(df[c].isna().sum())
            missing[c] = {"缺失行数": miss, "缺失率%": round(miss / n * 100, 2) if n else 0}

    blank_unit = df[df["使用单位"] == ""] if "使用单位" in df.columns else df.iloc[0:0]
    return {
        "总行数": n,
        "年份分布": {int(k): int(v) for k, v in df["数据年份"].value_counts().sort_index().items()},
        "业务日期范围": [
            str(df["业务日期"].min().date()) if pd.notna(df["业务日期"].min()) else None,
            str(df["业务日期"].max().date()) if pd.notna(df["业务日期"].max()) else None,
        ],
        "日期缺失行数": int(df["业务日期"].isna().sum()),
        "字段缺失": missing,
        "使用单位空白": {
            "行数": int(len(blank_unit)),
            "金额": float(blank_unit["验收金额"].sum()),
            "金额占比%": round(blank_unit["验收金额"].sum() / df["验收金额"].sum() * 100, 2)
            if df["验收金额"].sum() else 0,
            "类别构成": {k: round(v * 100, 1) for k, v in
                         blank_unit["物资类别"].value_counts(normalize=True).head(3).items()}
            if len(blank_unit) else {},
        },
    }
