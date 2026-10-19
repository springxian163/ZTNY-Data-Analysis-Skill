# -*- coding: utf-8 -*-
"""
Excel 分析看板生成器（11 工作表）

输入全部来自 analyzer / rules 的产出，不重新读原始数据、不重算指标——
保证 Excel 与 Word 报告口径完全一致。

用法：
    from excel_writer import write_excel
    path = write_excel(res, rule_out, quality, out_path, title="...")
"""
import os

import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# ---------------- 样式常量 ----------------
FONT_NAME = "微软雅黑"
C_PRIMARY = "1F4E79"
C_SUB = "2E5C8A"
C_ALERT = "C00000"
C_NOTE = "808080"

HDR_FILL = PatternFill("solid", fgColor=C_PRIMARY)
HDR_FONT = Font(color="FFFFFF", bold=True, size=10, name=FONT_NAME)
TITLE_FONT = Font(bold=True, size=14, color=C_PRIMARY, name=FONT_NAME)
DOC_TITLE_FONT = Font(bold=True, size=18, color=C_PRIMARY, name=FONT_NAME)
SUB_FONT = Font(bold=True, size=11, color=C_SUB, name=FONT_NAME)
BODY_FONT = Font(size=10, name=FONT_NAME)
BOLD_FONT = Font(size=10, bold=True, name=FONT_NAME)
ALERT_FONT = Font(size=9, italic=True, color=C_ALERT, name=FONT_NAME)
NOTE_FONT = Font(size=9, color=C_NOTE, name=FONT_NAME)

_THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)

NUM_FMT = "#,##0"
PCT_FMT = '0.0"%"'


# ---------------- 通用工具 ----------------
def _money(v):
    return "—" if v is None else float(v)


def _style_header(ws, row, ncols):
    for j in range(1, ncols + 1):
        cell = ws.cell(row=row, column=j)
        cell.fill = HDR_FILL
        cell.font = HDR_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDER


def _write_df(ws, dfx, start_row=1, start_col=1, num_cols=None, pct_cols=None):
    """写 DataFrame 并返回下一行号"""
    cols = list(dfx.columns)
    for j, name in enumerate(cols):
        ws.cell(row=start_row, column=start_col + j, value=str(name))
    _style_header(ws, start_row, len(cols))

    for i, (_, r) in enumerate(dfx.iterrows(), start=1):
        for j, name in enumerate(cols):
            v = r[name]
            if isinstance(v, np.integer):
                v = int(v)
            elif isinstance(v, np.floating):
                v = float(v)
            elif isinstance(v, (float, np.floating)) and np.isnan(v):
                v = None
            cell = ws.cell(row=start_row + i, column=start_col + j, value=v)
            cell.font = BODY_FONT
            cell.border = BORDER
            if name in (pct_cols or []):
                cell.number_format = PCT_FMT
            elif isinstance(v, (int, float)) and not isinstance(v, bool):
                cell.number_format = NUM_FMT
                cell.alignment = Alignment(horizontal="right")
    return start_row + len(dfx)


def _set_widths(ws, mapping):
    for col, w in mapping.items():
        ws.column_dimensions[col].width = w


def _note(ws, row, text, col=1, font=None):
    c = ws.cell(row=row, column=col, value=text)
    c.font = font or NOTE_FONT
    return c


def _new_sheet(wb, name, first=False):
    ws = wb.active if first else wb.create_sheet(name)
    ws.title = name
    ws.sheet_view.showGridLines = False
    return ws


# ---------------- 各工作表 ----------------
def _sheet_overview(wb, res, rule_out):
    t, meta = res["totals"], res["meta"]
    ov, conc = res["overdue"], res["supplier_concentration"]
    ws = _new_sheet(wb, "01 核心概览", first=True)

    ws["A1"] = res.get("_title", "采购验收数据分析报告")
    ws["A1"].font = DOC_TITLE_FONT
    ws["A2"] = (f"数据范围：{meta['data_range'][0]} 至 {meta['data_range'][1]} ｜ "
                f"分析口径：{meta['year_prev']} vs {meta['year_cur']} 可比区间（{meta['month_range']}）｜ "
                f"数据源：{' / '.join(res.get('_files', []))}")
    ws["A2"].font = NOTE_FONT

    rows = [
        ("验收总金额（本期）", t["amount_cur"], "元", f"同比 {rule_out['kpis'][0]['note'].replace('同比 ', '')}"),
        ("验收总金额（基期）", t["amount_prev"], "元", "基期"),
        ("同比增长", t["amount_yoy"], "%", f"增减 {t['amount_delta']:,.0f} 元".replace("增减 -", "减少 ").replace("增减 ", "增加 ")),
        ("验收单笔数（本期）", t["orders_cur"], "笔", f"同比 {_yoy_str(t['orders_yoy'])}"),
        ("验收单笔数（基期）", t["orders_prev"], "笔", "基期"),
        ("平均单笔金额（本期）", t["avg_order_cur"], "元", f"同比 {_yoy_str(t['avg_order_yoy'])}"),
        ("平均单笔金额（基期）", t["avg_order_prev"], "元", "基期"),
        ("合作供应商数（本期）", t["suppliers_cur"], "家", f"基期 {t['suppliers_prev']} 家"),
        ("采购物料种类（本期）", t["materials_cur"], "种", f"基期 {t['materials_prev']} 种"),
        ("逾期交付率（本期）", ov["cur"]["rate"], "%", f"基期 {ov['prev']['rate']:.1f}%"),
        ("平均逾期天数（本期）", ov["cur"]["avg_days"], "天", f"基期 {ov['prev']['avg_days']:.1f} 天"),
        ("CR5 集中度（本期）", conc["cur"]["cr5"], "%", f"基期 {conc['prev']['cr5']:.1f}%"),
        ("全年预估金额", t["annual_forecast"], "元", f"按 {meta['month_range']} 线性外推"),
    ]

    r = 4
    ws.cell(row=r, column=1, value="关键指标").font = SUB_FONT
    r += 1
    for j, h in enumerate(["指标", "数值", "单位", "说明"]):
        ws.cell(row=r, column=1 + j, value=h)
    _style_header(ws, r, 4)

    for name, val, unit, note in rows:
        r += 1
        ws.cell(row=r, column=1, value=name).font = BODY_FONT
        cv = ws.cell(row=r, column=2, value=round(val, 2) if isinstance(val, float) else val)
        cv.font = BOLD_FONT
        cv.number_format = "#,##0.00" if unit == "%" else NUM_FMT
        cv.alignment = Alignment(horizontal="right")
        ws.cell(row=r, column=3, value=unit).font = BODY_FONT
        ws.cell(row=r, column=4, value=note).font = NOTE_FONT
        for j in range(1, 5):
            ws.cell(row=r, column=j).border = BORDER

    r += 2
    _note(ws, r, "核心结论：" + rule_out["headline"], font=Font(size=10, bold=True, color=C_PRIMARY, name=FONT_NAME))
    _set_widths(ws, {"A": 30, "B": 20, "C": 8, "D": 30})
    return ws


def _yoy_str(v):
    return "—" if v is None else f"{v:+.1f}%"


def _sheet_monthly(wb, res):
    meta = res["meta"]
    yc, yp = meta["year_cur"], meta["year_prev"]
    ws = _new_sheet(wb, "02 月度趋势")
    ws["A1"] = "月度采购金额趋势（可比区间）"
    ws["A1"].font = TITLE_FONT

    df = pd.DataFrame([
        {"月份": m["month"], f"{yp}年": m["prev"], f"{yc}年": m["cur"],
         "同比%": m["yoy"] if m["yoy"] is not None else 0}
        for m in res["monthly"]
    ])
    end = _write_df(ws, df, start_row=3, num_cols=[f"{yp}年", f"{yc}年"], pct_cols=["同比%"])

    ch = LineChart()
    ch.title = f"月度验收金额对比（{yp} vs {yc}）"
    ch.style = 12
    ch.y_axis.title = "金额（元）"
    ch.x_axis.title = "月份"
    ch.add_data(Reference(ws, min_col=2, max_col=3, min_row=3, max_row=3 + len(df)),
                titles_from_data=True)
    ch.set_categories(Reference(ws, min_col=1, min_row=4, max_row=3 + len(df)))
    ch.height, ch.width = 9, 20
    ws.add_chart(ch, "F3")

    ms = res["monthly_summary"]
    if ms["max_yoy"] is not None:
        _note(ws, end + 2,
              f"月度同比区间 {ms['min_yoy']:+.1f}% ~ {ms['max_yoy']:+.1f}%，"
              f"波动{'剧烈，采购节奏受项目排期影响明显' if ms['volatile'] else '处于正常范围'}。")
    _set_widths(ws, {"A": 10, "B": 18, "C": 18, "D": 14})
    return ws


def _sheet_category(wb, res):
    meta = res["meta"]
    yc, yp = meta["year_cur"], meta["year_prev"]
    ws = _new_sheet(wb, "03 物资类别")
    ws["A1"] = "物资类别结构分析"
    ws["A1"].font = TITLE_FONT

    df = pd.DataFrame([
        {"物资类别": c["name"], f"{yp}年": c["prev"], f"{yc}年": c["cur"],
         "同比%": c["yoy"] if c["yoy"] is not None else 0, f"{yc}占比%": c["share"]}
        for c in res["categories"]
    ])
    _write_df(ws, df, start_row=3, num_cols=[f"{yp}年", f"{yc}年"],
              pct_cols=["同比%", f"{yc}占比%"])

    ch = BarChart()
    ch.type = "col"
    ch.title = "各物资类别金额对比"
    ch.y_axis.title = "金额（元）"
    ch.add_data(Reference(ws, min_col=2, max_col=3, min_row=3, max_row=3 + len(df)),
                titles_from_data=True)
    ch.set_categories(Reference(ws, min_col=1, min_row=4, max_row=3 + len(df)))
    ch.height, ch.width = 10, 20
    ws.add_chart(ch, "G3")
    _set_widths(ws, {"A": 14, "B": 16, "C": 16, "D": 12, "E": 12, "F": 12})
    return ws


def _sheet_supplier(wb, res):
    meta = res["meta"]
    yc, yp = meta["year_cur"], meta["year_prev"]
    conc, churn = res["supplier_concentration"], res["supplier_churn"]
    ws = _new_sheet(wb, "04 供应商分析")
    ws["A1"] = "供应商集中度与变化"
    ws["A1"].font = TITLE_FONT

    df = pd.DataFrame([
        {"年度": f"{yp}年", "供应商数量": conc["prev"]["count"],
         "CR1%": conc["prev"]["cr1"], "CR5%": conc["prev"]["cr5"],
         "CR10%": conc["prev"]["cr10"], "CR20%": conc["prev"]["cr20"]},
        {"年度": f"{yc}年", "供应商数量": conc["cur"]["count"],
         "CR1%": conc["cur"]["cr1"], "CR5%": conc["cur"]["cr5"],
         "CR10%": conc["cur"]["cr10"], "CR20%": conc["cur"]["cr20"]},
    ])
    end = _write_df(ws, df, start_row=3, num_cols=["供应商数量"],
                    pct_cols=["CR1%", "CR5%", "CR10%", "CR20%"])

    d = round(conc["cur"]["cr5"] - conc["prev"]["cr5"], 1)
    _note(ws, end + 2,
          f"解读：CR5 由 {conc['prev']['cr5']:.1f}% 变为 {conc['cur']['cr5']:.1f}%（{d:+.1f}pp）；"
          f"两年均合作 {churn['both']} 家，新增 {churn['new']} 家、流失 {churn['lost']} 家"
          f"（明细见「04A 新增供应商」「04B 退出供应商」）。",
          font=ALERT_FONT)

    r = end + 4
    ws.cell(row=r, column=1, value="TOP25 供应商（按本期金额）").font = SUB_FONT
    df2 = pd.DataFrame([
        {"供应商名称": s["name"], f"{yp}年": s["prev"], f"{yc}年": s["cur"],
         "同比%": s["yoy"] if s["yoy"] is not None else "新增", f"占{yc}%": s["share"]}
        for s in res["top_suppliers"]
    ])
    end2 = _write_df(ws, df2, start_row=r + 1, num_cols=[f"{yp}年", f"{yc}年"],
                     pct_cols=["同比%", f"占{yc}%"])

    r = end2 + 2
    ws.cell(row=r, column=1, value="细分品类集中度（本期）").font = SUB_FONT
    df3 = pd.DataFrame([
        {"类别": c["name"], "金额": c["amount"], "供应商数": c["suppliers"],
         "CR3%": c["cr3"], "TOP1供应商": c["top1_name"], "TOP1占比%": c["top1_share"]}
        for c in res["category_concentration"]
    ])
    _write_df(ws, df3, start_row=r + 1, num_cols=["金额", "供应商数"],
              pct_cols=["CR3%", "TOP1占比%"])
    _set_widths(ws, {"A": 32, "B": 16, "C": 16, "D": 14, "E": 26, "F": 14})
    return ws


def _sheet_supplier_churn(wb, res, kind):
    """
    新增 / 退出供应商完整名单（04A / 04B）。

    单独成表的理由：集中度与 TOP25 只看头部，而供应商更替反映的是渠道结构变化，
    需要逐家可核对——采购部门可直接当工作底稿用。
    """
    meta = res["meta"]
    yc, yp = meta["year_cur"], meta["year_prev"]
    churn = res["supplier_churn"]
    mr = meta["month_range"]
    is_new = kind == "new"
    data = res["supplier_new"] if is_new else res["supplier_lost"]
    name_cn = "新增" if is_new else "退出"

    ws = _new_sheet(wb, f"04{'A' if is_new else 'B'} {name_cn}供应商")

    if is_new:
        ws["A1"] = f"{yc} 年新增供应商名单（{mr} 可比区间）"
        sub = (f"共 {churn['new']} 家，合计验收金额 {churn['new_amount']:,.0f} 元，"
               f"占本期总金额 {churn['new_share']:.1f}%；单家中位数 {churn['new_median']:,.0f} 元。"
               f"口径：{yp} 年{ mr }无验收记录、{yc} 年{mr}有记录。")
        headers = ["序号", "供应商名称", f"{yc}金额（元）", "占本期%", "验收笔数", "物料数",
                   "平均单笔（元）", "主营类别", "主营占比%", "首次验收日期"]
        keys = ["seq", "name", "amount", "share", "orders", "materials",
                "avg_order", "top_cat", "top_cat_share", "first_date"]
    else:
        ws["A1"] = f"{yc} 年退出供应商名单（{mr} 可比区间）"
        sub = (f"共 {churn['lost']} 家，{yp} 年合计验收金额 {churn['lost_amount']:,.0f} 元，"
               f"占基期总金额 {churn['lost_share']:.1f}%；单家中位数 {churn['lost_median']:,.0f} 元。"
               f"口径：{yp} 年{mr}有验收记录、{yc} 年同区间无记录。")
        headers = ["序号", "供应商名称", f"{yp}金额（元）", "占基期%", "验收笔数", "物料数",
                   "平均单笔（元）", "主营类别", "主营占比%", "最后验收日期", "停供天数"]
        keys = ["seq", "name", "amount", "share", "orders", "materials",
                "avg_order", "top_cat", "top_cat_share", "last_date", "idle_days"]

    ws["A1"].font = TITLE_FONT
    c = ws["A2"]
    c.value = sub
    c.font = NOTE_FONT
    c.alignment = Alignment(wrap_text=True, vertical="top")
    ws.row_dimensions[2].height = 30

    # 表头
    hr = 4
    for j, h in enumerate(headers):
        ws.cell(row=hr, column=j + 1, value=h)
    _style_header(ws, hr, len(headers))

    for i, s in enumerate(data, start=1):
        vals = []
        for k in keys:
            vals.append(i if k == "seq" else s.get(k))
        for j, v in enumerate(vals):
            cell = ws.cell(row=hr + i, column=j + 1, value=v)
            cell.font = BODY_FONT
            cell.border = BORDER
            head = headers[j]
            if head.endswith("%"):
                cell.number_format = PCT_FMT
            elif isinstance(v, (int, float)) and not isinstance(v, bool):
                cell.number_format = NUM_FMT
                cell.alignment = Alignment(horizontal="right")

    last = hr + len(data)
    # 合计行
    ws.cell(row=last + 1, column=1, value="合计").font = BOLD_FONT
    tc = ws.cell(row=last + 1, column=3, value=churn["new_amount" if is_new else "lost_amount"])
    tc.font = BOLD_FONT
    tc.number_format = NUM_FMT
    pc = ws.cell(row=last + 1, column=4, value=churn["new_share" if is_new else "lost_share"])
    pc.font = BOLD_FONT
    pc.number_format = PCT_FMT
    oc = ws.cell(row=last + 1, column=5,
                 value=sum(s["orders"] for s in data))
    oc.font = BOLD_FONT
    oc.number_format = NUM_FMT
    for j in range(1, len(headers) + 1):
        ws.cell(row=last + 1, column=j).border = BORDER

    if is_new:
        big = [s for s in data if (s["amount"] or 0) >= 1e6]
        if big:
            _note(ws, last + 3,
                  f"注：其中 {len(big)} 家新增供应商单家金额超 100 万元，"
                  f"合计 {sum(s['amount'] for s in big):,.0f} 元，"
                  f"最大为 {big[0]['name']}（{big[0]['amount']:,.0f} 元）。"
                  f"建议对这部分新供应商补充准入与履约评估。", font=ALERT_FONT)
    else:
        late = [s for s in data if s["last_date"] and s["last_date"][:7] >= f"{yp}-07"]
        if data:
            _note(ws, last + 3,
                  f"注：{len(late)} 家的最后验收日期落在 {yp} 年下半年"
                  f"（上年仍有采购、本期完全停供），需确认是否有替代来源或已终止合作；"
                  f"其余 {len(data) - len(late)} 家在 {yp} 上半年即已停供，多与项目结束或供应商整合有关。"
                  f"全部退出的停供天数中位数为 {churn['lost_idle_median']:.0f} 天。", font=ALERT_FONT)
        _note(ws, last + 4,
              "说明：退出仅表示本期无可比区间的验收记录，不代表合作终止——"
              "长周期设备类供应商可能因无采购需求而自然缺席，需与采购部门确认。")

    ws.freeze_panes = f"A{hr + 1}"
    ws.auto_filter.ref = f"A{hr}:{get_column_letter(len(headers))}{max(last, hr + 1)}"
    _set_widths(ws, {"A": 6, "B": 44, "C": 18, "D": 10, "E": 10, "F": 8,
                     "G": 14, "H": 12, "I": 11, "J": 14, "K": 10})
    return ws


def _sheet_overdue(wb, res):
    meta = res["meta"]
    yc, yp = meta["year_cur"], meta["year_prev"]
    ws = _new_sheet(wb, "05 逾期交付")
    ws["A1"] = "逾期交付分析"
    ws["A1"].font = TITLE_FONT

    df = pd.DataFrame([
        {"逾期区间": d["range"], f"{yp}年占比%": d["prev"], f"{yc}年占比%": d["cur"]}
        for d in res["overdue_distribution"]
    ])
    end = _write_df(ws, df, start_row=3, pct_cols=[f"{yp}年占比%", f"{yc}年占比%"])

    ch = BarChart()
    ch.type = "col"
    ch.title = "逾期天数分布对比"
    ch.y_axis.title = "占比（%）"
    ch.add_data(Reference(ws, min_col=2, max_col=3, min_row=3, max_row=3 + len(df)),
                titles_from_data=True)
    ch.set_categories(Reference(ws, min_col=1, min_row=4, max_row=3 + len(df)))
    ch.height, ch.width = 8, 16
    ws.add_chart(ch, "E3")

    r = end + 2
    ws.cell(row=r, column=1, value=f"逾期最严重供应商 TOP10（{yc}年，按逾期单数）").font = SUB_FONT
    df2 = pd.DataFrame([
        {"供应商名称": s["name"], "逾期笔数": s["count"],
         "平均逾期天数": s["avg_days"], "涉及金额": s["amount"]}
        for s in res["overdue_suppliers"]
    ])
    end2 = _write_df(ws, df2, start_row=r + 1, num_cols=["逾期笔数", "涉及金额"])

    severe = [s for s in res["overdue_suppliers"] if (s["avg_days"] or 0) > 60]
    if severe:
        names = "、".join(s["name"][:14] for s in severe[:3])
        _note(ws, end2 + 2,
              f"注意：{names} 平均逾期超 60 天，属长期异常，建议核查合同到货条款是否失真",
              font=ALERT_FONT)
    _set_widths(ws, {"A": 40, "B": 16, "C": 16, "D": 16})
    return ws


def _sheet_price(wb, res):
    meta = res["meta"]
    yc, yp = meta["year_cur"], meta["year_prev"]
    p = res["price"]
    ws = _new_sheet(wb, "06 价格异动")
    ws["A1"] = f"同物料单价异动分析（两年均有采购且各自数量 ≥ 10）"
    ws["A1"].font = TITLE_FONT

    df = pd.DataFrame({
        "指标": ["可比物料数", "平均涨幅%", "中位涨幅%", "涨幅>20%物料数",
                 f"涨幅>20%涉及金额（{yc}）", "价格下降物料占比%"],
        "数值": [p["comparable"], p["mean_change"], p["median_change"],
                 p["up20_count"], p["up20_amount"], p["down_share"]],
    })
    end = _write_df(ws, df, start_row=3, num_cols=["数值"])

    r = end + 2
    ws.cell(row=r, column=1,
            value=f"涨幅 TOP25（涨幅 > 20%，按 {yc} 年采购金额排序）").font = SUB_FONT
    df2 = pd.DataFrame([
        {"物料编码": u["code"], "物料描述": u["desc"],
         f"单价{yp}": u["prev_price"], f"单价{yc}": u["cur_price"],
         "涨幅%": u["change"], f"{yc}数量": u.get("qty"), "金额": u["amount"]}
        for u in p["top_up"]
    ])
    end2 = _write_df(ws, df2, start_row=r + 1,
                     num_cols=[f"单价{yp}", f"单价{yc}", f"{yc}数量", "金额"], pct_cols=["涨幅%"])
    _note(ws, end2 + 1,
          f"注：本期共 {p['up20_count']:,} 种物料涨幅超 20%，涉及金额 {p['up20_amount']:,.0f} 元；"
          f"上表为其中金额最大的 25 种。")

    r = end2 + 4
    ws.cell(row=r, column=1,
            value=f"降价 TOP25（降幅 ≥ 5%，按 {yc} 年采购金额排序）").font = SUB_FONT
    df4 = pd.DataFrame([
        {"物料编码": u["code"], "物料描述": u["desc"],
         f"单价{yp}": u["prev_price"], f"单价{yc}": u["cur_price"],
         "降幅%": u["change"], f"{yc}数量": u.get("qty"), "金额": u["amount"]}
        for u in p["top_down"]
    ])
    end4 = _write_df(ws, df4, start_row=r + 1,
                     num_cols=[f"单价{yp}", f"单价{yc}", f"{yc}数量", "金额"], pct_cols=["降幅%"])
    if len(df4):
        md = p["max_drop"]
        drop_txt = ("个别品种降幅超过 80%，多为计量单位或规格口径变更，需单独核查"
                    if md is not None and md <= -80 else f"单种最大降幅 {md:.1f}%")
        _note(ws, end4 + 1,
              f"注：本期共 {p['down_count']:,} 种物料单价同比下降，其中降幅达 5% 以上的 "
              f"{p['down5_count']:,} 种、涉及金额 {p['down5_amount']:,.0f} 元，{drop_txt}。"
              f"上表为其中金额最大的 25 种，可作为同类物料议价的横向基准。")

    r = end4 + 4
    ws.cell(row=r, column=1, value="同规格多供应商价差（按涉及金额排序）").font = SUB_FONT
    df3 = pd.DataFrame([
        {"物料描述": g["desc"], "供应商数": g["suppliers"],
         "最低单价": g["min_price"], "最高单价": g["max_price"],
         "价差倍数": g["ratio"], "涉及金额": g["amount"]}
        for g in res["price_gaps"]
    ]) if res["price_gaps"] else pd.DataFrame(columns=["物料描述", "供应商数", "最低单价", "最高单价", "价差倍数", "涉及金额"])
    end3 = _write_df(ws, df3, start_row=r + 1, start_col=2,
                     num_cols=["最低单价", "最高单价", "涉及金额"])

    if res["price_gaps"]:
        g = res["price_gaps"][0]
        _note(ws, end3 + 2,
              f"重点：「{g['desc']}」同规格单价 {g['min_price']:.1f}—{g['max_price']:.1f} 元，"
              f"相差 {g['ratio']:.1f} 倍，是最值得优先议价的品类",
              font=ALERT_FONT)
    _set_widths(ws, {"A": 14, "B": 46, "C": 13, "D": 13, "E": 11, "F": 13, "G": 16})
    return ws


def _sheet_unit(wb, res):
    meta = res["meta"]
    yc, yp = meta["year_cur"], meta["year_prev"]
    ws = _new_sheet(wb, "07 矿点分布")
    ws["A1"] = "各使用单位（矿点）采购分布"
    ws["A1"].font = TITLE_FONT

    df = pd.DataFrame([
        {"使用单位": u["name"], f"{yp}年": u["prev"], f"{yc}年": u["cur"],
         "同比%": u["yoy"] if u["yoy"] is not None else 0, f"{yc}占比%": u["share"]}
        for u in res["units"]
    ])
    end = _write_df(ws, df, start_row=3, num_cols=[f"{yp}年", f"{yc}年"],
                    pct_cols=["同比%", f"{yc}占比%"])

    blank = res.get("blank_unit")
    if blank:
        _note(ws, end + 2,
              f"说明：「使用单位」为空的记录本期金额 {blank['cur']:,.0f} 元，"
              f"占总金额 {blank['share']:.1f}%，主要为设备类资产直供，不宜按矿点分摊。")
    _set_widths(ws, {"A": 26, "B": 16, "C": 16, "D": 14, "E": 14})
    return ws


def _sheet_conclusions(wb, rule_out):
    ws = _new_sheet(wb, "08 分析结论")
    ws["A1"] = "分析结论与建议"
    ws["A1"].font = Font(bold=True, size=16, color=C_PRIMARY, name=FONT_NAME)

    r = 3
    ws.cell(row=r, column=1, value="一、核心结论").font = SUB_FONT
    r += 1
    c = ws.cell(row=r, column=1, value=rule_out["headline"])
    c.font = BOLD_FONT
    c.alignment = Alignment(wrap_text=True, vertical="top")
    ws.row_dimensions[r].height = 32
    r += 2

    ws.cell(row=r, column=1, value="二、分项发现").font = SUB_FONT
    r += 1
    for f in rule_out["findings"]:
        mark = {"danger": "【严重】", "warn": "【关注】", "good": "【良好】", "info": "　　"}[f["level"]]
        head = ws.cell(row=r, column=1, value=f"{mark}{f['title']}")
        head.font = Font(size=10, bold=True, name=FONT_NAME,
                         color={"danger": C_ALERT, "warn": "B36B00"}.get(f["level"], C_PRIMARY))
        r += 1
        for fact in f["facts"]:
            cell = ws.cell(row=r, column=1, value="　· " + fact)
            cell.font = BODY_FONT
            cell.alignment = Alignment(wrap_text=True, vertical="top")
            ws.row_dimensions[r].height = 20 if len(fact) < 55 else 34
            r += 1
        if f.get("implication"):
            cell = ws.cell(row=r, column=1, value="　→ " + f["implication"])
            cell.font = Font(size=10, italic=True, color=C_SUB, name=FONT_NAME)
            cell.alignment = Alignment(wrap_text=True, vertical="top")
            ws.row_dimensions[r].height = 30
            r += 1
        r += 1

    if rule_out["problems"]:
        ws.cell(row=r, column=1, value="三、问题与改进建议").font = SUB_FONT
        r += 1
        for i, pr in enumerate(rule_out["problems"], 1):
            ws.cell(row=r, column=1,
                    value=f"问题{i}：{pr['title']}").font = Font(size=10, bold=True, name=FONT_NAME)
            r += 1
            for label, key in (("现状", "evidence"), ("建议", "action"), ("收益", "benefit")):
                val = pr.get(key)
                if not val:
                    continue
                cell = ws.cell(row=r, column=1, value=f"　{label}：{val}")
                cell.font = BODY_FONT
                cell.alignment = Alignment(wrap_text=True, vertical="top")
                ws.row_dimensions[r].height = 20 if len(val) < 55 else 34
                r += 1
            r += 1

    if rule_out["actions"]:
        ws.cell(row=r, column=1, value="四、下阶段工作重点").font = SUB_FONT
        r += 1
        for i, a in enumerate(rule_out["actions"], 1):
            cell = ws.cell(row=r, column=1, value=f"{i}. {a}")
            cell.font = BODY_FONT
            cell.alignment = Alignment(wrap_text=True, vertical="top")
            ws.row_dimensions[r].height = 20 if len(a) < 55 else 34
            r += 1

    ws.column_dimensions["A"].width = 130
    return ws


def _sheet_quality(wb, res, quality):
    meta = res["meta"]
    ws = _new_sheet(wb, "09 数据说明")
    ws["A1"] = "数据口径与质量说明"
    ws["A1"].font = TITLE_FONT

    fq = quality["字段缺失"]
    notes = [
        ("数据源", " / ".join(res.get("_files", ["—"]))),
        ("总记录数", f"{quality['总行数']:,} 行 × 29 个字段"),
        ("时间范围", f"{quality['业务日期范围'][0]} 至 {quality['业务日期范围'][1]}"),
        ("年份分布", "；".join(f"{y} 年 {n:,} 行" for y, n in quality["年份分布"].items())),
        ("同比口径", meta["compare_note"]),
        ("日期取值", "主日期优先取「编制验收单时间」，缺失时回退「验收申请时间」；原字段为 Excel 序列号，已转换"),
        ("金额口径", "「验收金额」为含税验收金额，单位：元"),
        ("逾期口径", "「逾期天数」> 0 判定为逾期，取实际到货日期与合同到货日期之差；已剔除极端异常值"),
        ("缺失情况", "；".join(
            f"{k} 缺失 {v['缺失行数']} 行（{v['缺失率%']}%）"
            for k, v in fq.items() if v["缺失行数"] > 0) or "核心字段零缺失"),
        ("使用单位空白", f"{quality['使用单位空白']['行数']} 行，"
                       f"占金额 {quality['使用单位空白']['金额占比%']}%；"
                       f"类别构成：{'、'.join(f'{k} {v}%' for k, v in quality['使用单位空白']['类别构成'].items())}"),
        ("价格分析", "仅纳入两年均有采购且各自数量 ≥ 10 的物料，确保可比；单价为加权均价（金额/数量）"),
        ("榜单口径", "涨幅榜：涨幅 > 20% 按本期金额排序前 25；降价榜：降幅 ≥ 5% 按本期金额排序前 25；"
                    "供应商榜：按本期金额排序前 25。榜单均按金额而非变化幅度排序，目的是让「金额影响」优先于「极端值」"),
        ("已知局限", "①验收金额为验收口径，非合同签订口径，反映实际到货而非采购决策；"
                    "②物料描述未做规格标准化，同物异码可能影响价格对比；"
                    "③最后一期数据不完整月份已从同比区间剔除"),
    ]
    r = 3
    for k, v in notes:
        ws.cell(row=r, column=1, value=k).font = BOLD_FONT
        cell = ws.cell(row=r, column=2, value=v)
        cell.font = BODY_FONT
        cell.alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[r].height = 34
        r += 1
    _set_widths(ws, {"A": 16, "B": 110})
    return ws


# ---------------- 入口 ----------------
def write_excel(res, rule_out, quality, out_path, title=None, files=None):
    """
    生成 9 工作表分析看板。

    参数
      res      : analyzer.analyze() 的返回值（含 _frames）
      rule_out : rules.build() 的返回值
      quality  : loader.quality_report() 的返回值
      out_path : 输出 .xlsx 路径
      title    : 报告标题（可选）
      files    : 数据源文件名列表（可选）
    """
    res = dict(res)
    if title:
        res["_title"] = title
    if files:
        res["_files"] = [os.path.basename(f) for f in files]

    wb = Workbook()
    _sheet_overview(wb, res, rule_out)
    _sheet_monthly(wb, res)
    _sheet_category(wb, res)
    _sheet_supplier(wb, res)
    _sheet_supplier_churn(wb, res, "new")
    _sheet_supplier_churn(wb, res, "lost")
    _sheet_overdue(wb, res)
    _sheet_price(wb, res)
    _sheet_unit(wb, res)
    _sheet_conclusions(wb, rule_out)
    _sheet_quality(wb, res, quality)

    os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
    try:
        wb.save(out_path)
    except PermissionError:
        raise PermissionError(
            f"无法写入 {out_path}\n"
            f"该文件很可能正被 Excel / WPS 打开。请关闭后重试，"
            f"或用 --out 指定另一个输出目录。"
        ) from None
    return out_path
