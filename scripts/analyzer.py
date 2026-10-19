# -*- coding: utf-8 -*-
"""
采购验收分析引擎

输入：loader.load_dataset() 产出的 DataFrame
输出：analysis dict，结构稳定，供 rules / excel / docx 三个渲染器消费。

同比口径（用户确认）：自动取两期重叠的月份区间（如 2026年1-10月 vs 2025年1-10月），
自动削平到可比区间，避免 2026 只到 10 月时直接全年对比造成低估。
"""
import pandas as pd
import numpy as np
from collections import OrderedDict

OVERDUE_SENTINEL = -3000  # 小于此值的逾期天数视为脏数据


def _pct(new, old):
    """同比百分比，old 为 0 时返回 None"""
    if old is None or pd.isna(old) or old == 0:
        return None
    return round((new - old) / old * 100, 2)


def _r(v, n=2):
    if v is None or (isinstance(v, float) and np.isnan(v)):
        return None
    return round(float(v), n)


def _supplier_profiles(frame, names, amt_total, anchor=None):
    """
    按供应商聚合画像，供「新增 / 退出供应商名单」使用。

    frame   : 该年度的可比区间数据（新增用 cur，退出用 prev）
    names   : 需要出名单的供应商名称集合
    anchor  : 退出名单传入本期数据截止日，用于计算「停供天数」
    """
    if not names or frame is None or frame.empty:
        return []
    d = frame[frame["供应商名称"].isin(names)]
    out = []
    for name, g in d.groupby("供应商名称"):
        cats = g.groupby("物资类别")["验收金额"].sum().sort_values(ascending=False)
        amt = float(g["验收金额"].sum())
        last, first = g["业务日期"].max(), g["业务日期"].min()
        item = {
            "name": str(name),
            "amount": _r(amt, 0),
            "share": round(amt / amt_total * 100, 2) if amt_total else 0,
            "orders": int(len(g)),
            "materials": int(g["物料编码"].nunique()),
            "avg_order": _r(amt / len(g), 0) if len(g) else None,
            "top_cat": str(cats.index[0]) if len(cats) else "",
            "top_cat_share": round(cats.iloc[0] / cats.sum() * 100, 1) if cats.sum() else 0,
            "first_date": str(first.date()) if pd.notna(first) else "",
            "last_date": str(last.date()) if pd.notna(last) else "",
        }
        if anchor is not None and pd.notna(last):
            item["idle_days"] = int((anchor - last).days)
        out.append(item)
    out.sort(key=lambda x: -(x["amount"] or 0))
    return out


def determine_compare_months(df: pd.DataFrame):
    """
    决定对比年份与月份区间。
    返回 (year_cur, year_prev, months, note)
    """
    years = sorted(int(y) for y in df["数据年份"].unique())
    if len(years) < 2:
        raise ValueError(f"需要至少两年数据才能做同比，当前只有: {years}")

    year_prev, year_cur = years[-2], years[-1]

    cur = df[df["数据年份"] == year_cur]["业务日期"].dropna()
    prev = df[df["数据年份"] == year_prev]["业务日期"].dropna()

    cur_max = cur.max()
    # 当前年份的自然年末（12/31）与实际数据截止
    complete_to_month = cur_max.month
    # 若当月数据极不完整（如仅 1 日的数据，金额占比异常低），回退一个月
    month_df = df[(df["数据年份"] == year_cur) & (df["月序"] == complete_to_month)]
    month_share = month_df["验收金额"].sum() / df[df["数据年份"] == year_cur]["验收金额"].sum() \
        if df[df["数据年份"] == year_cur]["验收金额"].sum() else 0
    note = ""
    if month_share < 0.02 and complete_to_month > 1:
        note = (f"{year_cur}年{complete_to_month}月数据仅覆盖至 {cur_max.day} 日，"
                f"占比 {month_share*100:.2f}%，已从同比区间剔除")
        complete_to_month -= 1

    # 基准年份必须覆盖到该月，否则向下取整
    prev_max_month = prev.max().month if len(prev) else 0
    months = list(range(1, min(complete_to_month, prev_max_month) + 1))

    if not note:
        note = f"{year_cur} 年数据覆盖至 {cur_max.strftime('%Y-%m-%d')}，同比统一取 1—{months[-1]} 月可比区间"

    return year_cur, year_prev, months, note


def analyze(df: pd.DataFrame) -> dict:
    res = OrderedDict()
    year_cur, year_prev, months, cmp_note = determine_compare_months(df)

    res["meta"] = {
        "year_cur": year_cur,
        "year_prev": year_prev,
        "months": months,
        "month_range": f"1—{months[-1]}月",
        "compare_note": cmp_note,
        "total_rows": int(len(df)),
        "data_range": [
            str(df["业务日期"].min().date()),
            str(df["业务日期"].max().date()),
        ],
    }

    cmp = df[(df["月序"].isin(months)) & (df["数据年份"].isin([year_cur, year_prev]))].copy()
    cur = cmp[cmp["数据年份"] == year_cur]
    prev = cmp[cmp["数据年份"] == year_prev]

    # ================= 1. 总量 =================
    amt_cur, amt_prev = cur["验收金额"].sum(), prev["验收金额"].sum()
    res["totals"] = {
        "amount_cur": _r(amt_cur, 0),
        "amount_prev": _r(amt_prev, 0),
        "amount_yoy": _pct(amt_cur, amt_prev),
        "amount_delta": _r(amt_cur - amt_prev, 0),
        "orders_cur": int(len(cur)),
        "orders_prev": int(len(prev)),
        "orders_yoy": _pct(len(cur), len(prev)),
        "avg_order_cur": _r(amt_cur / len(cur)) if len(cur) else None,
        "avg_order_prev": _r(amt_prev / len(prev)) if len(prev) else None,
        "avg_order_yoy": _pct(amt_cur / len(cur) if len(cur) else 0,
                              amt_prev / len(prev) if len(prev) else 0),
        "suppliers_cur": int(cur["供应商名称"].nunique()),
        "suppliers_prev": int(prev["供应商名称"].nunique()),
        "materials_cur": int(cur["物料编码"].nunique()),
        "materials_prev": int(prev["物料编码"].nunique()),
        "annual_forecast": _r(amt_cur / len(months) * 12, 0) if months else None,
        "full_year_prev": _r(df[df["数据年份"] == year_prev]["验收金额"].sum(), 0),
    }

    # ================= 2. 月度 =================
    mon = cmp.pivot_table(index="月序", columns="数据年份", values="验收金额",
                          aggfunc="sum").fillna(0)
    monthly = []
    for m in months:
        p = float(mon.loc[m, year_prev]) if m in mon.index and year_prev in mon.columns else 0
        c = float(mon.loc[m, year_cur]) if m in mon.index and year_cur in mon.columns else 0
        monthly.append({"month": f"{m}月", "prev": _r(p, 0), "cur": _r(c, 0), "yoy": _pct(c, p)})
    res["monthly"] = monthly

    yoys = [x["yoy"] for x in monthly if x["yoy"] is not None]
    res["monthly_summary"] = {
        "max_yoy": max(yoys) if yoys else None,
        "min_yoy": min(yoys) if yoys else None,
        "volatile": (max(yoys) - min(yoys)) > 100 if yoys else False,
    }

    # ================= 3. 物资类别 =================
    catp = cmp.pivot_table(index="物资类别", columns="数据年份", values="验收金额",
                           aggfunc="sum").fillna(0)
    cats = []
    total_cur_cat = catp[year_cur].sum() if year_cur in catp.columns else 0
    for name in catp.index:
        c = float(catp.loc[name, year_cur]) if year_cur in catp.columns else 0
        p = float(catp.loc[name, year_prev]) if year_prev in catp.columns else 0
        cats.append({
            "name": name, "prev": _r(p, 0), "cur": _r(c, 0), "yoy": _pct(c, p),
            "share": round(c / total_cur_cat * 100, 1) if total_cur_cat else 0,
        })
    cats.sort(key=lambda x: -(x["cur"] or 0))
    res["categories"] = cats

    # ================= 4. 供应商 =================
    sup = {}
    for yr in (year_cur, year_prev):
        d = cmp[cmp["数据年份"] == yr]
        s = d.groupby("供应商名称")["验收金额"].sum().sort_values(ascending=False)
        tot = s.sum()
        sup[yr] = {
            "count": int(len(s)),
            "cr1": round(s.iloc[0] / tot * 100, 1) if tot and len(s) else 0,
            "cr5": round(s.iloc[:5].sum() / tot * 100, 1) if tot else 0,
            "cr10": round(s.iloc[:10].sum() / tot * 100, 1) if tot else 0,
            "cr20": round(s.iloc[:20].sum() / tot * 100, 1) if tot else 0,
            "series": s,
        }
    res["supplier_concentration"] = {
        "cur": {k: v for k, v in sup[year_cur].items() if k != "series"},
        "prev": {k: v for k, v in sup[year_prev].items() if k != "series"},
    }

    sc, sp = sup[year_cur]["series"], sup[year_prev]["series"]
    top = pd.DataFrame({"prev": sp, "cur": sc}).fillna(0)
    top["yoy"] = ((top["cur"] - top["prev"]) / top["prev"].replace(0, np.nan) * 100).round(1)
    top["share"] = (top["cur"] / top["cur"].sum() * 100).round(1) if top["cur"].sum() else 0
    top = top.sort_values("cur", ascending=False)
    res["top_suppliers"] = [
        {"name": str(n), "prev": _r(r["prev"], 0), "cur": _r(r["cur"], 0),
         "yoy": _r(r["yoy"], 1), "share": _r(r["share"], 1)}
        for n, r in top.head(25).iterrows()
    ]
    res["supplier_churn"] = {
        "both": int(((top["prev"] > 0) & (top["cur"] > 0)).sum()),
        "new": int(((top["prev"] == 0) & (top["cur"] > 0)).sum()),
        "lost": int(((top["prev"] > 0) & (top["cur"] == 0)).sum()),
        "total": int(len(top)),
    }

    # ---------- 新增 / 退出供应商完整名单（Excel 单独成表） ----------
    new_names = [str(n) for n in top[(top["prev"] == 0) & (top["cur"] > 0)].index]
    lost_names = [str(n) for n in top[(top["prev"] > 0) & (top["cur"] == 0)].index]
    cur_end = cur["业务日期"].max()
    res["supplier_new"] = _supplier_profiles(cur, new_names, amt_cur)
    res["supplier_lost"] = _supplier_profiles(prev, lost_names, amt_prev, anchor=cur_end)

    new_amt = sum(s["amount"] or 0 for s in res["supplier_new"])
    lost_amt = sum(s["amount"] or 0 for s in res["supplier_lost"])
    res["supplier_churn"].update({
        "new_amount": _r(new_amt, 0),
        "new_share": round(new_amt / amt_cur * 100, 1) if amt_cur else 0,
        "new_median": _r(pd.Series([s["amount"] or 0 for s in res["supplier_new"]]).median(), 0)
        if res["supplier_new"] else None,
        "lost_amount": _r(lost_amt, 0),
        "lost_share": round(lost_amt / amt_prev * 100, 1) if amt_prev else 0,
        "lost_median": _r(pd.Series([s["amount"] or 0 for s in res["supplier_lost"]]).median(), 0)
        if res["supplier_lost"] else None,
        "lost_idle_median": _r(pd.Series([s.get("idle_days") or 0 for s in res["supplier_lost"]]).median(), 0)
        if res["supplier_lost"] else None,
    })

    # 各品类集中度
    cat_conc = []
    for cat in ["材料", "备件", "移动设备", "固定设备"]:
        dc = cur[cur["物资类别"] == cat]
        if len(dc) == 0:
            continue
        s = dc.groupby("供应商名称")["验收金额"].sum().sort_values(ascending=False)
        tot = s.sum()
        cat_conc.append({
            "name": cat, "amount": _r(tot, 0), "suppliers": int(len(s)),
            "cr3": round(s.iloc[:3].sum() / tot * 100, 1) if tot else 0,
            "top1_name": str(s.index[0]) if len(s) else "",
            "top1_share": round(s.iloc[0] / tot * 100, 1) if tot else 0,
        })
    res["category_concentration"] = cat_conc

    # 大额订单
    big = cur[cur["验收金额"] > 1e6]
    res["big_orders"] = {
        "count": int(len(big)),
        "amount": _r(big["验收金额"].sum(), 0),
        "share": round(big["验收金额"].sum() / amt_cur * 100, 1) if amt_cur else 0,
        "top": [
            {"name": str(n), "count": int(r["size"]), "amount": _r(r["sum"], 0)}
            for n, r in big.groupby("供应商名称")["验收金额"].agg(["size", "sum"])
            .sort_values("sum", ascending=False).head(8).iterrows()
        ],
    }

    # ================= 5. 逾期 =================
    ov = cmp[cmp["逾期天数"].notna() & (cmp["逾期天数"] > OVERDUE_SENTINEL)].copy()
    overdue = {}
    for yr in (year_cur, year_prev):
        d = ov[ov["数据年份"] == yr]
        pos = d[d["逾期天数"] > 0]
        overdue[yr] = {
            "valid": int(len(d)),
            "overdue_count": int(len(pos)),
            "rate": round(len(pos) / len(d) * 100, 1) if len(d) else 0,
            "avg_days": _r(pos["逾期天数"].mean(), 1) if len(pos) else 0,
            "amount": _r(pos["验收金额"].sum(), 0),
            "severe_share": round((pos["逾期天数"] > 60).mean() * 100, 1) if len(pos) else 0,
        }
    res["overdue"] = {"cur": overdue[year_cur], "prev": overdue[year_prev]}

    bins = [-1, 0, 7, 15, 30, 60, 1e9]
    labels = ["未逾期", "1-7天", "8-15天", "16-30天", "31-60天", "60天以上"]
    ov["逾期区间"] = pd.cut(ov["逾期天数"], bins=bins, labels=labels)
    dist = pd.crosstab(ov["逾期区间"], ov["数据年份"], normalize="columns") * 100
    res["overdue_distribution"] = [
        {"range": str(lbl),
         "prev": _r(dist.loc[lbl, year_prev], 1) if lbl in dist.index and year_prev in dist.columns else 0,
         "cur": _r(dist.loc[lbl, year_cur], 1) if lbl in dist.index and year_cur in dist.columns else 0}
        for lbl in labels
    ]

    d_pos = ov[(ov["数据年份"] == year_cur) & (ov["逾期天数"] > 0)]
    sup_ov = d_pos.groupby("供应商名称").agg(
        逾期笔数=("逾期天数", "size"), 平均逾期=("逾期天数", "mean"), 涉及金额=("验收金额", "sum")
    ).sort_values("逾期笔数", ascending=False)
    res["overdue_suppliers"] = [
        {"name": str(n), "count": int(r["逾期笔数"]), "avg_days": _r(r["平均逾期"], 1),
         "amount": _r(r["涉及金额"], 0)}
        for n, r in sup_ov.head(10).iterrows()
    ]

    # 验收效率
    eff = cmp.copy()
    eff["间隔"] = (eff["业务日期"] - eff["验收申请日期"]).dt.days
    res["accept_efficiency"] = {
        "cur_avg": _r(eff[eff["数据年份"] == year_cur]["间隔"].mean(), 1),
        "prev_avg": _r(eff[eff["数据年份"] == year_prev]["间隔"].mean(), 1),
        "max": _r(eff[eff["数据年份"] == year_cur]["间隔"].max(), 0),
    }

    # ================= 6. 价格异动 =================
    p = cmp.groupby(["物料编码", "数据年份"]).agg(
        金额=("验收金额", "sum"), 数量=("验收数量", "sum")).reset_index()
    p = p[p["数量"] > 0]
    p["单价"] = p["金额"] / p["数量"]
    w = p.pivot(index="物料编码", columns="数据年份", values=["单价", "数量", "金额"])
    w.columns = [f"{a}{b}" for a, b in w.columns]
    w = w.dropna(subset=[f"单价{year_prev}", f"单价{year_cur}"])
    w = w[(w[f"数量{year_prev}"] >= 10) & (w[f"数量{year_cur}"] >= 10)]
    w["涨幅"] = ((w[f"单价{year_cur}"] - w[f"单价{year_prev}"]) / w[f"单价{year_prev}"] * 100)
    name_map = cmp.drop_duplicates("物料编码").set_index("物料编码")["物料描述"]
    w["物料描述"] = name_map.reindex(w.index)

    def _price_rows(frame, n=25):
        return [
            {"code": str(i), "desc": str(r["物料描述"])[:40],
             "prev_price": _r(r[f"单价{year_prev}"]), "cur_price": _r(r[f"单价{year_cur}"]),
             "change": _r(r["涨幅"], 1), "amount": _r(r[f"金额{year_cur}"], 0),
             "qty": _r(r[f"数量{year_cur}"], 0)}
            for i, r in frame.head(n).iterrows()
        ]

    up = w[w["涨幅"] > 20].sort_values(f"金额{year_cur}", ascending=False)
    down_all = w[w["涨幅"] < 0]
    # 降幅 ≥ 5% 才算「显著下降」：0.1% 级别的变动属正常议价波动，纳入会稀释榜单
    down = down_all[down_all["涨幅"] <= -5].sort_values(f"金额{year_cur}", ascending=False)
    res["price"] = {
        "comparable": int(len(w)),
        "mean_change": _r(w["涨幅"].mean(), 1) if len(w) else None,
        "median_change": _r(w["涨幅"].median(), 1) if len(w) else None,
        "down_share": round((w["涨幅"] < 0).mean() * 100, 1) if len(w) else 0,
        "up20_count": int(len(up)),
        "up20_amount": _r(up[f"金额{year_cur}"].sum(), 0),
        "down_count": int(len(down_all)),
        "down5_count": int(len(down)),
        "down5_amount": _r(down[f"金额{year_cur}"].sum(), 0),
        "down_amount": _r(down_all[f"金额{year_cur}"].sum(), 0),
        "max_drop": _r(w["涨幅"].min(), 1) if len(w) else None,
        "top_up": _price_rows(up),
        "top_down": _price_rows(down),
    }

    # 同规格多供应商价差（成本风险的核心证据）
    c2 = cur.copy()
    c2["单价"] = c2["验收金额"] / c2["验收数量"].replace(0, np.nan)
    spec = c2.groupby(["物料描述", "供应商名称"]).agg(
        均价=("单价", "mean"), 数量=("验收数量", "sum"), 金额=("验收金额", "sum")).reset_index()
    spec = spec[(spec["数量"] >= 20) & spec["均价"].notna()]
    gaps = []
    for desc, g in spec.groupby("物料描述"):
        if g["供应商名称"].nunique() < 2:
            continue
        lo, hi = g["均价"].min(), g["均价"].max()
        if lo > 0 and hi / lo >= 1.5:
            gaps.append({
                "desc": str(desc)[:40], "suppliers": int(g["供应商名称"].nunique()),
                "min_price": _r(lo), "max_price": _r(hi),
                "ratio": _r(hi / lo, 1), "amount": _r(g["金额"].sum(), 0),
            })
    gaps.sort(key=lambda x: -(x["amount"] or 0))
    res["price_gaps"] = gaps[:8]

    # 品类级价格异动识别（如截齿）
    cur_c = cur.copy()
    cur_c["关键词"] = cur_c["物料描述"].str.slice(0, 4)
    kw = cur_c.groupby("关键词").agg(金额=("验收金额", "sum"), 笔数=("验收金额", "size"))
    kw_prev = prev.copy()
    kw_prev["关键词"] = kw_prev["物料描述"].str.slice(0, 4)
    kw_p = kw_prev.groupby("关键词")["验收金额"].agg(["sum", "size"])
    kw = kw.join(kw_p, rsuffix="_prev")
    kw = kw[(kw["笔数"] >= 20) & (kw["sum"] >= 5e6)]
    kw["yoy"] = (kw["金额"] - kw["sum"]) / kw["sum"].replace(0, np.nan) * 100
    hot = kw[kw["yoy"] > 50].sort_values("金额", ascending=False).head(5)
    res["hot_keywords"] = [
        {"keyword": str(i), "cur": _r(r["金额"], 0), "prev": _r(r["sum"], 0),
         "yoy": _r(r["yoy"], 1)} for i, r in hot.iterrows()
    ]

    # ================= 7. 使用单位 =================
    unitp = cmp.pivot_table(index="使用单位", columns="数据年份", values="验收金额",
                            aggfunc="sum").fillna(0)
    units = []
    for name in unitp.index:
        c = float(unitp.loc[name, year_cur]) if year_cur in unitp.columns else 0
        pv = float(unitp.loc[name, year_prev]) if year_prev in unitp.columns else 0
        units.append({
            "name": name if name else "（未填/设备直供）",
            "prev": _r(pv, 0), "cur": _r(c, 0), "yoy": _pct(c, pv),
            "share": round(c / amt_cur * 100, 1) if amt_cur else 0,
        })
    units.sort(key=lambda x: -(x["cur"] or 0))
    res["units"] = units[:15]
    res["blank_unit"] = next((u for u in units if u["name"].startswith("（未填")), None)

    # ================= 8. 长尾与小额 =================
    small = cmp[cmp["验收金额"] < 1000]
    res["tail"] = {
        "small_count_cur": int(len(small[small["数据年份"] == year_cur])),
        "small_share_cur": round(len(small[small["数据年份"] == year_cur]) / len(cur) * 100, 1)
        if len(cur) else 0,
        "small_amount_share": round(
            small[small["数据年份"] == year_cur]["验收金额"].sum() / amt_cur * 100, 1
        ) if amt_cur else 0,
        "once_material": int((cur.groupby("物料编码").size() == 1).sum()),
        "once_share": round((cur.groupby("物料编码").size() == 1).sum() /
                            cur["物料编码"].nunique() * 100, 1) if cur["物料编码"].nunique() else 0,
    }

    # 协议类型
    ag = cmp.pivot_table(index="协议类型", columns="数据年份", values="验收金额", aggfunc="sum")
    res["agreement"] = [
        {"name": str(i),
         "prev": _r(ag.loc[i, year_prev], 0) if year_prev in ag.columns else 0,
         "cur": _r(ag.loc[i, year_cur], 0) if year_cur in ag.columns else 0}
        for i in ag.index
    ]

    res["_frames"] = {"cmp": cmp, "cur": cur, "prev": prev, "price_table": w}
    return res


def strip_frames(res: dict) -> dict:
    """移除 DataFrame，便于 JSON 序列化"""
    return {k: v for k, v in res.items() if not k.startswith("_")}
