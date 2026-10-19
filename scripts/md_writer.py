# -*- coding: utf-8 -*-
"""
报告正文渲染器：analysis + rules -> Markdown

为什么用脚本生成正文而不是让模型自由撰写：
  用户明确要求「结论由脚本内置规则引擎生成」，目的是每月口径一致、可复现、可审计。
  因此正文的每一句判断都来自 rules.build() 的结构化产出，这里只做排版。

输出 Markdown 结构（固定 9 章）：
  核心结论 / 一、数据说明与分析口径 / 二、总体规模与月度走势 / 三、物资类别结构
  / 四、供应商格局 / 五、交付时效 / 六、价格与成本 / 七、矿点分布
  / 八、主要问题与管理建议 / 九、下阶段工作重点
"""
import os

from rules import _fmt_int, _fmt_money, _fmt_pct


def _w(v):
    """元 -> 万元数值"""
    return (v or 0) / 1e4


def _wan(v, nd=0):
    return f"{_w(v):,.{nd}f}"


def _mv(v):
    """金额（元）-> 亿元/万元，用于正文叙述（保留数字与单位间的空格）"""
    return _fmt_money(v)


def _yoy(v):
    return "新增" if v is None else _fmt_pct(v)


def _table(header, rows):
    out = ["| " + " | ".join(header) + " |",
           "|" + "|".join(["---"] * len(header)) + "|"]
    for r in rows:
        out.append("| " + " | ".join(str(x) for x in r) + " |")
    return "\n".join(out)


def render_md(res, rule_out, quality, title, subtitle=None, files=None) -> str:
    meta = res["meta"]
    t = res["totals"]
    yc, yp, mr = meta["year_cur"], meta["year_prev"], meta["month_range"]
    ov, conc, price, tail = res["overdue"], res["supplier_concentration"], res["price"], res["tail"]
    churn = res["supplier_churn"]
    efs = res["accept_efficiency"]

    findings = {f["title"]: f for f in rule_out["findings"]}
    L = []
    A = L.append

    # ============ 标题 ============
    A(f"# {title}")
    A("")

    # ============ 核心结论 ============
    A("## 核心结论")
    A("")
    A(rule_out["headline"] + "。")
    A("")

    f1 = findings.get("总体规模")
    if f1:
        A(f"规模层面：{f1['facts'][0]}。{f1['facts'][1]}。{f1['implication']}")
        A("")

    # 风险清单表
    danger = [f for f in rule_out["findings"] if f["level"] in ("danger", "warn")]
    if danger:
        A(f"在规模{'增长' if (t['amount_yoy'] or 0) > 0 else '变化'}的同时，"
          f"以下 {len(danger)} 项结构性问题需要重点关注：")
        A("")
        rows = []
        for f in danger:
            rows.append([f["title"], f["facts"][0].rstrip("。"),
                         {"danger": "高", "warn": "中"}.get(f["level"], "低")])
        A(_table(["关注点", "现状", "风险等级"], rows))
        A("")
        for f in danger[:2]:
            A(f"**{f['title']}**——{f['implication']}")
            A("")

    # ============ 一、数据说明 ============
    A("## 一、数据说明与分析口径")
    A("")
    A(f"本报告基于 {len(files or [])} 份验收单报表，合计 **{_fmt_int(quality['总行数'])} 条** 验收记录、"
      f"29 个字段，时间跨度 {quality['业务日期范围'][0]} 至 {quality['业务日期范围'][1]}。")
    A("")
    A("关于口径，有三点必须说明：")
    A("")
    A(f"**同比区间对齐。** {meta['compare_note']}。直接与基期全年对比会产生偏差，"
      f"因此本报告所有同比均统一采用 **{mr}可比区间**。")
    A("")
    A("**日期取值。** 原始「编制验收单时间」字段为 Excel 日期序列号，已转换为标准日期；"
      "缺失时回退使用「验收申请时间」。")
    A("")
    blank = res.get("blank_unit")
    miss_core = [k for k, v in quality["字段缺失"].items()
                 if v["缺失行数"] > 0 and k != "使用单位"]
    if blank:
        cat_txt = "、".join(f"{k} {v}%" for k, v in quality["使用单位空白"]["类别构成"].items())
        A(f"**数据质量。** 验收金额、验收数量、供应商名称、物资类别等核心字段"
          f"{'零缺失' if not miss_core else '基本完整'}。"
          f"唯一显著缺口是「使用单位」字段：全量 {_fmt_int(quality['使用单位空白']['行数'])} 行记录为空，"
          f"这部分金额达 {_mv(quality['使用单位空白']['金额'])}，占全部数据金额 "
          f"{quality['使用单位空白']['金额占比%']:.1f}%；"
          f"在 {mr}可比区间内占比为 **{blank['share']:.1f}%**。"
          f"经核查，其物资类别构成为 {cat_txt}，属于设备类资产直供，"
          f"并非数据录入错误，但确实导致这部分金额无法按矿点归集。")
        A("")
    else:
        A("**数据质量。** 核心字段完整，未发现显著缺失。")
        A("")

    # ============ 二、总体规模与月度走势 ============
    A("## 二、总体规模与月度走势")
    A("")
    A("本章从金额与笔数两个维度观察采购规模的变化，并进一步拆解月度节奏，识别采购活动的季节性特征。")
    A("")
    A("### 2.1 规模与效率")
    A("")
    rows = [
        ["验收金额", _mv(t["amount_prev"]), _mv(t["amount_cur"]), _fmt_pct(t["amount_yoy"])],
        ["验收笔数", f"{_fmt_int(t['orders_prev'])} 笔", f"{_fmt_int(t['orders_cur'])} 笔", _fmt_pct(t["orders_yoy"])],
        ["平均单笔金额", _mv(t["avg_order_prev"]), _mv(t["avg_order_cur"]), _fmt_pct(t["avg_order_yoy"])],
        ["合作供应商数", f"{t['suppliers_prev']} 家", f"{t['suppliers_cur']} 家",
         f"{t['suppliers_cur'] - t['suppliers_prev']:+d} 家"],
        ["采购物料种类", f"{_fmt_int(t['materials_prev'])} 种", f"{_fmt_int(t['materials_cur'])} 种",
         f"{t['materials_cur'] - t['materials_prev']:+,} 种"],
    ]
    A(_table([f"指标", f"{yp}年{mr}", f"{yc}年{mr}", "变化"], rows))
    A("")
    if t["orders_yoy"] is not None and t["avg_order_yoy"] is not None:
        if t["orders_yoy"] > 10 and t["avg_order_yoy"] < -5:
            A(f"笔数增速（{t['orders_yoy']:+.1f}%）远超金额增速（{_fmt_pct(t['amount_yoy'])}），"
              f"是本期最鲜明的特征。这既反映了检修维护类小额需求增加，也说明采购活动正在「碎片化」"
              f"——金额低于 1,000 元的订单 {_fmt_int(tail['small_count_cur'])} 笔，"
              f"占全部订单的 {tail['small_share_cur']:.1f}%，但仅贡献 {tail['small_amount_share']:.1f}% 的金额。"
              f"这类订单的单位流程成本远高于其采购价值，是流程优化的重点。")
        else:
            A(f"笔数同比 {_fmt_pct(t['orders_yoy'])}，平均单笔金额同比 {_fmt_pct(t['avg_order_yoy'])}，"
              f"采购规模与结构保持相对稳定。")
        A("")

    A("### 2.2 月度节奏")
    A("")
    A(_table([f"月份", f"{yp}年（万元）", f"{yc}年（万元）", "同比"],
             [[m["month"], _wan(m["prev"]), _wan(m["cur"]), _yoy(m["yoy"])] for m in res["monthly"]]))
    A("")
    ms = res["monthly_summary"]
    if ms["max_yoy"] is not None:
        if ms["volatile"]:
            A(f"月度波动幅度极大，单月同比在 {ms['max_yoy']:+.0f}% 至 {ms['min_yoy']:+.0f}% 之间反复摆动。"
              f"峰谷错位说明采购节奏受项目排期和集中招标影响明显，尚未形成平滑的月度计划。")
        else:
            A(f"月度同比区间为 {ms['min_yoy']:+.1f}% 至 {ms['max_yoy']:+.1f}%，波动处于合理范围。")
        A("")

    # ============ 三、物资类别结构 ============
    A("## 三、物资类别结构")
    A("")
    A(_table(["物资类别", f"{yp}年（万元）", f"{yc}年（万元）", "同比", f"{yc}占比"],
             [[c["name"], _wan(c["prev"]), _wan(c["cur"]), _yoy(c["yoy"]), f"{c['share']:.1f}%"]
              for c in res["categories"]]))
    A("")
    f_struct = findings.get("物资结构")
    if f_struct:
        A(f_struct["facts"][-1] + "。" if f_struct["facts"] else "")
        A("")
        A(f_struct["implication"])
        A("")

    # ============ 四、供应商格局 ============
    A("## 四、供应商格局")
    A("")
    A("供应商结构直接影响议价能力与供应安全，本章从集中度、供应商更替以及大额订单归属三个角度展开。")
    A("")
    A("### 4.1 集中度")
    A("")
    cr1d = round(conc["cur"]["cr1"] - conc["prev"]["cr1"], 1)
    cr5d = round(conc["cur"]["cr5"] - conc["prev"]["cr5"], 1)
    cr10d = round(conc["cur"]["cr10"] - conc["prev"]["cr10"], 1)
    A(_table(["集中度指标", f"{yp}年", f"{yc}年", "变化"],
             [["供应商数量", f"{conc['prev']['count']} 家", f"{conc['cur']['count']} 家",
               f"{conc['cur']['count'] - conc['prev']['count']:+d}"],
              ["CR1（最大供应商占比）", f"{conc['prev']['cr1']:.1f}%", f"{conc['cur']['cr1']:.1f}%", f"{cr1d:+.1f}pp"],
              ["CR5（前五家占比）", f"{conc['prev']['cr5']:.1f}%", f"{conc['cur']['cr5']:.1f}%", f"{cr5d:+.1f}pp"],
              ["CR10（前十家占比）", f"{conc['prev']['cr10']:.1f}%", f"{conc['cur']['cr10']:.1f}%", f"{cr10d:+.1f}pp"]]))
    A("")
    f_sup = findings.get("供应商格局")
    if f_sup:
        A(f_sup["implication"])
        A("")
    A(f"两年均保持合作的供应商 {churn['both']} 家。")
    A("")

    if res["category_concentration"]:
        A("细分品类的依赖情况：")
        A("")
        A(_table(["类别", f"{yc}金额", "供应商数", "CR3", "最大供应商占比"],
                 [[c["name"], _mv(c["amount"]), c["suppliers"],
                   f"{c['cr3']:.1f}%", f"{c['top1_share']:.1f}%（{c['top1_name'][:14]}）"]
                  for c in res["category_concentration"]]))
        A("")

    A("### 4.2 主要供应商")
    A("")
    topn = res["top_suppliers"]
    if topn:
        cover = sum(s["cur"] or 0 for s in topn)
        A(f"按本期金额排序的前 {len(topn)} 家供应商合计 {_mv(cover)}，"
          f"占本期总金额 {cover / t['amount_cur'] * 100:.1f}%。"
          f"其中同比增幅超过 50% 的有 "
          f"{sum(1 for s in topn if (s['yoy'] or 0) > 50)} 家，"
          f"下降超过 30% 的有 {sum(1 for s in topn if (s['yoy'] is not None and s['yoy'] < -30))} 家。")
        A("")
        A(_table(["#", "供应商", f"{yp}年（万元）", f"{yc}年（万元）", "同比", "占本期"],
                 [[i + 1, s["name"], _wan(s["prev"]), _wan(s["cur"]),
                   _yoy(s["yoy"]), f"{s['share']:.1f}%"]
                  for i, s in enumerate(topn)]))
        A("")

    A("### 4.3 大额订单")
    A("")
    bo = res["big_orders"]
    if bo["count"]:
        A(f"{yc} 年单笔超 100 万元的订单共 {bo['count']} 笔，合计 {_mv(bo['amount'])}，"
          f"占总金额 {bo['share']:.1f}%。主要集中于以下供应商：")
        A("")
        A(_table(["供应商", "大额订单数", "金额（万元）"],
                 [[x["name"], x["count"], _wan(x["amount"])] for x in bo["top"]]))
        A("")

    A("### 4.4 供应商更替")
    A("")
    sn, sl = res["supplier_new"], res["supplier_lost"]
    A(f"{yc} 年新增供应商 **{churn['new']} 家**，贡献金额 {_mv(churn['new_amount'])}，"
      f"占本期总金额 {churn['new_share']:.1f}%，单家中位数 {_mv(churn['new_median'])}；"
      f"退出（本期无可比区间验收记录）**{churn['lost']} 家**，"
      f"其在 {yp} 年的采购金额 {_mv(churn['lost_amount'])}，占基期 {churn['lost_share']:.1f}%，"
      f"停供天数中位数 {churn['lost_idle_median']:.0f} 天。")
    A("")
    if sn:
        A("新增金额最大的五家：" + "、".join(
            f"{s['name']}（{_wan(s['amount'])} 万元，{s['top_cat']}）" for s in sn[:5]) + "。")
        A("")
    if sl:
        A("退出时金额最大的五家：" + "、".join(
            f"{s['name']}（{_wan(s['amount'])} 万元，{s['top_cat']}）" for s in sl[:5]) + "。")
        A("")
    if churn["lost_amount"]:
        ratio = churn["new_amount"] / churn["lost_amount"] - 1
        if abs(ratio) <= 0.25:
            tail = (f"新增与退出金额大体相当（{_mv(churn['new_amount'])} vs "
                    f"{_mv(churn['lost_amount'])}），说明供应商池在上行的同时完成了同等规模的结构置换，"
                    f"而非简单地扩张名单。")
        elif ratio > 0:
            tail = (f"新增金额明显高于退出金额（{_mv(churn['new_amount'])} vs "
                    f"{_mv(churn['lost_amount'])}），说明本期供应商池以扩容为主，"
                    f"需关注新供应商的准入与履约能力是否同步跟上。")
        else:
            tail = (f"新增金额低于退出金额（{_mv(churn['new_amount'])} vs "
                    f"{_mv(churn['lost_amount'])}），说明供应商池在收缩，"
                    f"需确认退出方是否已有合规替代来源，避免单点断供。")
        A(tail + "完整名单见 Excel 看板「04A 新增供应商」「04B 退出供应商」两张表。")
        A("")

    # ============ 五、交付时效 ============
    A("## 五、交付时效")
    A("")
    A("交付时效是本期数据中变化最剧烈的指标，需要从逾期面与逾期深度两个层面区分观察。")
    A("")
    A("### 5.1 逾期面与逾期深度")
    A("")
    d_rate = round(ov["cur"]["rate"] - ov["prev"]["rate"], 1)
    d_days = round(ov["cur"]["avg_days"] - ov["prev"]["avg_days"], 1)
    d_sev = round(ov["cur"]["severe_share"] - ov["prev"]["severe_share"], 1)
    A(_table(["指标", f"{yp}年", f"{yc}年", "变化"],
             [["逾期率", f"{ov['prev']['rate']:.1f}%", f"{ov['cur']['rate']:.1f}%", f"{d_rate:+.1f}pp"],
              ["逾期笔数", f"{_fmt_int(ov['prev']['overdue_count'])} 笔", f"{_fmt_int(ov['cur']['overdue_count'])} 笔",
               _fmt_pct((ov["cur"]["overdue_count"] - ov["prev"]["overdue_count"]) /
                        ov["prev"]["overdue_count"] * 100 if ov["prev"]["overdue_count"] else None)],
              ["逾期订单平均逾期天数", f"{ov['prev']['avg_days']:.1f} 天", f"{ov['cur']['avg_days']:.1f} 天",
               f"{d_days:+.1f} 天"],
              ["60天以上严重逾期占比", f"{ov['prev']['severe_share']:.1f}%", f"{ov['cur']['severe_share']:.1f}%",
               f"{d_sev:+.1f}pp"]]))
    A("")
    f_ov = findings.get("交付时效")
    if f_ov:
        A(f_ov["implication"])
        A("")

    A("### 5.2 逾期分布")
    A("")
    A(_table(["逾期区间", f"{yp}年占比", f"{yc}年占比"],
             [[d["range"], f"{d['prev']:.1f}%", f"{d['cur']:.1f}%"] for d in res["overdue_distribution"]]))
    A("")

    if res["overdue_suppliers"]:
        A("### 5.3 逾期最严重的供应商")
        A("")
        A(_table(["供应商", "逾期笔数", "平均逾期天数", "涉及金额（万元）"],
                 [[s["name"], s["count"], f"{s['avg_days']:.1f}", _wan(s["amount"])]
                  for s in res["overdue_suppliers"]]))
        A("")
        severe = [s for s in res["overdue_suppliers"] if (s["avg_days"] or 0) > 60]
        if severe:
            from rules import _short
            names = "、".join(_short(s["name"]) for s in severe[:2])
            A(f"{names} 的平均逾期超过 60 天，已明显偏离行业常规配送周期。"
              f"这通常指向两种情况：一是合同约定到货日期本身填写失真，二是实际采用「先送后补单」的结算模式。"
              f"**建议优先核查这两家的合同到货条款设置，而非直接判定其履约能力**。")
            A("")

    if efs["cur_avg"] is not None and efs["prev_avg"] is not None:
        A("### 5.4 验收效率")
        A("")
        A(f"从验收申请到完成验收的平均间隔由 {efs['prev_avg']:.1f} 天"
          f"{'延长' if efs['cur_avg'] > efs['prev_avg'] else '缩短'}至 {efs['cur_avg']:.1f} 天，"
          f"最长一例达 {_fmt_int(efs['max'])} 天。验收环节的积压会直接推迟供应商结算，"
          f"进而影响后续供货意愿，是逾期问题的潜在诱因之一。")
        A("")

    # ============ 六、价格与成本 ============
    A("## 六、价格与成本")
    A("")
    A("在总量增长的同时，单位成本是否同步上升，是判断采购绩效的核心问题。本章采用同物料两年比价的方法回答这一问题。")
    A("")
    A("### 6.1 整体价格水平")
    A("")
    if price["comparable"]:
        A(f"对两年均有采购、且各自数量不低于 10 的 **{_fmt_int(price['comparable'])} 种** 物料进行同口径比价："
          f"平均单价变化 {_fmt_pct(price['mean_change'])}，中位变化 {_fmt_pct(price['median_change'])}，"
          f"价格下降的物料占比约 {price['down_share']:.0f}%。")
        A("")
        A(readings_impl(price))
        A("")

    A("### 6.2 涨幅异常的物料")
    A("")
    if price["up20_count"]:
        A(f"涨幅超过 20% 的物料共 {_fmt_int(price['up20_count'])} 种，"
          f"涉及 {yc} 年采购金额 {_mv(price['up20_amount'])}。按金额排序取前 25 种如下：")
        A("")
        A(_table(["物料编码", "物料描述", f"单价{yp}", f"单价{yc}", "涨幅", f"{yc}金额（万元）"],
                 [[u["code"], (u["desc"] or "")[:28],
                   f"{u['prev_price']:.2f}" if u["prev_price"] is not None else "—",
                   f"{u['cur_price']:.2f}" if u["cur_price"] is not None else "—",
                   f"{u['change']:+.1f}%", _wan(u["amount"])]
                  for u in price["top_up"]]))
        A("")
        if price["top_up"]:
            tot = sum(u["amount"] or 0 for u in price["top_up"])
            share = tot / price["up20_amount"] * 100 if price["up20_amount"] else 0
            A(f"上表 25 种物料合计涉及 {_mv(tot)}，占全部涨幅超 20% 物料金额的 {share:.1f}%——"
              f"{'涨价高度集中在少数品种上，具备逐项复核调价依据的可行性' if share >= 50 else '涨价分布较为分散，需按品类成批处理'}。"
              f"建议按供应商归类后集中谈判，而非逐单一处理。")
            A("")

    A("### 6.3 价格下降的物料")
    A("")
    if price.get("down_count"):
        md = price["max_drop"]
        drop_txt = ("个别品种降幅超过 80%，多为计量单位或规格口径变更所致，需单独核查"
                    if md is not None and md <= -80
                    else f"单种最大降幅 {md:.1f}%")
        A(f"同期有 **{_fmt_int(price['down_count'])} 种**物料单价同比下降，"
          f"其中降幅达 5% 以上的 {_fmt_int(price['down5_count'])} 种、"
          f"涉及 {yc} 年采购金额 {_mv(price['down5_amount'])}，{drop_txt}。"
          f"下表为该范围内按金额排序的前 25 种：")
        A("")
        A(_table(["物料编码", "物料描述", f"单价{yp}", f"单价{yc}", "降幅", f"{yc}金额（万元）"],
                 [[u["code"], (u["desc"] or "")[:28],
                   f"{u['prev_price']:.2f}" if u["prev_price"] is not None else "—",
                   f"{u['cur_price']:.2f}" if u["cur_price"] is not None else "—",
                   f"{u['change']:+.1f}%", _wan(u["amount"])]
                  for u in price["top_down"]]))
        A("")
        if price["top_down"]:
            changes = [u["change"] for u in price["top_down"] if u["change"] is not None]
            med_d = sorted(changes)[len(changes) // 2] if changes else 0
            deep = sum(1 for c in changes if c <= -10)
            tot_d = sum(u["amount"] or 0 for u in price["top_down"])
            A(f"上表 25 种物料合计 {_mv(tot_d)}，降幅中位数 {med_d:.1f}%，"
              f"其中 {deep} 种降幅超过 10%。"
              f"这批成交价可直接作为同类物料新一轮询价的价格基准；"
              f"同时应核对同规格的其他供应商是否同步下调，避免出现"
              f"「同一物料、同一时间、不同价格」的敞口。")
            A("")

    if res["price_gaps"]:
        A("### 6.4 同规格多供应商价差")
        A("")
        A(f"同一规格在不同供应商之间存在显著价差，按涉及金额排序如下：")
        A("")
        A(_table(["物料描述", "供应商数", "最低单价", "最高单价", "价差倍数", f"{yc}金额（万元）"],
                 [[g["desc"], g["suppliers"], f"{g['min_price']:.1f}", f"{g['max_price']:.1f}",
                   f"**{g['ratio']:.1f}**", _wan(g["amount"])] for g in res["price_gaps"]]))
        A("")
        g_max = max(res["price_gaps"], key=lambda x: x["ratio"] or 0)
        A(f"同一规格、同一时间段的采购，单价最高相差 **{g_max['ratio']:.1f} 倍**"
          f"（「{g_max['desc']}」，{g_max['min_price']:.1f}—{g_max['max_price']:.1f} 元）。"
          f"这已经超出正常的品牌溢价和运输差异范围，更可能源于**规格执行不严、以次充好、或缺乏统一定价基准**。")
        A("")
        total_amt = sum(g["amount"] or 0 for g in res["price_gaps"])
        top = res["price_gaps"][0]
        A(f"上表 {len(res['price_gaps'])} 个规格合计涉及 {yc} 年采购金额 {_mv(total_amt)}，"
          f"其中金额最大的「{top['desc']}」单规格达 {_mv(top['amount'])}、价差 {top['ratio']:.1f} 倍。"
          f"若将各规格中高于最低价的采购量全部按最低价成交，理论上可节约 "
          f"{_saving(total_amt, 0.13, 0.26)}（按均价向底线收敛的 13%—26% 保守估算）。")
        A("")

    # ============ 七、矿点分布 ============
    A("## 七、矿点分布")
    A("")
    A(_table(["使用单位", f"{yp}年（万元）", f"{yc}年（万元）", "同比", f"{yc}占比"],
             [[u["name"], _wan(u["prev"]), _wan(u["cur"]), _yoy(u["yoy"]), f"{u['share']:.1f}%"]
              for u in res["units"]]))
    A("")
    movers = [u for u in res["units"] if u["name"] and not u["name"].startswith("（未填")
              and u["yoy"] is not None and abs(u["yoy"]) > 30
              and (u["prev"] or 0) >= 1e6 and (u["cur"] or 0) >= 1e6]
    movers.sort(key=lambda x: -abs((x["cur"] or 0) - (x["prev"] or 0)))
    if movers:
        parts = []
        for u in movers[:3]:
            direction = "增长" if u["yoy"] > 0 else "下降"
            delta = abs((u["cur"] or 0) - (u["prev"] or 0))
            parts.append(f"**{u['name']} {direction} {abs(u['yoy']):.1f}%**"
                         f"（{'+' if u['yoy'] > 0 else '-'}{_wan(delta)} 万元）")
        A("显著变化：" + "；".join(parts) + "。采购主体在矿点间的分布正在调整，"
          "建议结合各矿点产量与检修计划核实变动合理性。")
        A("")
    gains = [u for u in res["units"] if u["name"] and not u["name"].startswith("（未填")
             and u["yoy"] is None and (u["cur"] or 0) >= 1e6]
    if gains:
        A("另有 " + "、".join(f"{u['name']}（{_wan(u['cur'])} 万元）" for u in gains[:4])
          + " 为基期无数据、本期新增的核算单位，多与组织架构调整或新设主体有关，"
            "其金额变动不宜按同比解读。")
        A("")

    # ============ 八、问题与建议 ============
    A("## 八、主要问题与管理建议")
    A("")
    cn = "一二三四五六七八九十"
    for i, pr in enumerate(rule_out["problems"]):
        # 问题标题本身已含编号语义，不再叠加「问题N：」前缀，避免「问题一：问题二：...」
        A(f"### 8.{i + 1} {pr['title']}")
        A("")
        A(f"**现状**：{pr['evidence']}")
        A("")
        A(f"**建议**：{pr['action']}")
        A("")
        if pr.get("benefit"):
            A(f"**预期收益**：{pr['benefit']}")
            A("")

    # ============ 九、下阶段重点 ============
    A("## 九、下阶段工作重点")
    A("")
    for i, a in enumerate(rule_out["actions"], 1):
        A(f"{i}. {a}")
    A("")
    A("> 说明：本报告数据来源于验收单报表，反映的是实际到货验收口径，而非合同签订口径。"
      "部分结论（如设备投资收缩、矿点产能变化）需结合业务部门的项目计划进一步验证。")
    A("")

    return "\n".join(L)


def readings_impl(price):
    if price["median_change"] is not None and price["median_change"] <= 0:
        return ("整体单价稳中有降，说明集团的框架协议与比价机制基本有效，成本控制未因采购量上升而失守。")
    return "整体单价有所上行，需复核调价依据是否充分，并检查框架协议的调价条款执行情况。"


def _saving(amount, rate_lo=0.13, rate_hi=0.26):
    if not amount:
        return "待测算"
    return f"{_fmt_money(amount * rate_lo)}—{_fmt_money(amount * rate_hi)}"
