# -*- coding: utf-8 -*-
"""
结论规则引擎

设计原则：完全确定性、可复现。同一份数据跑两次，输出一字不差。
所有判定阈值集中在 THRESHOLDS，可按需调整而无需改代码逻辑。

规则输出格式：
  {
    "findings": [ {level, title, facts[], implication} ],   # 分析发现（正文各章引用）
    "problems": [ {title, evidence, action, benefit} ],      # 问题与建议
    "actions":  [ str ],                                     # 下阶段工作重点
    "headline": str,                                          # 一句话核心结论
    "kpis":     [ {label, value, unit, note} ],               # 摘要卡片
  }
"""
import math

THRESHOLDS = {
    "concentration_drop_good": 3.0,    # CR5 下降超过 3pp 视为结构改善
    "concentration_high": 30.0,        # 单品类 TOP1 占比 > 30% 视为依赖
    "overdue_worsen_days": 3.0,        # 平均逾期天数恶化超过 3 天视为警告
    "overdue_severe": 10.0,            # 60天以上占比 > 10% 视为严重
    "price_gap_alert": 2.0,            # 同规格价差 > 2 倍视为异常
    "hot_growth": 50.0,                # 品类金额增幅 > 50% 视为异动
    "small_order_share": 40.0,         # 小额订单笔数占比 > 40% 视为流程负担
    "yoy_flat": 1.5,                   # |同比| < 1.5% 视为基本持平
    "yoy_slight": 5.0,                 # 1.5% ≤ |同比| < 5% 视为小幅波动
}


def _fmt_money(v, unit="元"):
    """金额自适应：亿 / 万 / 元

    自适应精度（避免 7.61 亿被舍成 8 亿、1.42 万被舍成 1 万）：
      ≥ 10 亿  -> 1 位小数
      ≥ 1 亿   -> 2 位小数
      ≥ 1000万 -> 0 位小数
      ≥ 100万  -> 1 位小数
      ≥ 1 万   -> 2 位小数
      其余      -> 0 位小数
    """
    if v is None:
        return "—"
    a = abs(v)
    if a >= 1e9:
        return f"{v/1e8:.1f} 亿元"
    if a >= 1e8:
        return f"{v/1e8:.2f} 亿元"
    if a >= 1e7:
        return f"{v/1e4:,.0f} 万元"
    if a >= 1e6:
        return f"{v/1e4:,.1f} 万元"
    if a >= 1e4:
        return f"{v/1e4:,.2f} 万元"
    return f"{v:,.0f} 元"


def _fmt_int(v):
    return f"{int(v):,}" if v is not None else "—"


def _fmt_pct(v, sign=True):
    if v is None:
        return "—"
    return f"{v:+.1f}%" if sign else f"{v:.1f}%"


def _cn(n):
    CN = "零一二三四五六七八九十"
    return CN[n] if 0 <= n <= 10 else str(n)


def build(res: dict) -> dict:
    t = res["totals"]
    meta = res["meta"]
    ov = res["overdue"]
    price = res["price"]
    conc = res["supplier_concentration"]
    th = THRESHOLDS

    cur_y, prev_y = meta["year_cur"], meta["year_prev"]
    mr = meta["month_range"]

    findings = []
    problems = []
    actions = []

    # ---------- 1. 总量判断 ----------
    yoy = t["amount_yoy"]
    if yoy is None:
        trend = "无法比较"
    elif abs(yoy) < th["yoy_flat"]:
        trend = "基本持平"
    elif abs(yoy) < th["yoy_slight"]:
        trend = "小幅增长" if yoy > 0 else "小幅下降"
    elif yoy > 0:
        trend = "稳步增长"
    else:
        trend = "有所下降"

    # 笔数与单笔的反向关系
    scale_note = ""
    if t["orders_yoy"] and t["avg_order_yoy"]:
        if t["orders_yoy"] > 10 and t["avg_order_yoy"] < -5:
            scale_note = (f"验收笔数增长 {t['orders_yoy']:.1f}%，但平均单笔金额下降 "
                          f"{abs(t['avg_order_yoy']):.1f}%，呈「笔数增、单笔小」特征，"
                          f"采购活动趋于碎片化")
        elif t["orders_yoy"] < -5 and t["avg_order_yoy"] > 5:
            scale_note = (f"验收笔数下降 {abs(t['orders_yoy']):.1f}%，但平均单笔金额上升 "
                          f"{t['avg_order_yoy']:.1f}%，采购集中度在提升")

    findings.append({
        "level": "info",
        "title": "总体规模",
        "facts": [
            f"{cur_y}年{mr}验收金额 {_fmt_money(t['amount_cur'])}，"
            f"{prev_y}年同期 {_fmt_money(t['amount_prev'])}，"
            f"同比 {_fmt_pct(yoy)}（{'+' if (t['amount_delta'] or 0) >= 0 else ''}"
            f"{_fmt_money(t['amount_delta'])}）",
            f"验收笔数 {_fmt_int(t['orders_cur'])} 笔，同比 {_fmt_pct(t['orders_yoy'])}；"
            f"平均单笔 {_fmt_money(t['avg_order_cur'])}，同比 {_fmt_pct(t['avg_order_yoy'])}",
            f"按{mr}节奏线性外推，{cur_y} 全年预计 {_fmt_money(t['annual_forecast'])}",
        ] + ([scale_note] if scale_note else []),
        "implication": _implication_total(yoy, t),
    })

    # ---------- 2. 结构变化 ----------
    growth_cats = [c for c in res["categories"] if (c["yoy"] or 0) > 10 and (c["prev"] or 0) > 0]
    shrink_cats = [c for c in res["categories"] if (c["yoy"] or 0) < -10 and (c["prev"] or 0) > 0]
    if growth_cats or shrink_cats:
        facts = []
        if growth_cats:
            names = "、".join(f"{c['name']}（{_fmt_pct(c['yoy'])}）" for c in growth_cats[:3])
            facts.append(f"增长品类：{names}")
        if shrink_cats:
            names = "、".join(f"{c['name']}（{_fmt_pct(c['yoy'])}）" for c in shrink_cats[:3])
            facts.append(f"收缩品类：{names}")

        cons, capex = _split_categories(res["categories"])
        if cons and capex:
            if cons["yoy"] > 0 and capex["yoy"] < 0:
                facts.append(
                    f"消耗性物资（材料+备件）合计 {_fmt_pct(cons['yoy'])}，"
                    f"资本性支出（设备+大型部件）合计 {_fmt_pct(capex['yoy'])}，"
                    f"呈现「消耗扩张、投资收缩」格局")
                findings.append({
                    "level": "warn",
                    "title": "物资结构",
                    "facts": facts,
                    "implication": (
                        "资本性支出收缩的同时消耗性物资扩张，通常指向存量设备维护强度上升。"
                        "建议与设备管理部门的检修计划交叉验证，判断是压减投资所致，"
                        "还是老旧设备故障率上升所致。"),
                })
        if not any(f["title"] == "物资结构" for f in findings):
            findings.append({
                "level": "info", "title": "物资结构", "facts": facts,
                "implication": "各品类结构变化幅度有限，采购结构保持稳定。",
            })

    # ---------- 3. 供应商集中度 ----------
    cr5_cur, cr5_prev = conc["cur"]["cr5"], conc["prev"]["cr5"]
    delta = round(cr5_cur - cr5_prev, 1)   # 负数代表集中度下降
    churn = res["supplier_churn"]
    facts = [
        f"合作供应商 {conc['cur']['count']} 家（去年 {conc['prev']['count']} 家），"
        f"两年均合作 {churn['both']} 家，新增 {churn['new']} 家"
        f"（{_fmt_money(churn['new_amount'])}，占本期 {churn['new_share']:.1f}%）、"
        f"退出 {churn['lost']} 家（{_fmt_money(churn['lost_amount'])}）",
        f"CR5 由 {cr5_prev:.1f}% 变为 {cr5_cur:.1f}%（{delta:+.1f}pp），"
        f"CR10 由 {conc['prev']['cr10']:.1f}% 变为 {conc['cur']['cr10']:.1f}%",
    ]
    high_dep = [c for c in res["category_concentration"] if c["top1_share"] > th["concentration_high"]]
    if high_dep:
        d = high_dep[0]
        facts.append(f"但 {d['name']} 类对单一供应商依赖度仍达 {d['top1_share']:.1f}%"
                     f"（{d['top1_name']}），CR3 为 {d['cr3']:.1f}%")

    if delta <= -th["concentration_drop_good"]:
        level, impl = "good", "头部依赖度明显下降，供应商结构趋于分散，抗风险能力增强。"
    elif delta >= th["concentration_drop_good"]:
        level, impl = "warn", "头部集中度上升，议价能力可能被反向制约，需关注单一供应商风险。"
    else:
        level, impl = "info", "集中度基本稳定。"
    if high_dep:
        impl += f" 细分品类的依赖问题依然存在，建议对 {high_dep[0]['name']} 类培育第二供应源。"
    findings.append({"level": level, "title": "供应商格局", "facts": facts, "implication": impl})

    # ---------- 4. 交付时效（最重要） ----------
    d_days = round(ov["cur"]["avg_days"] - ov["prev"]["avg_days"], 1)
    d_rate = round(ov["cur"]["rate"] - ov["prev"]["rate"], 1)
    d_severe = round(ov["cur"]["severe_share"] - ov["prev"]["severe_share"], 1)
    facts = [
        f"逾期率 {ov['prev']['rate']:.1f}% → {ov['cur']['rate']:.1f}%（{d_rate:+.1f}pp），"
        f"逾期笔数 {_fmt_int(ov['prev']['overdue_count'])} → {_fmt_int(ov['cur']['overdue_count'])}",
        f"逾期订单平均逾期天数 {ov['prev']['avg_days']:.1f} 天 → {ov['cur']['avg_days']:.1f} 天"
        f"（{d_days:+.1f} 天）",
        f"60 天以上严重逾期占比 {ov['prev']['severe_share']:.1f}% → "
        f"{ov['cur']['severe_share']:.1f}%（{d_severe:+.1f}pp）",
    ]
    ef = res["accept_efficiency"]
    if ef["cur_avg"] is not None and ef["prev_avg"] is not None:
        facts.append(f"验收申请到完成的平均间隔 {ef['prev_avg']:.1f} 天 → {ef['cur_avg']:.1f} 天，"
                     f"最长 {_fmt_int(ef['max'])} 天")

    if d_days > th["overdue_worsen_days"] or ov["cur"]["severe_share"] > th["overdue_severe"]:
        impl = ("逾期率看似平稳，但深度指标明显恶化——逾期正从偶发延误转变为部分供应商的常态。"
                "需从合同到货条款的真实性入手核查，而非仅考核供应商履约能力。")
        level = "danger"
        problems.append({
            "title": "逾期交付深度恶化",
            "evidence": f"平均逾期 {ov['prev']['avg_days']:.1f} → {ov['cur']['avg_days']:.1f} 天，"
                        f"严重逾期占比 {ov['prev']['severe_share']:.1f}% → {ov['cur']['severe_share']:.1f}%",
            "action": _overdue_action(res),
            "benefit": "压缩平均逾期天数，降低停工待料风险",
        })
    elif d_rate > 3:
        impl = "逾期率上升，需关注交付能力。"
        level = "warn"
    else:
        impl = "交付表现基本稳定。"
        level = "good"
    findings.append({"level": level, "title": "交付时效", "facts": facts, "implication": impl})

    # ---------- 5. 价格 ----------
    facts = [
        f"两年可比物料 {_fmt_int(price['comparable'])} 种，单价中位变化 "
        f"{_fmt_pct(price['median_change'])}，价格下降的物料占 {price['down_share']:.1f}%",
    ]
    if price["up20_count"]:
        facts.append(f"涨幅超 20% 的物料 {_fmt_int(price['up20_count'])} 种，"
                     f"涉及 {cur_y} 年金额 {_fmt_money(price['up20_amount'])}")
    gaps = res["price_gaps"]
    if gaps:
        g = gaps[0]
        facts.append(f"同规格不同供应商单价差异最大的为「{g['desc']}」，"
                     f"最低 {g['min_price']:.1f} 元、最高 {g['max_price']:.1f} 元，"
                     f"相差 {g['ratio']:.1f} 倍，涉及 {_fmt_money(g['amount'])}")

    if price["median_change"] is not None and price["median_change"] <= 0:
        impl = "整体价格水平稳中有降，框架协议与比价机制基本有效。"
        level = "good"
    else:
        impl = "整体价格上行，需复核调价依据。"
        level = "warn"

    if gaps and gaps[0]["ratio"] >= th["price_gap_alert"]:
        impl += (f" 但同规格价差达 {gaps[0]['ratio']:.1f} 倍已超出正常品牌与运输差异范围，"
                 f"更可能源于规格执行不严或缺乏统一价格基线。")
        level = "danger"
        problems.append({
            "title": f"「{gaps[0]['desc']}」类物资价格失控",
            "evidence": f"同规格 {gaps[0]['suppliers']} 家供应商，单价 {gaps[0]['min_price']:.1f}–"
                        f"{gaps[0]['max_price']:.1f} 元，差 {gaps[0]['ratio']:.1f} 倍",
            "action": "按规格建立统一价格基线库，归口 2—3 家供应商集中议价；"
                      "同时统一物料命名规则，杜绝以高规格名义采购低规格料",
            "benefit": _benefit_estimate(gaps[0]["amount"]),
        })
    findings.append({"level": level, "title": "价格与成本", "facts": facts, "implication": impl})

    # ---------- 6. 长尾与小额 ----------
    tail = res["tail"]
    if tail["small_share_cur"] > th["small_order_share"]:
        findings.append({
            "level": "warn", "title": "长尾与小额订单",
            "facts": [
                f"金额低于 1,000 元的订单 {_fmt_int(tail['small_count_cur'])} 笔，"
                f"占笔数 {tail['small_share_cur']:.1f}%，但金额仅占 {tail['small_amount_share']:.1f}%",
                f"{cur_y} 年采购的 {_fmt_int(price and res['totals']['materials_cur'] or 0)} 种物料中，"
                f"仅采购 1 次的有 {_fmt_int(tail['once_material'])} 种（{tail['once_share']:.0f}%）",
            ],
            "implication": "小额高频订单的单位流程成本远高于采购价值，长尾物料是编码膨胀的主因。",
        })
        problems.append({
            "title": "小额订单流程成本过高",
            "evidence": f"{_fmt_int(tail['small_count_cur'])} 笔订单金额低于 1,000 元，"
                        f"占笔数 {tail['small_share_cur']:.1f}%，金额仅占 {tail['small_amount_share']:.1f}%",
            "action": "对高频低值物资推行「框架协议 + 集中配送 + 定期结算」，"
                      "减少单次审批与验收动作",
            "benefit": "降低流程人力成本，释放采购与验收人员产能",
        })

    # ---------- 7. 数据治理 ----------
    blank = res.get("blank_unit")
    if blank and blank["share"] > 5:
        blank_money = _fmt_money(blank["cur"]).replace(" ", "")
        findings.append({
            "level": "warn", "title": "数据治理",
            "facts": [f"「使用单位」为空的记录在可比区间内金额 {blank_money}，"
                      f"占同期总金额 {blank['share']:.1f}%，无法归集到矿点核算"],
            "implication": "这部分金额长期游离在矿点考核体系之外，建议将使用单位设为必填并做有效性校验。",
        })
        problems.append({
            "title": "使用单位字段缺失",
            "evidence": f"可比区间内 {blank['share']:.1f}% 的金额无法归集到矿点",
            "action": "验收单中「使用单位」设为必填项并做有效性校验；设备直供类明确标注归属矿点",
            "benefit": "打通矿点级成本核算",
        })

    # ---------- 8. 集中度风险（若有高依赖） ----------
    if high_dep:
        d = high_dep[0]
        problems.append({
            "title": f"{d['name']}类供应商依赖度过高",
            "evidence": f"TOP1 {d['top1_name']} 占 {d['top1_share']:.1f}%，CR3 达 {d['cr3']:.1f}%",
            "action": f"对 {d['name']} 类设置单一供应商份额上限（建议不超过 30%），"
                      f"培育第二、第三供应源",
            "benefit": "防止议价权反转，保障供应连续性",
        })

    # ---------- 9. 其他建议 ----------
    if tail["once_share"] > 50:
        problems.append({
            "title": "物料编码膨胀",
            "evidence": f"{tail['once_share']:.0f}% 的物料编码仅采购 1 次",
            "action": "推进物料标准化，对同类低值耗材做特征值合并，压缩编码总量",
            "benefit": "降低比价难度与库存 SKU 数量",
        })

    if ef["cur_avg"] and ef["prev_avg"] and ef["cur_avg"] - ef["prev_avg"] > 0.5:
        actions.append(f"压缩验收环节积压，验收间隔已由 {ef['prev_avg']:.1f} 天延长至 "
                       f"{ef['cur_avg']:.1f} 天")

    # ---------- 10. 下阶段重点 ----------
    priority = []
    if gaps:
        priority.append(f"启动「{gaps[0]['desc']}」专项比价，30 日内形成统一价格基线，"
                        f"季度内完成供应商归口")
    if level_overdue := (d_days > th["overdue_worsen_days"]):
        priority.append("核查逾期最严重供应商的合同到货条款真实性，"
                        "区分条款失真与履约能力问题，一季度内完成条款修订")
    if tail["once_share"] > 50:
        priority.append(f"清理 {_fmt_int(tail['once_material'])} 个仅采购 1 次的物料编码，推进标准化")
    if blank and blank["share"] > 5:
        priority.append("将「使用单位」纳入验收单必填校验，打通矿点成本核算")
    priority.append("建立月度采购分析机制，将同比、集中度、逾期率、价格异动四项指标纳入常态化监控")
    actions = priority + actions

    # ---------- headline & KPI ----------
    headline = (f"{cur_y}年{mr}验收金额 {_fmt_money(t['amount_cur'])}，"
                f"同比 {_fmt_pct(yoy)}，{trend}")
    if scale_note:
        headline += f"；{scale_note.split('，')[0]}"

    kpis = [
        {"label": f"{cur_y}年{mr}验收金额", "value": _fmt_money(t["amount_cur"]).replace(" ", ""),
         "unit": "", "note": f"同比 {_fmt_pct(yoy)}"},
        {"label": "验收笔数", "value": _fmt_int(t["orders_cur"]), "unit": "笔",
         "note": f"同比 {_fmt_pct(t['orders_yoy'])}"},
        {"label": "合作供应商", "value": _fmt_int(t["suppliers_cur"]), "unit": "家",
         "note": f"CR5 {conc['cur']['cr5']:.1f}%"},
        {"label": "平均逾期天数", "value": f"{ov['cur']['avg_days']:.1f}", "unit": "天",
         "note": f"去年 {ov['prev']['avg_days']:.1f} 天"},
        {"label": "单价中位变化", "value": _fmt_pct(price["median_change"]), "unit": "",
         "note": f"可比物料 {_fmt_int(price['comparable'])} 种"},
    ]

    return {
        "headline": headline,
        "trend": trend,
        "findings": findings,
        "problems": problems,
        "actions": actions,
        "kpis": kpis,
        "alert_count": sum(1 for f in findings if f["level"] == "danger"),
    }


def _implication_total(yoy, t):
    if yoy is None:
        return "缺少可比的基期数据，无法判断增长趋势。"
    if yoy > 15:
        return "采购规模快速扩张，需同步关注预算执行进度与供应链承载能力。"
    if yoy > 0:
        return "规模温和增长，处于预算可控范围。"
    if yoy > -10:
        return "规模略有回落，需确认是否为需求下降或采购节奏推迟。"
    return "规模明显收缩，建议核查是否存在项目延期或预算压减。"


def _split_categories(cats):
    """拆分为消耗性与资本性两组"""
    cons = [c for c in cats if c["name"] in ("材料", "备件")]
    capex = [c for c in cats if c["name"] in ("移动设备", "固定设备", "大型部件")]
    out = []
    for grp in (cons, capex):
        if not grp:
            out.append(None)
            continue
        cur = sum(c["cur"] or 0 for c in grp)
        prev = sum(c["prev"] or 0 for c in grp)
        out.append({"cur": cur, "prev": prev,
                    "yoy": round((cur - prev) / prev * 100, 1) if prev else 0})
    return out[0], out[1]


def _short(name, n=12):
    """供应商名称简写：优先在后缀词前断开，保留「气体」「工贸」这类有辨识度的字眼"""
    s = str(name)
    # 从长到短尝试，取第一个长度合理的前缀
    for suffix in ("生产销售", "有限责任公司", "有限公司", "供应站", "分公司", "集团"):
        i = s.find(suffix)
        if 0 < i <= n:
            return s[:i]
    # 后缀在很后面，退而取「气体/工贸/商贸/科技/装备」等业务词结尾
    for kw in ("气体", "工贸", "商贸", "矿业", "煤机", "装备", "科技", "机电", "液压"):
        i = s.find(kw)
        if 0 < i <= n + 2:
            return s[:i + len(kw)]
    return s[:n] + "…" if len(s) > n else s


def _overdue_action(res):
    """根据逾期供应商特征生成针对性建议"""
    sup = res["overdue_suppliers"][:5]
    severe = [s for s in sup if s["avg_days"] > 60]
    base = "对平均逾期超过 60 天的供应商重新核定合同到货条款"
    if severe:
        names = "、".join(_short(s["name"]) for s in severe[:2])
        base += f"（如 {names}，平均逾期超 60 天，疑为「先送后补单」模式，应改为按实际配送批次开单）"
    base += "；同时引入逾期扣款条款，将逾期率纳入供应商年度评价体系"
    return base


def _benefit_estimate(amount):
    """按价差收敛到均价的保守估计：约 13%—26% 的节约空间"""
    if not amount:
        return "压缩价差后可释放议价空间"
    lo, hi = amount * 0.13, amount * 0.26
    return f"保守估算年节约 {_fmt_money(lo)}—{_fmt_money(hi)}"
