# -*- coding: utf-8 -*-
"""
采购验收数据分析 · 主入口

用法：
  # 手动指定文件
  python run.py --files "D:/path/验收单报表2025.xlsx" "D:/path/验收单报表2026.xlsx" --out "D:/out"

  # 自动扫描目录（按文件名中的年份识别，同年取最新修改的）
  python run.py --dir "D:/仲泰/仲泰2026/10月/数据分析" --out "D:/out"

  # 只产 Excel 或只产 Word
  python run.py --dir ... --out ... --skip-docx
  python run.py --dir ... --out ... --skip-excel

产物（在 --out 目录下）：
  采购数据分析报告_YYYY-YYYY.xlsx   11 工作表看板
  采购数据分析报告_YYYY-YYYY.docx   Word 版报告
  stage1/final_draft.md             正文中间产物
  stage2/formatted-report.html      排版中间产物
  analysis_result.json              全部结构化结果（便于复核与二次开发）
"""
import argparse
import json
import os
import shutil
import subprocess
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import analyzer
import excel_writer
import html_writer
import loader
import md_writer
import rules

# html-to-docx 转换器（tencent-docx 插件，可用 HTML2DOCX_PY 覆盖）
DOCX_PY = os.environ.get("HTML2DOCX_PY") or os.path.expanduser(
    "~/.venv-html-to-docx/Scripts/python.exe")
DOCX_PLUGIN = os.path.expanduser(
    "~/.workbuddy/plugins/cache/workbuddy-builtin/tencent-docx")


def _find_docx_scripts():
    """定位 html_to_docx 的 scripts 目录"""
    if not os.path.isdir(DOCX_PLUGIN):
        return None
    for ver in sorted(os.listdir(DOCX_PLUGIN), reverse=True):
        p = os.path.join(DOCX_PLUGIN, ver, "skills", "html-to-docx", "scripts")
        if os.path.isdir(p):
            return p
    return None


def step(msg):
    print(f"\n=== {msg} ===", flush=True)


def main():
    ap = argparse.ArgumentParser(description="采购验收数据分析")
    ap.add_argument("--files", nargs="*", help="验收单报表文件路径（可多个）")
    ap.add_argument("--dir", help="自动扫描该目录下的 验收单报表*.xlsx")
    ap.add_argument("--out", required=True, help="输出目录")
    ap.add_argument("--title", default=None, help="报告标题")
    ap.add_argument("--org", default=None, help="集团/单位名称，用于封面")
    ap.add_argument("--skip-excel", action="store_true")
    ap.add_argument("--skip-docx", action="store_true")
    ap.add_argument("--keep-html", action="store_true", help="保留中间 HTML 便于检查")
    args = ap.parse_args()

    out_dir = os.path.abspath(args.out)
    os.makedirs(out_dir, exist_ok=True)

    # ---------- 1. 定位文件 ----------
    step("1/6 定位数据文件")
    if args.files:
        files = [os.path.abspath(f) for f in args.files]
        for f in files:
            if not os.path.exists(f):
                raise FileNotFoundError(f)
    elif args.dir:
        found = loader.discover_files(args.dir)
        files = list(found.values())
        for y, p in found.items():
            print(f"  {y} 年 -> {os.path.basename(p)}")
    else:
        raise SystemExit("必须指定 --files 或 --dir")

    print(f"  共 {len(files)} 个文件")

    # ---------- 2. 读取与清洗 ----------
    step("2/6 读取与清洗")
    df = loader.load_dataset(files)
    quality = loader.quality_report(df)
    print(f"  合并 {len(df):,} 行，年份分布 {quality['年份分布']}")
    print(f"  日期范围 {quality['业务日期范围'][0]} 至 {quality['业务日期范围'][1]}")
    if quality["日期缺失行数"]:
        print(f"  ⚠ 有 {quality['日期缺失行数']} 行无法解析业务日期，已从分析中排除")

    # ---------- 3. 分析 ----------
    step("3/6 指标分析")
    res = analyzer.analyze(df)
    meta = res["meta"]
    print(f"  对比区间：{meta['year_prev']} vs {meta['year_cur']}（{meta['month_range']}）")
    print(f"  {meta['compare_note']}")

    # ---------- 4. 结论 ----------
    step("4/6 规则引擎生成结论")
    rule_out = rules.build(res)
    print(f"  核心结论：{rule_out['headline']}")
    print(f"  发现 {len(rule_out['findings'])} 项 / 问题 {len(rule_out['problems'])} 个 "
          f"/ 建议 {len(rule_out['actions'])} 条 / 高危 {rule_out['alert_count']} 项")
    for f in rule_out["findings"]:
        tag = {"danger": "[严重]", "warn": "[关注]", "good": "[良好]", "info": "[信息]"}[f["level"]]
        print(f"    {tag} {f['title']}")

    # ---------- 5. 输出 ----------
    step("5/6 生成报告文件")
    yc, yp = meta["year_cur"], meta["year_prev"]
    base = f"采购数据分析报告_{yp}-{yc}"
    org = args.org or ""
    title = args.title or f"{org}采购验收数据分析报告（{yp}—{yc}）"

    # Excel
    xlsx_path = None
    if not args.skip_excel:
        xlsx_path = os.path.join(out_dir, f"{base}.xlsx")
        excel_writer.write_excel(res, rule_out, quality, xlsx_path,
                                 title=title, files=files)
        print(f"  Excel: {xlsx_path} ({os.path.getsize(xlsx_path)/1024:.0f} KB)")

    # JSON（全量结构化结果）
    json_path = os.path.join(out_dir, "analysis_result.json")
    payload = {
        "meta": meta,
        "totals": res["totals"],
        "monthly": res["monthly"],
        "categories": res["categories"],
        "supplier_concentration": res["supplier_concentration"],
        "supplier_churn": res["supplier_churn"],
        "top_suppliers": res["top_suppliers"],
        "supplier_new": res["supplier_new"],
        "supplier_lost": res["supplier_lost"],
        "category_concentration": res["category_concentration"],
        "big_orders": res["big_orders"],
        "overdue": res["overdue"],
        "overdue_distribution": res["overdue_distribution"],
        "overdue_suppliers": res["overdue_suppliers"],
        "accept_efficiency": res["accept_efficiency"],
        "price": res["price"],
        "price_gaps": res["price_gaps"],
        "units": res["units"],
        "blank_unit": res["blank_unit"],
        "tail": res["tail"],
        "agreement": res["agreement"],
        "quality": quality,
        "conclusions": rule_out,
    }
    with open(json_path, "w", encoding="utf-8") as fh:
        json.dump(payload, fh, ensure_ascii=False, indent=2, default=str)
    print(f"  JSON : {json_path}")

    docx_path = None
    if not args.skip_docx:
        # Markdown 正文
        stage1 = os.path.join(out_dir, "stage1")
        stage2 = os.path.join(out_dir, "stage2")
        os.makedirs(stage1, exist_ok=True)
        os.makedirs(stage2, exist_ok=True)

        md = md_writer.render_md(res, rule_out, quality, title,
                                 files=[os.path.basename(f) for f in files])
        md_path = os.path.join(stage1, "final_draft.md")
        open(md_path, "w", encoding="utf-8").write(md)
        print(f"  MD   : {md_path} ({len(md)} 字符)")

        # 执行摘要段落
        summary = _build_summary(res, rule_out)
        meta_lines = [
            f"数据来源：SAP 验收单报表（{yp} / {yc}）",
            f"分析样本：{quality['总行数']:,} 条验收记录",
            f"分析口径：{meta['month_range']}可比区间",
            f"编制日期：{yc} 年 {_now_month()} 月",
        ]
        html_path = html_writer.write_html(
            md, stage2, title=title,
            cover_title=f"{org}采购验收数据分析报告".replace("集团采购", "集团<br>采购") if org else "采购验收数据分析报告",
            subtitle=f"{yp}—{yc} 年对比分析（{meta['month_range']}可比区间）",
            meta_lines=meta_lines,
            kpis=rule_out["kpis"],
            summary_paragraphs=summary,
        )
        print(f"  HTML : {html_path}")

        issues = html_writer.self_check(html_path)
        if issues:
            print("  ⚠ 排版自检发现问题：")
            for i in issues:
                print(f"      {i}")
        else:
            print("  排版自检：通过")

        # docx 转换
        docx_path = os.path.join(out_dir, f"{base}.docx")
        engine = _convert(html_path, docx_path)
        if engine:
            print(f"  DOCX : {docx_path} ({os.path.getsize(docx_path)/1024:.0f} KB) "
                  f"[{engine}]")
        else:
            print("  ⚠ DOCX 转换失败，已保留 HTML，可用 Word 直接打开或另存")
            docx_path = None

        if not args.keep_html and docx_path:
            shutil.rmtree(stage1, ignore_errors=True)
            shutil.rmtree(stage2, ignore_errors=True)

    step("6/6 完成")
    print(f"  输出目录：{out_dir}")
    return {
        "xlsx": xlsx_path,
        "docx": docx_path,
        "json": json_path,
        "headline": rule_out["headline"],
    }


def _now_month():
    import datetime
    return datetime.date.today().month


def _build_summary(res, rule_out):
    """由规则结果拼装执行摘要段落"""
    t, meta = res["totals"], res["meta"]
    yc = meta["year_cur"]
    danger = [f for f in rule_out["findings"] if f["level"] == "danger"]
    warn = [f for f in rule_out["findings"] if f["level"] == "warn"]

    p1 = (f"{yc}年{meta['month_range']}，累计完成验收金额 **{rules._fmt_money(t['amount_cur']).replace(' ', '')}**，"
          f"较基期同期 **{rules._fmt_pct(t['amount_yoy'])}**。")
    orders = f"金额变化主要来自采购笔数（{rules._fmt_pct(t['orders_yoy'])}）的扩张" \
        if (t["orders_yoy"] or 0) > (t["amount_yoy"] or 0) else \
        f"金额变化主要由单价水平（{rules._fmt_pct(t['avg_order_yoy'])}）驱动"
    p1 += f"{orders}，平均单笔金额 {rules._fmt_money(t['avg_order_prev']).replace(' ', '')} → " \
          f"{rules._fmt_money(t['avg_order_cur']).replace(' ', '')}。"
    if t["annual_forecast"]:
        p1 += f"按当前节奏外推，全年验收金额预计 {rules._fmt_money(t['annual_forecast']).replace(' ', '')}。"

    if danger or warn:
        items = "、".join(f"**{f['title']}**" for f in (danger + warn)[:3])
        p2 = f"规模变化的背后，{items} 是本期最需要关注的结构性问题。具体证据与改进建议见正文各章。"
    else:
        p2 = "各项指标整体平稳，未发现需要立即处置的重大风险。"

    return [p1, p2]


def _convert(html_path, docx_path):
    """优先用 tencent-docx 的 html-to-docx 插件，不可用时降级到内置转换器。

    返回实际使用的引擎名称（"html-to-docx" / "builtin"），失败返回 None。
    """
    if _convert_plugin(html_path, docx_path):
        return "html-to-docx"
    try:
        import docx_fallback
        docx_fallback.convert(html_path, docx_path)
        if os.path.exists(docx_path):
            return "builtin"
    except Exception as e:
        print(f"  内置转换器异常：{e}")
    return None


def _convert_plugin(html_path, docx_path):
    scripts = _find_docx_scripts()
    if not scripts:
        return False
    if not os.path.exists(DOCX_PY):
        return False

    cmd = [DOCX_PY, "-m", "html_to_docx", "convert", html_path,
           "-o", docx_path, "--page-size", "a4"]
    try:
        r = subprocess.run(cmd, cwd=scripts, capture_output=True, text=True, timeout=300)
        if r.returncode != 0:
            print("  转换器输出：", (r.stderr or r.stdout)[-500:])
            return False
        return os.path.exists(docx_path)
    except Exception as e:
        print(f"  转换异常：{e}")
        return False


if __name__ == "__main__":
    main()
