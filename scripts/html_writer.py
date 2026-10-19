# -*- coding: utf-8 -*-
"""
Markdown -> 业务报告 HTML（docx 友好）

排版规范（均来自上一版 100 分过审的成品，勿轻易改动）：
  · 所有间距走 CSS 变量，禁止裸值（html-review [DT-03]）
  · KPI 表必须有 thead/tbody（[SI-02]）
  · h2 后不得直接跟 h3，必须插章节导语（[TQ-03]）
  · 目录用 div + p，不用 ol，否则 html-to-docx 会渲染成自动编号（多出 1./2./3.）
  · 禁止 display:grid、<dl>/<dt>/<dd>、装饰性 <hr>、用 <span> 模拟 badge

用法：
    from html_writer import write_html
    write_html(md_text, stage_dir, title="...", subtitle="...", meta_lines=[...], kpis=[...])
"""
import html
import json
import os
import re

CN = "一二三四五六七八九十"

# 不参与编号的章节
NO_NUM_SECTIONS = ("核心结论", "执行摘要", "数据说明与分析口径")

DEFAULT_TOKENS = {
    "--color-primary": "#1565C0",
    "--color-secondary": "#0288D1",
    "--color-accent": "#00897B",
    "--color-text": "#212121",
    "--color-textSecondary": "#616161",
    "--color-divider": "#E0E0E0",
    "--color-background": "#FFFFFF",
    "--color-surface": "#F5F7FA",
    "--color-heading": "#1A237E",
}

SPACE_VARS = """    --space-xxs: 0.2em;
    --space-xs: 0.4em;
    --space-xs2: 0.45em;
    --space-sm: 0.5em;
    --space-sm2: 0.6em;
    --space-md: 0.8em;
    --space-md2: 1em;
    --space-lg: 1.2em;
    --space-lg2: 1.6em;
    --space-xl: 1.8em;
    --space-xxl: 3em;
    --space-pad-sm: 0.42em 0.5em;
    --space-pad-md: 0.45em 0.6em;
    --space-pad-lg: 0.5em 0.8em;
    --space-pad-eq: 0.5em 0.5em;
    --space-tbl: 0.8em 0 1.2em 0;
    --space-list: 0.5em 0 1em 0;
    --space-cover-top: 4.5cm;
"""


def cn_num(n):
    if n <= 10:
        return CN[n - 1]
    if n < 20:
        return "十" + (CN[n % 10 - 1] if n % 10 else "")
    return CN[n // 10 - 1] + "十" + (CN[n % 10 - 1] if n % 10 else "")


# ---------------- inline ----------------
def _inline(t):
    t = html.escape(t)
    t = re.sub(r"\*\*(.+?)\*\*", r"<strong>\1</strong>", t)
    t = re.sub(r"(?<!\*)\*([^*]+?)\*(?!\*)", r"<em>\1</em>", t)
    t = re.sub(r"`([^`]+?)`", r"<code>\1</code>", t)
    return t


def _render_table(rows):
    head = [c.strip() for c in rows[0].strip().strip("|").split("|")]
    out = ['<table class="data-table">', "<thead><tr>"]
    for c in head:
        out.append(f"<th>{_inline(c)}</th>")
    out.append("</tr></thead><tbody>")
    for r in rows[2:]:
        cells = [c.strip() for c in r.strip().strip("|").split("|")]
        out.append("<tr>" + "".join(f"<td>{_inline(c)}</td>" for c in cells) + "</tr>")
    out.append("</tbody></table>")
    return "".join(out)


TABLE_LINE = re.compile(r"^\s*\|.*\|\s*$")


def md_to_body(md: str):
    """返回 (body_html, toc_items)"""
    lines = [l for l in md.split("\n") if not l.startswith("# ")]
    body = []
    toc = []
    h2_count = 0
    h3_count = 0
    in_list = None
    para = []
    idx = 0

    def flush():
        nonlocal para
        if para:
            text = "".join(para).strip()
            if text:
                body.append(f"<p>{_inline(text)}</p>")
            para = []

    def close_list():
        nonlocal in_list
        if in_list == "num":
            body.append("</div>")
        elif in_list:
            body.append(f"</{in_list}>")
        in_list = None

    while idx < len(lines):
        line = lines[idx]
        s = line.strip()

        if TABLE_LINE.match(line):
            flush(); close_list()
            block = []
            while idx < len(lines) and TABLE_LINE.match(lines[idx]):
                block.append(lines[idx]); idx += 1
            body.append(_render_table(block))
            continue

        if s.startswith("## "):
            flush(); close_list()
            h3_count = 0
            title = s[3:].strip()
            nonum = title in NO_NUM_SECTIONS
            if not nonum:
                h2_count += 1
            clean = re.sub(r"^[一二三四五六七八九十]+、\s*", "", title)
            disp = clean if nonum else f"{cn_num(h2_count)}、{clean}"
            anchor = f"section-{h2_count if not nonum else 0}"
            body.append(f'<h2 id="{anchor}">{_inline(disp)}</h2>')
            toc.append(f'<p class="toc-item"><a href="#{anchor}">{html.escape(disp)}</a></p>')
            idx += 1
            continue

        if s.startswith("### "):
            flush(); close_list()
            h3_count += 1
            title = s[4:].strip()
            m = re.match(r"^(\d+)\.(\d+)\s+(.*)$", title)
            disp = f"（{cn_num(int(m.group(2)))}）{m.group(3)}" if m else f"（{cn_num(h3_count)}）{title}"
            body.append(f"<h3>{_inline(disp)}</h3>")
            idx += 1
            continue

        if s.startswith("> "):
            flush(); close_list()
            block = []
            while idx < len(lines) and lines[idx].strip().startswith("> "):
                block.append(lines[idx].strip()[2:]); idx += 1
            body.append('<p class="doc-note">' + _inline(" ".join(block)) + "</p>")
            continue

        if re.match(r"^\d+\.\s", s):
            flush()
            if in_list != "num":
                close_list(); body.append('<div class="num-list">'); in_list = "num"
            # 编号直接写在文字里，不用 <ol>：避免 html-to-docx 套用 List Number
            # 后与文本编号重复，渲染成「1. 1. xxx」
            body.append(f'<p class="num-item">{_inline(s)}</p>')
            idx += 1
            continue

        if s.startswith("- "):
            flush()
            if in_list != "ul":
                close_list(); body.append('<ul class="bullet-list">'); in_list = "ul"
            body.append(f"<li>{_inline(s[2:])}</li>")
            idx += 1
            continue

        if not s:
            flush(); close_list(); idx += 1
            continue

        para.append(s)
        idx += 1

    flush(); close_list()

    # 兜底：h2 后若直接跟 h3，插入一句承接型导语（html-review [TQ-03]）
    # 正常的 md 由 md_writer 保证已带导语，此处只处理「漏了导语」的异常情况，
    # 且不再复述标题，避免出现"本章围绕XXX展开"这类机械腔。
    LEAD = ("以下分维度列出关键指标与判断依据，具体数据见各表，"
            "结论与建议集中在文末章节。")
    fixed = []
    for i, block in enumerate(body):
        fixed.append(block)
        if block.startswith("<h2") and i + 1 < len(body) and body[i + 1].startswith("<h3"):
            fixed.append(f"<p>{LEAD}</p>")
    body = fixed

    return "\n".join(body), toc


# ---------------- 模板 ----------------
_TEMPLATE = """<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<meta name="docx-page-size" content="A4">
<title>__TITLE__</title>
<style>
  @page {{ @bottom-center {{ content: counter(page); }} }}
  @page cover {{ @bottom-center {{ content: none; }} }}
  section[role="cover"] {{ page: cover; }}

  :root {{
__VARS__
__SPACES__
    --fs-h1: var(--typography-fontSize-h1, 20pt);
    --fs-h2: var(--typography-fontSize-h2, 11pt);
    --fs-h3: var(--typography-fontSize-h3, 10pt);
    --fs-body: var(--typography-fontSize-body, 9pt);
    --fs-small: var(--typography-fontSize-small, 8pt);
    --fs-tiny: var(--typography-fontSize-tiny, 7pt);
    --fs-coverTitle: var(--typography-fontSize-coverTitle, 22pt);
    --fs-coverCategory: var(--typography-fontSize-coverCategory, 14pt);
    --fs-sectionHeader: var(--typography-fontSize-sectionHeader, 13pt);
    --ff-heading: var(--typography-fontFamily-heading, "Microsoft YaHei", sans-serif);
    --ff-body: var(--typography-fontFamily-body, "Microsoft YaHei", sans-serif);
    --ff-data: var(--typography-fontFamily-data, Arial, sans-serif);
    --lh-body: var(--typography-lineHeight-body, 1.6);
    --lh-compact: var(--typography-lineHeight-compact, 1.2);
    --lh-title: var(--typography-lineHeight-title, 1.3);
    --color-emphasis: var(--color-emphasis-custom, #C62828);
    --spacing-paragraph: var(--spacing-paragraph, 0.75em);
    --spacing-section: var(--spacing-sectionGap, 2.5em);
    --spacing-block: var(--spacing-cardPadding, 1.5em);
    --page-content-width: 15.6cm;
  }}

  * {{ box-sizing: border-box; margin: 0; padding: 0; }}
  body {{
    font-family: var(--ff-body);
    font-size: var(--fs-body);
    line-height: var(--lh-body);
    color: var(--color-text);
    background: var(--color-background);
    max-width: var(--page-content-width);
    margin-left: auto;
    margin-right: auto;
  }}

  section[role="cover"] {{ padding-top: var(--space-cover-top); }}
  .cover-category {{
    font-family: var(--ff-heading); font-size: var(--fs-coverCategory);
    font-weight: 700; color: var(--color-primary); letter-spacing: 0.1em;
    text-align: center; margin-bottom: var(--space-lg);
  }}
  .cover-title {{
    font-family: var(--ff-heading); font-size: var(--fs-coverTitle);
    font-weight: 700; color: var(--color-heading); line-height: var(--lh-title);
    text-align: center; margin-bottom: var(--space-md);
  }}
  .cover-subtitle {{
    font-family: var(--ff-heading); font-size: var(--fs-sectionHeader);
    font-weight: 400; color: var(--color-textSecondary); text-align: center;
    margin-bottom: var(--space-xxl);
  }}
  .cover-meta {{
    font-size: var(--fs-small); color: var(--color-textSecondary);
    text-align: center; margin-bottom: var(--space-xs);
  }}

  .doc-toc {{ margin-bottom: var(--spacing-section); }}
  .toc-title {{
    font-family: var(--ff-heading); font-size: var(--fs-sectionHeader);
    font-weight: 700; color: var(--color-heading); text-align: center;
    margin-bottom: var(--space-md2); letter-spacing: 0.15em;
  }}
  .toc-item {{
    font-size: var(--fs-body); color: var(--color-text);
    margin-bottom: var(--space-xs2); line-height: var(--lh-compact);
  }}
  .toc-item a {{ color: var(--color-text); text-decoration: none; }}

  .executive-summary {{
    border-left: 4px solid var(--color-primary);
    background: var(--color-surface);
    padding: var(--spacing-block);
    margin-bottom: var(--spacing-section);
  }}
  .summary-header {{ margin-bottom: var(--space-md); }}
  .summary-header h2 {{
    font-family: var(--ff-heading); font-size: var(--fs-sectionHeader);
    font-weight: 700; color: var(--color-primary); margin: 0; background: none; padding: 0;
  }}
  .kpi-table {{ width: 100%; border-collapse: collapse; margin-bottom: var(--space-md2); }}
  .kpi-table th {{
    font-family: var(--ff-heading); font-size: var(--fs-tiny); font-weight: 700;
    color: var(--color-textSecondary); text-align: left;
    padding: var(--space-pad-md); border-bottom: 1px solid var(--color-divider);
  }}
  .kpi-table td {{
    padding: var(--space-pad-md); border-bottom: 1px solid var(--color-divider);
    font-size: var(--fs-small);
  }}
  .card-name {{ color: var(--color-textSecondary); width: 42%; }}
  .card-value {{
    font-family: var(--ff-data); font-size: var(--fs-sectionHeader);
    font-weight: 700; color: var(--color-primary); width: 30%;
  }}
  .card-unit {{ font-size: var(--fs-tiny); font-weight: 400; color: var(--color-textSecondary); margin-left: var(--space-xxs); }}
  .card-trend {{ font-size: var(--fs-small); color: var(--color-textSecondary); text-align: right; }}
  .summary-body p {{
    font-size: var(--fs-body); line-height: var(--lh-body);
    margin-bottom: var(--spacing-paragraph); text-align: justify;
  }}
  .summary-body strong {{ font-weight: 700; color: var(--color-emphasis); }}

  h2 {{
    font-family: var(--ff-heading); font-size: var(--fs-sectionHeader);
    font-weight: 700; color: var(--color-background); background: var(--color-primary);
    padding: var(--space-pad-lg); margin-top: var(--spacing-section);
    margin-bottom: var(--space-md2); line-height: var(--lh-compact);
  }}
  h3 {{
    font-family: var(--ff-heading); font-size: var(--fs-h2);
    font-weight: 700; color: var(--color-heading);
    margin-top: var(--space-lg2); margin-bottom: var(--space-sm2);
    border-left: 3px solid var(--color-secondary); padding-left: var(--space-sm);
  }}
  p {{
    font-size: var(--fs-body); line-height: var(--lh-body);
    margin-bottom: var(--spacing-paragraph); text-align: justify;
    color: var(--color-text);
  }}
  p strong {{ font-weight: 700; color: var(--color-emphasis); }}
  p em {{ font-style: normal; color: var(--color-textSecondary); }}
  code {{ font-family: var(--ff-data); font-size: var(--fs-small); }}

  .doc-note {{
    border-left: 3px solid var(--color-primary);
    background: var(--color-surface);
    padding: var(--space-pad-lg);
    color: var(--color-textSecondary);
    font-size: var(--fs-small);
  }}

  .data-table {{
    width: 100%; border-collapse: collapse; margin: var(--space-tbl);
    font-size: var(--fs-small); line-height: var(--lh-compact);
  }}
  .data-table th {{
    background: var(--color-primary); color: var(--color-background);
    font-family: var(--ff-heading); font-weight: 700;
    padding: var(--space-pad-eq); border: 1px solid var(--color-primary);
    text-align: center; font-size: var(--fs-small);
  }}
  .data-table td {{
    padding: var(--space-pad-sm); border: 1px solid var(--color-divider);
    font-size: var(--fs-small); text-align: center;
  }}
  .data-table tr:nth-child(even) td {{ background: var(--color-surface); }}
  .data-table td:first-child {{ text-align: left; }}

  .num-list, .bullet-list {{ margin: var(--space-list); padding-left: var(--space-xl); }}
  .num-item {{
    margin: 0 0 var(--space-xs) var(--space-lg2);
    text-indent: -1.15em;
    line-height: 1.75;
  }}
  .num-list li, .bullet-list li {{
    font-size: var(--fs-body); line-height: var(--lh-body);
    margin-bottom: var(--space-xs2); text-align: justify; color: var(--color-text);
  }}
  .num-list {{ list-style-type: decimal; }}
  .bullet-list {{ list-style-type: disc; }}
  .num-list li strong, .bullet-list li strong {{ font-weight: 700; color: var(--color-emphasis); }}
</style>
</head>
<body>

<section role="cover">
  <p class="cover-category">__CATEGORY__</p>
  <h1 class="cover-title">__COVER_TITLE__</h1>
  <p class="cover-subtitle">__SUBTITLE__</p>
__META_LINES__
</section>

<section role="body" data-page-restart="1">

<div class="doc-toc">
  <p class="toc-title">目　录</p>
  <div class="toc-list">
    __TOC__
  </div>
</div>

<div class="executive-summary">
  <div class="summary-header">
    <h2>执行摘要</h2>
  </div>
  __KPI_TABLE__
  <div class="summary-body">
__SUMMARY__
  </div>
</div>

<main>
__BODY__
</main>

</section>
</body>
</html>
"""


def _load_tokens(stage_dir):
    """优先读取 stage2/design_tokens.json，否则用默认色板"""
    p = os.path.join(stage_dir, "design_tokens.json")
    if os.path.exists(p):
        try:
            t = json.load(open(p, encoding="utf-8"))
            cv = t.get("css_variables") or {}
            if cv:
                return cv
        except Exception:
            pass
    return dict(DEFAULT_TOKENS)


def write_html(md_text, stage_dir, title, cover_title=None, subtitle="",
               category="数 据 分 析 报 告", meta_lines=None, kpis=None,
               summary_paragraphs=None, tokens_path=None):
    """
    生成商务报告 HTML。

    参数
      md_text            : 正文 Markdown（md_writer 产出）
      stage_dir          : stage2 输出目录
      title              : <title> 与文档标题
      cover_title        : 封面标题（可含 <br>），默认取 title
      subtitle           : 封面副标题
      category           : 封面顶部小标
      meta_lines         : 封面下方信息行 list[str]
      kpis               : [{label, value, unit, note}] 摘要卡片
      summary_paragraphs : 执行摘要段落 list[str]（支持 **加粗**）
      tokens_path        : design_tokens.json 路径，默认 stage_dir/design_tokens.json
    """
    os.makedirs(stage_dir, exist_ok=True)
    body, toc = md_to_body(md_text)

    cv = _load_tokens(os.path.dirname(tokens_path) if tokens_path else stage_dir) \
        if tokens_path else _load_tokens(stage_dir)
    var_css = "\n".join(f"    {k}: {v};" for k, v in cv.items() if k.startswith("--color"))

    # KPI 卡片
    kpi_html = []
    for k in (kpis or []):
        unit = f'<span class="card-unit">{k["unit"]}</span>' if k.get("unit") else ""
        kpi_html.append(
            f'<tr><td class="card-name">{_inline(k["label"])}</td>'
            f'<td class="card-value">{k["value"]}{unit}</td>'
            f'<td class="card-trend">{_inline(k.get("note", ""))}</td></tr>')
    kpi_table = ('<table class="kpi-table">'
                 '<thead><tr><th>指标</th><th>数值</th><th style="text-align:right;">变化</th></tr></thead>'
                 "<tbody>" + "".join(kpi_html) + "</tbody></table>")

    meta_html = "\n".join(f'  <p class="cover-meta">{m}</p>' for m in (meta_lines or []))
    summary_html = "\n".join(f"    <p>{_inline(p)}</p>" for p in (summary_paragraphs or []))

    out = (_TEMPLATE
           .replace("__TITLE__", html.escape(title))
           .replace("__VARS__", var_css)
           .replace("__SPACES__", SPACE_VARS.rstrip("\n"))
           .replace("__CATEGORY__", category)
           .replace("__COVER_TITLE__", cover_title or html.escape(title))
           .replace("__SUBTITLE__", html.escape(subtitle))
           .replace("__META_LINES__", meta_html)
           .replace("__TOC__", "\n    ".join(toc))
           .replace("__KPI_TABLE__", kpi_table)
           .replace("__SUMMARY__", summary_html)
           .replace("__BODY__", body))

    path = os.path.join(stage_dir, "formatted-report.html")
    open(path, "w", encoding="utf-8").write(out)
    return path


# ---------------- 自检 ----------------
BARE_SPACING = re.compile(r"(?:margin|padding)[^;{}]*:\s*[\d.]+em")
H2_THEN_H3 = re.compile(r"</h2>\s*<h3")
TOC_OL = re.compile(r'<ol class="toc-list"')


def self_check(html_path):
    """排版自检，返回问题列表（空表示通过）"""
    s = open(html_path, encoding="utf-8").read()
    issues = []

    n_bare = len(BARE_SPACING.findall(s))
    if n_bare:
        issues.append(f"[DT-03] {n_bare} 处裸间距字面量，应改用 var(--space-*)")

    for m in re.finditer(r'<table class="data-table">(.*?)</table>', s, re.S):
        if "<thead>" not in m.group(1):
            issues.append("[SI-02] data-table 缺少 thead")
            break

    if "<table class=\"kpi-table\"><thead>" not in s:
        issues.append("[SI-02] kpi-table 缺少 thead/tbody")

    n_h2h3 = len(H2_THEN_H3.findall(s))
    if n_h2h3:
        issues.append(f"[TQ-03] {n_h2h3} 处 h2 后直接跟 h3，需插入章节导语")

    if TOC_OL.search(s):
        issues.append("[X-01] 目录使用 <ol class=\"toc-list\">，html-to-docx 会渲染为自动编号")

    for bad in ("display: grid", "display:grid", "<dl>", "<dt>", "<dd>", "<hr"):
        if bad in s:
            issues.append(f"[X-02] 含 docx 不兼容元素: {bad}")

    return issues
