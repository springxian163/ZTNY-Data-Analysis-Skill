# -*- coding: utf-8 -*-
"""把 html_writer._TEMPLATE 导出为可独立查看的模板文件"""
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__))))
import html_writer

root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
p = os.path.join(root, "assets", "report_template.html")
os.makedirs(os.path.dirname(p), exist_ok=True)
with open(p, "w", encoding="utf-8") as f:
    f.write(html_writer._TEMPLATE)
print("已导出:", p, os.path.getsize(p), "bytes")
