# -*- coding: utf-8 -*-
"""
一键把本仓库安装为 WorkBuddy / Claude 技能。

用法：
    python scripts/install_skill.py
    python scripts/install_skill.py --target "D:/my/skills" --name ztny-analysis
    python scripts/install_skill.py --dry-run

默认安装到 ~/.workbuddy/skills/procurement-analysis（SKILL.md 的 name 字段），
已存在时会先备份为 <name>.bak.<时间戳>。
"""
import argparse
import datetime
import os
import shutil
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
COPY_DIRS = ("scripts", "references", "assets")
COPY_FILES = ("SKILL.md", "README.md", "LICENSE", "requirements.txt", "CHANGELOG.md")
IGNORE = shutil.ignore_patterns("__pycache__", "*.pyc", "*.pyo")


def main():
    ap = argparse.ArgumentParser(description="安装 ZTNY 采购分析技能")
    ap.add_argument("--target", default=os.path.expanduser("~/.workbuddy/skills"),
                    help="技能根目录，默认 ~/.workbuddy/skills")
    ap.add_argument("--name", default="procurement-analysis",
                    help="技能目录名，默认 procurement-analysis")
    ap.add_argument("--dry-run", action="store_true", help="只打印将要执行的操作")
    args = ap.parse_args()

    dest = os.path.join(os.path.abspath(args.target), args.name)

    print(f"源目录：{ROOT}")
    print(f"目标  ：{dest}")

    if args.dry_run:
        for d in COPY_DIRS:
            print(f"  + {d}/")
        for f in COPY_FILES:
            if os.path.exists(os.path.join(ROOT, f)):
                print(f"  + {f}")
        return 0

    if os.path.isdir(dest):
        stamp = datetime.datetime.now().strftime("%Y%m%d%H%M%S")
        backup = f"{dest}.bak.{stamp}"
        shutil.move(dest, backup)
        print(f"已备份旧版本到 {backup}")

    os.makedirs(dest, exist_ok=True)
    for d in COPY_DIRS:
        src = os.path.join(ROOT, d)
        if os.path.isdir(src):
            shutil.copytree(src, os.path.join(dest, d), ignore=IGNORE)
    for f in COPY_FILES:
        src = os.path.join(ROOT, f)
        if os.path.isfile(src):
            shutil.copy2(src, os.path.join(dest, f))

    print(f"\n✅ 安装完成：{dest}")
    print("重启会话后，用自然语言说「分析一下 X 月的采购数据」即可触发。")
    return 0


if __name__ == "__main__":
    sys.exit(main())
