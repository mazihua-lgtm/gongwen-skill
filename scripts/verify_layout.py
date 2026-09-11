"""公文版式核验 — 基于 pdftotext -bbox 的坐标检查(比视觉模型可靠)。

    python3 verify_layout.py <file.pdf>

检查项(容差 ±6pt):
1. 标题基本居中(左空 ≈ 右空)
2. 正文首行缩进:左边距 70.87pt(2.5cm)+ 缩进 32.03pt(1.13cm)≈ 102.9pt
3. 页码「— 1 —」存在

需要 poppler 的 pdftotext(brew install poppler)。
"""

from __future__ import annotations

import re
import subprocess
import sys
import xml.etree.ElementTree as ET

LEFT_MARGIN = 70.87   # 2.5cm
INDENT = 32.03        # 1.13cm
EXPECTED_X = LEFT_MARGIN + INDENT  # ≈ 102.9pt
TOL = 6.0


def bbox_pages(path: str) -> list[ET.Element]:
    out = subprocess.run(
        ["pdftotext", "-bbox", path, "-"],
        capture_output=True, text=True, check=True,
    ).stdout
    root = ET.fromstring(re.sub(r'xmlns="[^"]+"', "", out))  # 剥掉默认命名空间
    return root.findall(".//page")


def main() -> int:
    if len(sys.argv) != 2:
        print(__doc__)
        return 2
    try:
        pages = bbox_pages(sys.argv[1])
    except FileNotFoundError:
        print("SKIP 需要 pdftotext: brew install poppler")
        return 2
    except subprocess.CalledProcessError as e:
        print(f"FAIL pdftotext 失败: {e.stderr.strip()[:120]}(输入必须是 PDF)")
        return 1

    page_w = float(pages[0].get("width"))
    ok = True

    # 首页逐词 (text, xMin, yMin, xMax)
    pw = [
        (w.text or "", float(w.get("xMin")), float(w.get("yMin")), float(w.get("xMax")))
        for w in pages[0].findall(".//word")
    ]
    if not pw:
        print("FAIL 首页无文字")
        return 1

    # 1) 标题居中:取最顶行
    top_y = min(y for _, _, y, _ in pw)
    line = sorted([(t, x0, x1) for t, x0, y, x1 in pw if y - top_y < 2.0], key=lambda r: r[1])
    gap_l = line[0][1]
    gap_r = page_w - max(x1 for _, _, x1 in line)
    if abs(gap_l - gap_r) <= TOL + 4:
        print(f"PASS 标题居中 (左空 {gap_l:.1f}pt / 右空 {gap_r:.1f}pt)")
    else:
        print(f"FAIL 标题不居中 (左空 {gap_l:.1f}pt / 右空 {gap_r:.1f}pt)")
        ok = False

    # 2) 正文首行缩进:标题行之下,行首 x ≈ 102.9pt
    body_xs = [x0 for _, x0, y, _ in pw if y > top_y + 10]
    indented = any(abs(x - EXPECTED_X) <= TOL for x in body_xs)
    print(("PASS " if indented else "FAIL ") + f"首行缩进 ≈{EXPECTED_X:.1f}pt (采样 {len(body_xs)} 词)")
    ok = ok and indented

    # 3) 页码(全文任意页)
    all_pages = [
        (w.text or "")
        for p in pages
        for w in p.findall(".//word")
    ]
    has_pagenum = bool(re.search(r"—\s*1\s*—", " ".join(all_pages)))
    print(("PASS " if has_pagenum else "FAIL ") + "页码「— 1 —」存在")
    ok = ok and has_pagenum

    print("\n== 版式核验:", "PASS ==" if ok else "FAIL ==")
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
