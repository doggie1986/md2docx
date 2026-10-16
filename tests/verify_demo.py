# -*- coding: utf-8 -*-
"""端到端核对：拿真样例 ``samples/demo.md`` 走一遍命令行，51 项判据全读**落盘的 .docx**。

跑法（不参与 ``unittest discover``，文件名不是 ``test_*``）::

    python tests/verify_demo.py

三方对照：① demo.md 里逐行数出来的（写在下面 EXPECTED，注释标了行号）；
② 命令行回执念的；③ zip 包里 document.xml 数出来的。
三方对得上，"公式在 Word 里是原生对象"这句话才算有证据。
"""
from __future__ import annotations

import re
import subprocess
import sys
import tempfile
import time
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

ROOT = Path(__file__).resolve().parent.parent
W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
M = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"
PY = sys.executable

#: demo.md 逐行数出来的（代码段里的 `$$…$$`、`\[…\]` 不算公式，所以第 54 行那两处不在此列）
EXPECTED = {
    "inline": 14,      # 8/10/12/13/15 各 2/2/1/2/1，48-51 各 1，59 一条，105 一条（这条是故意写坏的）
    "display": 5,      # 21、23-25、29-30、58（在单元格里）、103
    "tables": 3,       # 36-40、46-51、56-59
    "headings": 10,    # 1、6、17、32、34、44、61、79、90、98
    "code_blocks": 1,  # 83-85
    "failed": 1,       # 105 的 $\foo \bar$
}
EXPECTED["total"] = EXPECTED["inline"] + EXPECTED["display"]
EXPECTED["ok"] = EXPECTED["total"] - EXPECTED["failed"]

ROWS: list[tuple[str, str, str, bool]] = []


def check(name: str, expected: object, actual: object) -> None:
    ok = str(expected) == str(actual)
    ROWS.append((name, str(expected), str(actual), ok))


def note(name: str, actual: object) -> None:
    ROWS.append((name, "（只念，不判）", str(actual), True))


def text_of(element: ET.Element) -> str:
    return "".join(element.itertext())


PARENTS: dict[int, ET.Element] = {}


def index_parents(root: ET.Element) -> None:
    PARENTS.clear()
    for parent in root.iter():
        for child in list(parent):
            PARENTS[id(child)] = parent


def jc_of(element: ET.Element | None) -> str:
    """段落/表格 Properties 里 ``w:jc`` 那一格（正文段落在 pPr 下，单元格里在 tcPr 下）。"""
    if element is None:
        return ""
    node = element.find("%s%s" % (W, "jc"))
    if node is None:
        props = element.find("%s%s" % (W, "pPr"))
        node = props.find("%s%s" % (W, "jc")) if props is not None else None
    return node.get(W + "val", "") if node is not None else ""


def ancestors(node: ET.Element) -> list[str]:
    tags = []
    current = PARENTS.get(id(node))
    while current is not None:
        tags.append(current.tag)
        current = PARENTS.get(id(current))
    return tags


def load(path: Path) -> dict:
    with zipfile.ZipFile(str(path)) as archive:
        names = list(archive.namelist())
        document = ET.fromstring(archive.read("word/document.xml"))
        styles = archive.read("word/styles.xml").decode("utf-8")
    index_parents(document)
    return {
        "names": names,
        "document": document,
        "styles": styles,
        "maths": document.findall(".//%soMath" % M),
        "paras": document.findall(".//%soMathPara" % M),
        "tables": document.findall(".//%stbl" % W),
    }


def receipt_numbers(text: str) -> dict:
    main = re.search(
        r"块 (\d+) 个｜表格 (\d+) 张｜公式 (\d+) 条（转成 (\d+)，失败 (\d+)）｜超链接 (\d+) 个"
        r"｜图片 (\d+) 处（嵌进 (\d+)，没成 (\d+)）", text)
    parse = re.search(
        r"解析计数：块 (\d+)｜标题 (\d+)｜行内公式 (\d+)｜显示式 (\d+)｜代码块 (\d+)｜表格 (\d+)", text)
    keys = ["blocks", "tables", "math_total", "math_ok", "math_failed", "links",
            "images", "images_embedded", "images_skipped"]
    out = {k: int(v) for k, v in zip(keys, main.groups())} if main else {}
    if parse:
        out.update({k: int(v) for k, v in zip(
            ["p_blocks", "p_headings", "p_inline", "p_display", "p_code", "p_tables"], parse.groups())})
    return out


def run_cli(argv: list[str], cwd: Path) -> tuple[int, str, str, float]:
    started = time.time()
    process = subprocess.run([PY, *argv], cwd=str(cwd), capture_output=True,
                             text=True, encoding="utf-8", errors="replace")
    return process.returncode, process.stdout, process.stderr, time.time() - started


def main() -> int:
    work = Path(tempfile.mkdtemp(prefix="md2docx-b5-"))
    demo = ROOT / "samples" / "demo.md"
    target = work / "新建" / "demo.docx"  # 父目录不存在，看命令行会不会自己建

    code, out, err, seconds = run_cli(
        [str(ROOT / "md2docx.py"), str(demo), str(target), "--body-font", "宋体", "--body-size", "12"],
        cwd=work)  # 从别的工作目录跑：样式表靠自动定位，不靠 cwd
    print("=== 首轮回执（退出码 %d，%.2f 秒）===" % (code, seconds))
    print(out.rstrip())
    if err.strip():
        print("--- stderr ---")
        print(err.rstrip())
    print("=== 取证 ===")

    check("① 从别的工作目录跑 md2docx.py 的退出码", 0, code)
    check("② 输出父目录不存在也建出来了", True, target.is_file())
    numbers = receipt_numbers(out)
    pkg = load(target)

    # —— 三方对照：手工数 / 回执 / 文件
    check("③ 回执的行内条数 = demo.md 里数出的", EXPECTED["inline"], numbers.get("p_inline"))
    check("④ 回执的显示式条数 = demo.md 里数出的", EXPECTED["display"], numbers.get("p_display"))
    check("⑤ 回执的标题/表格/代码块", "10/3/1", "%s/%s/%s" % (
        numbers.get("p_headings"), numbers.get("p_tables"), numbers.get("p_code")))
    check("⑥ 回执「公式」= 行内 + 显示式", EXPECTED["total"], numbers.get("math_total"))
    check("⑦ 回执「失败」= 故意写坏的那一条", EXPECTED["failed"], numbers.get("math_failed"))
    check("⑧ 文件里 m:oMath 条数 = 回执「转成」", EXPECTED["ok"], len(pkg["maths"]))
    check("⑨ 文件里 m:oMathPara 条数 = 显示式（坏的那条是行内）", EXPECTED["display"], len(pkg["paras"]))
    check("⑩ 文件里 w:tbl 条数 = 回执「表格」", numbers.get("tables"), len(pkg["tables"]))
    check("⑪ 自检那一句念的是「一致」", True, "与内存里转成的条数一致" in out and "不一致" not in out)

    # —— 结构不变式
    off_center = [text_of(node)[:20] for node in pkg["paras"] if jc_of(PARENTS.get(id(node))) != "center"]
    check("⑫ 每条 m:oMathPara 所在段落都居中（不居中的那几条列出来）", 0, len(off_center))
    in_cells = sum(1 for node in pkg["maths"] if (W + "tc") in ancestors(node))
    check("⑬ 单元格里的公式条数（祖先链含 w:tc：48-51 四条文内 + 59 一条 + 58 一条显示式）", 6, in_cells)
    runs = pkg["document"].findall(".//%sr" % M)
    no_font = [text_of(run)[:20] for run in runs
               if (run.find("%s%s/%s%s" % (W, "rPr", W, "rFonts")) is None
                   or run.find("%s%s/%s%s" % (W, "rPr", W, "rFonts")).get(W + "ascii") != "Cambria Math")]
    check("⑭ 每个 m:r 都声明 Cambria Math（共 %d 个 m:r，缺的列出来）" % len(runs), 0, len(no_font))
    loose = [text_of(node) for node in pkg["document"].findall(".//%st" % M) if node.get(XML_SPACE) != "preserve"]
    check("⑮ 每个 m:t 都带 xml:space=preserve", 0, len(loose))
    media = [n for n in pkg["names"] if n.startswith("word/media")]
    check("⑯ 本地那张真进了包（媒体文件 1 个）且回执把三格分开报了（处数/嵌进/没成）",
          "1 / 3 / 1 / 2",
          "%d / %d / %d / %d" % (len(media), numbers.get("images"),
                                 numbers.get("images_embedded"), numbers.get("images_skipped")))

    # —— 样式真的落进了 styles.xml
    check("⑰ 正文中文字体走 w:eastAsia=宋体", True, 'w:eastAsia="宋体"' in pkg["styles"])
    check("⑱ 正文 12pt = w:sz 24（半磅）", True, '<w:sz w:val="24"/>' in pkg["styles"])
    check("⑲ 标题字体 黑体", True, 'w:eastAsia="黑体"' in pkg["styles"])
    normal = re.search(r'<w:style [^>]*w:styleId="Normal">.*?</w:style>', pkg["styles"], re.S)
    head1 = re.search(r'<w:style [^>]*w:styleId="Heading1">.*?</w:style>', pkg["styles"], re.S)
    check("⑳ Normal / Heading1 里没有主题字体与主题色残留", "0 / 0",
          "%d / %d" % (normal.group(0).count("Theme") + normal.group(0).count("themeColor"),
                       head1.group(0).count("Theme") + head1.group(0).count("themeColor")))
    sizes = [int(v) for v in re.findall(r'<w:sz w:val="(\d+)"', head1.group(0))]
    check("㉑ Heading1 字号比正文大（> 24 半磅）", True, any(v > 24 for v in sizes))

    # —— 正文内容
    body = text_of(pkg["document"])
    for probe in ["电力工程材料统计——转换样例", "钢芯铝绞线", "已发布的清单不得更改。", "综合单价", "高低压柜"]:
        check("㉒ 原文照在正文里：%s" % probe, True, probe in body)
    check("㉓ 图片那三处各有说法：嵌成的不留占位、缺的念出查过的位置、远程的念「这不是本机的文件」",
          True,
          "line.dia" in body
          and "这台机器上没找到这个文件" in body
          and "这不是本机的文件" in body
          and "图片没嵌进去：单线示意图" not in body)
    check("㉔ 坏公式退回原文（连定界符一起写回正文，好让人找到那一格）", True, "$\\foo \\bar$" in body)
    highlights = pkg["document"].findall(".//%s%s" % (W, "highlight"))
    check("㉕ 坏公式带黄色底纹（黄底只给坏的那一条）", "1 处 / 全是 yellow", "%d 处 / %s" % (
        len(highlights), "全是 yellow" if {h.get(W + "val") for h in highlights} == {"yellow"} else
        "混进了 %s" % sorted({h.get(W + "val") for h in highlights} - {"yellow"})))
    check("㉖ 定界符没漏进正文（只有第 54 行代码段里 \`$$…$$\` 的 4 个 $ + 退回原文那 2 个）", "6 / 0 / 1",
          "%d / %d / %d" % (body.count("$"), body.count(r"\("), body.count(r"\[")))
    paragraph_jc = ".//%s%s/%s%s/%s%s" % (W, "p", W, "pPr", W, "jc")
    right = sum(1 for cell in pkg["document"].findall(".//%s%s" % (W, "tc"))
                if any(j.get(W + "val") == "right" for j in cell.findall(paragraph_jc)))
    centered_cells = sum(1 for cell in pkg["document"].findall(".//%s%s" % (W, "tc"))
                         if any(j.get(W + "val") == "center" for j in cell.findall(paragraph_jc)))
    check("㉗ 数字列按 --: 右对齐、:-: 那一列居中（两类都得有）", True, right > 0 and centered_cells > 0)
    note("㉗-实测 右对齐 / 居中的格子数", "%d / %d" % (right, centered_cells))

    # —— python-docx 读回（包合法 + 样式名对得上）
    import docx

    reopened = docx.Document(str(target))
    used = {p.style.name for p in reopened.paragraphs}
    check("㉘ python-docx 能打开这份包", True, len(reopened.paragraphs) > 0)
    check("㉙ 标题落到 Heading 1/2/3", "全在", "%s" % (
        "全在" if {"Heading 1", "Heading 2", "Heading 3"} <= used else "缺 %s" % (
            {"Heading 1", "Heading 2", "Heading 3"} - used)))
    check("㉚ 表格数（python-docx 数一遍）", len(pkg["tables"]), len(reopened.tables))

    # —— 重跑同一份稿子
    again = work / "demo2.docx"
    code2, out2, _, _ = run_cli([str(ROOT / "md2docx.py"), str(demo), str(again)], cwd=work)
    pkg2 = load(again)
    check("㉛ 重跑一次仍成功", 0, code2)
    check("㉜ 两次文件里的公式条数不变", "%d/%d" % (len(pkg["maths"]), len(pkg["paras"])),
          "%d/%d" % (len(pkg2["maths"]), len(pkg2["paras"])))
    note("㉝ 两次字节数（核心属性里有时间戳，只念）", "%d vs %d" % (target.stat().st_size, again.stat().st_size))

    # —— 坏公式的两档与其余退出码
    bad = work / "bad.md"
    bad.write_text("正常 $a+b$，坏掉 $\\foo \\bar$。\n", encoding="utf-8")
    code3, out3, _, _ = run_cli([str(ROOT / "md2docx.py"), str(bad), str(work / "默认档.docx")], cwd=work)
    pkg3 = load(work / "默认档.docx")
    check("㉞ 默认档：一条坏公式不拦整份文档", 0, code3)
    check("㉟ 默认档：好的那条仍转成原生对象", 1, len(pkg3["maths"]))
    strict = work / "strict.docx"
    code4, out4, _, _ = run_cli([str(ROOT / "md2docx.py"), str(bad), str(strict), "--strict"], cwd=work)
    check("㊱ --strict：退出码单独一档", 5, code4)
    check("㊲ --strict：判失败就不留半成品在盘上", False, strict.exists())
    code5, out5, _, _ = run_cli(
        [str(ROOT / "md2docx.py"), str(bad), str(work / "无样式表.docx"),
         "--xslt-path", str(work / "无此文件.XSL")], cwd=work)
    check("㊳ 样式表找不到 → 退出码 3", 3, code5)
    check("㊴ 且「解决办法」只念一遍（回归：曾重复过）", 1, out5.count("解决办法"))
    code6, _, _, _ = run_cli([str(ROOT / "md2docx.py"), str(work / "没有这个.md"), str(work / "x.docx")], cwd=work)
    check("㊵ 输入读不进 → 退出码 4", 4, code6)
    code7, out7, _, _ = run_cli([str(ROOT / "md2docx.py"), str(bad), str(bad)], cwd=work)
    check("㊶ 输入与输出同名 → 退出码 2 且不改稿子", "2 / True", "%d / %s" % (code7, "$\\foo \\bar$" in bad.read_text(encoding="utf-8")))

    # —— 规模化：20 张表 × 15 行 = 300 条公式（表头行与数据行之间不能有空行，那会断成两张表）
    tables = []
    for i in range(20):
        rows = ["## 第 %d 张表" % (i + 1), "", "| 项 | 式 |", "| --- | --- |"]
        rows.extend("| 第%d行 | $q_{%d_%d} = %d \\times 2$ |" % (j + 1, i, j, j) for j in range(15))
        tables.append("\n".join(rows))
    big = work / "big.md"
    big.write_text("\n\n".join(tables) + "\n", encoding="utf-8")
    big_target = work / "big.docx"
    code8, out8, _, took = run_cli([str(ROOT / "md2docx.py"), str(big), str(big_target)], cwd=work)
    pkg8 = load(big_target)
    check("㊷ 300 条公式的稿子退出码 0", 0, code8)
    check("㊸ 300 条全部原生落进文件", 300, len(pkg8["maths"]))
    check("㊹ 20 张表都在", 20, len(pkg8["tables"]))
    check("㊺ 全部在单元格内（一条都没越界）", 0, sum(1 for node in pkg8["maths"] if (W + "tc") not in ancestors(node)))
    note("㊻ 300 条那一份的耗时与体积", "%.2f 秒｜%d 字节" % (took, big_target.stat().st_size))

    print()
    diff = [row for row in ROWS if not row[3]]
    for name, expected, actual, ok in ROWS:
        print("%s %s｜期望 %s｜实测 %s" % ("OK  " if ok else "差异", name, expected, actual))
    print("\n临时目录：%s" % work)
    print("合计 比对项 %d｜差异 %d" % (len(ROWS), len(diff)))
    return 0 if not diff else 1


if __name__ == "__main__":
    sys.exit(main())
