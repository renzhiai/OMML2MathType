# -*- coding: utf-8 -*-
"""OMML2MathType —— 把 Word 内置公式（OMML）批量转换成 MathType 公式对象

原理（两段式，缺一不可）：
  ① Python 层：把每个 m:oMath 换成纯文本 run "$TeX$"。
     必须在 XML 层换 —— 在 Word 里直接改公式的 Range.Text 仍落在数学区内，
     触发不了 MathType 的 TeX 切换。
  ② Word 层：逐段定位 "$...$"，选中后执行 MathType 宏
     MathTypeCommands.UILib.MTCommand_TeXToggle → 生成 Equation.DSMT4 对象。

用法：
  python omml2mathtype.py 文档.docx                # 输出 文档_MathType.docx
  python omml2mathtype.py 文档.docx --inplace      # 覆盖原文件（自动备份）
  python omml2mathtype.py 文档.docx --dry-run      # 只列出公式的 TeX，不动文件
  python omml2mathtype.py 文档.docx -o out.docx --max-width 340
"""
from __future__ import annotations

import argparse
import json
import os
import re
import shutil
import sys
import time
import traceback
from datetime import datetime

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from omml2tex import omml_to_tex  # noqa: E402

MATH_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"
TOGGLE = "MathTypeCommands.UILib.MTCommand_TeXToggle"
SENTINEL = "\uff04"          # 全角 ＄，用于临时屏蔽正文里原有的 $ 符号
DEFAULT_PROGID = "Equation.DSMT4"


# ----------------------------------------------------------------- 基础工具
def _in_math(el) -> bool:
    p = el.getparent()
    while p is not None:
        if p.tag == MATH_NS + "oMath":
            return True
        p = p.getparent()
    return False


def log(msg: str = "", quiet: bool = False):
    if not quiet:
        print(msg, flush=True)


# ----------------------------------------------------------------- ① Python 层
def collect_tex(doc: Document):
    return [omml_to_tex(om) for om in doc.element.body.iter(MATH_NS + "oMath")]


def shield_dollars(doc: Document) -> int:
    """正文里原有的 $ 换成全角 ＄，避免被当成公式定界符"""
    n = 0
    for t in doc.element.body.iter(qn("w:t")):
        if not t.text or "$" not in t.text or _in_math(t):
            continue
        n += t.text.count("$")
        t.text = t.text.replace("$", SENTINEL)
        t.set(qn("xml:space"), "preserve")
    return n


def replace_math_with_text(doc: Document, texs) -> None:
    oms = list(doc.element.body.iter(MATH_NS + "oMath"))
    if len(oms) != len(texs):
        raise RuntimeError("公式数量不一致（%d vs %d），文件可能已被修改" % (len(oms), len(texs)))
    for om, tex in zip(oms, texs):
        parent = om.getparent()
        target = parent if parent.tag == MATH_NS + "oMathPara" else om
        holder = target.getparent()
        idx = holder.index(target)
        run = OxmlElement("w:r")
        t = OxmlElement("w:t")
        t.set(qn("xml:space"), "preserve")
        t.text = "$%s$" % tex
        run.append(t)
        holder.remove(target)
        holder.insert(idx, run)


def prepare(src: str, work: str, quiet=False) -> list:
    """生成中间文件：所有 OMML 变成 $TeX$ 纯文本"""
    doc = Document(src)
    texs = collect_tex(doc)
    if not texs:
        raise RuntimeError("文档里没有找到 Word 公式（OMML），无需转换")
    shielded = shield_dollars(doc)
    if shielded:
        log("  提示：正文原有 %d 个 $ 符号，已临时屏蔽，转换后自动还原" % shielded, quiet)
    replace_math_with_text(doc, texs)
    doc.save(work)
    return texs


# ----------------------------------------------------------------- ② Word 层
def _connect_word():
    import win32com.client as win32
    app = win32.DispatchEx("Word.Application")
    app.Visible = False
    app.DisplayAlerts = 0
    return app


def toggle_all(work: str, progid: str, quiet=False):
    """逐段把 $TeX$ 转成 MathType 对象"""
    app = _connect_word()
    doc = app.Documents.Open(os.path.abspath(work))
    ok = fail = 0
    fails = []
    total_para = doc.Paragraphs.Count
    try:
        last_report = time.time()
        for i in range(1, total_para + 1):
            rng = doc.Paragraphs(i).Range
            text = rng.Text
            if "$" not in text:
                continue
            tries = 0
            while "$" in text and tries < 8:
                tries += 1
                a = text.find("$")
                b = text.find("$", a + 1)
                if b < 0:
                    break
                piece = text[a:b + 1]
                r = doc.Range(rng.Start + a, rng.Start + b + 1)
                r.Select()
                n0 = doc.InlineShapes.Count
                try:
                    app.Run(TOGGLE)
                    time.sleep(0.08)
                except Exception as e:
                    fails.append({"para": i, "tex": piece[:120], "error": str(e)[:120]})
                    fail += 1
                    break
                if doc.InlineShapes.Count > n0:
                    ok += 1
                    rng = doc.Paragraphs(i).Range
                    text = rng.Text
                else:
                    fails.append({"para": i, "tex": piece[:120], "error": "未生成对象"})
                    fail += 1
                    break
            if not quiet and time.time() - last_report > 3:
                log("  … 已转换 %d 个（当前第 %d/%d 段）" % (ok, i, total_para))
                last_report = time.time()

        # 还原被临时屏蔽的 $（此时文档里已不存在公式定界符）
        try:
            f = doc.Content.Find
            f.ClearFormatting()
            f.Replacement.ClearFormatting()
            f.Text = SENTINEL
            f.Replacement.Text = "$"
            f.Forward = True
            f.Wrap = 1  # wdFindContinue
            f.Execute(Replace=2)  # wdReplaceAll
        except Exception:
            pass
        doc.Save()
    finally:
        doc.Close(False)
        app.Quit()
    return ok, fail, fails


# ----------------------------------------------------------------- 审计
def audit_doc(path: str, progid: str, max_width=None):
    app = _connect_word()
    doc = app.Documents.Open(os.path.abspath(path))
    try:
        ps = doc.PageSetup
        body_w = ps.PageWidth - ps.LeftMargin - ps.RightMargin
        limit = max_width or round(body_w * 0.82, 1)
        items, wide = [], []
        for i in range(1, doc.InlineShapes.Count + 1):
            s = doc.InlineShapes(i)
            try:
                pid = s.OLEFormat.ProgID
            except Exception:
                continue
            if pid != progid:
                continue
            w, h = round(s.Width, 1), round(s.Height, 1)
            label = ""
            try:
                tx = s.Range.Paragraphs(1).Range.Text
                m = re.search(r"（(\d+\.\d+)）", tx)
                if m:
                    label = m.group(1)
            except Exception:
                pass
            items.append({"index": i, "label": label, "width": w, "height": h})
            if w > limit:
                wide.append({"index": i, "label": label, "width": w})
        pages = doc.ComputeStatistics(2)
    finally:
        doc.Close(False)
        app.Quit()
    return {"body_width": round(body_w, 1), "limit": limit,
            "equations": items, "wide": wide, "pages": pages}


# ----------------------------------------------------------------- 入口
def main():
    ap = argparse.ArgumentParser(
        prog="omml2mathtype",
        description="把 Word 内置公式（OMML）批量转换为 MathType 公式对象")
    ap.add_argument("input", help="输入 .docx 路径")
    ap.add_argument("-o", "--output", help="输出 .docx 路径")
    ap.add_argument("--inplace", action="store_true", help="覆盖原文件（自动备份）")
    ap.add_argument("--dry-run", action="store_true", help="只输出公式的 TeX，不修改文件")
    ap.add_argument("--no-backup", action="store_true", help="--inplace 时不备份")
    ap.add_argument("--progid", default=DEFAULT_PROGID, help="MathType 的 OLE ProgID")
    ap.add_argument("--max-width", type=float, default=None,
                    help="公式宽度告警阈值（磅），默认取版心的 82%%")
    ap.add_argument("--keep-work", action="store_true", help="保留中间文件")
    ap.add_argument("--report", help="审计报告 JSON 输出路径")
    ap.add_argument("-q", "--quiet", action="store_true", help="精简输出")
    a = ap.parse_args()

    src = os.path.abspath(a.input)
    if not os.path.isfile(src):
        sys.exit("找不到文件：%s" % src)

    ts = datetime.now().strftime("%Y%m%d-%H%M%S")
    stem, ext = os.path.splitext(src)
    if a.inplace:
        out = src
    elif a.output:
        out = os.path.abspath(a.output)
    else:
        out = "%s_MathType%s" % (stem, ext or ".docx")

    # ---- dry run
    if a.dry_run:
        doc = Document(src)
        texs = collect_tex(doc)
        for i, t in enumerate(texs, 1):
            print("%02d | %s" % (i, t))
        print("---- 共 %d 个公式" % len(texs))
        return

    t0 = time.time()
    log("输入：%s" % src)
    log("输出：%s" % out)

    blog = os.path.join(HERE, "_run.log")
    work = os.path.join(HERE, "_work_%s.docx" % ts)

    try:
        log("\n[1/4] 提取公式并转为 TeX 文本 …")
        texs = prepare(src, work, a.quiet)
        log("      共 %d 个公式" % len(texs))

        log("[2/4] 调用 MathType 逐条转换（需要已安装 MathType 且 Word 加载项启用）…")
        ok, fail, fails = toggle_all(work, a.progid, a.quiet)
        log("      成功 %d，失败 %d" % (ok, fail))

        log("[3/4] 审计公式尺寸 …")
        rep = audit_doc(work, a.progid, a.max_width)
        log("      页数 %d｜公式对象 %d｜版心 %.1f pt｜告警阈值 %.1f pt"
            % (rep["pages"], len(rep["equations"]), rep["body_width"], rep["limit"]))
        if rep["wide"]:
            log("      ⚠ 有 %d 个公式超出版心，可能把式号挤到下一行："
                % len(rep["wide"]))
            for w in rep["wide"][:10]:
                tag = ("式 (%s)" % w["label"]) if w.get("label") else ("第 %d 个对象" % w["index"])
                log("         %s 宽 %.1f pt，约超出 %.1f pt" % (tag, w["width"], w["width"] - rep["limit"]))
        else:
            log("      ✓ 没有超宽公式")

        # ---- 落盘
        log("[4/4] 写入目标文件 …")
        if a.inplace and not a.no_backup:
            bak = "%s.bak-%s%s" % (stem, ts, ext or ".docx")
            shutil.copy2(src, bak)
            log("      已备份：%s" % bak)
        shutil.copy2(work, out)
        log("      完成：%s" % out)

        report = {
            "input": src, "output": out, "equations": len(texs),
            "converted": ok, "failed": fail, "failures": fails,
            "pages": rep["pages"], "body_width": rep["body_width"],
            "limit": rep["limit"], "wide": rep["wide"],
            "seconds": round(time.time() - t0, 1), "time": ts,
        }
        rp = a.report or os.path.join(HERE, "_report_%s.json" % ts)
        with open(rp, "w", encoding="utf-8") as f:
            json.dump(report, f, ensure_ascii=False, indent=2)
        log("      报告：%s" % rp)

        if fails:
            with open(blog, "a", encoding="utf-8") as f:
                f.write("== %s %s\n" % (ts, src))
                for x in fails:
                    f.write("  段%s | %s | %s\n" % (x["para"], x["tex"], x["error"]))
            log("\n有 %d 条失败，明细见 %s" % (fail, blog))
            log("可直接用 --dry-run 查看对应公式的 TeX 是否合法")
        log("\n完成，用时 %.1f 秒。" % (time.time() - t0))
    finally:
        if not a.keep_work and os.path.exists(work):
            try:
                os.remove(work)
            except OSError:
                pass


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        sys.exit("\n已中断")
    except Exception:
        traceback.print_exc()
        sys.exit(1)
