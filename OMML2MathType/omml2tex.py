# -*- coding: utf-8 -*-
"""OMML(Office Math Markup Language) → LaTeX 转换核心

只依赖 lxml，可单独 import 使用：

    from omml2tex import omml_to_tex
    from docx import Document
    d = Document("a.docx")
    M = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"
    print(omml_to_tex(next(d.element.body.iter(M + "oMath"))))
"""
from __future__ import annotations

import re

from lxml import etree

M = "http://schemas.openxmlformats.org/officeDocument/2006/math"
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"m": M, "w": W}

# 希腊字母与常见数学符号
SYM = {
    "\u03b1": r"\alpha", "\u03b2": r"\beta", "\u03b3": r"\gamma", "\u03b4": r"\delta",
    "\u03b5": r"\varepsilon", "\u03b6": r"\zeta", "\u03b7": r"\eta", "\u03b8": r"\theta",
    "\u03b9": r"\iota", "\u03ba": r"\kappa", "\u03bb": r"\lambda", "\u03bc": r"\mu",
    "\u03bd": r"\nu", "\u03be": r"\xi", "\u03c0": r"\pi", "\u03c1": r"\rho",
    "\u03c3": r"\sigma", "\u03c4": r"\tau", "\u03c5": r"\upsilon", "\u03c6": r"\varphi",
    "\u03c7": r"\chi", "\u03c8": r"\psi", "\u03c9": r"\omega",
    "\u0393": r"\Gamma", "\u0394": r"\Delta", "\u0398": r"\Theta", "\u039b": r"\Lambda",
    "\u039e": r"\Xi", "\u03a0": r"\Pi", "\u03a3": r"\Sigma", "\u03a6": r"\Phi",
    "\u03a8": r"\Psi", "\u03a9": r"\Omega",
    "\u00b7": r"\cdot", "\u2212": "-", "\u2013": "-", "\u00d7": r"\times",
    "\u00f7": r"\div", "\u2264": r"\leq", "\u2265": r"\geq", "\u2260": r"\neq",
    "\u2248": r"\approx", "\u2261": r"\equiv", "\u221d": r"\propto",
    "\u00b1": r"\pm", "\u2213": r"\mp", "\u221e": r"\infty",
    "\u2192": r"\rightarrow", "\u2190": r"\leftarrow", "\u2194": r"\leftrightarrow",
    "\u21d2": r"\Rightarrow", "\u21d4": r"\Leftrightarrow",
    "\u2208": r"\in", "\u2209": r"\notin", "\u2282": r"\subset", "\u2286": r"\subseteq",
    "\u222a": r"\cup", "\u2229": r"\cap", "\u2200": r"\forall", "\u2203": r"\exists",
    "\u2202": r"\partial", "\u2207": r"\nabla", "\u222b": r"\int", "\u2211": r"\sum",
    "\u220f": r"\prod", "\u221a": r"\surd", "\u00b0": r"^{\circ}",
    "\u2026": r"\ldots", "\u22ef": r"\cdots", "\u2032": r"'", "\u2033": r"''",
    "\u2044": r"/", "\u2215": r"/", "\u00a0": " ",
}

# 这些函数名要用正体，且前面带反斜杠
FUNC = {
    "sin", "cos", "tan", "cot", "sec", "csc", "arcsin", "arccos", "arctan",
    "sinh", "cosh", "tanh", "coth", "exp", "ln", "log", "lg",
    "max", "min", "lim", "sup", "inf", "det", "dim", "arg", "deg", "gcd",
}

# 重音符 → TeX 命令
ACC_CMD = {
    "\u0307": r"\dot", "\u0308": r"\ddot", "\u0305": r"\bar", "\u0304": r"\bar",
    "\u20d7": r"\vec", "\u0302": r"\hat", "\u0303": r"\tilde",
    "\u0306": r"\breve", "\u030c": r"\check", "\u0301": r"\acute", "\u0300": r"\grave",
}

# 括号映射
LEFT = {"(": r"\left(", "[": r"\left[", "{": r"\left\{", "|": r"\left|",
        "\u27e8": r"\left\langle", "\u230a": r"\left\lfloor", "\u2308": r"\left\lceil"}
RIGHT = {")": r"\right)", "]": r"\right]", "}": r"\right\}", "|": r"\right|",
         "\u27e9": r"\right\rangle", "\u230b": r"\right\rfloor", "\u2309": r"\right\rceil"}


def _txt(el) -> str:
    return "".join(t.text or "" for t in el.iter("{%s}t" % M))


def _esc(s: str) -> str:
    s = s.replace("\\", r"\backslash ")
    for k in ("{", "}", "%", "&", "#", "$"):
        s = s.replace(k, "\\" + k)
    s = s.replace("_", r"\_")
    s = s.replace("~", r"\sim ")
    s = s.replace("^", r"\hat{}")
    return s


def _word(w: str) -> str:
    if not w:
        return ""
    if w.strip() == "":
        return " "
    if w in FUNC:
        return "\\" + w + " "
    if w == "d":  # 微分算子 d，正体
        return r"\mathrm{d}"
    if re.fullmatch(r"[A-Za-z]{2,}", w):
        return r"\mathrm{%s}" % w
    return _esc(w)


def tok(s: str) -> str:
    """把一段 m:t 纯文本转成 TeX（符号逐个替换，连续字母合并处理）"""
    out, buf = [], ""
    for ch in s:
        if ch in SYM:
            if buf:
                out.append(_word(buf))
                buf = ""
            out.append(SYM[ch])
            if SYM[ch] and not SYM[ch].endswith("}"):
                out.append(" ")
        else:
            buf += ch
    if buf:
        out.append(_word(buf))
    return "".join(out)


def conv(el, up: bool = False) -> str:
    """递归把 OMML 元素转成 TeX 片段"""
    if el is None:
        return ""
    tag = etree.QName(el).localname

    if tag in ("oMath", "oMathPara", "e", "num", "den", "sub", "sup",
               "fName", "lim", "deg", "box", "borderBox"):
        return "".join(conv(c) for c in el)

    if tag == "r":
        rpr = el.find("m:rPr", NS)
        sty = ""
        if rpr is not None:
            s = rpr.find("m:sty", NS)
            if s is not None:
                sty = s.get("{%s}val" % M) or ""
        content = "".join(t.text or "" for t in el.findall("m:t", NS))
        if sty in ("p", "b"):  # 正体 / 粗正体
            if content in FUNC:
                return "\\" + content + " "
            if re.fullmatch(r"[A-Za-z]{2,}", content):
                return r"\mathrm{%s}" % content
        return tok(content)

    if tag == "t":
        return tok(el.text or "")

    if tag == "f":
        return r"\frac{%s}{%s}" % (conv(el.find("m:num", NS)), conv(el.find("m:den", NS)))

    if tag == "sSub":
        return "{%s}_{%s}" % (conv(el.find("m:e", NS)), conv(el.find("m:sub", NS)))

    if tag == "sSup":
        return "{%s}^{%s}" % (conv(el.find("m:e", NS)), conv(el.find("m:sup", NS)))

    if tag == "sSubSup":
        return "{%s}_{%s}^{%s}" % (conv(el.find("m:e", NS)),
                                   conv(el.find("m:sub", NS)),
                                   conv(el.find("m:sup", NS)))

    if tag == "sPre":
        return "{}_{%s}^{%s}{%s}" % (conv(el.find("m:sub", NS)),
                                     conv(el.find("m:sup", NS)),
                                     conv(el.find("m:e", NS)))

    if tag == "rad":
        deg = el.find("m:deg", NS)
        body = conv(el.find("m:e", NS))
        if deg is not None and _txt(deg).strip():
            return r"\sqrt[%s]{%s}" % (conv(deg), body)
        return r"\sqrt{%s}" % body

    if tag == "nary":
        pr = el.find("m:naryPr", NS)
        chr_ = "\u222b"
        if pr is not None:
            c = pr.find("m:chr", NS)
            if c is not None:
                chr_ = c.get("{%s}val" % M) or chr_
        op = { "\u2211": r"\sum", "\u220f": r"\prod", "\u222b": r"\int",
               "\u222c": r"\iint", "\u222e": r"\oint", "\u22c3": r"\bigcup",
               "\u22c2": r"\bigcap" }.get(chr_, SYM.get(chr_, r"\int"))
        s = op
        sub, sup = el.find("m:sub", NS), el.find("m:sup", NS)
        if sub is not None and _txt(sub).strip():
            s += "_{%s}" % conv(sub)
        if sup is not None and _txt(sup).strip():
            s += "^{%s}" % conv(sup)
        return s + " " + conv(el.find("m:e", NS))

    if tag == "d":
        pr = el.find("m:dPr", NS)
        beg, end, sep = "(", ")", None
        if pr is not None:
            b, e, s = pr.find("m:begChr", NS), pr.find("m:endChr", NS), pr.find("m:sepChr", NS)
            if b is not None:
                beg = b.get("{%s}val" % M)
            if e is not None:
                end = e.get("{%s}val" % M)
            if s is not None:
                sep = s.get("{%s}val" % M)
        parts = [conv(c) for c in el.findall("m:e", NS)]
        if sep is None:
            inner = "".join(parts)
        else:
            sp = sep if sep.strip() else ","
            sp = {"\u2212": "-", "\u00b7": r"\cdot", "\u00d7": r"\times",
                  "\u2192": r"\rightarrow", "\u21a6": r"\mapsto"}.get(sp, SYM.get(sp, sp))
            inner = (" %s " % sp).join(parts)
        lb, rb = LEFT.get(beg, ""), RIGHT.get(end, "")
        if beg == "" and end == "":
            return "{%s}" % inner
        if beg == "{" and not rb:
            rb = r"\right."
        if not lb and not rb:
            return inner
        return "%s %s %s" % (lb, inner, rb)

    if tag == "eqArr":
        rows = [conv(e) for e in el.findall("m:e", NS)]
        return r"\begin{array}{l}%s\end{array}" % r" \\ ".join(rows)

    if tag == "m":
        rows, ncol = [], 1
        for mr in el.findall("m:mr", NS):
            cells = [conv(e) for e in mr.findall("m:e", NS)]
            ncol = max(ncol, len(cells))
            rows.append(" & ".join(cells))
        return r"\begin{array}{%s}%s\end{array}" % ("l" * ncol, r" \\ ".join(rows))

    if tag == "acc":
        pr = el.find("m:accPr", NS)
        ch = "\u0307"
        if pr is not None:
            c = pr.find("m:chr", NS)
            if c is not None:
                ch = c.get("{%s}val" % M) or ch
        cmd = ACC_CMD.get(ch, r"\dot")
        return "%s{%s}" % (cmd, conv(el.find("m:e", NS)))

    if tag == "bar":
        pr = el.find("m:barPr", NS)
        pos = "top"
        if pr is not None:
            p = pr.find("m:pos", NS)
            if p is not None:
                pos = p.get("{%s}val" % M) or "top"
        cmd = r"\overline" if pos == "top" else r"\underline"
        return "%s{%s}" % (cmd, conv(el.find("m:e", NS)))

    if tag == "func":
        return "%s %s" % (conv(el.find("m:fName", NS)), conv(el.find("m:e", NS)))

    if tag == "limLow":
        return "{%s}_{%s}" % (conv(el.find("m:e", NS)), conv(el.find("m:lim", NS)))

    if tag == "limUpp":
        return "{%s}^{%s}" % (conv(el.find("m:e", NS)), conv(el.find("m:lim", NS)))

    if tag == "groupChr":
        pr = el.find("m:groupChrPr", NS)
        ch, pos = "\u23df", "bot"
        if pr is not None:
            c, p = pr.find("m:chr", NS), pr.find("m:pos", NS)
            if c is not None:
                ch = c.get("{%s}val" % M) or ch
            if p is not None:
                pos = p.get("{%s}val" % M) or pos
        cmd = r"\underbrace" if pos == "bot" else r"\overbrace"
        return "%s{%s}" % (cmd, conv(el.find("m:e", NS)))

    if tag == "phant":
        return conv(el.find("m:e", NS))

    # 兜底：顺序拼接子节点
    return "".join(conv(c) for c in el)


def omml_to_tex(el) -> str:
    """OMML 元素 → 单行 TeX 字符串"""
    s = conv(el)
    s = re.sub(r"[ \t]{2,}", " ", s)
    s = s.replace("{ }", "").strip()
    return s


if __name__ == "__main__":
    import sys
    from docx import Document  # noqa: E402

    path = sys.argv[1] if len(sys.argv) > 1 else None
    if not path:
        sys.exit("用法: python omml2tex.py 文档.docx")
    doc = Document(path)
    n = 0
    for om in doc.element.body.iter("{%s}oMath" % M):
        n += 1
        print("%02d | %s" % (n, omml_to_tex(om)))
    print("---- 共 %d 个公式" % n)
