#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""把整理后的文章生成“逐字带拼音”的 Word 文档（Word 拼音指南 ruby 格式）。

- 正文每个汉字上方标注拼音（适合青少年阅读）
- 标点、数字、字母不注音
- 输出：pinyin_doc.docx
"""
import html
import json
import os
import re
import sys

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Cm, Mm, Pt
from pypinyin import Style, load_phrases_dict, load_single_dict, pinyin

HERE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(HERE, "article_source.json")          # 可被命令行参数覆盖
OUT = os.path.join(HERE, "带拼音大字版.docx")              # 可被命令行参数覆盖
# 词表查找顺序：脚本同目录的 phrases_extra.json → 技能包 assets/多音字词表.json
_CANDIDATES = [
    os.path.join(HERE, "phrases_extra.json"),
    os.path.join(os.path.dirname(HERE), "assets", "多音字词表.json"),
]
PHRASES_FILE = next((p for p in _CANDIDATES if os.path.exists(p)), _CANDIDATES[0])

# ---------------------------------------------------------------- 多音字短语校正
PHRASES = {
    "空难": [["kōng"], ["nàn"]],
    "遇难": [["yù"], ["nàn"]],
    "灾难": [["zāi"], ["nàn"]],
    "系好": [["jì"], ["hǎo"]],
    "应力": [["yìng"], ["lì"]],
    "强劲": [["qiáng"], ["jìng"]],
    "盐分": [["yán"], ["fèn"]],
    "止裂条": [["zhǐ"], ["liè"], ["tiáo"]],
    "粘合剂": [["nián"], ["hé"], ["jì"]],
    "空乘": [["kōng"], ["chéng"]],
    "重伤": [["zhòng"], ["shāng"]],
    "载有": [["zài"], ["yǒu"]],
    "迫降": [["pò"], ["jiàng"]],
    "只剩": [["zhǐ"], ["shèng"]],
    "只有": [["zhǐ"], ["yǒu"]],
    "只是": [["zhǐ"], ["shì"]],
    "吹得": [["chuī"], ["de"]],
    "感觉": [["gǎn"], ["jué"]],
    "觉得": [["jué"], ["de"]],
    "热情地": [["rè"], ["qíng"], ["de"]],
    "稳稳地": [["wěn"], ["wěn"], ["de"]],
    "为了保护": [["wèi"], ["le"], ["bǎo"], ["hù"]],
    "为万一": [["wèi"], ["wàn"], ["yī"]],
    "结实": [["jiē"], ["shi"]],
    "阿罗哈": [["ā"], ["luó"], ["hā"]],
    "克拉拉贝尔": [["kè"], ["lā"], ["lā"], ["bèi"], ["ěr"]],
    "兰辛": [["lán"], ["xīn"]],
    "米歇尔": [["mǐ"], ["xiē"], ["ěr"]],
    "佐藤": [["zuǒ"], ["téng"]],
    "富田": [["fù"], ["tián"]],
    "本田": [["běn"], ["tián"]],
    "希洛": [["xī"], ["luò"]],
    "檀香山": [["tán"], ["xiāng"], ["shān"]],
    "茂宜岛": [["mào"], ["yí"], ["dǎo"]],
    "机鼻": [["jī"], ["bí"]],
    "起落架": [["qǐ"], ["luò"], ["jià"]],
    "防冲撞": [["fáng"], ["chōng"], ["zhuàng"]],
    "记录了": [["jì"], ["lù"], ["le"]],
    "了当": [["le"], ["dāng"]],
    "了事故": [["le"], ["shì"], ["gù"]],
    "得多": [["de"], ["duō"]],
    "压力差": [["yā"], ["lì"], ["chā"]],
    "处在": [["chǔ"], ["zài"]],
    "飞机长期": [["fēi"], ["jī"], ["cháng"], ["qī"]],
    "昏倒": [["hūn"], ["dǎo"]],
    "冲倒": [["chōng"], ["dǎo"]],
}

# 可选：合并同目录下的增补词表 phrases_extra.json（跨期积累的多音字/人名/专名）
# 使用方法：把「多音字总表.json」复制为本期目录的 phrases_extra.json，并把本期新词追加进去。
_EXTRA = PHRASES_FILE
if os.path.exists(_EXTRA):
    try:
        with open(_EXTRA, encoding="utf-8") as _f:
            _extra = json.load(_f)
        PHRASES.update(_extra)
        print(f"merged phrases_extra.json: +{len(_extra)} entries")
    except Exception as _e:  # noqa: BLE001
        print("WARN: failed to load phrases_extra.json:", _e)

load_phrases_dict(PHRASES)

# 单字“长”默认读 cháng（“机长、船长”等由词组词库保证读 zhǎng）。
# 否则“70多米长、航程不长、继续加长”会被误标成 zhǎng。
load_single_dict({ord("长"): "cháng,zhǎng"})

CJK = re.compile(r"([\u3400-\u4dbf\u4e00-\u9fff]+)")
LATIN_PINYIN = re.compile(r"^[A-Za-züÜāáǎàēéěèīíǐìōóǒòūúǔùǖǘǚǜêÊ]+$")


def annotate(text):
    """返回 [(char, pinyin_or_None), ...]，与输入逐字对齐。

    额外统一规则：
    - “一”“不”按变调规则标注（一→yí/yì/yī，不→bú/bù），全文一致；
    - “个”统一 gè；“儿”（儿化）标 r。
    """
    out = []
    tones = []  # 声调数字，用于“一/不”变调判断
    for i, part in enumerate(CJK.split(text)):
        if not part:
            continue
        if i % 2 == 1:  # 汉字块
            py_list = pinyin(part, style=Style.TONE)
            py_num = pinyin(part, style=Style.TONE3)
            for ch, py, pn in zip(part, py_list, py_num):
                p0 = py[0] if py else ""
                n0 = pn[0] if pn else ""
                m = re.search(r"(\d)$", n0 or "")
                tones.append(int(m.group(1)) if m else None)
                out.append((ch, p0 if LATIN_PINYIN.match(p0) else None))
        else:
            for ch in part:
                tones.append(None)
                out.append((ch, None))

    # “一/不”变调
    for i, (ch, py) in enumerate(out):
        if not py:
            continue
        nxt = tones[i + 1] if i + 1 < len(tones) else None
        if i + 1 < len(out) and out[i + 1][0] == "个":
            nxt = 4  # “个”按 gè 处理
        if ch == "一" and py.lower().startswith("y"):
            if nxt == 4:
                out[i] = (ch, "yí")
            elif nxt in (1, 2, 3):
                out[i] = (ch, "yì")
            else:
                out[i] = (ch, "yī")
        elif ch == "不" and py.lower().startswith("b"):
            out[i] = (ch, "bú" if nxt == 4 else "bù")
        elif ch == "个":
            out[i] = (ch, "gè")
        elif ch == "儿":
            out[i] = (ch, "r")

    # 结构助词“得”读 de：动词/形容词 + 得 + 补语（“飞得更快、拖得够呛、走得又准”）。
    # 只处理未命中词库的单字“得”；“得到、得意、值得”等不会被误伤。
    COMPLEMENT_HEADS = set("更很够又远近快慢好坏多少早晚准稳对错太极真格非停住起来下完动翻晕忙比")
    for i, (ch, py) in enumerate(out):
        if ch == "得" and py != "de" and 0 < i < len(out) - 1:
            prev, nxt_ch = out[i - 1], out[i + 1][0]
            if prev[1] is not None and nxt_ch in COMPLEMENT_HEADS:
                out[i] = (ch, "de")

    return out


# ---------------------------------------------------------------- 样式参数
SIZES = {  # kind: (base_half_pt, ruby_half_pt)
    "title": (48, 22),   # 24pt / 11pt
    "h2": (44, 20),      # 22pt / 10pt
    "p": (40, 16),       # 20pt / 8pt
}

RUBY_TMPL = (
    '<w:r {ns}>'
    "<w:ruby>"
    "<w:rubyPr>"
    '<w:rubyAlign w:val="center"/>'
    '<w:hps w:val="{rt_sz}"/>'
    '<w:hpsRaise w:val="{raise_v}"/>'
    '<w:hpsBaseText w:val="{base_sz}"/>'
    '<w:lid w:val="zh-CN"/>'
    "<w:dirty/>"
    "</w:rubyPr>"
    "<w:rt><w:r><w:rPr>"
    '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>'
    "{bold}<w:sz w:val=\"{rt_sz}\"/><w:szCs w:val=\"{rt_sz}\"/>"
    "</w:rPr><w:t xml:space=\"preserve\">{py}</w:t></w:r></w:rt>"
    "<w:rubyBase><w:r><w:rPr>"
    '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="楷体" w:hint="eastAsia"/>'
    "{bold}<w:sz w:val=\"{base_sz}\"/><w:szCs w:val=\"{base_sz}\"/>"
    "</w:rPr><w:t xml:space=\"preserve\">{ch}</w:t></w:r></w:rubyBase>"
    "</w:ruby>"
    "</w:r>"
)


def make_ruby(ch, py, base_sz, rt_sz, bold):
    xml = RUBY_TMPL.format(
        ns=nsdecls("w"),
        rt_sz=rt_sz,
        base_sz=base_sz,
        raise_v=base_sz - 3,
        py=html.escape(py),
        ch=html.escape(ch),
        bold="<w:b/><w:bCs/>" if bold else "",
    )
    return parse_xml(xml)


def style_normal_run(run, base_sz, bold):
    run.font.size = Pt(base_sz / 2)
    run.bold = bold
    run.font.name = "Times New Roman"
    rfonts = run._element.get_or_add_rPr().rFonts
    rfonts.set(qn("w:eastAsia"), "楷体")
    rfonts.set(qn("w:hint"), "eastAsia")


def add_block(doc, kind, text, space_after=None):
    base_sz, rt_sz = SIZES[kind]
    if kind == "p":
        base_sz, rt_sz = SIZES["p"]
    bold = kind in ("title", "h2")

    p = doc.add_paragraph()
    pf = p.paragraph_format
    # 全篇单倍行距：2026-09-22 晟哥打印实测确定（拼音行高已由 ruby 撑开，
    # 多倍行距反而会把行间距撑得过大）。标题/小节/正文统一单倍。
    if kind == "title":
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        pf.line_spacing = 1.0
        pf.space_after = Pt(space_after if space_after is not None else 22)
    elif kind == "h2":
        pf.line_spacing = 1.0
        pf.space_before = Pt(18)
        pf.space_after = Pt(10)
    else:
        pf.line_spacing = 1.0
        pf.space_after = Pt(10)
        pf.first_line_indent = Pt(base_sz)  # 首行缩进两个字符

    buf = []

    def flush_buf():
        if buf:
            r = p.add_run("".join(buf))
            style_normal_run(r, base_sz, bold)
            buf.clear()

    for ch, py in annotate(text):
        if py is None:
            buf.append(ch)
        else:
            flush_buf()
            p._p.append(make_ruby(ch, py, base_sz, rt_sz, bold))
    flush_buf()
    return p


def add_page_number(doc):
    """页脚居中页码（PAGE 域，Word/WPS 中自动编号）。

    2026-09-22 晟哥要求：每个页面标页码，方便孩子阅读。
    """
    sec = doc.sections[0]
    footer = sec.footer
    p = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run()
    fld1 = OxmlElement("w:fldChar")
    fld1.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = "PAGE"
    fld2 = OxmlElement("w:fldChar")
    fld2.set(qn("w:fldCharType"), "end")
    run._r.append(fld1)
    run._r.append(instr)
    run._r.append(fld2)
    run.font.size = Pt(12)
    run.font.name = "Times New Roman"


def patch_font_table(path, primary="楷体", alt="汉仪楷体简"):
    """在 fontTable.xml 中为「楷体」添加替代字体名（WPS 自带的汉仪楷体简）。"""
    import shutil
    import zipfile

    tmp = path + ".tmp"
    zin = zipfile.ZipFile(path, "r")
    zout = zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED)
    for item in zin.infolist():
        data = zin.read(item.filename)
        if item.filename == "word/fontTable.xml":
            s = data.decode("utf-8")
            if f'w:name="{primary}"' not in s:
                entry = (f'<w:font w:name="{primary}">'
                         f'<w:altName w:val="{alt}"/>'
                         f'<w:charset w:val="86"/>'
                         f'<w:family w:val="auto"/>'
                         f'<w:pitch w:val="variable"/>'
                         f'</w:font>')
                s = s.replace("</w:fonts>", entry + "</w:fonts>")
            data = s.encode("utf-8")
        zout.writestr(item, data)
    zout.close()
    zin.close()
    shutil.move(tmp, path)


def main():
    global SRC, OUT
    import argparse

    ap = argparse.ArgumentParser(description="把文章 JSON 生成为逐字带拼音的 Word 文档")
    ap.add_argument("source", nargs="?", default=SRC, help="article_source.json 路径")
    ap.add_argument("-o", "--output", default=None, help="输出 docx 路径")
    ap.add_argument("--phrases", default=None, help="额外的多音字词表（JSON）")
    args = ap.parse_args()

    SRC = args.source
    if args.output:
        OUT = args.output
    elif args.source != os.path.join(HERE, "article_source.json"):
        OUT = os.path.join(os.path.dirname(os.path.abspath(args.source)), "带拼音大字版.docx")
    if args.phrases:
        global PHRASES
        with open(args.phrases, encoding="utf-8") as f:
            PHRASES.update(json.load(f))
        load_phrases_dict(PHRASES)

    with open(SRC, encoding="utf-8") as f:
        data = json.load(f)
    blocks = data["blocks"]

    doc = Document()
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Mm(210), Mm(297)
    sec.top_margin = sec.bottom_margin = Cm(2.2)
    sec.left_margin = sec.right_margin = Cm(2.0)

    st = doc.styles["Normal"]
    st.font.name = "Times New Roman"
    st.font.size = Pt(14)
    rpr = st.element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        from docx.oxml import OxmlElement

        rfonts = OxmlElement("w:rFonts")
        rpr.insert(0, rfonts)
    rfonts.set(qn("w:eastAsia"), "楷体")
    rfonts.set(qn("w:ascii"), "Times New Roman")
    rfonts.set(qn("w:hAnsi"), "Times New Roman")

    doc.core_properties.title = blocks[0]["text"]

    # 修正 python-docx 模板里 settings.xml 的 zoom 元素（补齐必需的 percent 属性）
    zoom = doc.settings.element.find(qn("w:zoom"))
    if zoom is not None:
        if zoom.get(qn("w:val")) is not None:
            del zoom.attrib[qn("w:val")]
        zoom.set(qn("w:percent"), "100")

    for i, blk in enumerate(blocks):
        if blk["kind"] == "title":
            nxt_is_title = i + 1 < len(blocks) and blocks[i + 1]["kind"] == "title"
            add_block(doc, "title", blk["text"], space_after=4 if nxt_is_title else 22)
        else:
            add_block(doc, blk["kind"], blk["text"])

    add_page_number(doc)  # 页脚居中页码
    doc.save(OUT)
    patch_font_table(OUT)
    print("saved:", OUT)

    # -------- 校验 1：重新解析，逐段对比汉字是否完整 --------
    doc2 = Document(OUT)
    expected = [blk["text"] for blk in blocks]
    ok = len(expected) == len(doc2.paragraphs)
    if not ok:
        print("PARAGRAPH COUNT MISMATCH:", len(expected), "vs", len(doc2.paragraphs))
    for blk_text, para in zip(expected, doc2.paragraphs):
        got = []
        for elem in para._p.iter():
            tag = elem.tag
            if tag == qn("w:rubyBase"):
                for t in elem.findall(".//" + qn("w:t")):
                    if t.text:
                        got.append(t.text)
            elif tag == qn("w:t"):
                # 跳过 ruby 内部的 t（已处理），只取普通 run 的 t
                # 普通 t 的父级链里没有 ruby
                anc = elem
                inside_ruby = False
                while anc is not None:
                    if anc.tag == qn("w:ruby"):
                        inside_ruby = True
                        break
                    anc = anc.getparent()
                if not inside_ruby and elem.text:
                    got.append(elem.text)
        got_s = "".join(got)
        if got_s != blk_text:
            ok = False
            print("MISMATCH in block:", blk_text[:30])
            print("  got :", got_s[:60])
            print("  want:", blk_text[:60])
    print("VERIFY_TEXT_COMPLETE:", ok)

    # -------- 校验 2：导出拼音审查报告 --------
    review_path = os.path.join(os.path.dirname(os.path.abspath(OUT)), "拼音审查稿.txt")
    with open(review_path, "w", encoding="utf-8") as f:
        total_hanzi = 0
        for blk in blocks:
            ann = annotate(blk["text"])
            total_hanzi += sum(1 for _, py in ann if py)
            f.write(f"【{blk['kind']}】" + "".join(
                (f"{ch}[{py}]" if py else ch) for ch, py in ann
            ) + "\n\n")
        f.write(f"TOTAL_HANZI_ANNOTATED={total_hanzi}\n")
    print("review written:", review_path)

    # 抽查易错字所在词
    watch = "重难系应供载结得地着了为只几率分露冲强长差折兴处转曲薄血便"
    for blk in blocks:
        text = blk["text"]
        ann = annotate(text)
        chars = [ch for ch, _ in ann]
        for i, (ch, py) in enumerate(ann):
            if py and ch in watch:
                ctx = "".join(chars[max(0, i - 3):i + 4])
                print(f"watch: {ctx}  ->  {ch}={py}")


if __name__ == "__main__":
    main()
