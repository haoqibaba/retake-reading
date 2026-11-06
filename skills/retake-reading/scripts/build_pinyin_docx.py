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

try:
    from docx import Document
except ImportError:
    print("缺少依赖，请先运行：python3 -m pip install python-docx", file=sys.stderr)
    raise SystemExit(1)
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Cm, Mm, Pt

# ---- Windows 兼容：控制台默认 GBK，输出中文会 UnicodeEncodeError ----
import sys as _sys
for _stream in (_sys.stdout, _sys.stderr):
    try:
        _stream.reconfigure(encoding="utf-8", errors="replace")
    except (AttributeError, ValueError):
        pass

try:
    from pypinyin import Style, load_phrases_dict, load_single_dict, pinyin
except ImportError:
    print("缺少依赖，请先运行：python3 -m pip install pypinyin", file=sys.stderr)
    raise SystemExit(1)

HERE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(HERE, "article_source.json")          # 可被命令行参数覆盖
OUT = os.path.join(HERE, "带拼音大字版.docx")              # 可被命令行参数覆盖
# 词表查找顺序：脚本同目录的 phrases_extra.json → 技能包 assets/phrases.json
_CANDIDATES = [
    os.path.join(HERE, "phrases_extra.json"),
    os.path.join(os.path.dirname(HERE), "assets", "phrases.json"),
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

def _validate_phrases(d):
    """词表结构校验：字数和拼音数必须一致，每个拼音必须是非空字符串。

    这两条不满足时，拼音会整体错位、丢字，或者输出单个字母，且不报错。
    """
    bad = []
    for k, v in d.items():
        if not isinstance(v, list):
            bad.append(f"  {k!r}: 值必须是列表")
            continue
        if len(k) != len(v):
            bad.append(f"  {k!r}: 字 {len(k)} 个，拼音 {len(v)} 个")
            continue
        for item in v:
            # 只接受 [["bō"], ["yīn"]] 这种嵌套写法。
            # 扁平的 ["bō","yīn"] 会被 pypinyin 误解析成单字母 b、y。
            if not isinstance(item, list):
                bad.append(f'  {k!r}: 拼音项必须写成 [["音"]] 的嵌套形式，不能是普通字符串')
                break
            if len(item) != 1 or not isinstance(item[0], str) or not item[0].strip():
                bad.append(f'  {k!r}: 拼音项格式不对，应为 [["yīn"]]')
                break
    return bad


# 可选：合并同目录下的增补词表 phrases_extra.json（跨期积累的多音字/人名/专名）
_EXTRA = PHRASES_FILE
if os.path.exists(_EXTRA):
    try:
        with open(_EXTRA, encoding="utf-8") as _f:
            _extra = json.load(_f)
        bad = _validate_phrases(_extra)
        if bad:
            print("ERROR: 词表条目有问题，会导致整篇拼音错位：", file=sys.stderr)
            print("\n".join(bad), file=sys.stderr)
            sys.exit(1)
        # 合并而不是替换：词表可能分成「内置 + 增补」两份
        PHRASES.update(_extra)
        print(f"merged {os.path.basename(_EXTRA)}: +{len(_extra)} entries")
    except SystemExit:
        raise
    except json.JSONDecodeError as _e:
        print(f"ERROR: 词表 JSON 格式错误（{os.path.basename(_EXTRA)}）：第 {_e.lineno} 行 {_e.msg}",
              file=sys.stderr)
        sys.exit(1)
    except SystemExit:
        raise
    except Exception as _e:  # noqa: BLE001
        print(f"ERROR: 词表加载失败（{os.path.basename(_EXTRA)}）：{_e}", file=sys.stderr)
        sys.exit(1)

load_phrases_dict(PHRASES)

# 单字“长”默认读 cháng（“机长、船长”等由词组词库保证读 zhǎng）。
# 否则“70多米长、航程不长、继续加长”会被误标成 zhǎng。
load_single_dict({ord("长"): "cháng,zhǎng"})

CJK = re.compile(r"([\u3400-\u4dbf\u4e00-\u9fff]+)")
LATIN_PINYIN = re.compile(r"^[A-Za-züÜāáǎàēéěèīíǐìōóǒòūúǔùǖǘǚǜêÊ]+$")


def _phrase_lock(text):
    """标出被词组词表覆盖的位置（最长匹配），这些位置后处理规则不再改动。

    这样用户往词表里加词才能真正生效——否则规则会把词表的正确读音又覆盖掉。
    """
    n = len(text)
    locked = [False] * n
    i = 0
    keys = sorted(PHRASES.keys(), key=len, reverse=True)
    while i < n:
        for k in keys:
            if text.startswith(k, i):
                for j in range(i, i + len(k)):
                    locked[j] = True
                i += len(k)
                break
        else:
            i += 1
    return locked


def annotate(text):
    """返回 [(char, pinyin_or_None), ...]，与输入逐字对齐。

    规则顺序很重要：先取基础读音，再用规则修正常见错误；
    但被词组词表锁定的位置一律不动，保证「加进词表就一定生效」。
    """
    out = []
    tones = []
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

    chars = [t[0] for t in out]
    n = len(out)

    locked = [False] * n
    pos = 0
    for i, part in enumerate(CJK.split(text)):
        if i % 2 == 1 and part:
            block_lock = _phrase_lock(part)
            for k, v in enumerate(block_lock):
                if pos + k < n:
                    locked[pos + k] = v
            pos += len(part)
        elif part:
            pos += len(part)

    def is_hanzi(c):
        return c is not None and re.match(r"[\u3400-\u4dbf\u4e00-\u9fff]", c) is not None

    # ---------- “一”的变调 ----------
    # 只有“一 + 量词/名词”才变调。序数、数字串、年月、词尾一律读本调 yī。
    NUM_STR = set("一二三四五六七八九十零〇")      # 数字串内部（一九四七、十一）
    NUM_ANY = set("一二三四五六七八九十百千万亿两零〇")
    for i in range(n):
        ch, py = out[i]
        if ch != "一" or not py or not py.lower().startswith("y") or locked[i]:
            continue
        prev_ch = chars[i - 1] if i > 0 else None
        next_ch = chars[i + 1] if i + 1 < n else None
        # 数字串：后字是数字（一九四七）
        if next_ch in NUM_STR:
            out[i] = (ch, "yī"); continue
        # 序数 / 数字串：前字是“第/初”或任何数字
        if prev_ch in set("第初") or prev_ch in NUM_ANY:
            out[i] = (ch, "yī"); continue
        # 词尾：统一、唯一、万一、同一（“同一”前面是“如/相/等/共”时不算词尾）
        if prev_ch is not None and is_hanzi(next_ch) is False:
            two = (prev_ch or "") + "一"
            if two in ("统一", "唯一", "万一", "归一", "划一", "均一", "纯一"):
                out[i] = (ch, "yī"); continue
            if two == "同一" and (i < 2 or chars[i - 2] not in set("如相等等共")):
                out[i] = (ch, "yī"); continue
        # 常规变调
        nxt = tones[i + 1] if i + 1 < n else None
        if i + 1 < n and chars[i + 1] == "个":
            nxt = 4
        out[i] = (ch, "yí" if nxt == 4 else ("yì" if nxt in (1, 2, 3) else "yī"))

    # ---------- “不”的变调 ----------
    for i in range(n):
        ch, py = out[i]
        if ch != "不" or not py or not py.lower().startswith("b") or locked[i]:
            continue
        nxt = tones[i + 1] if i + 1 < n else None
        out[i] = (ch, "bú" if nxt == 4 else "bù")

    # ---------- “个”统一 gè ----------
    for i in range(n):
        ch, py = out[i]
        if ch == "个" and py and not locked[i]:
            out[i] = (ch, "gè")

    # ---------- “儿”：先查儿化词，再看后一个字 ----------
    # 「儿子/儿童/婴儿/女儿/幼儿/儿科/儿孙/儿歌」都是 ér。
    # 儿化只出现在「点儿/事儿/玩儿」这类固定词，或句末的单字：
    # 注意 chars 里包含标点，所以「后一个字是标点」也算句末。
    ERHUA_WORDS = ("点儿", "事儿", "玩儿", "劲儿", "味儿", "空儿", "片儿", "块儿",
                   "活儿", "伴儿", "座儿", "串儿", "这儿", "那儿", "哪儿", "一会儿",
                   "这会儿", "那会儿", "走神儿", "没准儿", "压根儿", "个头儿", "烟儿",
                   "劲儿", "火儿", "法儿", "几声儿", "话儿", "歌儿", "腿儿", "根儿")
    ER_NON = set("子童女科孙歌谣园时话")
    for i in range(n):
        ch, py = out[i]
        if ch != "儿" or not py or locked[i]:
            continue
        # 1) 先看是否命中儿化词（最长匹配）
        hit_erhua = False
        for w in sorted(ERHUA_WORDS, key=len, reverse=True):
            if w.endswith("儿") and i - len(w) + 1 >= 0:
                if "".join(chars[i - len(w) + 1:i + 1]) == w:
                    hit_erhua = True
                    break
        if hit_erhua:
            out[i] = (ch, "r"); continue
        # 2) 后一个字是汉字（儿子/儿童/儿歌）→ ér
        next_ch = chars[i + 1] if i + 1 < n else None
        if next_ch is not None and next_ch in ER_NON:
            out[i] = (ch, "ér"); continue
        # 3) 后面紧跟另一个汉字（不是标点）→ 多半是词内（幼儿教育）
        if next_ch is not None and is_hanzi(next_ch):
            out[i] = (ch, "ér"); continue
        # 4) 句末/词尾单字「儿」：前面是婴/女/幼/男/胎/孤 这类名词性字 → ér
        prev_ch = chars[i - 1] if i > 0 else None
        if prev_ch in set("婴女幼男胎孤弃宠产保育"):
            out[i] = (ch, "ér"); continue
        # 5) 其余句末单字「儿」按儿化处理
        out[i] = (ch, "r")

    # ---------- “得” ----------
    # 默认读 de（结构助词/轻声）：飞得快、跑得远、值得、懂得、晓得、记得
    # 只有已知复合词读 dé：获得、取得、赢得、难得、心得、得到、得意、得益、博得、求得
    DE_DE = ("获得", "取得", "赢得", "难得", "心得", "得到", "得意", "得益",
             "博得", "求得", "得奖", "得分", "得利", "得体", "得知", "得以", "只得")
    DE_DE_CHARS = set("获取赢难心博求")
    COMPLEMENT_HEADS = set("更很够又远近快慢好坏多少早晚准稳对错太极真格非停住起来"
                           "下完动翻晕忙比高饱巧干净整齐漂亮痛快开心难过着急")
    for i in range(n):
        ch, py = out[i]
        if ch != "得" or not py or locked[i]:
            continue
        prev_ch = chars[i - 1] if i > 0 else None
        next_ch = chars[i + 1] if i + 1 < n else None
        # 已知复合词 → dé（词表已经能覆盖一部分，这里兜底）
        if prev_ch in DE_DE_CHARS:
            out[i] = (ch, "dé"); continue
        # 动词 + 得 + 补语 → de
        if next_ch is not None and next_ch in COMPLEMENT_HEADS:
            out[i] = (ch, "de"); continue
        # 前面有字、后面是动词/形容词（我得走了 → děi 属特例，这里保守保留原音）
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


def patch_font_table(path, primary="楷体", alt="汉仪楷体简", alt2="STKaiti"):
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
                         f'<w:altName w:val="{alt},{alt2}"/>'
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
        if not os.path.exists(args.phrases):
            print(f"ERROR: 找不到词表文件：{args.phrases}", file=sys.stderr); sys.exit(1)
        try:
            with open(args.phrases, encoding="utf-8") as f:
                _ph = json.load(f)
        except json.JSONDecodeError as e:
            print(f"ERROR: 词表 JSON 格式错误，第 {e.lineno} 行：{e.msg}", file=sys.stderr)
            sys.exit(1)
        _bad = _validate_phrases(_ph)
        if _bad:
            print("ERROR: 词表条目有问题，会导致整篇拼音错位：", file=sys.stderr)
            print("\n".join(_bad), file=sys.stderr)
            sys.exit(1)
        PHRASES.update(_ph)
        load_phrases_dict(PHRASES)

    if not os.path.exists(SRC):
        print(f"ERROR: 找不到输入文件：{SRC}", file=sys.stderr); sys.exit(1)
    try:
        with open(SRC, encoding="utf-8") as f:
            data = json.load(f)
    except json.JSONDecodeError as e:
        print(f"ERROR: JSON 格式有误（第 {e.lineno} 行）：{e.msg}", file=sys.stderr); sys.exit(1)

    if not isinstance(data, dict) or "blocks" not in data:
        print('ERROR: JSON 顶层必须是 {"blocks": [...]}', file=sys.stderr); sys.exit(1)
    blocks = data["blocks"]
    if not isinstance(blocks, list) or not blocks:
        print("ERROR: blocks 不能为空", file=sys.stderr); sys.exit(1)

    ALLOWED = {"title", "h2", "p"}
    for i, b in enumerate(blocks):
        if not isinstance(b, dict) or "kind" not in b or "text" not in b:
            print(f'ERROR: 第 {i+1} 块缺少 kind 或 text 字段', file=sys.stderr); sys.exit(1)
        if b["kind"] not in ALLOWED:
            print(f'ERROR: 第 {i+1} 块的 kind="{b["kind"]}" 不合法，只能填 title / h2 / p', file=sys.stderr)
            sys.exit(1)

    # 输出目录不存在就建
    out_dir = os.path.dirname(os.path.abspath(OUT))
    os.makedirs(out_dir, exist_ok=True)

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

    doc.core_properties.title = blocks[0]["text"][:100]  # 保守截断，避免超 Word 元数据上限

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
    if not ok:
        print("ERROR: 文本完整性校验失败，上方有 MISMATCH 明细", file=sys.stderr)
        sys.exit(1)

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
    # -------- 校验 3：拼音疑问清单（给家长看，不是给程序员看）--------
    # 按「字」分组：同一个字如果文中出现两种读音，才值得让家长确认。
    # 只出现一种读音的字，说明上下文一致，不用看。
    # 不再用固定白名单——那样会漏掉「一」「要」这类真歧义。
    # 改为遍历所有已注音的字，凡在文中出现两种以上读音的，都值得让家长看一眼。
    # 只排除两类噪音：极高频虚词（的/了/着/地 这类几乎不会有歧义），
    # 以及「一」「不」——它们的变调规则已针对性处理过，列出来只会淹没重点。
    NOISE = set("的了着地得吗呢吧啊不一个")
    PUNC = "，。！？；：、（）《》“”‘’…—"

    by_char = {}   # ch -> {py: [(para, ctx, count)]}
    for bi, blk in enumerate(blocks):
        text = blk["text"]
        ann = annotate(text)
        chars = [ch for ch, _ in ann]
        locks = [False] * len(chars)
        pos = 0
        for part in re.split(r"([\u3400-\u4dbf\u4e00-\u9fff]+)", text):
            if not part:
                continue
            if re.match(r"[\u3400-\u4dbf\u4e00-\u9fff]", part):
                bl = _phrase_lock(part)
                for k, v in enumerate(bl):
                    if pos + k < len(locks):
                        locks[pos + k] = v
            pos += len(part)
        for i, (ch, py) in enumerate(ann):
            if not py or ch in NOISE or locks[i]:
                continue
            a = i
            while a > 0 and i - a < 10 and chars[a - 1] not in PUNC:
                a -= 1
            b = i + 1
            while b < len(chars) and b - i < 10 and chars[b] not in PUNC:
                b += 1
            ctx = "".join(chars[a:i]) + "【" + ch + "】" + "".join(chars[i + 1:b])
            by_char.setdefault(ch, {}).setdefault(py, {})
            slot = by_char[ch][py]
            if ctx not in slot:
                slot[ctx] = [bi + 1, 0]
            slot[ctx][1] += 1

    # 只要「一个字出现两种以上读音」的，那才是真需要人判断的
    ambiguous = {ch: v for ch, v in by_char.items() if len(v) > 1}

    qpath = os.path.join(os.path.dirname(os.path.abspath(OUT)), "拼音疑问清单.txt")
    with open(qpath, "w", encoding="utf-8") as f:
        f.write("拼音疑问清单\n")
        f.write("=" * 46 + "\n\n")
        f.write("下面这些字，在同一篇文章里有两种以上读音。\n")
        f.write("我按上下文各选了一个，请你读一读，看哪句不对劲。\n")
        f.write("（只出现一种读音的字不会列出来，那种基本不会错）\n\n")

        if not ambiguous:
            f.write("（没有需要确认的字）\n")
        for n, (ch, readings) in enumerate(sorted(ambiguous.items()), 1):
            f.write(f"{n}. 「{ch}」—— 文中出现 {len(readings)} 种读音\n")
            for py, ctxs in sorted(readings.items(), key=lambda kv: -sum(c[1] for c in kv[1].values())):
                total = sum(c[1] for c in ctxs.values())
                f.write(f"     读 {py}（{total} 处）\n")
                for ctx, (para, cnt) in list(ctxs.items())[:2]:
                    f.write(f"        例：{ctx}\n")
            f.write("\n")

        f.write("=" * 46 + "\n")
        f.write("发现读错的，回一句「某某词读某某音」就行，\n")
        f.write("我会加进词表重新生成，以后整篇都不会再错。\n")

    print(f"拼音疑问清单：{len(ambiguous)} 个字有歧义 → {qpath}")


if __name__ == "__main__":
    main()
