#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""把字幕文件清洗成纯文本。自动判断格式，不用你操心。

支持三种输入：
  1. SRT（带序号和时间轴）—— 常见于 B 站字幕网站导出
  2. VTT（WebVTT，网页字幕）—— 常见于 YouTube
  3. 纯文本 —— 已经是一段一段的文字

用法：
    python3 clean_subtitle.py 字幕.srt
    python3 clean_subtitle.py 字幕.txt 输出.txt

判断规则：文件里出现 `-->`（有时间轴）就按字幕格式解析，否则当纯文本直接用。
解析时按「序号行 / 时间轴行 / 正文行」的位置关系处理，不会把正文里的
「747」「1969」这种纯数字当成序号删掉。
"""
import os
import re
import sys

# ---- Windows 兼容：控制台默认 GBK，输出中文会 UnicodeEncodeError ----
import sys as _sys
for _stream in (_sys.stdout, _sys.stderr):
    try:
        _stream.reconfigure(encoding="utf-8", errors="replace")
    except (AttributeError, ValueError):
        pass



def looks_like_subtitle(raw):
    return "-->" in raw


TIMELINE = re.compile(
    r"(?:\d{1,2}:)?\d{1,2}:\d{2}[.,]\d{1,3}\s*-->\s*(?:\d{1,2}:)?\d{1,2}:\d{2}[.,]\d{1,3}"
)


def clean_subtitle(raw):
    """按字幕结构清洗：去掉文件头、序号行、时间轴行、注释块、样式标签。"""
    # VTT 文件头（WEBVTT 及其后的元信息行），跳到第一个空行
    if raw.lstrip().startswith("WEBVTT"):
        parts = raw.split("\n\n", 1)
        raw = parts[1] if len(parts) > 1 else ""

    lines = []
    for block in re.split(r"\n\s*\n", raw.strip()):
        rows = [r.rstrip() for r in block.split("\n") if r.strip()]
        if not rows:
            continue
        # 跳过 NOTE / STYLE / REGION 块（VTT 的注释与样式）
        if rows[0].strip().upper().startswith(("NOTE", "STYLE", "REGION")):
            continue

        # 丢掉 cue 标识行：下一行是时间轴，且本行不是正文
        # （SRT 是纯数字序号，VTT 可能是 intro 这种任意 cue 名）
        if len(rows) > 1 and TIMELINE.search(rows[1]):
            rows = rows[1:]

        # 丢掉时间轴行：可能是第一行，也可能前面还有 cue 名
        # （YouTube 的 VTT 常见「intro 00:00:00.380 --> ...」）
        kept = []
        for r in rows:
            if TIMELINE.search(r):
                continue
            kept.append(r)
        rows = kept

        text = " ".join(r.strip() for r in rows)
        text = re.sub(r"<[^>]+>", "", text).strip()
        if text:
            # 滚动字幕常把同一句拆成多行重复，相邻重复去重
            if not lines or lines[-1] != text:
                lines.append(text)
    return lines


def clean_plain(raw):
    """纯文本：一行一段，去掉空行。"""
    return [l.strip() for l in raw.split("\n") if l.strip()]


def read_text(path):
    """按 BOM → utf-8 → gb18030 的顺序判断编码。

    关键：没有 BOM 就不按 UTF-16 解。否则一段短 GBK 中文会被解成乱码
    （UTF-16 解释下汉字数量少，但长度够短，容易被误判为有效）。
    """
    data = open(path, "rb").read()

    # 1) 有 BOM，按 BOM 来
    if data[:2] in (b"\xff\xfe", b"\xfe\xff"):
        try:
            return data.decode("utf-16"), "utf-16"
        except UnicodeDecodeError:
            pass

    # 2) UTF-8（含 BOM）
    try:
        text = data.decode("utf-8-sig")
        if "\x00" not in text:
            return text, "utf-8"
    except UnicodeDecodeError:
        pass

    # 3) GB18030（兼容 GBK / GB2312）
    try:
        text = data.decode("gb18030")
        if "\x00" not in text and "\ufffd" not in text:
            return text, "gb18030"
    except UnicodeDecodeError:
        pass

    # 4) 都不行，容错解码并提示
    return data.decode("utf-8", errors="replace"), "utf-8(容错)"


def clean_file(path):
    raw, enc = read_text(path)
    if looks_like_subtitle(raw):
        return clean_subtitle(raw), f"字幕格式（{enc}）"
    return clean_plain(raw), f"纯文本（{enc}）"


def main():
    if len(sys.argv) < 2:
        print(__doc__, file=sys.stderr)
        sys.exit(1)

    src = sys.argv[1]
    if not os.path.exists(src):
        print(f"ERROR: 找不到文件：{src}", file=sys.stderr)
        sys.exit(1)

    lines, kind = clean_file(src)
    if len(sys.argv) > 2:
        out = sys.argv[2]
    else:
        base = os.path.splitext(src)[0]
        out = base + ".clean.txt"
        # 万一和原文件同名，再加后缀
        if os.path.abspath(out) == os.path.abspath(src):
            out = base + ".cleaned.txt"

    with open(out, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))

    hanzi = sum(1 for c in "".join(lines) if "\u4e00" <= c <= "\u9fff")
    print(f"识别为：{kind}")
    if "容错" in kind:
        print("WARN: 编码可能不是常见格式，文字若有乱码请另存为 UTF-8 再试", file=sys.stderr)
    print(f"{len(lines)} 段 / {hanzi} 个汉字 → {out}")

    if not lines:
        print("WARN: 没解析出任何内容，检查文件格式", file=sys.stderr)
        sys.exit(1)
    if hanzi < 200:
        print("WARN: 内容偏少（不到 200 字），确认一下字幕是否完整", file=sys.stderr)


if __name__ == "__main__":
    main()
