import os
import re
import time
import platform
import threading
import configparser
import subprocess
from datetime import datetime

import requests
import pyperclip
import keyboard
from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.shared import Pt, RGBColor
from plyer import notification


# ============================================================
# 配置区
# ============================================================
ESHELPER_EXE = r"C:\Program Files\EsHelper\eshelper.exe"
# ============================================================


# 创建配置解析器
config = configparser.ConfigParser()

# 读取配置文件 (如果不存在则创建默认值)
if os.path.exists('config.ini'):
    config.read('config.ini', encoding='utf-8')
    ESHELPER_EXE = config.get('Settings', 'eshelper_path')
else:
    # 兜底方案：如果文件丢失，使用默认路径
    ESHELPER_EXE = r"C:\Program Files\EsHelper\eshelper.exe"

processing_lock = threading.Lock()

# 词条序号：线程安全，每次保存自动递增
_entry_counter      = 1
_entry_counter_lock = threading.Lock()


def next_entry_number() -> int:
    global _entry_counter
    with _entry_counter_lock:
        n = _entry_counter
        _entry_counter += 1
        return n


# ── 词性缩写对照表 ───────────────────────────────────────────
POS_MAP_ES = {
    "verb":              "v.",
    "transitive verb":   "tr.",
    "intransitive verb": "intr.",
    "reflexive verb":    "Prnl.",
    "pronominal verb":   "Prnl.",
    "auxiliary verb":    "v.aux.",
    "noun":              "n.",
    "masculine noun":    "m.",
    "feminine noun":     "f.",
    "proper noun":       "n.prop.",
    "adjective":         "adj.",
    "adverb":            "adv.",
    "preposition":       "prep.",
    "conjunction":       "conj.",
    "pronoun":           "pron.",
    "interjection":      "interj.",
    "exclamation":       "interj.",
    "article":           "art.",
    "numeral":           "num.",
    "participle":        "part.",
    "past participle":   "p.p.",
}

POS_MAP_EN = {
    "verb":              "v.",
    "transitive verb":   "Vt.",
    "intransitive verb": "Vi.",
    "auxiliary verb":    "v.aux.",
    "modal verb":        "v.mod.",
    "noun":              "n.",
    "proper noun":       "n.prop.",
    "adjective":         "adj.",
    "adverb":            "adv.",
    "preposition":       "prep.",
    "conjunction":       "conj.",
    "pronoun":           "pron.",
    "interjection":      "interj.",
    "exclamation":       "interj.",
    "numeral":           "num.",
    "article":           "art.",
    "abbreviation":      "abbr.",
    "plural":            "pl.",
}


# ── 桌面路径 ─────────────────────────────────────────────────
def get_desktop_path() -> str:
    if platform.system() == "Windows":
        try:
            import winreg
            key = winreg.OpenKey(
                winreg.HKEY_CURRENT_USER,
                r"Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders",
            )
            val, _ = winreg.QueryValueEx(key, "Desktop")
            return os.path.expandvars(val)
        except Exception:
            pass
    return os.path.expanduser("~/Desktop")


# ── 自动编号文档：学习笔记 → 学习笔记2 → 学习笔记3 … ────────
def get_new_doc_info() -> tuple:
    desktop = get_desktop_path()
    path = os.path.join(desktop, "学习笔记.docx")
    if not os.path.exists(path):
        return path, "学习笔记"
    idx = 2
    while True:
        title = f"学习笔记{idx}"
        path  = os.path.join(desktop, f"{title}.docx")
        if not os.path.exists(path):
            return path, title
        idx += 1


DOC_PATH, DOC_TITLE = get_new_doc_info()


# ── Word 字体辅助 ─────────────────────────────────────────────
def _apply_font(run, bold=False, italic=False, size_pt=15, color=None):
    """等线字体，同时覆盖拉丁（font.name）和中文（eastAsia）两个属性。"""
    run.font.name   = "等线"
    run.font.size   = Pt(size_pt)
    run.font.bold   = bold
    run.font.italic = italic
    if color:
        run.font.color.rgb = color
    rPr    = run._r.get_or_add_rPr()
    rFonts = rPr.get_or_add_rFonts()
    rFonts.set(qn("w:eastAsia"), "等线")


def add_mixed_runs(paragraph, text: str, size_pt=15, italic=False):
    """
    拉丁字母段加粗，其余不加粗，全部等线 size_pt pt。
    re.split 保留捕获组，得到交替的「非拉丁」和「拉丁」片段。
    """
    for seg in re.split(r"([A-Za-z]+)", text):
        if not seg:
            continue
        is_latin = bool(re.fullmatch(r"[A-Za-z]+", seg))
        run = paragraph.add_run(seg)
        _apply_font(run, bold=is_latin, italic=italic, size_pt=size_pt)


# ── 标题段落底部蓝色横线 ─────────────────────────────────────
def _add_blue_bottom_border(paragraph):
    """
    在段落底部添加蓝色单线边框，视觉上形成标题下方的分隔线。
    w:sz 单位为 1/8 pt，值 16 ≈ 2pt 粗。
    """
    pPr    = paragraph._p.get_or_add_pPr()
    pBdr   = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"),   "single")
    bottom.set(qn("w:sz"),    "16")         # 2pt 粗
    bottom.set(qn("w:space"), "4")          # 与文字间距
    bottom.set(qn("w:color"), "0070C0")     # 与标题同色蓝
    pBdr.append(bottom)
    pPr.append(pBdr)


# ── 新建文档：蓝色标题 + 蓝色底线 ───────────────────────────
def create_doc_with_title(title: str) -> Document:
    doc = Document()
    p   = doc.add_paragraph()

    # 标题文字：等线 19pt 加粗蓝色
    run = p.add_run(title)
    _apply_font(run, bold=True, size_pt=19,
                color=RGBColor(0x00, 0x70, 0xC0))

    # 标题段落底部蓝色横线
    _add_blue_bottom_border(p)

    return doc


# ── 退出时追加时间戳 ─────────────────────────────────────────
def append_timestamp():
    if not os.path.exists(DOC_PATH):
        return
    try:
        doc = Document(DOC_PATH)
        p   = doc.add_paragraph()
        now = datetime.now().strftime("%Y/%m/%d  %H:%M")
        run = p.add_run(now)
        _apply_font(run, italic=True, size_pt=13)
        doc.save(DOC_PATH)
        print(f" 已记录时间：{now}")
    except Exception as e:
        print(f"⚠️ 写入时间戳失败：{e}")


# ── Google 翻译词典 API（dj=1 命名字段）─────────────────────
def get_google_definition(text: str, source_lang: str) -> list:
    """
    dj=1：返回 dict 结构，data["dict"] 包含完整词性分组。
    dt=bd：请求词典数据（仅对单词/短词组有效）。
    返回格式化列表，例如 ["tr.使恼火，惹怒", "Prnl.讨厌"]。
    """
    is_word = len(text.split()) <= 4
    pos_map = POS_MAP_ES if source_lang == "es" else POS_MAP_EN

    params = [
        ("client", "gtx"),
        ("sl",     source_lang),
        ("tl",     "zh-CN"),
        ("dt",     "t"),
        ("dj",     "1"),
        ("q",      text),
    ]
    if is_word:
        params.append(("dt", "bd"))

    try:
        resp = requests.get(
            "https://translate.googleapis.com/translate_a/single",
            params=params,
            timeout=6,
        )
        resp.raise_for_status()
        data = resp.json()
    except Exception as e:
        print(f"⚠️ Google 翻译请求失败：{e}")
        return ["（释义获取失败）"]

    # 解析词典条目
    if isinstance(data, dict) and is_word:
        lines = []
        for entry in data.get("dict", []):
            pos_raw = entry.get("pos", "")
            terms   = entry.get("terms", [])
            if not terms:
                continue
            pos_label = pos_map.get(pos_raw.lower(), pos_raw)
            meanings  = "，".join(str(t) for t in terms[:5])
            lines.append(f"{pos_label}{meanings}")
        if lines:
            return lines

    # 降级：普通翻译
    if isinstance(data, dict):
        plain = "".join(s.get("trans", "") for s in data.get("sentences", []))
    else:
        try:
            plain = "".join(seg[0] for seg in data[0] if seg[0])
        except Exception:
            plain = ""
    return [plain] if plain else ["（释义获取失败）"]


# ── 保存词条到 Word 文档 ─────────────────────────────────────
def save_to_doc(text: str, def_lines: list) -> int:
    """
    格式：1| Fastidiar —— tr.使恼火，使得讨厌 || Prnl.讨厌
    序号由全局计数器自动递增，线程安全。
    返回本条的序号供终端打印使用。
    """
    n            = next_entry_number()
    word_display = text[0].upper() + text[1:]
    defs_str     = " || ".join(def_lines)
    full_line    = f"{n}| {word_display} \u2014\u2014 {defs_str}"

    if os.path.exists(DOC_PATH):
        doc = Document(DOC_PATH)
    else:
        doc = create_doc_with_title(DOC_TITLE)

    p = doc.add_paragraph()
    add_mixed_runs(p, full_line, size_pt=15)
    doc.save(DOC_PATH)
    return n


# ── 打开西语助手 ─────────────────────────────────────────────
def open_in_xiyuzhushou(word: str):
    if os.path.exists(ESHELPER_EXE):
        try:
            subprocess.Popen([ESHELPER_EXE, "-w", word])
            print(f"   📖 西语助手已启动")
        except Exception as e:
            print(f"   ⚠️ 启动西语助手失败：{e}")
    else:
        print(f"   ⚠️ 未找到西语助手：{ESHELPER_EXE}")


# ── 核心处理 ─────────────────────────────────────────────────
def process(source_lang: str, open_app: bool = False):
    if not processing_lock.acquire(blocking=False):
        return

    try:
        pyperclip.copy("")
        time.sleep(0.15)
        keyboard.send("ctrl+c")

        text = ""
        for _ in range(10):
            time.sleep(0.05)
            content = pyperclip.paste().strip()
            if content:
                text = content
                break

        if not text:
            print("⚠️ 未检测到选中文本！")
            return

        # 异步启动西语助手（仅 S+D+C）
        if open_app:
            threading.Thread(
                target=open_in_xiyuzhushou, args=(text,), daemon=True
            ).start()

        # 获取释义
        def_lines    = get_google_definition(text, source_lang)
        word_display = text[0].upper() + text[1:]
        defs_str     = " || ".join(def_lines)

        # 写入文档（同时拿到序号）
        n = save_to_doc(text, def_lines)

        # ── 终端输出（清晰展示序号 + 释义）────────────────
        print(f"\n{'─' * 48}")
        print(f"  {n}| {word_display}  ——  {defs_str}")
        print(f"{'─' * 48}")
        print(f"  ✅ 已保存到：{DOC_PATH}\n")

        # 桌面通知
        msg = f"{n}| {word_display} —— {defs_str}"
        notification.notify(
            title="✅ 笔记已保存",
            message=msg[:200],
            app_name="语言学习助手",
            timeout=3,
        )

    except PermissionError:
        tip = "文档正被 Word 打开，请关闭后再试！"
        print(f"❌ {tip}")
        notification.notify(title="⚠️ 保存失败", message=tip, timeout=5)
    except Exception as e:
        print(f"❌ 出错：{e}")
    finally:
        processing_lock.release()


# ── 入口 ─────────────────────────────────────────────────────
if __name__ == "__main__":
    print("=" * 52)
    print(" 语言学习助手已启动")
    print(f" 本次笔记：{DOC_PATH}")
    print()
    print("  S+D+C  →  西班牙语：西语助手查词 + 保存笔记")
    print("  S+C    →  西班牙语：Google 翻译  + 保存笔记")
    print("  E+C    →  英语：    Google 翻译  + 保存笔记")
    print("  Esc    →  退出（自动记录时间）")
    print("=" * 52)

    if os.path.exists(ESHELPER_EXE):
        print(f"✅ 西语助手就绪：{ESHELPER_EXE}")
    else:
        print(f"⚠️  未找到西语助手：{ESHELPER_EXE}")
    print()

    keyboard.add_hotkey("s+d+c", lambda: process("es", open_app=True))
    keyboard.add_hotkey("s+c",   lambda: process("es", open_app=False))
    keyboard.add_hotkey("e+c",   lambda: process("en", open_app=False))

    try:
        keyboard.wait("esc")
    except KeyboardInterrupt:
        pass

    append_timestamp()
    print("\n 已退出")
