#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
zotero_cite.py - 在 .docx 中写入 Zotero 可刷新的引用域

在 Word 文档中插入 Zotero Word 插件能够识别的域（field），使文档可以在
Word 里通过 Zotero 插件刷新引文、增删条目后自动更新参考文献表、并可一键
切换引文样式。

域格式依据 Zotero 官方发行版源码 chrome/content/zotero/xpcom/integration.js：
    引用域    ADDIN ZOTERO_ITEM CSL_CITATION {json}
    文献表域  ADDIN ZOTERO_BIBL {json} CSL_BIBLIOGRAPHY
条目解析优先级（刷新时）：citationItem.uris -> key+libraryID -> id -> 内嵌 itemData

用法：
    python zotero_cite.py inspect  --docx FILE
    python zotero_cite.py cite     --docx IN --out OUT --keys KEY1,KEY2 [--after-para N] [--label "[2]"]
    python zotero_cite.py biblio   --docx IN --out OUT [--after-para N]
    python zotero_cite.py style    --docx FILE
    python zotero_cite.py setstyle --docx IN --out OUT --style gb7714

文档样式偏好（Zotero → Document Preferences）存在 docProps/custom.xml 的
ZOTERO_PREF_n 自定义属性里，Zotero 按 255 字符一段切分（Word 自定义属性长度上限）。
本脚本可按同样的规则改写其中 <style id="..."/>，从而在不打开 Word 的情况下
把整篇文档的引文样式换掉；随后在 Word 里 Refresh 一次即整体重排。
"""

import argparse
import html
import json
import os
import random
import re
import sqlite3
import string
import sys
import zipfile

from lxml import etree

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from zotero_paths import default_db  # noqa: E402

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
XML_NS = "http://www.w3.org/XML/1998/namespace"
XML_SPACE = "{%s}space" % XML_NS
CITATION_SCHEMA = (
    "https://github.com/citation-style-language/schema/raw/master/csl-citation.json"
)
DEFAULT_DB = None  # 运行时由 zotero_paths.default_db() 探测，见 connect_library()
DOC_PART = "word/document.xml"
CUSTOM_PART = "docProps/custom.xml"

CP_NS = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
VT_NS = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"
FMTID = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"
PREF_NAME_RE = re.compile(r"^ZOTERO_PREF_(\d+)$")
CHUNK = 255  # Word 自定义属性字符串长度上限，Zotero 按此切分
STYLE_BASE = "http://www.zotero.org/styles/"

STYLE_ALIASES = {
    "gb7714": "china-national-standard-gb-t-7714-2015-numeric",
    "gbt7714": "china-national-standard-gb-t-7714-2015-numeric",
    "gb-t-7714": "china-national-standard-gb-t-7714-2015-numeric",
    "gb7714-numeric": "china-national-standard-gb-t-7714-2015-numeric",
    "gb7714-author-date": "china-national-standard-gb-t-7714-2015-author-date",
    "afm": "advanced-functional-materials",
    "acs": "american-chemical-society",
    "apa": "apa",
    "ieee": "ieee",
    "nature": "nature",
}


def w(tag):
    return "{%s}%s" % (W_NS, tag)


def el(tag, **attrs):
    e = etree.Element(w(tag))
    for k, v in attrs.items():
        e.set(w(k), v)
    return e


def rand_id(n=8):
    alphabet = string.ascii_letters + string.digits
    return "".join(random.choice(alphabet) for _ in range(n))


# --------------------------------------------------------------------------
# Zotero 本地库查询（只读，immutable 模式，可在 Zotero 运行时读取）
# --------------------------------------------------------------------------
def connect_library(db_path=None):
    db_path = db_path or os.environ.get("ZOTERO_SQLITE") or default_db()
    uri = "file:%s?immutable=1" % db_path.replace("\\", "/")
    con = sqlite3.connect(uri, uri=True, timeout=8)
    con.row_factory = sqlite3.Row
    return con


def resolve_keys(keys, db_path=None):
    """把 Zotero item key 解析为 (library_id, user_id, {key: itemID})"""
    con = connect_library(db_path)
    c = con.cursor()
    row = c.execute(
        "SELECT libraryID FROM libraries WHERE type='user' ORDER BY libraryID LIMIT 1"
    ).fetchone()
    library_id = row[0] if row else 1
    row = c.execute("SELECT userID FROM users LIMIT 1").fetchone()
    user_id = row[0] if row else None
    ids = {}
    for k in keys:
        r = c.execute(
            "SELECT itemID FROM items WHERE key=? AND libraryID=?", (k, library_id)
        ).fetchone()
        ids[k] = r[0] if r else None
    con.close()
    return library_id, user_id, ids


def item_title(item_id, db_path=None):
    con = connect_library(db_path)
    r = con.execute(
        """SELECT v.value FROM itemData d
           JOIN itemDataValues v ON v.valueID = d.valueID
           JOIN fields f ON f.fieldID = d.fieldID AND f.fieldName='title'
           WHERE d.itemID=? LIMIT 1""",
        (item_id,),
    ).fetchone()
    con.close()
    return r[0] if r else ""


# --------------------------------------------------------------------------
# 域构造
# --------------------------------------------------------------------------
def run_text(text, superscript=False):
    r = el("r")
    if superscript:
        rPr = el("rPr")
        rPr.append(el("vertAlign", val="superscript"))
        r.append(rPr)
    t = el("t")
    t.set(XML_SPACE, "preserve")
    t.text = text
    r.append(t)
    return r


def run_field_char(kind):
    r = el("r")
    r.append(el("fldChar", fldCharType=kind))
    return r


def run_instr(code):
    r = el("r")
    it = el("instrText")
    it.set(XML_SPACE, "preserve")
    it.text = code
    r.append(it)
    return r


def field_runs(instr_code, display_runs):
    """begin + instrText + separate + 显示内容 + end"""
    runs = [run_field_char("begin"), run_instr(instr_code), run_field_char("separate")]
    runs.extend(display_runs)
    runs.append(run_field_char("end"))
    return runs


def build_citation_code(
    keys,
    user_id,
    item_ids,
    label,
    citation_id=None,
    note_index=0,
    superscript=True,
):
    """构造  ADDIN ZOTERO_ITEM CSL_CITATION {json}  的域代码"""
    citation_items = []
    for k in keys:
        item = {"uris": ["http://zotero.org/users/%s/items/%s" % (user_id, k)]}
        iid = item_ids.get(k)
        if iid:
            item["id"] = iid
        citation_items.append(item)

    payload = {
        "citationID": citation_id or rand_id(),
        "properties": {
            "unsorted": False,
            "formattedCitation": (
                "\\super %s\\nosupersub{}" % label if superscript else label
            ),
            "plainCitation": label,
            "noteIndex": note_index,
        },
        "citationItems": citation_items,
        "schema": CITATION_SCHEMA,
    }
    body = json.dumps(payload, ensure_ascii=False, separators=(",", ":"))
    return " ADDIN ZOTERO_ITEM CSL_CITATION " + body


def build_biblio_code():
    payload = {"uncited": [], "omitted": [], "custom": []}
    body = json.dumps(payload, ensure_ascii=False, separators=(",", ":"))
    return " ADDIN ZOTERO_BIBL " + body + " CSL_BIBLIOGRAPHY "


# --------------------------------------------------------------------------
# docx 读写
# --------------------------------------------------------------------------
def load_docx(path):
    z = zipfile.ZipFile(path)
    root = etree.fromstring(z.read(DOC_PART))
    return z, root


def save_docx(zin, root, dst):
    body = etree.tostring(root, encoding="unicode")
    data = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' + body
    ).encode("utf-8")
    with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        for info in zin.infolist():
            if info.filename == DOC_PART:
                zout.writestr(info, data)
            else:
                zout.writestr(info, zin.read(info.filename))


# --------------------------------------------------------------------------
# 文档样式偏好（docProps/custom.xml 的 ZOTERO_PREF_n）
# --------------------------------------------------------------------------
def read_prefs(zin):
    """拼接所有 ZOTERO_PREF_n 分块，返回完整 prefs 字符串（无则 None）"""
    try:
        raw = zin.read(CUSTOM_PART).decode("utf-8")
    except KeyError:
        return None
    parts = re.findall(
        r'<property[^>]*name="ZOTERO_PREF_(\d+)"[^>]*>\s*<vt:lpwstr>(.*?)</vt:lpwstr>',
        raw,
        re.S,
    )
    if not parts:
        return None
    parts.sort(key=lambda x: int(x[0]))
    return "".join(html.unescape(v) for _, v in parts)


def current_style_id(zin):
    prefs = read_prefs(zin)
    if not prefs:
        return None, None
    m = re.search(r'<style\s+id="([^"]*)"([^/>]*)/?>', prefs)
    if not m:
        return prefs, None
    return prefs, m.group(1)


def resolve_style_id(value):
    """把用户输入解析成 CSL 样式 id"""
    v = value.strip()
    if v.lower().endswith(".csl") and os.path.exists(v):
        txt = open(v, encoding="utf-8").read()
        m = re.search(r"<id>\s*([^<\s]+)\s*</id>", txt)
        if not m:
            raise SystemExit("无法从 %s 中解析 <id>" % v)
        return m.group(1).strip()
    if v.startswith("http://") or v.startswith("https://"):
        return v
    key = v.lower().replace(" ", "").replace("/", "")
    if key in STYLE_ALIASES:
        return STYLE_BASE + STYLE_ALIASES[key]
    return STYLE_BASE + v


def rewrite_prefs(zin, dst, new_prefs):
    """按 255 字符重新分块写回 ZOTERO_PREF_n，保留其它自定义属性"""
    ns = {"cp": CP_NS, "vt": VT_NS}
    root = etree.fromstring(zin.read(CUSTOM_PART))
    chunks = [new_prefs[i:i + CHUNK] for i in range(0, len(new_prefs), CHUNK)] or [""]

    existing = []
    for prop in root.findall("cp:property", ns):
        m = PREF_NAME_RE.match(prop.get("name") or "")
        if m:
            existing.append((int(m.group(1)), prop))
    existing.sort(key=lambda x: x[0])

    # 复用已有的属性元素承载分块，多退少补
    used_pids = set()
    for prop in root.findall("cp:property", ns):
        pid = prop.get("pid")
        if pid and prop.get("name") not in (None,) and not PREF_NAME_RE.match(prop.get("name") or ""):
            used_pids.add(int(pid))

    reused, extra = existing[:len(chunks)], existing[len(chunks):]
    for i, chunk in enumerate(chunks):
        if i < len(reused):
            prop = reused[i][1]
        else:
            pid = 2
            all_pids = {
                int(p.get("pid")) for p in root.findall("cp:property", ns) if p.get("pid")
            }
            while pid in all_pids:
                pid += 1
            prop = etree.SubElement(root, "{%s}property" % CP_NS)
            prop.set("fmtid", FMTID)
            prop.set("pid", str(pid))
            prop.set("name", "ZOTERO_PREF_%d" % (i + 1))
        for child in list(prop):
            prop.remove(child)
        lp = etree.SubElement(prop, "{%s}lpwstr" % VT_NS)
        lp.text = chunk
    for _, prop in extra:
        root.remove(prop)

    data = etree.tostring(
        root, encoding="utf-8", xml_declaration=True, standalone=True
    )
    with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        for info in zin.infolist():
            if info.filename == CUSTOM_PART:
                zout.writestr(info, data)
            else:
                zout.writestr(info, zin.read(info.filename))


def set_style(zin, style_id):
    prefs, old = current_style_id(zin)
    if prefs is None:
        raise SystemExit(
            "文档中没有 ZOTERO_PREF 属性：说明它尚未被 Zotero 插件处理过。\n"
            "请先在 Word 里用 Zotero 插件插入任意一条引用（建立文档偏好）后再试。"
        )
    m = re.search(r'(<style\s+id=")([^"]*)(")', prefs)
    if not m:
        raise SystemExit("prefs 中未找到 <style id=.../> 元素")
    new_prefs = prefs[: m.start(2)] + style_id + prefs[m.end(2):]
    return old, new_prefs


def paragraphs(root):
    body = root.find(w("body"))
    return [p for p in body if p.tag == w("p")]


def make_paragraph(runs):
    p = el("p")
    for r in runs:
        p.append(r)
    return p


def insert_paragraph_after(root, index, para):
    """在第 index 个段落之后插入（index=None 表示追加到文档末尾）"""
    body = root.find(w("body"))
    ps = paragraphs(root)
    if index is None or index >= len(ps) - 1:
        ps[-1].addnext(para)
    else:
        ps[index].addnext(para)


# --------------------------------------------------------------------------
# 命令实现
# --------------------------------------------------------------------------
def split_fields(code):
    """把一个段落里拼接起来的多个域代码拆开，返回 [(kind, payload_str)]"""
    out = []
    for part in re.split(r"(?=ADDIN ZOTERO_)", code):
        part = part.strip()
        if not part:
            continue
        if "CSL_CITATION" in part:
            body = part.split("CSL_CITATION", 1)[1]
            out.append(("citation", _first_json(body)))
        elif "CSL_BIBLIOGRAPHY" in part:
            out.append(("bibliography", _first_json(part.split("CSL_BIBLIOGRAPHY", 1)[0]
                                                    .split("ZOTERO_BIBL", 1)[-1])))
        else:
            out.append(("other", ""))
    return out


def _first_json(s):
    s = s.strip()
    start = s.find("{")
    if start < 0:
        return ""
    depth = 0
    for i in range(start, len(s)):
        if s[i] == "{":
            depth += 1
        elif s[i] == "}":
            depth -= 1
            if depth == 0:
                return s[start:i + 1]
    return ""


def cmd_inspect(args):
    z, root = load_docx(args.docx)
    count = 0
    for i, p in enumerate(paragraphs(root)):
        instrs = p.findall(".//" + w("instrText"))
        if not instrs:
            continue
        code = "".join(x.text or "" for x in instrs)
        texts = "".join((t.text or "") for t in p.findall(".//" + w("t")))
        fields = split_fields(code)
        count += len(fields)
        print("[段 %d] %d 个域" % (i, len(fields)))
        for kind, payload in fields:
            print("  - %s" % kind)
            if kind == "citation" and payload:
                try:
                    data = json.loads(payload)
                    print("    citationID: %s" % data.get("citationID"))
                    print("    plainCitation: %s"
                          % data.get("properties", {}).get("plainCitation"))
                    for ci in data.get("citationItems", []):
                        uris = ci.get("uris", [])
                        print("    item: %s" % (",".join(u.rsplit("/", 1)[-1] for u in uris)
                                                or ci.get("id")))
                except Exception as e:
                    print("    (JSON 解析失败: %s)" % e)
            elif kind == "bibliography":
                print("    %s" % (payload[:60] or "(空 payload)"))
        print("    显示文本: %s" % repr(texts[:60]))
    print("共 %d 个域" % count)
    z.close()


def cmd_cite(args):
    keys = [k.strip() for k in args.keys.split(",") if k.strip()]
    library_id, user_id, item_ids = resolve_keys(keys, args.db)
    missing = [k for k, v in item_ids.items() if not v]
    if missing:
        print("警告：以下 key 不在本地库中（将仅写入 uris）: %s" % ", ".join(missing))
    if user_id is None and not args.user_id:
        raise SystemExit("无法从本地库确定 userID，请用 --user-id 指定")
    user_id = args.user_id or user_id

    label = args.label or ("[%s]" % ",".join(["?"] * len(keys)))
    code = build_citation_code(
        keys,
        user_id,
        item_ids,
        label,
        superscript=not args.no_superscript,
    )
    runs = field_runs(code, [run_text(label, superscript=not args.no_superscript)])
    para = make_paragraph(runs)

    z, root = load_docx(args.docx)
    insert_paragraph_after(root, args.after_para, para)
    save_docx(z, root, args.out)
    z.close()
    print("已写入引用域 -> %s" % args.out)
    print("  keys=%s  userID=%s  label=%s" % (",".join(keys), user_id, label))
    print("  域长度 %d 字符" % len(code))


def cmd_biblio(args):
    code = build_biblio_code()
    para = make_paragraph(field_runs(code, [run_text("参考文献表（请在 Word 中刷新）")]))
    z, root = load_docx(args.docx)
    insert_paragraph_after(root, args.after_para, para)
    save_docx(z, root, args.out)
    z.close()
    print("已写入文献表域 -> %s" % args.out)


def cmd_style(args):
    z = zipfile.ZipFile(args.docx)
    prefs, style_id = current_style_id(z)
    if prefs is None:
        print("文档中没有 ZOTERO_PREF 属性（未经过 Zotero 插件处理）")
    else:
        print("当前样式 id : %s" % style_id)
        print("样式短名    : %s" % (style_id.rsplit("/", 1)[-1] if style_id else "?"))
        print("prefs 长度  : %d 字符" % len(prefs))
        print("原始分块数  : %d（每块 255 字符）" % (
            (len(prefs) + 254) // 255))
    z.close()


def cmd_setstyle(args):
    z = zipfile.ZipFile(args.docx)
    style_id = resolve_style_id(args.style)
    old_prefs, old = current_style_id(z)
    _, new_prefs = set_style(z, style_id)
    rewrite_prefs(z, args.out, new_prefs)
    z.close()
    print("样式已改写 -> %s" % args.out)
    print("  原样式: %s" % old)
    print("  新样式: %s" % style_id)
    print("  prefs %d -> %d 字符" % (len(old_prefs), len(new_prefs)))
    print("  提示：在 Word 中打开后点 Zotero 插件的 Refresh，全篇引文与文献表将按新样式重排。")


# --------------------------------------------------------------------------
# 内联渲染（rewrite）：把带标记的正文写成「含引用域的段落」
#   标记：{{cite:KEY1,KEY2|[2–4]}}  引用域（label 省略则按出现序自动编号）
#         {{it:...}}               斜体      {{bd:...}} 加粗
#         {{ph:...}}               占位符（红字 + 黄底，自动加 ⟦⟧）
# --------------------------------------------------------------------------
MARKUP_RE = re.compile(r"\{\{(cite|it|bd|ph|color):(.*?)\}\}", re.S)

DEFAULT_RPR = (
    '<w:rPr xmlns:w="%s"><w:rFonts w:ascii="Times New Roman" '
    'w:hAnsi="Times New Roman" w:cs="Times New Roman"/></w:rPr>' % W_NS
)

EMPTY_CUSTOM = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
    '<Properties xmlns="%s" xmlns:vt="%s"/>' % (CP_NS, VT_NS)
)
CT_OVERRIDE = (
    '<Override PartName="/docProps/custom.xml" ContentType='
    '"application/vnd.openxmlformats-officedocument.custom-properties+xml"/>'
)
REL_CUSTOM = (
    '<Relationship Id="rIdZoteroCustomProps" Type="http://schemas.openxmlformats.org/'
    'officeDocument/2006/relationships/custom-properties" Target="docProps/custom.xml"/>'
)


def body_paragraphs(root):
    body = root.find(w("body"))
    return [p for p in body if p.tag == w("p")]


def clone_rpr(p):
    """取段落 p 的 rPr（pPr/rPr 或首个 run 的 rPr），无则用默认 Times New Roman"""
    pPr = p.find(w("pPr"))
    if pPr is not None:
        rpr = pPr.find(w("rPr"))
        if rpr is not None:
            return etree.fromstring(etree.tostring(rpr))
    for r in p.findall(w("r")):
        rpr = r.find(w("rPr"))
        if rpr is not None:
            return etree.fromstring(etree.tostring(rpr))
    return etree.fromstring(DEFAULT_RPR)


def _add_rpr_child(rpr, tag, **attrs):
    e = etree.SubElement(rpr, w(tag))
    for k, v in attrs.items():
        e.set(w(k), v)
    return e


def make_run(text, base_rpr, italic=False, bold=False, superscript=False,
             color=None, highlight=None):
    r = el("r")
    rpr = etree.fromstring(etree.tostring(base_rpr))
    if italic:
        _add_rpr_child(rpr, "i")
    if bold:
        _add_rpr_child(rpr, "b")
    if color:
        _add_rpr_child(rpr, "color", val=color)
    if highlight:
        _add_rpr_child(rpr, "highlight", val=highlight)
    if superscript:
        _add_rpr_child(rpr, "vertAlign", val="superscript")
    if len(rpr):
        r.append(rpr)
    t = el("t")
    t.set(XML_SPACE, "preserve")
    t.text = text
    r.append(t)
    return r


def build_para(markup, idx, user_id, item_ids, base_rpr, counter,
               cite_superscript=True):
    """把一段标记文本渲染成 run 序列"""
    runs = []
    pos = 0
    for m in MARKUP_RE.finditer(markup):
        if m.start() > pos:
            runs.append(make_run(markup[pos:m.start()], base_rpr))
        kind, body = m.group(1), m.group(2)
        if kind == "cite":
            keys_part, _, label = body.partition("|")
            keys = [k.strip() for k in keys_part.split(",") if k.strip()]
            if not label:
                counter[0] += 1
                label = "[%d]" % counter[0]
            code = build_citation_code(keys, user_id, item_ids, label)
            runs.extend(field_runs(code, [make_run(label, base_rpr, superscript=cite_superscript)]))
        elif kind == "it":
            runs.append(make_run(body, base_rpr, italic=True))
        elif kind == "bd":
            runs.append(make_run(body, base_rpr, bold=True))
        elif kind == "ph":
            runs.append(make_run("\u27e6" + body + "\u27e7", base_rpr,
                                 color="FF0000", highlight="yellow"))
        elif kind == "color":
            val, _, txt = body.partition("|")
            runs.append(make_run(txt, base_rpr, color=val))
        pos = m.end()
    if pos < len(markup):
        runs.append(make_run(markup[pos:], base_rpr))
    return runs


def write_prefs_into_root(root, prefs):
    """在 custom.xml 根元素里写入/更新 ZOTERO_PREF_n（按 255 字符分块）"""
    chunks = [prefs[i:i + CHUNK] for i in range(0, len(prefs), CHUNK)] or [""]
    ns = {"cp": CP_NS}
    existing = []
    for prop in root.findall("cp:property", ns):
        m = PREF_NAME_RE.match(prop.get("name") or "")
        if m:
            existing.append((int(m.group(1)), prop))
    existing.sort(key=lambda x: x[0])
    reused, extra = existing[:len(chunks)], existing[len(chunks):]
    for i, chunk in enumerate(chunks):
        if i < len(reused):
            prop = reused[i][1]
        else:
            pids = {int(p.get("pid")) for p in root.findall("cp:property", ns) if p.get("pid")}
            pid = 2
            while pid in pids:
                pid += 1
            prop = etree.SubElement(root, "{%s}property" % CP_NS)
            prop.set("fmtid", FMTID)
            prop.set("pid", str(pid))
            prop.set("name", "ZOTERO_PREF_%d" % (i + 1))
        for child in list(prop):
            prop.remove(child)
        lp = etree.SubElement(prop, "{%s}lpwstr" % VT_NS)
        lp.text = chunk
    for _, prop in extra:
        root.remove(prop)


def make_prefs(style_id, session_id=None):
    return (
        '<data data-version="3" zotero-version="9.0.6">'
        '<session id="%s"/>'
        '<style id="%s" hasBibliography="1" bibliographyStyleHasBeenSet="1"/>'
        '<prefs><pref name="fieldType" value="Field"/>'
        '<pref name="automaticJournalAbbreviations" value="true"/></prefs>'
        "</data>" % (session_id or rand_id(8), style_id)
    )


def save_package(src, dst, doc_bytes=None, prefs_str=None):
    zin = zipfile.ZipFile(src)
    names = zin.namelist()
    with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        for info in zin.infolist():
            n = info.filename
            if n == DOC_PART and doc_bytes is not None:
                zout.writestr(info, doc_bytes)
                continue
            if n == CUSTOM_PART and prefs_str is not None:
                continue  # 统一在末尾按需重写
            if n in ("[Content_Types].xml", "_rels/.rels") and prefs_str is not None:
                continue
            zout.writestr(info, zin.read(n))
        if prefs_str is not None:
            if CUSTOM_PART in names:
                custom_root = etree.fromstring(zin.read(CUSTOM_PART))
            else:
                custom_root = etree.fromstring(EMPTY_CUSTOM.encode("utf-8"))
            write_prefs_into_root(custom_root, prefs_str)
            zout.writestr(CUSTOM_PART, etree.tostring(
                custom_root, encoding="utf-8", xml_declaration=True, standalone=True))
            if CUSTOM_PART not in names:
                ct = zin.read("[Content_Types].xml").decode("utf-8")
                if "docProps/custom.xml" not in ct:
                    ct = ct.replace("</Types>", CT_OVERRIDE + "</Types>")
                zout.writestr("[Content_Types].xml", ct.encode("utf-8"))
                rels = zin.read("_rels/.rels").decode("utf-8")
                if "docProps/custom.xml" not in rels:
                    rels = rels.replace("</Relationships>", REL_CUSTOM + "</Relationships>")
                zout.writestr("_rels/.rels", rels.encode("utf-8"))
    zin.close()


def cmd_rewrite(args):
    spec = json.load(open(args.spec, encoding="utf-8"))

    z, root = load_docx(args.docx)
    ps = body_paragraphs(root)

    allkeys = []
    for para in spec.get("paragraphs", []):
        allkeys.extend(re.findall(r"\{\{cite:([^|}]+)", para.get("text", "")))
    keys = []
    for group in allkeys:
        keys.extend(k.strip() for k in group.split(",") if k.strip())
    library_id, user_id, item_ids = resolve_keys(keys, args.db)
    user_id = args.user_id or user_id
    missing = [k for k, v in item_ids.items() if not v]
    if missing:
        print("警告：%d 个 key 不在本地库中：%s" % (len(missing), ", ".join(missing)))
    print("userID=%s  引用条目 %d 个（去重 %d）" % (user_id, len(keys), len(set(keys))))

    counter = [0]
    for para in spec.get("paragraphs", []):
        i = para["index"]
        p = ps[i]
        base = clone_rpr(p)
        pPr = p.find(w("pPr"))
        for child in list(p):
            if child is not pPr:
                p.remove(child)
        for r in build_para(para["text"], i, user_id, item_ids, base, counter,
                            cite_superscript=not getattr(args, "cite_inline", False)):
            p.append(r)
        print("  已重写段 %d（%d 字符）" % (i, len(re.sub(r"\{\{.*?\}\}", "x", para["text"]))))

    # 文献表域：追加到正文末尾（sectPr 之前）
    bib = spec.get("bibliography")
    if bib:
        body = root.find(w("body"))
        sect = body.find(w("sectPr"))
        anchor = None
        for p in body_paragraphs(root):
            anchor = p
        h_rpr = etree.fromstring(DEFAULT_RPR)
        heading = make_paragraph([make_run(bib.get("heading", "References"), h_rpr, bold=True)])
        biblio = make_paragraph(field_runs(
            build_biblio_code(),
            [make_run("参考文献表（在 Word 中刷新后生成）", h_rpr)]))
        if anchor is not None:
            anchor.addnext(heading)
            heading.addnext(biblio)
        elif sect is not None:
            sect.addprevious(heading)
            sect.addprevious(biblio)
        print("  已追加文献表域（标题：%s）" % bib.get("heading", "References"))

    doc_bytes = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
                 + etree.tostring(root, encoding="unicode")).encode("utf-8")
    prefs_str = None
    if spec.get("style"):
        prefs_str = make_prefs(resolve_style_id(spec["style"]), spec.get("session"))
        print("  已写入文档引文样式：%s" % spec["style"])
    save_package(args.docx, args.out, doc_bytes, prefs_str)
    z.close()
    print("已输出 -> %s" % args.out)


def _render_markup(markup, idx, user_id, item_ids, base_rpr, counter,
                   cite_superscript=True):
    """把含 {{cite:}}/{{it:}}/{{bd:}}/{{ph:}}/{{color:}} 标记的文本渲染成 run 序列
    （rewrite 的 build_para 复用）。"""
    return build_para(markup, idx, user_id, item_ids, base_rpr, counter,
                      cite_superscript=cite_superscript)


def cmd_replace(args):
    """run 级定点文本替换：保留全部 Zotero 域（含已刷新的域结果），可另存新版本。

    --edit '{"index":59,"old":"原句","new":"新句"}'，可重复；old 必须在该段可见文本中
    唯一出现，且不得跨越引用域（fldChar/instrText）。
    new 可含 {{cite:KEY1,KEY2|[标签]}} / {{it:}} / {{bd:}} / {{ph:}} 标记，
    会渲染为真正的 Zotero 引用域与格式 run（key 须在库中，可用 --db 指定库）。
    段落插入形式：{"index":59,"insert_after":"带标记的整段文本","clone_from":3}
    —— 在段 59 之后插入新段落，段落属性/字体克隆自段 clone_from（缺省=index）。
    """
    edits = [json.loads(s) for s in args.edit]
    if args.edits_file:
        edits.extend(json.load(open(args.edits_file, encoding="utf-8")))

    z, root = load_docx(args.docx)
    ps = body_paragraphs(root)

    # 预解析全部 cite key，提前发现不在库中的 key
    allkeys = []
    for ed in edits:
        for field in ("new", "insert_after"):
            for m in re.finditer(r"\{\{cite:([^|}]+)", ed.get(field) or ""):
                allkeys.extend(k.strip() for k in m.group(1).split(",") if k.strip())
    user_id, item_ids = None, {}
    if allkeys:
        _, user_id, item_ids = resolve_keys(sorted(set(allkeys)), args.db)
        miss = [k for k, v in item_ids.items() if v is None]
        if miss:
            raise SystemExit("以下 key 不在 Zotero 库中：%s" % ",".join(miss))

    counter = [0]
    pending_inserts = []  # (seq, index, 新段落元素)

    for seq, ed in enumerate(edits):
        idx = ed["index"]
        if idx >= len(ps):
            raise SystemExit("段落 %d 不存在（共 %d 段）" % (idx, len(ps)))
        p = ps[idx]

        # ---- 形式二：整段插入 ----
        if "insert_after" in ed:
            src = ps[ed.get("clone_from", idx)]
            newp = el("p")
            src_ppr = src.find(w("pPr"))
            if src_ppr is not None:
                newp.append(etree.fromstring(etree.tostring(src_ppr)))
            base_rpr = clone_rpr(src)
            for rr in _render_markup(ed["insert_after"], idx, user_id,
                                     item_ids, base_rpr, counter,
                                     cite_superscript=not getattr(args, "cite_inline", False)):
                newp.append(rr)
            pending_inserts.append((seq, idx, newp))
            print("  段 %d 后将插入新段落（%d 字符，克隆自段 %d）"
                  % (idx, len(ed["insert_after"]), ed.get("clone_from", idx)))
            continue

        # ---- 形式一：run 级文本替换 ----
        runs = p.findall(w("r"))

        # 可参与替换的 run：含 w:t 且不带 fldChar / instrText
        spans = []
        parts = []
        pos = 0
        for r in runs:
            if r.find(w("fldChar")) is not None or r.find(w("instrText")) is not None:
                continue
            t = r.find(w("t"))
            if t is None or t.text is None:
                continue
            parts.append(t.text)
            spans.append((pos, pos + len(t.text), r))
            pos += len(t.text)
        text = "".join(parts)
        n = text.count(ed["old"])
        if n == 0:
            raise SystemExit("段 %d 未找到待替换文本：%r" % (idx, ed["old"][:60]))
        if n > 1:
            raise SystemExit("段 %d 待替换文本出现 %d 次，无法定位：%r" % (idx, n, ed["old"][:60]))
        s = text.index(ed["old"])
        e = s + len(ed["old"])

        hit = [(a, b, r) for a, b, r in spans if a < e and b > s]
        if not hit:
            raise SystemExit("段 %d 的替换范围未落在任何文本 run 上" % idx)
        for a, b, r in hit:  # 域不能被切断
            if r.find(w("fldChar")) is not None or r.find(w("instrText")) is not None:
                raise SystemExit("段 %d 的替换范围跨越了 Zotero 域，请改用 rewrite" % idx)

        new = ed["new"]
        if "{{" in new:
            # 标记渲染路径：head 留在首 hit run，rendered 插其后，tail 留在末 hit run
            a0, b0, r0 = hit[0]
            aN, bN, rN = hit[-1]
            rpr0 = r0.find(w("rPr"))
            base_rpr = (etree.fromstring(etree.tostring(rpr0))
                        if rpr0 is not None else clone_rpr(p))
            rendered = _render_markup(new, idx, user_id, item_ids, base_rpr, counter,
                                      cite_superscript=not getattr(args, "cite_inline", False))
            t0 = r0.find(w("t"))
            head = t0.text[: s - a0] if a0 < s else ""
            t0.text = head
            t0.set(XML_SPACE, "preserve")
            pos_in_p = list(p).index(r0)
            for k, rr in enumerate(rendered):
                p.insert(pos_in_p + 1 + k, rr)
            if rN is r0:
                tail = text[e:b0] if b0 > e else ""
                if tail:
                    p.insert(pos_in_p + 1 + len(rendered),
                             make_run(tail, base_rpr))
            else:
                tN = rN.find(w("t"))
                tN.text = text[e:bN] if bN > e else ""
                tN.set(XML_SPACE, "preserve")
                for a, b, r in hit[1:-1]:
                    p.remove(r)
                if (tN.text or "") == "":
                    p.remove(rN)
            if (t0.text or "") == "":
                p.remove(r0)
            print("  段 %d 已替换（含标记渲染）：%r -> %r"
                  % (idx, ed["old"][:36], re.sub(r"\{\{.*?\}\}", "", new)[:36]))
        else:
            for i, (a, b, r) in enumerate(hit):
                t = r.find(w("t"))
                head = t.text[: s - a] if a < s else ""
                tail = t.text[e - a:] if b > e else ""
                t.text = (head + new + tail) if i == 0 else (head + tail)
                t.set(XML_SPACE, "preserve")
            for a, b, r in hit:  # 清空后的 run 直接删掉
                t = r.find(w("t"))
                if (t.text or "") == "":
                    p.remove(r)
            print("  段 %d 已替换：%r -> %r（涉及 %d 个 run）"
                  % (idx, ed["old"][:40], new[:40], len(hit)))

    # 统一插入新段落：索引降序 + 同索引按登记逆序，避免位置漂移且保持书写顺序
    by_idx = {}
    for seq, idx, newp in pending_inserts:
        by_idx.setdefault(idx, []).append((seq, newp))
    for idx in sorted(by_idx, reverse=True):
        anchor = body_paragraphs(root)[idx]
        for seq, newp in sorted(by_idx[idx], reverse=True):
            anchor.addnext(newp)

    doc_bytes = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
                 + etree.tostring(root, encoding="unicode")).encode("utf-8")
    save_package(args.docx, args.out, doc_bytes)
    z.close()
    print("已输出 -> %s" % args.out)


def main():
    ap = argparse.ArgumentParser(description="Zotero 可刷新引用域工具")
    sub = ap.add_subparsers(dest="cmd", required=True)

    p1 = sub.add_parser("inspect", help="列出文档中的 Zotero 域")
    p1.add_argument("--docx", required=True)
    p1.set_defaults(func=cmd_inspect)

    p2 = sub.add_parser("cite", help="插入引用域")
    p2.add_argument("--docx", required=True)
    p2.add_argument("--out", required=True)
    p2.add_argument("--keys", required=True, help="逗号分隔的 Zotero item key")
    p2.add_argument("--after-para", type=int, default=None, help="在该段落索引后插入")
    p2.add_argument("--label", default=None, help="显示文本，默认 [1]")
    p2.add_argument("--no-superscript", action="store_true")
    p2.add_argument("--db", default=None, help="zotero.sqlite 路径")
    p2.add_argument("--user-id", default=None, help="覆盖 userID")
    p2.set_defaults(func=cmd_cite)

    p3 = sub.add_parser("biblio", help="插入文献表域")
    p3.add_argument("--docx", required=True)
    p3.add_argument("--out", required=True)
    p3.add_argument("--after-para", type=int, default=None)
    p3.set_defaults(func=cmd_biblio)

    p4 = sub.add_parser("style", help="查看文档当前的 Zotero 引文样式")
    p4.add_argument("--docx", required=True)
    p4.set_defaults(func=cmd_style)

    p5 = sub.add_parser("setstyle", help="改写文档的 Zotero 引文样式")
    p5.add_argument("--docx", required=True)
    p5.add_argument("--out", required=True)
    p5.add_argument(
        "--style",
        required=True,
        help="样式 id / 短名 / .csl 路径；内置别名 gb7714、gb7714-author-date、afm、acs、apa、ieee、nature",
    )
    p5.set_defaults(func=cmd_setstyle)

    p6 = sub.add_parser("rewrite", help="按 spec 把整段重写为含引用域/文献表域的段落")
    p6.add_argument("--docx", required=True)
    p6.add_argument("--out", required=True)
    p6.add_argument("--spec", required=True,
                    help='JSON：{"style":"afm","bibliography":{"heading":"References"},'
                         '"paragraphs":[{"index":4,"text":"...{{cite:KEY|[1]}}..."}]}')
    p6.add_argument("--db", default=None)
    p6.add_argument("--user-id", default=None)
    p6.add_argument("--cite-inline", action="store_true",
                    help="{{cite:}} 显示为行内文本（不加 vertAlign 上标），用于 CEJ/Elsevier 方括号行内样式")
    p6.set_defaults(func=cmd_rewrite)

    p7 = sub.add_parser("replace", help="按段落索引做 run 级定点替换（保留 Zotero 域，new 支持 {{cite:}} 等标记）")
    p7.add_argument("--docx", required=True)
    p7.add_argument("--out", required=True)
    p7.add_argument("--edit", action="append", default=[],
                    help='JSON：{"index":59,"old":"原句","new":"新句"} 或 {"index":59,"insert_after":"…","clone_from":3}；可重复')
    p7.add_argument("--edits-file", default=None, help="批量编辑 JSON 文件（对象数组，与 --edit 合并）")
    p7.add_argument("--db", default=None, help="zotero.sqlite 路径（{{cite:}} 的 key 解析用）")
    p7.add_argument("--cite-inline", action="store_true",
                    help="{{cite:}} 显示为行内文本（不加 vertAlign 上标），用于 CEJ/Elsevier 方括号行内样式")
    p7.set_defaults(func=cmd_replace)

    args = ap.parse_args()
    args.func(args)


if __name__ == "__main__":
    main()
