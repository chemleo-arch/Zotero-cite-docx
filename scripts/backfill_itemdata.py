# -*- coding: utf-8 -*-
"""backfill_itemdata.py — 为缺失 itemData 的 ZOTERO_ITEM 引用域补全 CSL 快照。

背景：zotero_cite.py 的 {{cite:}}/cite 命令写入的引用域只有 id+uris，没有 itemData
CSL 快照。Word 端 Zotero 插件 Refresh 时会跳过这类域——文献表不收录其引用的文献、
编号不重算，导致全文新旧编号混杂。本脚本扫描全文，把所有缺 itemData 的条目补全
（优先从文内其他域的快照收割，缺失的走 Zotero 本地 API 构建 CSL），并统一
properties 结构（formattedCitation/plainCitation/noteIndex + schema），使新域
与 Word 插件自建域结构完全一致。

用法（永远另存新文件，不覆盖原稿）：
    python backfill_itemdata.py --docx in.docx --out out.docx
前置：Zotero 客户端须在运行（本地 API 兜底用）；zotero.sqlite 路径按
ZOTERO_SQLITE / ZOTERO_DATA_DIR / 常见默认位置自动探测，也可用 --sqlite 指定。
"""
import argparse, json, os, re, sqlite3, sys, urllib.request
from docx import Document

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from zotero_paths import default_db

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'


def parse_instr(raw):
    m = re.search(r'ZOTERO_ITEM CSL_CITATION (\{.*\})\s*$', raw.strip(), re.S)
    return json.loads(m.group(1)) if m else None


def key_of(uri):
    m = re.search(r'/items/([A-Z0-9]{8})$', uri or '')
    return m.group(1) if m else None


def to_csl(d, uri):
    """Zotero API item -> CSL itemData（对齐 Word 插件域内快照的关键字段）"""
    csl = {'id': uri, 'type': {'journalArticle': 'article-journal'}.get(d.get('itemType'), d.get('itemType', 'document'))}
    for src, dst in (('title', 'title'), ('shortTitle', 'title-short'), ('publicationTitle', 'container-title'),
                     ('journalAbbreviation', 'journalAbbreviation'), ('volume', 'volume'), ('issue', 'issue'),
                     ('pages', 'page'), ('DOI', 'DOI'), ('ISSN', 'ISSN'), ('url', 'URL'),
                     ('language', 'language'), ('publisher', 'publisher'), ('abstractNote', 'abstract')):
        if d.get(src):
            csl[dst] = d[src]
    ym = re.search(r'(\d{4})', d.get('date') or '')
    if ym:
        csl['issued'] = {'date-parts': [[int(ym.group(1))]]}
    au = []
    for c in d.get('creators', []):
        au.append({'literal': c['name']} if 'name' in c else {'family': c.get('lastName', ''), 'given': c.get('firstName', '')})
    if au:
        csl['author'] = au
    return csl


def iter_fields(doc):
    """按 fldChar begin/separate/end 状态机跨 run 聚合域指令，yield (instrText元素列表, 完整指令文本)。

    关键：Word/Zotero 插件刷新后会把长域指令拆成多个 w:r/w:instrText run，
    逐 run 解析（或带 'ZOTERO_ITEM' in raw 守卫）会漏掉不含关键字的部分——
    必须先聚合成完整指令再解析。doc.element.body 含表格，覆盖全部域。"""
    state, buf = 'out', []
    for r in doc.element.body.iter(f'{W}r'):
        fld = r.find(f'{W}fldChar')
        if fld is not None:
            t = fld.get(f'{W}fldCharType')
            if t == 'begin':
                state, buf = 'code', []
            elif t == 'separate':
                if buf:
                    yield buf, ''.join(e.text or '' for e in buf)
                buf, state = [], 'disp'
            elif t == 'end':
                if buf:  # 无 separate 的简单域
                    yield buf, ''.join(e.text or '' for e in buf)
                buf, state = [], 'out'
            continue
        itn = r.find(f'{W}instrText')
        if itn is not None and state == 'code':
            buf.append(itn)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--docx', required=True)
    ap.add_argument('--out', required=True)
    ap.add_argument('--zotero-api', default='http://127.0.0.1:23119')
    ap.add_argument('--sqlite', default=None,
                    help='zotero.sqlite 路径；缺省按 ZOTERO_SQLITE / ZOTERO_DATA_DIR 探测')
    a = ap.parse_args()
    a.sqlite = a.sqlite or default_db()

    doc = Document(a.docx)

    # 1) 收割文内已有 itemData 快照 + schema 惯例
    harvest, ref_schema = {}, None
    for _els, raw in iter_fields(doc):
        if 'ZOTERO_ITEM' not in raw:
            continue
        d = parse_instr(raw)
        if not d:
            continue
        if not ref_schema and d.get('citationItems') and d['citationItems'][0].get('itemData'):
            ref_schema = d.get('schema')
        for it in d.get('citationItems', []):
            k = key_of((it.get('uris') or [''])[-1])
            if k and it.get('itemData'):
                harvest.setdefault(k, it['itemData'])

    con = sqlite3.connect(f'file:{a.sqlite}?immutable=1', uri=True)
    fixed = skipped = 0
    for els, raw in iter_fields(doc):
        if 'ZOTERO_ITEM' not in raw:
            continue
        d = parse_instr(raw)
        if not d:
            continue
        items = d.get('citationItems', [])
        if not items or all(it.get('itemData') for it in items):
            continue  # 已完整
        props = d.get('properties') or {}
        plain = props.get('plainCitation') or re.sub(r'\\super\s*|\{\}$|\s*\\nosupersub', '', props.get('formattedCitation', '')).strip()
        new_items = []
        for it in items:
            uri = (it.get('uris') or [''])[0]
            k = key_of(uri)
            idata = it.get('itemData') or harvest.get(k)
            if not idata and k:
                try:
                    req = urllib.request.Request(f'{a.zotero_api}/api/users/0/items/{k}?format=json')
                    idata = to_csl(json.loads(urllib.request.urlopen(req, timeout=10).read())['data'], uri)
                except Exception as ex:
                    print(f'  !! {k} API 失败: {ex}')
                    idata = {}
            row = con.execute('SELECT itemID FROM items WHERE key=?', (k,)).fetchone()
            new_items.append({'id': it.get('id', row[0] if row else 0),
                              'uris': it.get('uris', [uri]), 'itemData': idata or {}})
        new = {'citationID': d.get('citationID'), 'citationItems': new_items,
               'properties': {'formattedCitation': plain, 'plainCitation': plain, 'noteIndex': 0},
               'schema': d.get('schema') or ref_schema}
        prefix = raw[:raw.index('ZOTERO_ITEM')]
        # 指令被拆成多个 run 时：完整新指令写入首个元素，其余清空
        els[0].text = prefix + 'ZOTERO_ITEM CSL_CITATION ' + json.dumps(new, ensure_ascii=False, separators=(',', ':'))
        for e in els[1:]:
            e.text = ''
        fixed += 1
    con.close()
    doc.save(a.out)
    print(f'补全 {fixed} 个域, 已保存: {a.out}')
    if fixed:
        print('提醒: 回 Word 后需 Zotero -> Refresh 才会重算编号与文献表。')


if __name__ == '__main__':
    main()
