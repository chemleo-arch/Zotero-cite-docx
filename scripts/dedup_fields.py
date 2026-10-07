# -*- coding: utf-8 -*-
"""dedup_fields.py — Zotero 引用域内重复条目的域级去重。

场景：同一文献在 Zotero 库中存在两个条目（同 DOI、不同 key），正文不同位置
分别引用了两者，刷新后文献表出现重复条目、占两个编号。本脚本把引用 dup key
的 citationItem 整个替换为 keep key 的 item 对象（从文内已有域收割完整
id/uris/itemData），刷新后重复条目自动从文献表消失、编号重排。

keep/dup 判定：正文首次引用更早者保留（顺序编码制下即小编号）。
分组键：DOI 优先（小写），无 DOI 用规范化标题兜底。

用法（永远另存新文件，不覆盖原稿；先 --dry-run 确认分组无误）：
    python dedup_fields.py --docx in.docx --dry-run
    python dedup_fields.py --docx in.docx --out out.docx
"""
import argparse, json, re, copy
from collections import defaultdict
from docx import Document

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'


def iter_fields(doc):
    """fldChar begin/separate/end 状态机跨 run 聚合域指令（长指令会被插件拆成多 run）。"""
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
                if buf:
                    yield buf, ''.join(e.text or '' for e in buf)
                buf, state = [], 'out'
            continue
        itn = r.find(f'{W}instrText')
        if itn is not None and state == 'code':
            buf.append(itn)


def parse_instr(raw):
    m = re.search(r'ZOTERO_ITEM CSL_CITATION (\{.*\})\s*$', raw.strip(), re.S)
    return json.loads(m.group(1)) if m else None


def key_of(uri):
    m = re.search(r'/items/([A-Z0-9]{8})$', uri or '')
    return m.group(1) if m else None


def scan(doc):
    """返回 (fields, info)：fields=[(els, raw, parsed)], info=key -> {title,doi,first,n}"""
    fields, info = [], {}
    fidx = 0
    for els, raw in iter_fields(doc):
        if 'ZOTERO_ITEM' not in raw:
            continue
        d = parse_instr(raw)
        if not d:
            continue
        fidx += 1
        fields.append((els, raw, d))
        for it in d.get('citationItems', []):
            k = key_of((it.get('uris') or [''])[-1])
            if not k:
                continue
            idata = it.get('itemData') or {}
            if k not in info:
                info[k] = {'title': (idata.get('title') or '')[:80],
                           'doi': (idata.get('DOI') or '').lower().strip(),
                           'first': fidx, 'n': 0}
            info[k]['n'] += 1
    return fields, info


def find_dup_groups(info):
    groups = defaultdict(list)
    for k, v in info.items():
        gkey = 'doi:' + v['doi'] if v['doi'] else 'title:' + re.sub(r'\W+', '', v['title'].lower())[:60]
        groups[gkey].append(k)
    return {g: sorted(ks, key=lambda k: info[k]['first']) for g, ks in groups.items() if len(ks) > 1}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--docx', required=True)
    ap.add_argument('--out')
    ap.add_argument('--dry-run', action='store_true')
    a = ap.parse_args()

    doc = Document(a.docx)
    fields, info = scan(doc)
    dups = find_dup_groups(info)
    print('ZOTERO_ITEM fields: %d, unique keys: %d, dup groups: %d' % (len(fields), len(info), len(dups)))
    for g, ks in sorted(dups.items()):
        print('--- %s' % g)
        for k in ks:
            print('    %s first@field#%d cites=%d  %s' % (k, info[k]['first'], info[k]['n'], info[k]['title']))
        print('    => KEEP %s  DUP %s' % (ks[0], ks[1:]))
    if a.dry_run or not dups:
        return
    if not a.out:
        raise SystemExit('必须指定 --out（永不覆盖原稿）')

    dup2keep = {k: ks[0] for ks in dups.values() for k in ks[1:]}
    keep_items = {}
    for _els, _raw, d in fields:
        for it in d.get('citationItems', []):
            k = key_of((it.get('uris') or [''])[-1])
            if k in dup2keep.values() and k not in keep_items and it.get('itemData'):
                keep_items[k] = copy.deepcopy(it)
    missing = [k for k in dup2keep.values() if k not in keep_items]
    if missing:
        raise SystemExit('keep 条目缺 itemData 快照: %s' % missing)

    touched = replaced = 0
    for els, raw, d in fields:
        items = d.get('citationItems', [])
        hit = False
        for idx, it in enumerate(items):
            k = key_of((it.get('uris') or [''])[-1])
            if k in dup2keep:
                items[idx] = copy.deepcopy(keep_items[dup2keep[k]])
                replaced += 1
                hit = True
        if not hit:
            continue
        prefix = raw[:raw.index('ZOTERO_ITEM')]
        els[0].text = prefix + 'ZOTERO_ITEM CSL_CITATION ' + json.dumps(d, ensure_ascii=False, separators=(',', ':'))
        for e in els[1:]:
            e.text = ''
        touched += 1
    doc.save(a.out)
    print('touched %d fields, replaced %d items -> %s' % (touched, replaced, a.out))
    print('提醒: 回 Word 后需 Zotero -> Refresh 重算编号与文献表。')


if __name__ == '__main__':
    main()
