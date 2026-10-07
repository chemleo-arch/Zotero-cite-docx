---
name: zotero-cite-docx
description: 在 Word (.docx) 中写入 Zotero 可刷新的引用域（ADDIN ZOTERO_ITEM CSL_CITATION / ADDIN ZOTERO_BIBL CSL_BIBLIOGRAPHY），使文档能由 Word 里的 Zotero 插件刷新引文、自动重建参考文献表、一键切换引文样式；并可在不打开 Word 的情况下改换整篇文档的引文样式。用于需要"引用 Zotero 库中文献且格式可随样式切换"的中英文论文、基金申报书、学位论文写作。触发词：Zotero 引用、插入引用、参考文献表、引用域、可刷新引用、Zotero 样式、GB/T 7714 引用。
agent_created: true
---

# Zotero 可刷新引用域（.docx）

把 Zotero 的引用以**真正的域（field）**写进 `.docx`，而不是静态文本。写完后在 Word 里用
Zotero 插件点「Refresh」，引文编号、上标、参考文献表都会由 Zotero 按文档当前选定的样式重新渲染；
换期刊时只改一次文档样式，全篇引用连带文献表自动重排。

**已实测闭环**：脚本生成的引用域在 Word + Zotero 9.0.6 插件里 Refresh 通过——
新增引用正确显示为上标、文献表自动扩充。`uris` 定位方式确认有效。

## 何时用这个技能

- 用户要"按 Zotero 的格式引用文献""参考文献能自动更新""换刊换格式"
- 需要把 Zotero 库里的条目引用进 .docx（论文、标书、学位论文）
- 需要读取已有 .docx 里 Zotero 域的条目信息（`inspect`）
- 需要把整篇文档的引文样式换成另一种（`setstyle`），如投稿从某刊样式换成中文国标
- 需要在一份**已被 Zotero 刷新过**的稿子里改几句话而不破坏引用域（`replace`）

如果用户只要一份**静态**参考文献表（不需要在 Word 里刷新），直接按 CSL 规则排版即可，不需要本技能。

## 安装与依赖

```bash
pip install -r requirements.txt      # lxml（全部功能）+ python-docx（backfill/dedup 用）
```

前置条件：

1. 本机装有 **Zotero 客户端**（域只能在 Zotero 客户端里刷新；本技能不替代它）
2. Word 里装了 **Zotero 插件**（Zotero → 编辑 → 首选项 → 引用 → 文字处理软件）
3. 能访问本机 `zotero.sqlite`（只读即可，脚本用 `?immutable=1` 模式，**Zotero 开着也能读**）

## 运行环境与路径约定

- 脚本：`scripts/`
- **Zotero 数据目录**：优先取环境变量 `ZOTERO_DATA_DIR`（数据目录）或 `ZOTERO_SQLITE`
  （直接指向 `zotero.sqlite`）；未设置时由 `scripts/zotero_paths.py` 自动探测
  （Windows 依次试 `D:\Zotero`、`C:\Zotero`、`E:\Zotero`、`~/Zotero`；macOS/Linux 试 `~/Zotero`）。
  查当前解析结果：`python scripts/zotero_paths.py`
- 所有命令都可用 `--db <zotero.sqlite 路径>` 单独覆盖。
- **userID 自动从库中 `users` 表读取**，无需手工配置。

脚本以 `?immutable=1` 只读方式访问 `zotero.sqlite`，不需要关闭 Zotero，也不需要开启本地 API。

## 用法

```bash
PY=python
SC="scripts/zotero_cite.py"

# 1) 查看文档里已有的 Zotero 域（域类型 / citationID / 条目 key / 显示文本）
"$PY" "$SC" inspect --docx "稿件.docx"

# 2) 插入一条引用域（在第 0 个段落之后新增一段）
"$PY" "$SC" cite --docx "稿件.docx" --out "稿件_cited.docx" \
        --keys 22J52PUE --after-para 0 --label "[2]"

# 3) 插入参考文献表域（一篇文档只能有一个文献表域）
"$PY" "$SC" biblio --docx "稿件.docx" --out "稿件_biblio.docx"

# 4) 查看文档当前用的引文样式
"$PY" "$SC" style --docx "稿件.docx"

# 5) 整篇换引文样式
"$PY" "$SC" setstyle --docx "稿件.docx" --out "稿件_gb.docx" --style gb7714

# 6) 整段重写 + 内联引用域（推荐方式：一次把整篇正文写成含引用域的段落）
"$PY" "$SC" rewrite --docx "初稿.docx" --out "V2.docx" --spec spec.json

# 7) run 级定点替换（文档已被 Word 刷新过时用这个，不要用 rewrite）
"$PY" "$SC" replace --docx "V2.docx" --out "V3.docx" \
        --edit '{"index":59,"old":"they suppress the bacterial challenge, temper","new":"they temper"}'
```

多个条目用逗号分隔：`--keys KEY1,KEY2,KEY3`。
`--after-para N` 省略时追加到文档末尾。`--no-superscript` 用于非上标样式（如 APA）。

## `replace`：已刷新文档的定点修订（**优先于 rewrite**）

用户已在 Word 里 Refresh 过、或文档本就由 Zotero 插件维护时，**不要用 `rewrite` 整段重建**——
那会丢弃域结果缓存、用户接受的修订与 Word 侧格式。改用 `replace`：

```bash
"$PY" "$SC" replace --docx in.docx --out out.docx \
  --edit '{"index":8,"old":"lower both the tissue burden by ⟦NUM-6⟧ log₁₀ CFU per gram and ","new":"support "}'
```

- `--edit` 可重复；`index` 与 `rewrite` 同为顶层段落序号（0-based）。
- `old` 必须在该段**可见文本中唯一出现**，否则报错退出（不做模糊匹配）。同一标记在同段出现多次时，old 带前缀上下文消歧。
- 命中范围**不得跨越引用域**（`fldChar`/`instrText`）——跨越时脚本报错并提示改用 `rewrite`。
  域显示文本（如 `[1]`）不参与定位，避免把域切断。
- 替换后为空的 run（含红字黄底占位符）会被整体删除，段落 `pPr` 与其余 run 的 `rPr` 原样保留。
- **`new` 支持标记渲染**：`{{cite:KEY1,KEY2|[标签]}}` 渲染为真正的 Zotero 引用域，
  `{{it:}}` / `{{bd:}}` / `{{ph:}}` 同 `rewrite`。即在已刷新稿中也能**句级插入新引用域**
  （key 必须在库中，提前全量校验，缺 key 直接报错不产文件）。
- **`--cite-inline`**（replace/rewrite 通用）：`{{cite:}}` 显示 run 不加
  `vertAlign superscript`，用于 Elsevier/CEJ 等**行内方括号非上标**样式。缺省仍为上标。
  若文档既有域为非上标（检查：正文引用是否为行内 `[n]`），转换时务必加此参数，否则刷新前显示与全文不一致。
- **⚠️ 本脚本写入的引用域缺 `itemData` 快照**：`{{cite:}}`/`cite` 产的域只有 `id+uris`。
  部分环境（Word 端 Zotero 插件）**Refresh 会跳过这类域**——文献表不收录其引用文献、编号不重算，
  全文新旧编号混杂。凡用本脚本插域后，**建议**再跑配套补全脚本（另存新文件，永不覆盖原稿）：

  ```bash
  "$PY" scripts/backfill_itemdata.py --docx in.docx --out out.docx
  ```

  其逻辑：扫描全文缺 itemData 的域 → 从文内其他域快照收割 → 缺失走 Zotero 本地 API（客户端须运行）
  构建 CSL → 统一 properties（formattedCitation/plainCitation/noteIndex）与 schema，与插件自建域一致。
  跑完后回 Word Refresh 即可正常并入编号与文献表。
  **域解析必须跨 run 聚合**：Word/Zotero 插件刷新后会把长域指令拆成多个
  `w:r/w:instrText` run，逐 run 解析（或带 `'ZOTERO_ITEM' in raw` 守卫）会漏掉不含关键字的
  分片，导致 key/DOI 收集不完整、误判"文献无引用"。`backfill_itemdata.py` 已改用
  `fldChar begin/separate/end` 状态机聚合（`iter_fields`），自写脚本解析域时照此办理。

- **整段插入形式**：`{"index":59,"insert_after":"带标记的整段文本","clone_from":3}` ——
  在段 59 后插入新段落，段落属性与字体克隆自 `clone_from`（缺省 = `index`）。
  标题段克隆标题段（如 `"clone_from":3` 克隆「1. Introduction」的加粗格式），正文段克隆正文段。
  同一 `index` 多次 `insert_after` 时按登记顺序依次排布。适合补 Conclusion 等新章节。
- **批量模式**：`--edits-file edits.json`（对象数组，与 `--edit` 合并）。条目多或含特殊字符时
  务必用文件，不要用命令行拼接。生成 edits.json 时 old 锚点应从目标 docx 实时读取校验唯一性。
- **用法要点**：**另存新版本**（如 V2→V3）而不是原地覆盖——若用户此刻正用 Word 打开该文件
  （目录下有 `~$xxx.docx` 锁文件，或 `docProps/app.xml` 里 `Application=Microsoft Office Word`），
  原地写会被 Word 保存时覆盖。

**回归验证**：对同一份已刷新稿，`replace` 的定点替换输出与手写 lxml 脚本的输出
逐部件 md5 一致，引用域与文献表域未受影响。

## `rewrite`：内联引用域（整段渲染）

`cite` 只能在段落之间**新增**一个引用段落，无法把正文里的 `[1]` 就地换成域。实际写稿用 `rewrite`：
它按 spec 逐段重建段落内容，把标记展开成 run / 引用域，并可选写入文档引文样式与文献表域。

spec.json：

```json
{
  "style": "afm",
  "bibliography": {"heading": "References"},
  "paragraphs": [
    {"index": 4, "text": "…act on the envelope.{{cite:KEY1|[1]}} The species {{it:Escherichia coli}} …{{ph:NUM-1}}…"}
  ]
}
```

| 标记 | 展开为 |
|---|---|
| `{{cite:KEY1,KEY2\|[2–4]}}` | 引用域（`label` 省略则按全文出现序自动编号 `[n]`）+ 上标显示文本；`--cite-inline` 时为行内显示 |
| `{{it:文本}}` | 斜体 run（物种名、*ortho* 等） |
| `{{bd:文本}}` | 加粗 run |
| `{{ph:文本}}` | 占位符 run：红字 `FF0000` + 黄底，自动加 `⟦⟧` |
| `{{color:RRGGBB\|文本}}` | 指定颜色 run |

- `index` 是**正文顶层段落序号（0-based）**。段落的 `pPr`（缩进、样式）原样保留，
  `rPr` 取自该段原有 run（无则 Times New Roman）。
- `"style"` 可选：写入 `docProps/custom.xml` 的 `ZOTERO_PREF_n`（同 `setstyle` 的效果）。
  **文档此前从未被 Zotero 插件处理过时也适用**——脚本会自动新建 `docProps/custom.xml` 部件并补
  `[Content_Types].xml` Override 与 `_rels/.rels` 关系。这是 `setstyle` 做不到的（`setstyle` 要求 prefs 已存在）。
- `"bibliography"` 可选：在正文末尾（`sectPr` 之前）追加「标题 + 文献表域」。全稿装配时记得把文献表域移到 Conclusion 之后。
- 引用显示文本默认写上标 `w:vertAlign`；非上标样式用 `--cite-inline`。

**校验**：改完用 `inspect` 复核——它会把一段里拼接的多个域拆开逐条列出条目 key。

## 重复条目的域级去重（`dedup_fields.py`）

同一文献在库中有两个条目（同 DOI、不同 key）时，正文不同位置会分别引用两者，
刷新后文献表出现重复条目、占两个编号。`scripts/dedup_fields.py` 把引用 dup key
的 citationItem 整个替换为 keep key 的 item 对象（从文内已有域收割完整 id/uris/itemData），
刷新后重复条目自动消失、编号重排，**无需在 Zotero 库里合并条目**。

```bash
"$PY" scripts/dedup_fields.py --docx 稿件.docx --dry-run          # 先诊断：分组、keep/dup 判定
"$PY" scripts/dedup_fields.py --docx 稿件.docx --out 稿件_v2.docx  # 确认后替换，另存新版本
```

- 分组键：DOI（小写）优先，无 DOI 用规范化标题兜底；keep = 正文首次引用更早者（顺序编码制下即小编号）。
- 实测（一份 300+ 引用域、约 200 条文献的期刊返修稿）：19 域被触碰、19 个 item 替换，
  keys 204→198；段落可见文本/表格/媒体/custom.xml 零差异。
- 替换只动域指令 JSON，显示文本（separate/end 之间的编号缓存）不动，Refresh 后由 Zotero 重算。
- 验证方法：对输出文件回跑 `--dry-run`，dup groups 应为 0、域总数不变。

## 库中缺条目时怎么补

`uris` 走的是「库中已有条目」。若稿件引用的文献**不在本机库**里，必须先导入，否则刷新时 Zotero 报「条目未找到」。
Zotero 本地通道实测结论：

| 通道 | 结果 |
|---|---|
| `POST /api/users/0/items`、`POST/DELETE /api/users/0/collections\|items/<key>` | ❌ 本地 API 是**只读**的（400 "Endpoint does not support method" / 501 Not Implemented） |
| `POST /connector/saveItems` | ✅ 可写入，`{"sessionID","uri","items":[...],"singleFile":false}` |
| saveItems 里的 `collections` 字段 | ❌ 被忽略 |
| saveItems 里的 `target` 字段（"L1"/"C14"） | ❌ 被忽略 |
| saveItems 里的 `tags` 字段 | ✅ 生效 |
| **关闭客户端后直接写 `zotero.sqlite`**（字段级批量修正） | ✅ 可行 |

**sqlite 直写法**（适合批量字段规范化，如全大写刊名→Title Case+缩写）：
前置——Zotero 客户端必须关闭（运行中写有 WAL 冲突风险）；先用 online backup API 备份
（`sqlite3.connect('file:...?immutable=1',uri=True).backup(dst)`，运行中也是一致快照）。
字段存储结构：`itemData(itemID,fieldID,valueID)` → `itemDataValues(valueID,value)`，
**value 去重共享**——改字段值 = 查/建目标 valueID 后 `UPDATE itemData SET valueID` 指向，
不要改共享的旧 value 行；abbr 无行则 INSERT。fieldID 查 `fields` 表
（实测 publicationTitle=41、journalAbbreviation=85）。单事务提交+回滚保护，改完回读验证。
注意：直写不更新 items.dateModified，改动**不会同步**到 zotero.org（仅本地生效）。

因此**导入的条目会落到 Zotero 左侧当前选中的分类集**，且**无法用 API 指定或事后删除**。
正确做法：先让用户在 Zotero 里新建并点选目标分类集，再导入；导入后按 DOI 从
`/api/users/0/collections/<KEY>/items` 回读 item key（比按 dateAdded 分页更可靠，避免命中同 DOI 的旧条目）。

每条元数据建议先用 Crossref 抓（`https://api.crossref.org/works/<DOI>`）再转 Zotero item JSON，
字段映射：`journal-article→journalArticle`、`book-chapter→bookSection`；`container-title→publicationTitle`、
`page→pages`、`issued.date-parts→date`、`author[].family/given→creators`。

## `--style` 的取值

`--style` 接受四种写法：内置别名（`gb7714`、`gb7714-author-date`、`afm`、`acs`、`apa`、
`ieee`、`nature`）、Zotero 样式短名、完整 `http://www.zotero.org/styles/...` id、
或本地 `.csl` 文件路径（自动读其 `<id>`）。带空格与斜杠的写法（如 `GB/T 7714`）会先归一化。

例：AFM 样式（`advanced-functional-materials`）的引文是
`vertical-align="sup" prefix="[" suffix="]"`，即上标方括号编号；
GB/T 7714-2015 顺序编码制 id 为 `china-national-standard-gb-t-7714-2015-numeric`
（Zotero 样式库官方名称是完整长名，短名 `chinese-gb7714-*` 会返回 Style not found）。

## 找条目 key 的方法

库在 Zotero 数据目录下的 `zotero.sqlite`，可直接查（只读）：

```python
import sqlite3
con = sqlite3.connect("file:/path/to/Zotero/zotero.sqlite?immutable=1", uri=True)
# key -> itemID
con.execute("SELECT itemID FROM items WHERE key=? AND libraryID=1", (key,))
# 按标题检索
con.execute("""SELECT i.key, v.value FROM items i
  JOIN itemData d ON d.itemID=i.itemID
  JOIN itemDataValues v ON v.valueID=d.valueID
  JOIN fields f ON f.fieldID=d.fieldID AND f.fieldName='title'
  WHERE i.libraryID=1 AND v.value LIKE ?""", ("%关键词%",))
```

若本地 API 已开启（Zotero 设置 → 高级 → 允许本机应用通信），也可用
`GET http://127.0.0.1:23119/api/users/0/items?q=关键词` 检索，两者结果一致。

## 域格式要点（依据 Zotero 源码 `xpcom/integration.js`）

见 `references/zotero-field-spec.md`。最关键的三条：

1. 域代码前缀固定为 ` ADDIN ZOTERO_ITEM CSL_CITATION ` 和 ` ADDIN ZOTERO_BIBL `，
   `ADDIN ZOTERO_` 不可省略，否则 Word 不认作域、Zotero 也不认。
2. `citationItems[]` 里 **`uris` 是刷新时的首选定位方式**，格式
   `http://zotero.org/users/<userID>/items/<KEY>`（纯本地库用 `/users/local/<libraryID>/items/<KEY>`，
   群组库用 `/groups/<libraryID>/items/<KEY>`）。本脚本自动从库中读 userID。
   uris 解析失败且无内嵌 `itemData` 时，Zotero 会报"条目未找到"。
3. 上标靠**字符属性** `w:vertAlign val="superscript"`，不是 RTF 标记。

## 样式改写的原理（`setstyle`）

引文样式**不在域里**，而在文档的 `docProps/custom.xml`，以自定义属性
`ZOTERO_PREF_n` 存储：Word 自定义属性单串上限 255 字符，Zotero 就按 255 定长切分。
`setstyle` 把各块拼接 → 替换其中 `<style id="..."/>` → 按 255 重新分块写回。
已实测该规则可逐字节还原原文件，改写后 `word/document.xml` 及其余部件完全不变。

## 局限与注意

- 脚本写入的 `plainCitation`/显示文本只是刷新前的预览，刷新后以文档样式为准。
- 脚本默认不写内嵌 `itemData`（域更短）。部分环境实测未写 `itemData` 也能正常刷新；
  若目标机器上条目 key 不存在（跨账号共享文档），需用 `backfill_itemdata.py` 或改为携带完整 CSL `itemData` 兜底。
- 一篇文档只应有一个 `CSL_BIBLIOGRAPHY` 文献表域，重复插入会导致刷新异常。
- 写入后 `.docx` 必须在 **Word + Zotero 插件**里刷新一次才能得到正确编号与格式；
  仅凭脚本不会自动渲染。
- `setstyle` 要求目标样式的 `.csl` 已在本机样式目录（`<数据目录>/styles`）或可联网获取；
  且要求文档**已有** `ZOTERO_PREF`（即曾被 Zotero 插件处理过）。从未被处理过的文档请用
  `rewrite --spec`（带 `"style"`）代替，它会从零建立 `docProps/custom.xml`。
- `inspect` 会把同段内拼接的多个域拆开（`split_fields`）。
- 改样式后**不要**用第三方编辑器打开保存（可能重写 custom.xml）；直接交给 Word。
