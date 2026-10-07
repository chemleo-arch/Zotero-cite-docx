# Zotero Word 引用域格式规格

来源：Zotero 9.0.6 官方发行版的源码包 `<Zotero 安装目录>/app/omni.ja`（zip）内
`chrome/content/zotero/xpcom/integration.js`（约 3872 行）与
`integration/word-for-windows/components/zoteroWinWordIntegration.mjs`。
另有真实样本：由 Zotero 插件原生生成的一个 .docx（含引用域与文献表域）。

查看方式（无需装任何第三方组件，直接读官方发行版）：

```python
import zipfile
z = zipfile.ZipFile(r"<Zotero 安装目录>/app/omni.ja")
print(z.read("chrome/content/zotero/xpcom/integration.js").decode("utf-8"))
```

## 一、域代码

Zotero 内部 `rawCode`（`Field.setCode`，integration.js:2935/2937）：

| 类型 | 代码 |
|---|---|
| 引用 | `ITEM CSL_CITATION {json}` |
| 文献表 | `BIBL {json} CSL_BIBLIOGRAPHY` |

Word 中的完整形态（多出 `ADDIN ZOTERO_` 前缀，由 Windows 端
`libzoteroWinWordIntegration.dll` / `Zotero.dotm` 处理，JS 层不含该字面量）：

```
 ADDIN ZOTERO_ITEM CSL_CITATION {"citationID":"...","properties":{...},"citationItems":[...],"schema":"..."}
 ADDIN ZOTERO_BIBL {"uncited":[],"omitted":[],"custom":[]} CSL_BIBLIOGRAPHY
```

注意前导空格；解析入口 `Field.loadExisting`（2991-3013）按 `ITEM`/`CITATION`（兼容旧格式）
和 `BIBL` 前缀分派。

## 二、引用域 JSON

```json
{
  "citationID": "uzbBIAc8",
  "properties": {
    "unsorted": false,
    "formattedCitation": "\\super [1]\\nosupersub{}",
    "plainCitation": "[1]",
    "noteIndex": 0
  },
  "citationItems": [
    {
      "id": 3195,
      "uris": ["http://zotero.org/users/<userID>/items/Z6H53G2S"],
      "itemData": { "...CSL JSON..." }
    }
  ],
  "schema": "https://github.com/citation-style-language/schema/raw/master/csl-citation.json"
}
```

- `citationID`：8 位随机串，同一文档内唯一，否则 Zotero 会重排（integration.js:2205-2217）
- `formattedCitation`：上次渲染结果；上标样式下为 RTF `\super [1]\nosupersub{}`
- `citationItems[].locator` / `.label`：页码定位（如 `label="page"`）
- `itemData`：完整 CSL JSON，可省略（见第四节解析优先级）

## 三、文献表域 JSON

```json
{"uncited": [], "omitted": [], "custom": []}
```

- `uncited`：未在正文引用但要列入文献表的条目
- `omitted`：排除的条目
- `custom`：自定义条目文本
读取逻辑见 integration.js:3529-3551、3630-3650。

## 四、刷新时的条目解析优先级（关键）

`Zotero.Integration.Citation.loadItemData`（integration.js:3306-3388）：

1. **`citationItem.uris`** → `uriMap.getZoteroItemForURIs(uris)`（2817-…）
   - 先查内嵌条目 `embeddedItemsByURI`，再 `Zotero.URI.getURIItem(uri)`（内部 try/catch）
   - 再查被合并条目关系 `replacedItemPredicate`，再查 Mendeley URI 映射
2. 无 `uris` 时：`key` + `libraryID` → `Zotero.Items.getByLibraryAndKey()`
3. 或 `itemID` / `id` → `Zotero.Items.get(id)`
4. 全部失败：
   - 有 `itemData` → 作为**匿名内嵌条目**渲染（与本地库脱钩，不再随库更新）
   - 无 `itemData` → 弹窗重选，或抛 `MissingItemException`

**结论**：`uris` 是首选且最关键；只要 uris 有效，可不写 `itemData`。

### URI 形式

- 已登录同步账号：`http://zotero.org/users/<userID>/items/<KEY>`
- 纯本地库：`http://zotero.org/users/local/<libraryID>/items/<KEY>`
- 群组库：`http://zotero.org/groups/<libraryID>/items/<KEY>`

实测（一个登录了同步账号的本地库）：`libraries = [(1,'user')]`、`users = [(<userID>)]`，
样本域 uris 为 `http://zotero.org/users/<userID>/items/Z6H53G2S` → 用 **userID 形式**，
而非 `local/<libraryID>`。查法：

```python
con = sqlite3.connect("file:/path/to/zotero.sqlite?immutable=1", uri=True)
print(con.execute("SELECT libraryID, type FROM libraries").fetchall())
print(con.execute("SELECT userID FROM users").fetchall())
```

## 五、OOXML 结构

引用域（原生样本，段内）：

```xml
<w:r><w:fldChar w:fldCharType="begin"/></w:r>
<w:r><w:instrText xml:space="preserve"> ADDIN ZOTERO_ITEM CSL_CITATION {...}</w:instrText></w:r>
<w:r><w:fldChar w:fldCharType="separate"/></w:r>
<w:r><w:rPr><w:rFonts .../><w:kern w:val="0"/><w:szCs w:val="24"/>
        <w:vertAlign w:val="superscript"/></w:rPr><w:t>[</w:t></w:r>
<w:r><w:rPr>...<w:vertAlign w:val="superscript"/></w:rPr><w:t>1]</w:t></w:r>
<w:r><w:fldChar w:fldCharType="end"/></w:r>
```

- **上标由 `w:vertAlign w:val="superscript"` 实现**，不是 RTF。
- Zotero 原生会把 `instrText` 拆成多个 run（沿用了正文各字符的 `rFonts`）；
  单 run 同样合法，Word 与 Zotero 都能识别。
- 域的 `end` 可以与 `begin` **跨段落**（样本中文献表域 begin 在段 3、end 在段 4）。

文献表域：

```xml
<w:r><w:fldChar w:fldCharType="begin"/></w:r>
<w:r><w:instrText xml:space="preserve"> ADDIN ZOTERO_BIBL {"uncited":[],"omitted":[],"custom":[]} CSL_BIBLIOGRAPHY </w:instrText></w:r>
<w:r><w:fldChar w:fldCharType="separate"/></w:r>
  ... 格式化后的参考文献 runs（含 w:i 斜体、w:b 加粗、w:tab 缩进）...
<w:r><w:fldChar w:fldCharType="end"/></w:r>
```

## 六、文档样式偏好（docPrefs）：决定引文格式

引文样式**不在域里**，存在文档的 `docProps/custom.xml` 中，作为自定义属性
`ZOTERO_PREF_n`。这解释了"为什么换样式只改一处、全篇重排"。

```xml
<Properties xmlns=".../custom-properties" xmlns:vt=".../docPropsVTypes">
  <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2" name="ZOTERO_PREF_1">
    <vt:lpwstr>&lt;data data-version="3" zotero-version="9.0.6"&gt;...&lt;pref name="auto</vt:lpwstr>
  </property>
  <property fmtid="..." pid="3" name="ZOTERO_PREF_2">
    <vt:lpwstr>maticJournalAbbreviations" value="true"/&gt;&lt;/prefs&gt;&lt;/data&gt;</vt:lpwstr>
  </property>
</Properties>
```

**分块规则**：Word 自定义属性字符串上限 **255 字符**，Zotero 按 255 定长切分，
依次放入 `ZOTERO_PREF_1`、`_2`、…（`pid` 递增，`fmtid` 固定
`{D5CDD505-2E9C-101B-9397-08002B2CF9AE}`）。已实测：把各块按序拼接后按 255 重新切分，
**逐字节还原原文件**。切分由 Windows DLL（`setDocumentData`）完成，JS 层无此逻辑。

拼接后的完整内容（样本实测）：

```xml
<data data-version="3" zotero-version="9.0.6">
  <session id="6CLzisqe"/>
  <style id="http://www.zotero.org/styles/advanced-functional-materials"
         hasBibliography="1" bibliographyStyleHasBeenSet="1"/>
  <prefs>
    <pref name="fieldType" value="Field"/>
    <pref name="automaticJournalAbbreviations" value="true"/>
  </prefs>
</data>
```

- `<style id="...">` 即当前样式；值为 Zotero 样式库中的 canonical id（与 `.csl` 的 `<id>` 一致）。
- 改样式只需替换该 `id` 再按 255 分块写回 —— 脚本 `setstyle` 子命令做的就是这件事。
- 样式 id 对应的 `.csl` 需**已存在于 Zotero 样式目录**（`<数据目录>/styles/`）或可联网获取，
  否则 Zotero 刷新时找不到样式。
- 若文档中无 `ZOTERO_PREF_*`，说明它从未被 Zotero 插件处理过；需先在 Word 里插入一条引用。

## 七、其他坑

- **Mac Word "All Caps"** 会把域内 JSON 全部大写，Zotero 靠整数 `id` 恢复并丢弃 `uris`/`itemData`
  （integration.js:3073-3108）。因此 `id`（整数 itemID）值得一并写入作兜底。
- `Zotero.HTTPIntegrationClient` / `Server.Endpoints['/integration/macWordCommand']` 方向相反
  （Zotero 作为 client 去请求外部编辑器），**不能**用来让 Zotero 输出格式化文本。
- 本地 HTTP API（23119）没有 citeproc 端点，拿不到"已格式化的引文"；格式化只能由
  Zotero 在刷新文档时完成。
- `document.xml` 声明保持 `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n`。
