# zotero-cite-docx

Write **real, refreshable Zotero citation fields** into a Word `.docx` — instead of pasting static
text. After running the script, open the file in Word and hit Zotero's **Refresh**: citation numbers,
superscripts and the bibliography are re-rendered by Zotero according to the document's current style.
Switch journals or draft a Chinese manuscript and the whole document re-typesets from a single change.
It also reads the field format and the document-level style preference straight from your local Zotero
install, so no third-party plugin or connector is required.

把 Zotero 的引用以**真正的域（field）**写进 Word `.docx`，而不是静态文本。写完在 Word 里点一次
Zotero 插件的 **Refresh**，引文编号、上标与参考文献表就由 Zotero 按文档当前样式重新渲染；
换期刊、改投中文刊只需改一处，全篇重排。域格式与文档样式偏好均取自本机 Zotero 官方发行版源码，
不依赖任何第三方插件或连接器。

## Structure

- `SKILL.md` — skill instructions (WorkBuddy / Claude-style agent skill format)
- `.skill-metadata.yaml` — recommended queries for the skill UI (bilingual)
- `scripts/zotero_cite.py` — core tool: `inspect` / `cite` / `biblio` / `style` / `setstyle` / `rewrite` / `replace`
- `scripts/backfill_itemdata.py` — fill in the missing CSL `itemData` snapshot in citation fields
- `scripts/dedup_fields.py` — merge duplicate literature items at the field level (same DOI, two keys)
- `scripts/zotero_paths.py` — cross-platform locator for the Zotero data directory
- `references/zotero-field-spec.md` — the authoritative field format spec, derived from Zotero source

## Prerequisites

- Python 3.8+ and `pip install -r requirements.txt` (`lxml`, `python-docx`)
- **Zotero desktop client** installed — fields can only be refreshed by Zotero itself
- The **Zotero word-processor plugin** installed in Word
- Read access to `zotero.sqlite`. It is opened with `?immutable=1`, so **Zotero may stay running**
  and the local API does not need to be enabled.

Locate your `zotero.sqlite` (the data directory is configurable inside Zotero and is not always the
default one):

```bash
python scripts/zotero_paths.py          # prints candidates and the resolved path
export ZOTERO_DATA_DIR="/path/to/Zotero"   # or ZOTERO_SQLITE=/path/to/zotero.sqlite
```

## Usage

```bash
SC="scripts/zotero_cite.py"

python "$SC" inspect  --docx "manuscript.docx"
python "$SC" cite     --docx "manuscript.docx" --out "cited.docx" --keys 22J52PUE --after-para 0 --label "[2]"
python "$SC" biblio   --docx "manuscript.docx" --out "biblio.docx"
python "$SC" style    --docx "manuscript.docx"
python "$SC" setstyle --docx "manuscript.docx" --out "gb.docx" --style gb7714
python "$SC" rewrite  --docx "draft.docx" --out "v2.docx" --spec examples/rewrite-spec.example.json
python "$SC" replace  --docx "v2.docx" --out "v3.docx" --edits-file examples/replace-edits.example.json
```

Run `python "$SC" --help` or `python "$SC" <command> --help` for all options.

## What each command is for

| Command | Purpose |
|---|---|
| `inspect` | List every Zotero field in a document: type, citationID, cited item keys, cached display text |
| `cite` | Insert one citation field as a new paragraph |
| `biblio` | Insert the bibliography field (`CSL_BIBLIOGRAPHY`) |
| `style` | Show the document's current citation style and how the preference is chunked |
| `setstyle` | Change the whole document's citation style **without opening Word** |
| `rewrite` | Rebuild paragraphs from a JSON spec, expanding `{{cite:}}` / `{{it:}}` / `{{bd:}}` / `{{ph:}}` into real fields and runs |
| `replace` | Run-level surgical edits on an **already-refreshed** manuscript without breaking fields |

Prefer `replace` over `rewrite` once the document has been refreshed at least once — `rewrite` rebuilds
paragraphs and would discard the cached field results and your Word-side formatting.

## Inline markup (`rewrite` / `replace`)

| Markup | Expands to |
|---|---|
| `{{cite:KEY1,KEY2\|[2–4]}}` | a real citation field + superscript display text (auto-numbered when the label is omitted) |
| `{{it:text}}` | italic run |
| `{{bd:text}}` | bold run |
| `{{ph:text}}` | placeholder run, red on yellow, wrapped in `⟦⟧` |
| `{{color:RRGGBB\|text}}` | coloured run |

Use `--cite-inline` for styles whose in-text citations are inline brackets rather than superscripts
(Elsevier / CEJ and similar).

## Notes and caveats

- Everything is verified against **Zotero 9.0.6**: a script-written citation field refreshes correctly
  in Word, superscript and bibliography both update, and the field is structurally identical to one
  Zotero writes itself. Changing the style rewrites only `docProps/custom.xml` — `word/document.xml`
  and all other parts stay byte-identical.
- Fields written by `cite` / `{{cite:}}` carry `id` + `uris` but no `itemData` snapshot. Some
  environments need that snapshot; run `scripts/backfill_itemdata.py` afterwards to fill it in.
- A document must contain exactly one bibliography field.
- The `.docx` still has to be refreshed once inside **Word + Zotero plugin** — the script alone does
  not render anything.
- Never let a third-party editor re-save the file after changing the style; hand it to Word directly.
- Full format details, including the field parse priority Zotero uses on refresh, are in
  `references/zotero-field-spec.md`.

## License

MIT — see `LICENSE`.
