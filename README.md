# gongwen-skill

**公文 Word 生成技能包(SKILL.md 标准)** — 让 Claude Code / ZCode 等 70+ 兼容 agent 一键产出版式合规的中国党政机关公文 `.docx`(GB/T 9704 风格,依真实范文逐项定稿),可选 WPS 转 PDF 与坐标级版式核验。

A SKILL.md pack for generating Chinese official documents (公文) in Word with production-verified layout — XiaoBiaoSong title, FangSong body, exact 28pt line spacing, odd-right/even-left page numbers, plus WPS-based PDF conversion and coordinate-level layout verification.

## 内置什么

| 文件 | 作用 |
|---|---|
| `SKILL.md` | 技能入口:工作流、铁律、坑清单 |
| `scripts/build_docx.js` | docx 生成器(JSON 数据 → 公文 docx,docx-js) |
| `scripts/to_pdf.sh` | WPS `wpscli word2pdf` 封装(沙盒安全路径处理) |
| `scripts/verify_layout.py` | `pdftotext -bbox` 坐标核验:标题居中/首行缩进/页码 |
| `references/layout-spec.md` | 完整版式参数表(twips 级) |
| `templates/` + `examples/` | 数据模板与真实样例 |

## 版式要点(全表见 references/layout-spec.md)

- 页边距四边 2.5cm;标题方正小标宋二号居中;正文仿宋_GB2312 三号
- 一级标题黑体、二级楷体_GB2312,首行缩进 1.13cm(不顶格)
- 行距固定值 28 磅;页码「— 1 —」宋体四号,**奇数页居右、偶数页居左**

## 安装

把本仓库放入 agent 的技能目录(ZCode:`~/.agents/skills/gongwen-doc`;Claude Code:`.claude/skills/gongwen-doc`):

```bash
git clone https://github.com/mazihua-lgtm/gongwen-skill ~/.agents/skills/gongwen-doc
```

或仅复制 `SKILL.md` + `scripts/` + `templates/`。

## 使用

对 agent 说:

> 按 gongwen-doc 技能,把下面内容生成公文 Word:……

或手动走流程:

```bash
node scripts/build_docx.js examples/sample-data.json out.docx
# docx 缺失时:NODE_PATH=$(npm root -g) node scripts/build_docx.js …
bash scripts/to_pdf.sh out.docx
python3 scripts/verify_layout.py out.pdf
```

## 设计立场

1. **内容与版式分离** — agent 只写 JSON,版式由脚本保证,杜绝每次重新发挥
2. **OCR 铁律写入技能** — 图像识别的事实未经用户文字核对不得进交付稿
3. **坐标核验优先于目检** — `pdftotext -bbox` 拿精确 pt 坐标,比视觉模型可靠

## License

MIT
