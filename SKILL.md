---
name: gongwen-doc
description: 起草符合中国党政机关公文版式(GB/T 9704 风格,经实测定稿)的 Word 文档——通知/报告/函件/说明/请示等,一键产出 .docx,可选转 PDF 与版式核验。当用户要求"写公文/出正式文件/按公文格式交付 Word"时使用。
---

# 公文 Word 生成

把内容整理成 JSON 数据 → 生成版式合规的 .docx → 校验 → 可选转 PDF。

## ⚠️ 铁律:OCR 内容必须先经用户核对

从图片/PDF 识别出的**具体事实**(数字、日期、金额、人名、单位名)在用户以**文字形式确认**之前,一律不得写入交付稿。识别文本仅供起草参考,引用前必须让用户过目。历史教训:OCR 识别公文内容全错,直接进正式稿被批。

## 工作流

### 1. 整理内容为 JSON(不要把中文写进 JS)

完整字段见 `templates/document-data.json`,示例见 `examples/sample-data.json`:

```json
{
  "title": "关于××××的说明",
  "body": [
    { "type": "h1",   "text": "一、基本情况" },
    { "type": "h2",   "text": "(一)总体情况" },
    { "type": "para", "text": "正文内容……" }
  ],
  "sign": [
    { "text": "××单位" },
    { "text": "2026年9月11日" }
  ]
}
```

### 2. 生成 docx

```bash
node scripts/build_docx.js <data.json> <output.docx>
```

若 `require("docx")` 报错:优先 `npm install docx`;或用全局包 `NODE_PATH=$(npm root -g) node scripts/build_docx.js …`。

**坑**:中文弯引号(" " 《》)只能放 JSON 数据文件;JS 源码里一律 `\u` 转义。zsh + JS 双重转义是最大翻车点。

### 3. 校验

```bash
textutil -convert txt -stdout <output.docx>          # 读回内容核对文字
python3 scripts/verify_layout.py <output.pdf>        # 版式坐标核验(需先转 PDF,pdftotext)
```

### 4. 转 PDF(可选)

```bash
bash scripts/to_pdf.sh <input.docx> [output.pdf]
```

走 WPS wpscli(本机无 LibreOffice)。**坑**:WPS 沙盒不能写 `/tmp`,输出路径用工作目录再 `mv`。

## 版式定稿参数(实测对齐范文)

完整表见 `references/layout-spec.md`。要点:

| 元素 | 规格 |
|---|---|
| 页边距 | 四边均 2.5cm(1418 twips) |
| 标题 | 方正小标宋简体 二号(22pt) 居中 |
| 一级「一、」 | 黑体 三号(16pt) 首行缩进 1.13cm(641 twips) |
| 二级「(一)」 | 楷体_GB2312 三号 同缩进 |
| 正文 | 仿宋_GB2312 三号 首行缩进 两端对齐 |
| 行距 | 固定值 28 磅(560 twips, EXACT) |
| 落款 | 右对齐 仿宋 |
| 页码 | 「— 1 —」宋体四号(14pt),奇数页居右/偶数页居左 |

西文/数字用 Times New Roman(公文惯例)。

## 交付前自查

1. 文字已过用户核对(尤其 OCR 来源的事实)
2. `textutil` 读回通读一遍:错别字、标点全半角、成对引号
3. 版式核验通过(或人工开 PDF 目检:字体、缩进、页码方向)
4. 文件名规范:`关于××××的说明(20260911).docx` 类
