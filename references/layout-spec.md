# 公文版式定稿参数

来源:2026-08-27 依范文实测定稿。一切公文/报告/函件类 Word 交付默认按此执行。

## 页面

| 项 | 值 | docx twips |
|---|---|---|
| 页边距(四边) | 2.5cm | 1418 |
| 行距 | 固定值 28 磅 | 560, lineRule EXACT |
| 首行缩进 | 1.13cm | 641 |

## 字体字号

| 元素 | 字体(eastAsia) | 字号 | docx size(half-pt) | 其他 |
|---|---|---|---|---|
| 标题 | 方正小标宋简体 | 二号 22pt | 44 | 居中 |
| 一级「一、」 | 黑体 | 三号 16pt | 32 | 首行缩进 641,不顶格 |
| 二级「(一)」 | 楷体_GB2312 | 三号 16pt | 32 | 同上 |
| 正文 | 仿宋_GB2312 | 三号 16pt | 32 | 首行缩进 641,两端对齐 |
| 落款 | 仿宋_GB2312 | 三号 16pt | 32 | 右对齐 |
| 页码 | 宋体 | 四号 14pt | 28 | 「— 1 —」 |

西文与数字:Times New Roman(ascii 槽)。

## 页码

- 格式:`— N —`(一字线 + 空格 + 页码 + 空格 + 一字线)
- **奇数页居右,偶数页居左**
- 实现:`evenAndOddHeaderAndFooters: true` + default(奇)/even(偶)双 Footer
- docx SDK 的 Footer 必须用 `new Footer({children:[...]})` 包装

## 常见翻车点

1. **中文弯引号**(" " 《》):只能放 JSON 数据文件;JS 源码用 `\u` 转义。zsh 命令行内联中文引号是双重转义重灾区。
2. **NODE_PATH**:全局 docx 包需 `NODE_PATH=$(npm root -g)`;或项目内 `npm install docx`。
3. **WPS 沙盒不能写 /tmp**:`wpscli word2pdf` 输出路径用工作目录,完成后再 `mv`。
4. **本机无 LibreOffice**:转 PDF 一律走 WPS:`NO_PROXY="*" /Applications/wpsoffice.app/Contents/MacOS/wpscli word2pdf <in> -o <out>`。
5. **读回校验**:`textutil -convert txt -stdout file.docx`(系统自带);版式核验用 `pdftotext -bbox` 拿文字坐标(左边距 ≈70.9pt,正文首行 ≈102.9pt),比视觉模型可靠。
