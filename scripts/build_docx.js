/**
 * 公文 docx 生成器
 *   node build_docx.js <data.json> <output.docx>
 *
 * 数据结构见 templates/document-data.json。
 * 注意:所有中文内容走 JSON 文件,本源码内中文一律 \u 转义,
 * 避免 zsh/JS 双重转义把弯引号变成乱码。
 */
const fs = require("fs");
const {
  Document,
  Packer,
  Paragraph,
  TextRun,
  AlignmentType,
  Footer,
  PageNumber,
} = require("docx");

// ---- 版式定稿参数(2026-08 对齐范文实测)----
const PAGE_MARGIN = 1418; // 四边 2.5cm
const LINE_EXACT = 560; // 固定行距 28 磅
const FIRST_LINE_INDENT = 641; // 首行缩进 1.13cm
const SIZE_TITLE = 44; // 二号 22pt(half-points)
const SIZE_BODY = 32; // 三号 16pt
const SIZE_PAGENUM = 28; // 四号 14pt

const F_TITLE = { ascii: "Times New Roman", eastAsia: "\u65b9\u6b63\u5c0f\u6807\u5b8b\u7b80\u4f53" }; // 方正小标宋简体
const F_H1 = { ascii: "Times New Roman", eastAsia: "\u9ed1\u4f53" }; // 黑体
const F_H2 = { ascii: "Times New Roman", eastAsia: "\u6977\u4f53_GB2312" }; // 楷体_GB2312
const F_BODY = { ascii: "Times New Roman", eastAsia: "\u4eff\u5b8b_GB2312" }; // 仿宋_GB2312
const F_PAGENUM = { ascii: "Times New Roman", eastAsia: "\u5b8b\u4f53" }; // 宋体

const DASH = "\u2014"; // —

function bodyParagraph(text, font) {
  return new Paragraph({
    alignment: AlignmentType.JUSTIFIED,
    indent: { firstLine: FIRST_LINE_INDENT },
    spacing: { line: LINE_EXACT, lineRule: "exact" },
    children: [new TextRun({ text, font, size: SIZE_BODY })],
  });
}

function signParagraph(text) {
  return new Paragraph({
    alignment: AlignmentType.RIGHT,
    spacing: { line: LINE_EXACT, lineRule: "exact" },
    children: [new TextRun({ text, font: F_BODY, size: SIZE_BODY })],
  });
}

function build(data) {
  if (!data.title || !Array.isArray(data.body)) {
    throw new Error("data.json 需含 title(字符串)与 body(数组)");
  }

  const children = [
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { line: LINE_EXACT, lineRule: "exact", after: LINE_EXACT },
      children: [new TextRun({ text: data.title, font: F_TITLE, size: SIZE_TITLE })],
    }),
  ];

  for (const item of data.body) {
    if (item.type === "h1") children.push(bodyParagraph(item.text, F_H1));
    else if (item.type === "h2") children.push(bodyParagraph(item.text, F_H2));
    else children.push(bodyParagraph(item.text, F_BODY));
  }

  for (const s of data.sign || []) children.push(signParagraph(s.text));

  return new Document({
    evenAndOddHeaderAndFooters: true, // 页码奇右偶左的前提
    sections: [
      {
        properties: {
          page: {
            margin: { top: PAGE_MARGIN, right: PAGE_MARGIN, bottom: PAGE_MARGIN, left: PAGE_MARGIN },
          },
        },
        footers: {
          default: new Footer({
            children: [
              new Paragraph({
                alignment: AlignmentType.RIGHT, // 奇数页居右
                children: [
                  new TextRun({
                    children: [DASH + " ", PageNumber.CURRENT, " " + DASH],
                    font: F_PAGENUM,
                    size: SIZE_PAGENUM,
                  }),
                ],
              }),
            ],
          }),
          even: new Footer({
            children: [
              new Paragraph({
                alignment: AlignmentType.LEFT, // 偶数页居左
                children: [
                  new TextRun({
                    children: [DASH + " ", PageNumber.CURRENT, " " + DASH],
                    font: F_PAGENUM,
                    size: SIZE_PAGENUM,
                  }),
                ],
              }),
            ],
          }),
        },
        children,
      },
    ],
  });
}

const [, , dataPath, outPath] = process.argv;
if (!dataPath || !outPath) {
  console.error("用法: node build_docx.js <data.json> <output.docx>");
  process.exit(1);
}

const data = JSON.parse(fs.readFileSync(dataPath, "utf8"));
Packer.toBuffer(build(data)).then((buf) => {
  fs.writeFileSync(outPath, buf);
  console.log("OK " + outPath);
});
