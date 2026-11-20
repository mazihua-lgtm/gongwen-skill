import argparse
import json
import re
import unicodedata
from datetime import date
from decimal import Decimal
from difflib import SequenceMatcher
from pathlib import Path
from xml.etree import ElementTree
from zipfile import BadZipFile, ZipFile


NUMBER = r"(?<![\d.,])[+\-−]?(?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d+)?"
AMOUNT = re.compile(rf"({NUMBER})\s*(亿元|万元|千元|百元|元)")
RATIO = re.compile(rf"({NUMBER})\s*(个百分点|%|‰)")
COUNT = re.compile(
    rf"(?:(大约|约|接近|近|超过|逾|不足|不到|至少|至多|不超过)\s*)?"
    rf"({NUMBER})\s*(余|多)?\s*(人次|场次|家|件|名|人|条|项|对|期|场|笔|户|处|次|个(?!百分点|月))"
)
DATE = re.compile(
    r"(?<!\d)(\d{4})(?:年(\d{1,2})月(\d{1,2})日|([-/.])(\d{1,2})\4(\d{1,2})(?!\d))"
)
ARTICLE = r"第\s*[零〇一二三四五六七八九十百千两\d]+\s*[条款项](?:\s*之\s*[零〇一二三四五六七八九十百千两\d]+)?"
POLICY = re.compile(rf"《[^《》\n]+》(?:\s*{ARTICLE})*")
RESPONSIBILITY = re.compile(
    r"(?:由\s*(?P<by>[^\W_][^，。；、,;：:\s]{0,39}?)|"
    r"(?P<subject>[\u4e00-\u9fffA-Za-z0-9]{1,30}?(?:部门|支行|办公室|委员会|工作组|部|处|科|组)))"
    r"\s*(?:负责|牵头|承担|具体实施)"
)
BASIS_REVIEW = re.compile(
    r"(?:按照|参照|依据|依照|根据)\s*(?:国家|上级|本行)?\s*(?:有关|相关)\s*(?:规定|要求|政策)"
    r"|[\[【(]\s*待核实\s*[\]】)]"
)
AMOUNT_FACTORS = {"元": 1, "百元": 100, "千元": 1_000, "万元": 10_000, "亿元": 100_000_000}
CASE_STAGES = {
    "transfer_and_filing": ("移送和立案连写，需分开核对", r"移送(?:至)?公安(?:机关)?\s*(?:后)?\s*(?:已|尚未|未)?立案(?:侦查)?"),
    "transfer": ("移送", r"移送(?:至)?(?:公安(?:机关)?|检察(?:机关|院)|司法机关)|移送(?:的)?(?=\s*(?:\d+\s*件|案件|[，。；、,.!?！？;]|$))"),
    "police_filing": ("公安立案侦查表述", r"(?:由)?公安(?:机关)?\s*(?:已经|已|尚未|未|拟|决定|不予)?\s*立案(?:侦查)?|立案侦查"),
    "investigation": ("审查调查", r"立案审查调查|接受(?:纪律)?审查(?:和)?(?:监察)?调查|审查调查"),
    "filing": ("立案，主体需结合上下文核对", r"立案(?=\s*(?:\d|成功|[，。；、,.!?！？;]|$))"),
    "review_for_prosecution": ("审查起诉", r"(?:移送(?:至)?(?:检察机关|检察院)?\s*)?审查起诉"),
    "prosecution": ("提起公诉", r"提起公诉"),
    "judgment": ("法院判决或刑罚表述", r"(?:[一二]审\s*)?(?:(?:被|由)?法院\s*)?判(?:处(?:实刑|有期徒刑|无期徒刑|拘役|管制|死刑|缓刑|刑罚)?|决(?:无罪|有罪)?|刑)"),
    "discipline": ("处分表述，类型以原文为准", r"(?:给予|受到)?(?:党纪(?:政务)?|政务|纪律)处分|处分(?=\s*\d+\s*(?:人|名))"),
    "penalty": ("行政处罚", r"(?:作出|受到|给予)?行政处罚"),
    "clue": ("发现线索", r"(?:主动)?发现(?:的)?(?:问题)?线索"),
    "clue_check": ("线索核查", r"(?:问题)?线索\s*(?:正在|仍在|正|拟|尚未|未|持续|进一步)?\s*(?:核查|深挖)"),
}
CASE_STAGE_PATTERN = re.compile(
    r"(?P<qualifier>尚未被|尚未|还未|未予|未被|不予|未|正在|已由|已经|已|拟|待|计划|即将|将|成功|零)?\s*"
    + "(?:" + "|".join(f"(?P<{key}>{pattern})" for key, (_, pattern) in CASE_STAGES.items()) + ")"
)
WORD_NS = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
REVISION_TAGS = {WORD_NS + name for name in ("ins", "del", "moveFrom", "moveTo")}
MAX_DOCUMENT_XML_BYTES = 20 * 1024 * 1024


def read_draft(path):
    if path.suffix.lower() != ".docx":
        return path.read_text(encoding="utf-8-sig").splitlines(), False

    with ZipFile(path) as package, package.open("word/document.xml") as part:
        content = part.read(MAX_DOCUMENT_XML_BYTES + 1)
    if len(content) > MAX_DOCUMENT_XML_BYTES:
        raise ValueError("Word 正文 XML 超过 20 MiB，无法清点。")
    body = ElementTree.fromstring(content).find(WORD_NS + "body")
    if body is None:
        raise ValueError("未找到支持的 Word 正文。")
    if any(node.tag in REVISION_TAGS for node in body.iter()):
        raise ValueError("Word 正文含未处理的修订。请在副本中确认修订，或导出已确认的正文文本再审阅。")

    paragraphs = []
    text_tag = WORD_NS + "t"
    excluded = {WORD_NS + "txbxContent", WORD_NS + "drawing", WORD_NS + "pict"}
    whitespace = {WORD_NS + name for name in ("tab", "br", "cr")}

    def paragraph_text(node):
        if node.tag == text_tag:
            return node.text or ""
        if node.tag in whitespace:
            return " "
        if node.tag in excluded:
            return ""
        return "".join(map(paragraph_text, node))

    for paragraph in body.findall(WORD_NS + "p"):
        paragraphs.append(paragraph_text(paragraph))
    return paragraphs, True


def fold(text):
    return "".join(unicodedata.normalize("NFKC", char) for char in text)


def searchable(text):
    parts = [unicodedata.normalize("NFKC", char) for char in text]
    offsets = [index for index, part in enumerate(parts) for _ in part]
    return "".join(parts), offsets


def original_span(text, offsets, span):
    start, end = span
    return {"text": text[offsets[start]:offsets[end - 1] + 1], "column": offsets[start] + 1}


def decimal_value(text):
    return Decimal(text.replace(",", "").replace("−", "-"))


def date_value(match):
    year, month_cn, day_cn, _, month_num, day_num = match.groups()
    try:
        return date(int(year), int(month_cn or month_num), int(day_cn or day_num)).isoformat()
    except ValueError:
        return match.group()


def compact(text):
    return re.sub(r"\s+", "", text)


def case_mentions(line, text, offsets):
    mentions = []
    for match in CASE_STAGE_PATTERN.finditer(text):
        group = match.lastgroup
        start = match.start("qualifier") if match["qualifier"] else match.start(group)
        span = original_span(line, offsets, (start, match.end()))
        mentions.append((compact(match.group()), {**span, "stage": CASE_STAGES[group][0]}))
    return mentions


def extract(line, org_pattern):
    text, offsets = searchable(line)
    values = {
        "金额": [(decimal_value(match[1]) * AMOUNT_FACTORS[match[2]], match.span()) for match in AMOUNT.finditer(text)],
        "比例": [((decimal_value(match[1]), match[2]), match.span()) for match in RATIO.finditer(text)],
        "计数": [
            ((match[1], decimal_value(match[2]), match[3], match[4]), match.span())
            for match in COUNT.finditer(text)
            if not re.search(r"第\s*$", text[:match.start()])
        ],
        "日期": [(date_value(match), match.span()) for match in DATE.finditer(text)],
        "机构名称": [(match.group(), match.span()) for match in org_pattern.finditer(text)] if org_pattern else [],
        "责任主体": [
            (compact(match.group(group)), match.span(group))
            for match in RESPONSIBILITY.finditer(text)
            for group in ["by" if match["by"] else "subject"]
        ],
        "政策引用": [
            (compact(match.group()), match.span())
            for match in POLICY.finditer(text)
            if re.search(ARTICLE, match.group())
            or re.search(r"(?:法|条例|办法|规则|规定|细则|意见|指引|通知|决定|方案|规程|规范|制度|公约|标准)》$", match.group())
        ],
    }
    fields = {
        field: [(value, original_span(line, offsets, span)) for value, span in matches]
        for field, matches in values.items()
    }
    fields["案件阶段"] = case_mentions(line, text, offsets)
    return fields


def align_lines(original, revised):
    matcher = SequenceMatcher(None, list(map(fold, original)), list(map(fold, revised)), autojunk=False)
    for operation, old_start, old_end, new_start, new_end in matcher.get_opcodes():
        old_indexes = range(old_start, old_end)
        new_indexes = range(new_start, new_end)
        if operation == "equal" or (operation == "replace" and len(old_indexes) == len(new_indexes)):
            yield from zip(old_indexes, new_indexes)
        else:
            yield from ((index, None) for index in old_indexes)
            yield from ((None, index) for index in new_indexes)


def evidence(index, lines, matches):
    if index is None:
        return None
    return {"line": index + 1, "context": lines[index], "matches": [span for _, span in matches]}


def organization_pattern(orgs):
    names = sorted({fold(name.strip()) for name in orgs if name.strip()}, key=lambda name: (-len(name), name))
    return names, re.compile("|".join(map(re.escape, names))) if names else None


def basis_flags(lines):
    flags = []
    for number, line in enumerate(lines, 1):
        text, offsets = searchable(line)
        for match in BASIS_REVIEW.finditer(text):
            flags.append({"line": number, "context": line, **original_span(line, offsets, match.span())})
    return flags


def inspect(lines, orgs):
    names, org_pattern = organization_pattern(orgs)
    inventory = []
    for index, line in enumerate(lines):
        for field, matches in extract(line, org_pattern).items():
            if matches:
                inventory.append({"field": field, **evidence(index, lines, matches)})
    return {"organization_terms": names, "inventory": inventory, "basis_review": basis_flags(lines)}


def compare(original, revised, orgs):
    names, org_pattern = organization_pattern(orgs)
    changes = []
    for old_index, new_index in align_lines(original, revised):
        old_line = original[old_index] if old_index is not None else ""
        new_line = revised[new_index] if new_index is not None else ""
        if fold(old_line) == fold(new_line):
            continue
        old_fields = extract(old_line, org_pattern)
        new_fields = extract(new_line, org_pattern)
        for field, old_matches in old_fields.items():
            new_matches = new_fields[field]
            if [value for value, _ in old_matches] != [value for value, _ in new_matches]:
                changes.append({
                    "field": field,
                    "original": evidence(old_index, original, old_matches),
                    "revised": evidence(new_index, revised, new_matches),
                })

    return {"organization_terms": names, "changes": changes, "basis_review": basis_flags(revised)}


def code_span(text):
    fence = "`" * (1 + max((len(run) for run in re.findall(r"`+", text)), default=0))
    return f"{fence} {text} {fence}"


def show_location(item):
    stage = f"（{item['stage']}）" if "stage" in item else ""
    return f"第 {item['column']} 列 {code_span(item['text'])}{stage}"


def show_evidence(side):
    if side is None:
        return "无对应行"
    locations = "、".join(map(show_location, side["matches"]))
    return f"第 {side['line']} 行，{locations or '未识别到该类字段'}；原句：{code_span(side['context'])}"


def render(report):
    single = "inventory" in report
    items = report["inventory"] if single else report["changes"]
    header = "单稿字段清点" if single else "关键字段变化"
    lines = [
        "# 材料送审前" + ("单稿清单" if single else "对照清单"),
        "",
        "以下是待人工复核的线索，不是事实、政策有效性或送审结论。未提示不等于没有问题。",
        "",
    ]
    if report.get("docx_inputs"):
        lines += ["Word 输入只清点正文直属段落，行号对应段落序号（含空段），不是页码或版面行号。", ""]
    lines += [f"## {header}（{len(items)} 组）", ""]
    if not report["organization_terms"]:
        lines += ["未提供机构名称词表，跳过机构名称专项检查。", ""]
    if single:
        lines += ["只有一份稿件，以下是待核对的原文摘录，不是改动清单。", ""]
    for item in items:
        field = item["field"]
        if single:
            lines.append(f"- **{field}**：{show_evidence(item)}")
            continue
        suffix = {
            "政策引用": "：［待核实］查原文、条款与适用性",
            "案件阶段": "：［待核实］查主体、办理阶段及依据",
        }.get(field, "：请确认")
        lines += [
            f"- **{field}**{suffix}",
            f"  - 原稿：{show_evidence(item['original'])}",
            f"  - 改稿：{show_evidence(item['revised'])}",
        ]
    if not items:
        lines.append("规则未提取到字段。" if single else "规则未识别到变化。")
    if any(item["field"] == "案件阶段" for item in items):
        lines += ["", "案件阶段仅按原文用语提示。移送不等于公安立案，立案、处分、处罚、起诉和判决也不能互相替代。"]
    lines += ["", f"## 依据或待核实事项（{len(report['basis_review'])} 处）", ""]
    for item in report["basis_review"]:
        lines.append(
            f"- ［待核实］{'稿件' if single else '改稿'}第 {item['line']} 行、第 {item['column']} 列 "
            f"{code_span(item['text'])}；原句：{code_span(item['context'])}"
        )
    if not report["basis_review"]:
        lines.append("未检出已标记事项或规则覆盖的笼统依据用语。")
    return "\n".join(lines)


def main():
    parser = argparse.ArgumentParser(description="在本地只读清点一份或对照两份 Word 正文或 UTF-8 文本，不调用网络。")
    parser.add_argument("original", type=Path, help="单稿或原稿（.docx 或 UTF-8 文本）")
    parser.add_argument("revised", type=Path, nargs="?", help="改稿（.docx 或 UTF-8 文本）；不指定则只清点单稿")
    parser.add_argument("--org", action="append", default=[], help="机构全称或简称，可多次指定")
    parser.add_argument("--json", action="store_true", help="输出 JSON 而非 Markdown")
    options = parser.parse_args()
    try:
        original, original_is_docx = read_draft(options.original)
        revised, revised_is_docx = read_draft(options.revised) if options.revised else (None, False)
    except (OSError, ValueError, BadZipFile, KeyError, ElementTree.ParseError) as error:
        parser.exit(2, f"无法读取稿件：{error}\n")
    if not all(any(line.strip() for line in draft) for draft in (original, revised) if draft is not None):
        parser.exit(2, "所提供的每份稿件都必须包含非空文本。\n")
    report = inspect(original, options.org) if revised is None else compare(original, revised, options.org)
    docx_inputs = [label for label, is_docx in (("original", original_is_docx), ("revised", revised_is_docx)) if is_docx]
    if docx_inputs:
        report["docx_inputs"] = docx_inputs
    print(json.dumps(report, ensure_ascii=False, indent=2) if options.json else render(report))


if __name__ == "__main__":
    main()
