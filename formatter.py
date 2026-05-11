from __future__ import annotations

import re
import shutil
import subprocess
import tempfile
import zipfile
from dataclasses import dataclass
from pathlib import Path
from xml.etree import ElementTree as ET


NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}

SUPPORTED_INPUT_EXTENSIONS = {".docx", ".doc", ".rtf", ".txt", ".html", ".htm", ".odt"}
CONVERTIBLE_INPUT_EXTENSIONS = SUPPORTED_INPUT_EXTENSIONS - {".docx"}

for prefix, uri in NS.items():
    ET.register_namespace(prefix, uri)


def qn(name: str) -> str:
    prefix, tag = name.split(":")
    return f"{{{NS[prefix]}}}{tag}"


@dataclass
class ParagraphReport:
    index: int
    text: str
    role: str


@dataclass
class ValidationWarning:
    level: str
    message: str


@dataclass
class FormatResult:
    output_path: Path
    report: list[ParagraphReport]
    warnings: list[ValidationWarning]


@dataclass(frozen=True)
class FormatOptions:
    western_font: str = "Times New Roman"
    title_font: str = "方正公文小标宋"
    body_font: str = "方正公文仿宋"
    first_heading_font: str = "方正公文黑体"
    second_heading_font: str = "方正公文楷体"
    title_size_pt: float = 22
    body_size_pt: float = 16
    line_spacing_pt: float = 29.5
    title_line_spacing_pt: float = 33
    char_spacing_pt: float = 0.5


DEFAULT_FORMAT_OPTIONS = FormatOptions()


def format_options_from_mapping(values: dict[str, str]) -> FormatOptions:
    defaults = DEFAULT_FORMAT_OPTIONS
    return FormatOptions(
        western_font=clean_option_text(values.get("westernFont"), defaults.western_font),
        title_font=clean_option_text(values.get("titleFont"), defaults.title_font),
        body_font=clean_option_text(values.get("bodyFont"), defaults.body_font),
        first_heading_font=clean_option_text(values.get("firstHeadingFont"), defaults.first_heading_font),
        second_heading_font=clean_option_text(values.get("secondHeadingFont"), defaults.second_heading_font),
        title_size_pt=option_float(values.get("titleSizePt"), defaults.title_size_pt, 12, 36),
        body_size_pt=option_float(values.get("bodySizePt"), defaults.body_size_pt, 10.5, 22),
        line_spacing_pt=option_float(values.get("lineSpacingPt"), defaults.line_spacing_pt, 18, 40),
        title_line_spacing_pt=option_float(values.get("titleLineSpacingPt"), defaults.title_line_spacing_pt, 18, 44),
        char_spacing_pt=option_float(values.get("charSpacingPt"), defaults.char_spacing_pt, 0, 3),
    )


def clean_option_text(value: str | None, default: str) -> str:
    value = (value or "").strip()
    return value or default


def option_float(value: str | None, default: float, minimum: float, maximum: float) -> float:
    try:
        number = float(value) if value not in {None, ""} else default
    except ValueError:
        return default
    return min(max(number, minimum), maximum)


@dataclass
class NumberingLevel:
    start: int
    num_fmt: str
    lvl_text: str


@dataclass
class NumberingContext:
    num_levels: dict[tuple[str, int], NumberingLevel]
    counters: dict[tuple[str, int], int]


def format_docx(
    input_path: Path,
    original_name: str,
    output_dir: Path,
    options: FormatOptions | None = None,
) -> FormatResult:
    options = options or DEFAULT_FORMAT_OPTIONS
    input_extension = input_path.suffix.lower()
    if input_extension not in SUPPORTED_INPUT_EXTENSIONS:
        supported = "、".join(sorted(SUPPORTED_INPUT_EXTENSIONS))
        raise ValueError(f"暂不支持 {input_extension or '无扩展名'} 文件，请上传这些格式：{supported}")

    output_dir.mkdir(parents=True, exist_ok=True)
    base = Path(original_name).stem or "formatted"
    output_path = output_dir / f"{base}_公文格式.docx"

    with tempfile.TemporaryDirectory() as tmp:
        tmp_path = Path(tmp)
        source_docx = prepare_docx_input(input_path, tmp_path)
        with zipfile.ZipFile(source_docx, "r") as zin:
            zin.extractall(tmp_path)

        document_path = tmp_path / "word" / "document.xml"
        numbering_path = tmp_path / "word" / "numbering.xml"
        if not document_path.exists():
            raise ValueError("这个文件看起来不是有效的 Word 文档")

        tree = ET.parse(document_path)
        root = tree.getroot()
        body = root.find("w:body", NS)
        if body is None:
            raise ValueError("未找到 Word 正文内容")

        numbering_context = load_numbering_context(numbering_path)
        report = apply_government_document_format(body, numbering_context, options)
        warnings = validate_heading_numbering(report)
        ensure_page_number_footer(tmp_path, body, options)
        tree.write(document_path, encoding="UTF-8", xml_declaration=True)

        with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zout:
            for file_path in tmp_path.rglob("*"):
                if file_path.is_file():
                    arcname = file_path.relative_to(tmp_path).as_posix()
                    zout.write(file_path, arcname)

    return FormatResult(output_path=output_path, report=report, warnings=warnings)


def prepare_docx_input(input_path: Path, tmp_path: Path) -> Path:
    if input_path.suffix.lower() == ".docx":
        return input_path

    converted_path = tmp_path / "converted_input.docx"
    if shutil.which("textutil"):
        return convert_with_textutil(input_path, converted_path)
    if shutil.which("soffice") or shutil.which("libreoffice"):
        return convert_with_libreoffice(input_path, tmp_path, converted_path)
    raise ValueError("当前系统缺少文稿转换工具，请先另存为 .docx 后再上传。")


def convert_with_textutil(input_path: Path, converted_path: Path) -> Path:
    try:
        subprocess.run(
            [
                "textutil",
                "-convert",
                "docx",
                "-output",
                str(converted_path),
                str(input_path),
            ],
            check=True,
            capture_output=True,
            text=True,
            timeout=30,
        )
    except FileNotFoundError as exc:
        raise ValueError("当前系统缺少文稿转换工具，无法处理非 .docx 文件。") from exc
    except subprocess.TimeoutExpired as exc:
        raise ValueError("文稿转换超时，请先另存为 .docx 后再上传。") from exc
    except subprocess.CalledProcessError as exc:
        detail = (exc.stderr or exc.stdout or "").strip()
        message = "文稿转换失败，请先另存为 .docx 后再上传。"
        if detail:
            message += f"转换工具提示：{detail}"
        raise ValueError(message) from exc

    if not converted_path.exists():
        raise ValueError("文稿转换没有生成有效文件，请先另存为 .docx 后再上传。")
    return converted_path


def convert_with_libreoffice(input_path: Path, tmp_path: Path, converted_path: Path) -> Path:
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if soffice is None:
        raise ValueError("当前系统缺少 LibreOffice 转换工具，请先另存为 .docx 后再上传。")

    try:
        subprocess.run(
            [
                soffice,
                "--headless",
                "--convert-to",
                "docx",
                "--outdir",
                str(tmp_path),
                str(input_path),
            ],
            check=True,
            capture_output=True,
            text=True,
            timeout=60,
        )
    except subprocess.TimeoutExpired as exc:
        raise ValueError("文稿转换超时，请先另存为 .docx 后再上传。") from exc
    except subprocess.CalledProcessError as exc:
        detail = (exc.stderr or exc.stdout or "").strip()
        message = "文稿转换失败，请先另存为 .docx 后再上传。"
        if detail:
            message += f"转换工具提示：{detail}"
        raise ValueError(message) from exc

    generated_path = tmp_path / f"{input_path.stem}.docx"
    if generated_path.exists():
        generated_path.replace(converted_path)
    if not converted_path.exists():
        raise ValueError("文稿转换没有生成有效文件，请先另存为 .docx 后再上传。")
    return converted_path


def apply_government_document_format(
    body: ET.Element,
    numbering_context: NumberingContext,
    options: FormatOptions,
) -> list[ParagraphReport]:
    paragraphs = body.findall("w:p", NS)
    paragraph_items: list[tuple[int, ET.Element, str]] = []
    report: list[ParagraphReport] = []

    set_page_setup(body)

    for idx, paragraph in enumerate(paragraphs, start=1):
        text = normalize_paragraph_text(paragraph, numbering_context)
        if is_artifact_paragraph(text):
            body.remove(paragraph)
            continue
        if not text:
            apply_blank_paragraph_style(paragraph, options)
            continue
        paragraph_items.append((idx, paragraph, text))

    remove_trailing_blank_paragraphs(body)

    leading_attachment_indexes = detect_leading_attachment_indexes(paragraph_items)
    title_indexes = detect_title_indexes(paragraph_items, leading_attachment_indexes)
    byline_indexes = detect_byline_indexes(paragraph_items, title_indexes)
    meeting_info_indexes = detect_meeting_info_indexes(paragraph_items)
    signer_indexes = detect_signer_indexes(paragraph_items)
    addressee_indexes = detect_addressee_indexes(paragraph_items)
    numeric_heading_indexes = detect_numeric_heading_indexes(paragraph_items)
    implicit_first_level_numbers = detect_implicit_first_level_heading_numbers(
        paragraph_items,
        title_indexes | byline_indexes | meeting_info_indexes | signer_indexes | addressee_indexes,
    )

    for position, (idx, paragraph, text) in enumerate(paragraph_items, start=1):
        if idx in leading_attachment_indexes:
            role = "附件首页标识"
        elif idx in title_indexes:
            role = "主标题"
        elif idx in byline_indexes:
            role = "发言署名"
        elif idx in meeting_info_indexes:
            role = "会议材料信息"
        elif idx in signer_indexes:
            role = "落款"
        elif idx in addressee_indexes:
            role = "主送单位"
        elif idx in implicit_first_level_numbers:
            role = "一级标题"
            text = apply_missing_first_level_number(
                paragraph,
                implicit_first_level_numbers[idx],
                options,
            )
        elif idx in numeric_heading_indexes:
            role = "三级标题"
        else:
            role = classify_paragraph(text, position)
        apply_paragraph_style(paragraph, role, options)
        apply_run_style(paragraph, role, options)
        final_text = paragraph_text(paragraph).strip() or text
        report.append(ParagraphReport(index=idx, text=final_text[:60], role=role))

    if paragraph_items:
        title_paragraphs = [paragraph for idx, paragraph, _text in paragraph_items if idx in title_indexes]
        info_paragraphs = [
            paragraph
            for idx, paragraph, _text in paragraph_items
            if idx in byline_indexes or idx in meeting_info_indexes
        ]
        standardize_title_blank_lines(body, title_paragraphs, info_paragraphs, options)
        standardize_info_bottom_blank_line(body, info_paragraphs, options)
        addressee_paragraphs = [
            paragraph for idx, paragraph, _text in paragraph_items if idx in addressee_indexes
        ]
        remove_blank_lines_after_addressee(body, addressee_paragraphs)
        heading_paragraphs = [
            paragraph
            for item, (_idx, paragraph, _text) in zip(report, paragraph_items)
            if item.role in {"一级标题", "二级标题", "三级标题", "四级标题"}
        ]
        remove_blank_lines_after_headings(body, heading_paragraphs)

    return report


def paragraph_text(paragraph: ET.Element) -> str:
    return "".join(t.text or "" for t in paragraph.findall(".//w:t", NS))


def paragraph_display_text(paragraph: ET.Element, numbering_context: NumberingContext) -> str:
    prefix = resolve_paragraph_number_prefix(paragraph, numbering_context)
    text = paragraph_text(paragraph)
    if prefix and text and not text.startswith(prefix):
        return f"{prefix}{text}"
    return text


def validate_heading_numbering(report: list[ParagraphReport]) -> list[ValidationWarning]:
    warnings: list[ValidationWarning] = []
    first_level_expected = 1
    second_level_expected = 1
    second_level_active = False

    for item in report:
        if item.role != "一级标题" and item.role != "二级标题":
            continue
        first_number = extract_first_level_number(item.text)
        if first_number is not None:
            if first_number != first_level_expected:
                warnings.append(
                    ValidationWarning(
                        level="一级标题",
                        message=(
                            f"第 {item.index} 段一级标题编号为“{number_to_chinese(first_number)}、”，"
                            f"按顺序应为“{number_to_chinese(first_level_expected)}、”。"
                        ),
                    )
                )
            first_level_expected = first_number + 1
            second_level_expected = 1
            second_level_active = False
            continue

        if item.role != "二级标题":
            continue
        second_number = extract_second_level_number(item.text)
        if second_number is not None:
            if second_level_active and second_number != second_level_expected:
                warnings.append(
                    ValidationWarning(
                        level="二级标题",
                        message=(
                            f"第 {item.index} 段二级标题编号为“（{number_to_chinese(second_number)}）”，"
                            f"按顺序应为“（{number_to_chinese(second_level_expected)}）”。"
                        ),
                    )
                )
            elif not second_level_active and second_number != 1:
                warnings.append(
                    ValidationWarning(
                        level="二级标题",
                        message=(
                            f"第 {item.index} 段二级标题编号为“（{number_to_chinese(second_number)}）”，"
                            "当前一级标题下应从“（一）”开始。"
                        ),
                    )
                )
            second_level_expected = second_number + 1
            second_level_active = True

    return warnings


def load_numbering_context(numbering_path: Path) -> NumberingContext:
    if not numbering_path.exists():
        return NumberingContext(num_levels={}, counters={})

    tree = ET.parse(numbering_path)
    root = tree.getroot()
    abstract_levels: dict[tuple[str, int], NumberingLevel] = {}
    num_to_abstract: dict[str, str] = {}

    for abstract_num in root.findall("w:abstractNum", NS):
        abstract_id = abstract_num.get(qn("w:abstractNumId"))
        if not abstract_id:
            continue
        for lvl in abstract_num.findall("w:lvl", NS):
            ilvl = int(lvl.get(qn("w:ilvl"), "0"))
            start = lvl.find("w:start", NS)
            num_fmt = lvl.find("w:numFmt", NS)
            lvl_text = lvl.find("w:lvlText", NS)
            abstract_levels[(abstract_id, ilvl)] = NumberingLevel(
                start=int(start.get(qn("w:val"), "1")) if start is not None else 1,
                num_fmt=num_fmt.get(qn("w:val"), "decimal") if num_fmt is not None else "decimal",
                lvl_text=lvl_text.get(qn("w:val"), f"%{ilvl + 1}") if lvl_text is not None else f"%{ilvl + 1}",
            )

    for num in root.findall("w:num", NS):
        num_id = num.get(qn("w:numId"))
        abstract_ref = num.find("w:abstractNumId", NS)
        if num_id and abstract_ref is not None:
            num_to_abstract[num_id] = abstract_ref.get(qn("w:val"), "")

    num_levels = {
        (num_id, ilvl): level
        for num_id, abstract_id in num_to_abstract.items()
        for (candidate_abstract_id, ilvl), level in abstract_levels.items()
        if candidate_abstract_id == abstract_id
    }
    return NumberingContext(num_levels=num_levels, counters={})


def resolve_paragraph_number_prefix(paragraph: ET.Element, numbering_context: NumberingContext) -> str:
    ppr = paragraph.find("w:pPr", NS)
    if ppr is None:
        return ""
    num_pr = ppr.find("w:numPr", NS)
    if num_pr is None:
        return ""

    ilvl_node = num_pr.find("w:ilvl", NS)
    num_id_node = num_pr.find("w:numId", NS)
    if num_id_node is None:
        return ""

    num_id = num_id_node.get(qn("w:val"), "")
    ilvl = int(ilvl_node.get(qn("w:val"), "0")) if ilvl_node is not None else 0
    level = numbering_context.num_levels.get((num_id, ilvl))
    if level is None:
        return ""

    current = next_number_value(numbering_context, num_id, ilvl, level.start)
    result = level.lvl_text
    for depth in range(ilvl + 1):
        child_level = numbering_context.num_levels.get((num_id, depth))
        if child_level is None:
            continue
        value = numbering_context.counters.get((num_id, depth), child_level.start - 1)
        if value < child_level.start:
            value = child_level.start
        result = result.replace(f"%{depth + 1}", format_number_for_level(value, child_level.num_fmt))
    return result


def next_number_value(numbering_context: NumberingContext, num_id: str, ilvl: int, start: int) -> int:
    key = (num_id, ilvl)
    current = numbering_context.counters.get(key, start - 1) + 1
    numbering_context.counters[key] = current

    for candidate in list(numbering_context.counters):
        candidate_num_id, candidate_ilvl = candidate
        if candidate_num_id == num_id and candidate_ilvl > ilvl:
            del numbering_context.counters[candidate]
    return current


def format_number_for_level(value: int, num_fmt: str) -> str:
    if num_fmt in {"chineseCounting", "chineseCountingThousand"}:
        return number_to_chinese(value)
    if num_fmt == "decimal":
        return str(value)
    return str(value)


def extract_first_level_number(text: str) -> int | None:
    match = re.match(r"^([一二三四五六七八九十]+)、", text)
    if match:
        return chinese_number_to_int(match.group(1))
    arabic_match = re.match(r"^(\d+)[、.．]", text)
    if arabic_match:
        return int(arabic_match.group(1))
    return None


def extract_second_level_number(text: str) -> int | None:
    match = re.match(r"^（([一二三四五六七八九十]+)）", text)
    if not match:
        return None
    return chinese_number_to_int(match.group(1))


def chinese_number_to_int(text: str) -> int | None:
    digits = {
        "一": 1,
        "二": 2,
        "三": 3,
        "四": 4,
        "五": 5,
        "六": 6,
        "七": 7,
        "八": 8,
        "九": 9,
    }
    if text == "十":
        return 10
    if text.startswith("十") and len(text) == 2:
        return 10 + digits.get(text[1], 0)
    if text.endswith("十") and len(text) == 2:
        return digits.get(text[0], 0) * 10
    if "十" in text and len(text) == 3:
        return digits.get(text[0], 0) * 10 + digits.get(text[2], 0)
    return digits.get(text)


def number_to_chinese(number: int) -> str:
    digits = "零一二三四五六七八九"
    if number <= 0:
        return str(number)
    if number < 10:
        return digits[number]
    if number == 10:
        return "十"
    if number < 20:
        return "十" + digits[number % 10]
    if number % 10 == 0 and number < 100:
        return digits[number // 10] + "十"
    if number < 100:
        return digits[number // 10] + "十" + digits[number % 10]
    return str(number)


def normalize_paragraph_text(paragraph: ET.Element, numbering_context: NumberingContext) -> str:
    text = normalize_text(paragraph_display_text(paragraph, numbering_context))
    ppr = paragraph.find("w:pPr", NS)

    if ppr is not None:
        for child in list(ppr):
            ppr.remove(child)

    for child in list(paragraph):
        if child is not ppr:
            paragraph.remove(child)

    if text:
        run = ET.SubElement(paragraph, qn("w:r"))
        text_node = ET.SubElement(run, qn("w:t"))
        set_text(text_node, text)
    return text


def is_artifact_paragraph(text: str) -> bool:
    compact = re.sub(r"\s+", " ", text).strip()
    if not compact:
        return False
    upper = compact.upper()
    if "PAGE" not in upper:
        return False
    if "MERGEFORMAT" in upper:
        return True
    stripped = "".join(char for char in compact if char not in {"—", "-", " "})
    return bool(re.fullmatch(r"PAGE\\d?", stripped, re.I))


def remove_trailing_blank_paragraphs(body: ET.Element) -> None:
    children = list(body)
    sect_pr = body.find("w:sectPr", NS)
    while children:
        last = children[-1]
        if last is sect_pr:
            children = children[:-1]
            continue
        if is_empty_paragraph(last):
            body.remove(last)
            children = list(body)
            continue
        break


def normalize_text(text: str) -> str:
    text = text.replace("\u3000", " ")
    text = re.sub(r"[\t\r\n]+", " ", text)
    text = re.sub(r" {2,}", " ", text).strip()
    if is_likely_byline_text(text):
        return normalize_byline_spacing(text)
    text = re.sub(r"(?<=[\u4e00-\u9fff]) +(?=[\u4e00-\u9fff])", "", text)
    text = re.sub(r"(?<=[\u4e00-\u9fff]) +(?=[，。；：、！？）])", "", text)
    text = re.sub(r"(?<=[（(]) +(?=[\u4e00-\u9fff0-9])", "", text)
    return text


def classify_paragraph(text: str, non_empty_seen: int) -> str:
    if is_meeting_info_line(text, non_empty_seen):
        return "会议材料信息"
    if text.startswith("附件"):
        return "附件标注"
    if is_first_level_heading(text):
        return "一级标题"
    if is_arabic_first_level_heading(text):
        return "一级标题"
    if is_second_level_heading(text):
        return "二级标题"
    if is_numeric_heading(text):
        return "三级标题"
    if re.match(r"^（\d+）", text) and is_heading_only(text):
        return "四级标题"
    return "正文"


def detect_leading_attachment_indexes(paragraph_items: list[tuple[int, ET.Element, str]]) -> set[int]:
    if not paragraph_items:
        return set()
    first_idx, _paragraph, first_text = paragraph_items[0]
    return {first_idx} if first_text.startswith("附件") else set()


def detect_title_indexes(
    paragraph_items: list[tuple[int, ET.Element, str]],
    skip_indexes: set[int] | None = None,
) -> set[int]:
    if not paragraph_items:
        return set()
    skip_indexes = skip_indexes or set()

    candidates = [(idx, paragraph, text) for idx, paragraph, text in paragraph_items if idx not in skip_indexes]
    if not candidates:
        return set()

    title_indexes = {candidates[0][0]}
    early = candidates[1:4]
    for idx, _paragraph, text in early:
        if is_title_continuation(text):
            title_indexes.add(idx)
            continue
        break
    return title_indexes


def is_title_continuation(text: str) -> bool:
    if text.endswith(("：", ":")):
        return False
    if is_first_level_heading(text) or is_second_level_heading(text) or is_numeric_heading_candidate(text):
        return False
    if is_meeting_info_line(text, 2):
        return False
    if is_likely_byline_text(text):
        return False
    if text in {"交流发言", "发言材料", "讲话提纲", "讲话材料", "主持词"}:
        return True
    if len(text) <= 14 and text.endswith(("交流发言", "发言材料", "讲话")):
        return True
    if len(text) <= 24 and is_heading_like(text):
        return True
    return False


def detect_numeric_heading_indexes(paragraph_items: list[tuple[int, ET.Element, str]]) -> set[int]:
    numeric_heading_indexes: set[int] = set()
    for position, (idx, _paragraph, text) in enumerate(paragraph_items):
        if is_arabic_first_level_heading(text):
            continue
        if not is_numeric_heading_candidate(text):
            continue
        if is_numeric_heading(text):
            numeric_heading_indexes.add(idx)
            continue
        next_text = next_nonempty_text(paragraph_items, position + 1)
        if (
            next_text
            and len(text) <= 18
            and is_numeric_heading(text)
            and looks_like_body_text(next_text)
        ):
            numeric_heading_indexes.add(idx)
    return numeric_heading_indexes


def detect_implicit_first_level_heading_numbers(
    paragraph_items: list[tuple[int, ET.Element, str]],
    excluded_indexes: set[int],
) -> dict[int, int]:
    explicit_position: int | None = None
    explicit_number: int | None = None
    for position, (_idx, _paragraph, text) in enumerate(paragraph_items):
        number = extract_chinese_first_level_number(text)
        if number and number > 1:
            explicit_position = position
            explicit_number = number
            break

    if explicit_position is None or explicit_number is None:
        return {}

    candidates: list[int] = []
    for position in range(explicit_position):
        idx, _paragraph, text = paragraph_items[position]
        if idx in excluded_indexes:
            continue
        if is_implicit_first_level_heading_candidate(text):
            next_text = next_nonempty_text(paragraph_items, position + 1)
            if next_text and looks_like_body_text(next_text):
                candidates.append(idx)

    needed = explicit_number - 1
    if needed <= 0 or len(candidates) < needed:
        return {}

    selected = candidates[-needed:]
    return {idx: number for number, idx in enumerate(selected, start=1)}


def extract_chinese_first_level_number(text: str) -> int | None:
    match = re.match(r"^([一二三四五六七八九十]+)、", text)
    if not match:
        return None
    return chinese_number_to_int(match.group(1))


def is_implicit_first_level_heading_candidate(text: str) -> bool:
    if not is_heading_like(text):
        return False
    if len(text) > 18:
        return False
    if is_meeting_info_line(text, 2) or is_likely_byline_text(text) or is_addressee_line(text):
        return False
    if is_first_level_heading(text) or is_second_level_heading(text) or is_numeric_heading_candidate(text):
        return False
    return bool(re.search(r"[\u4e00-\u9fff]", text))


def next_nonempty_text(paragraph_items: list[tuple[int, ET.Element, str]], start: int) -> str | None:
    for _idx, _paragraph, text in paragraph_items[start:]:
        if text:
            return text
    return None


def is_numeric_heading_candidate(text: str) -> bool:
    return bool(re.match(r"^\d+[.．、]", text))


def is_numeric_heading(text: str) -> bool:
    match = re.match(r"^\d+[.．、](.+)$", text)
    if not match:
        return False
    title_text = match.group(1).strip()
    if any(mark in title_text for mark in "。；！？.;!?"):
        return False
    return len(text) <= 18


def looks_like_body_text(text: str) -> bool:
    return len(text) >= 30 or any(mark in text for mark in "。；！？.;!?")


def detect_byline_indexes(
    paragraph_items: list[tuple[int, ET.Element, str]],
    title_indexes: set[int],
) -> set[int]:
    if not title_indexes:
        return set()
    last_title_position = max(
        position for position, (idx, _paragraph, _text) in enumerate(paragraph_items) if idx in title_indexes
    )
    byline_indexes: set[int] = set()
    for idx, _paragraph, text in paragraph_items[last_title_position + 1 : last_title_position + 3]:
        if is_likely_byline_text(text):
            byline_indexes.add(idx)
            break
        if text.startswith(("一、", "二、", "三、", "（一）")) or len(text) > 40:
            break
    return byline_indexes


def is_likely_byline_text(text: str) -> bool:
    if not text or len(text) > 24:
        return False
    if any(mark in text for mark in "，。；：、！？,.;:!?"):
        return False
    return bool(re.search(r"(公司|集团|局|办|委|部|处|科|中心|园区|街道|部门) ?[\u4e00-\u9fff]{2,4}$", text))


def normalize_byline_spacing(text: str) -> str:
    text = re.sub(r" +", " ", text).strip()
    match = re.match(r"^(.+?(?:公司|集团|局|办|委|部|处|科|中心|园区|街道|部门)) ?([\u4e00-\u9fff]{2,4})$", text)
    if match:
        return f"{match.group(1)} {match.group(2)}"
    return text


def is_first_level_heading(text: str) -> bool:
    return bool(re.match(r"^[一二三四五六七八九十]+、", text)) and is_heading_like(text)


def is_arabic_first_level_heading(text: str) -> bool:
    return bool(re.match(r"^\d+、", text)) and is_heading_like(text)


def is_second_level_heading(text: str) -> bool:
    return bool(re.match(r"^（[一二三四五六七八九十]+）", text)) and is_heading_like(text)


def is_heading_like(text: str) -> bool:
    if len(text) > 48:
        return False
    return not any(mark in text for mark in "。；！？.;!?")


def is_meeting_info_line(text: str, non_empty_seen: int) -> bool:
    if non_empty_seen > 5:
        return False
    if re.fullmatch(r"（送审稿）", text):
        return True
    if re.fullmatch(r"[\u4e00-\u9fffA-Za-z0-9]{2,20}(局|办|委|部|处|科|中心|单位|公司)", text):
        return True
    if re.fullmatch(r"((\d{4}|[Xx]{2,4})年(\d{1,2}|[Xx]{1,2})月(\d{1,2}|[Xx]{1,2})日)", text):
        return True
    return False


def detect_meeting_info_indexes(paragraph_items: list[tuple[int, ET.Element, str]]) -> set[int]:
    meeting_indexes: set[int] = set()
    early = paragraph_items[1:5]

    for offset, (idx, _paragraph, text) in enumerate(early):
        if re.fullmatch(r"（送审稿）", text):
            meeting_indexes.add(idx)
            for next_idx, _next_paragraph, next_text in early[offset + 1 : offset + 3]:
                if is_meeting_info_line(next_text, offset + 3):
                    meeting_indexes.add(next_idx)
            break

    if not meeting_indexes and len(early) >= 2:
        first_idx, _first_paragraph, first_text = early[0]
        second_idx, _second_paragraph, second_text = early[1]
        if is_meeting_info_line(first_text, 2) and is_meeting_info_line(second_text, 3):
            meeting_indexes.update({first_idx, second_idx})

    return meeting_indexes


def is_heading_only(text: str) -> bool:
    return is_heading_like(text)


def detect_signer_indexes(paragraph_items: list[tuple[int, ET.Element, str]]) -> set[int]:
    signer_indexes: set[int] = set()
    tail = paragraph_items[-12:]
    date_pattern = re.compile(
        r"^((\d{4}|[一二三四五六七八九〇零]{4})年(\d{1,2}|[一二三四五六七八九十]{1,3})月(\d{1,2}|[一二三四五六七八九十]{1,3})日)$"
    )

    for tail_pos, (idx, _paragraph, text) in enumerate(tail):
        if not date_pattern.match(text):
            continue
        signer_indexes.add(idx)
        prev = previous_nonempty_paragraph(tail, tail_pos)
        if prev is not None:
            prev_idx, _prev_paragraph, prev_text = prev
            if is_likely_signer_name(prev_text):
                signer_indexes.add(prev_idx)

    if not signer_indexes:
        last_nonempty = previous_nonempty_paragraph(tail, len(tail))
        if last_nonempty is not None:
            last_idx, _last_paragraph, last_text = last_nonempty
            if is_likely_signer_name(last_text):
                signer_indexes.add(last_idx)
    return signer_indexes


def is_likely_signer_name(text: str) -> bool:
    if len(text) > 32:
        return False
    if any(mark in text for mark in "，。；：！？,.;:!?"):
        return False
    return bool(re.search(r"[\u4e00-\u9fff]", text))


def previous_nonempty_paragraph(
    items: list[tuple[int, ET.Element, str]],
    start_pos: int,
) -> tuple[int, ET.Element, str] | None:
    max_pos = min(start_pos - 1, len(items) - 1)
    for pos in range(max_pos, -1, -1):
        idx, paragraph, text = items[pos]
        if text.strip():
            return idx, paragraph, text
    return None


def detect_addressee_indexes(paragraph_items: list[tuple[int, ET.Element, str]]) -> set[int]:
    addressee_indexes: set[int] = set()
    candidates = paragraph_items[1:6]

    for idx, _paragraph, text in candidates:
        if not is_addressee_line(text):
            continue
        addressee_indexes.add(idx)
        break

    return addressee_indexes


def is_addressee_line(text: str) -> bool:
    if not text.endswith(("：", ":")):
        return False
    if len(text) > 24:
        return False
    if text.startswith("附件："):
        return False
    if is_first_level_heading(text) or is_second_level_heading(text) or is_numeric_heading_candidate(text):
        return False
    if is_meeting_info_line(text, 2) or is_likely_byline_text(text):
        return False
    body_marks = "，。；！？,.;!?"
    if any(mark in text[:-1] for mark in body_marks):
        return False
    return bool(re.search(r"[\u4e00-\u9fff]", text))


def get_or_add(parent: ET.Element, tag: str, before: str | None = None) -> ET.Element:
    child = parent.find(tag, NS)
    if child is not None:
        return child

    child = ET.Element(qn(tag))
    if before:
        for i, item in enumerate(list(parent)):
            if item.tag == qn(before):
                parent.insert(i, child)
                return child
    parent.append(child)
    return child


def clear_children(element: ET.Element, names: set[str]) -> None:
    qnames = {qn(name) for name in names}
    for child in list(element):
        if child.tag in qnames:
            element.remove(child)


def clear_all_children(element: ET.Element) -> None:
    for child in list(element):
        element.remove(child)


def apply_blank_paragraph_style(paragraph: ET.Element, options: FormatOptions) -> None:
    ppr = get_or_add(paragraph, "w:pPr", before="w:r")
    clear_all_children(ppr)

    spacing = ET.SubElement(ppr, qn("w:spacing"))
    spacing.set(qn("w:before"), "0")
    spacing.set(qn("w:after"), "0")
    spacing.set(qn("w:line"), points_to_twips(options.line_spacing_pt))
    spacing.set(qn("w:lineRule"), "exact")

    ind = ET.SubElement(ppr, qn("w:ind"))
    ind.set(qn("w:firstLineChars"), "0")


def standardize_title_blank_lines(
    body: ET.Element,
    title_paragraphs: list[ET.Element],
    info_paragraphs: list[ET.Element],
    options: FormatOptions,
) -> None:
    if not title_paragraphs:
        return
    children = list(body)
    first_title = title_paragraphs[0]
    last_title = title_paragraphs[-1]
    if first_title not in children or last_title not in children:
        return

    first_index = children.index(first_title)
    while first_index > 0 and is_empty_paragraph(children[first_index - 1]):
        body.remove(children[first_index - 1])
        children = list(body)
        first_index = children.index(first_title)

    for previous_title, next_title in zip(title_paragraphs, title_paragraphs[1:]):
        children = list(body)
        if previous_title not in children or next_title not in children:
            continue
        previous_index = children.index(previous_title)
        next_index = children.index(next_title)
        for element in children[previous_index + 1 : next_index]:
            if is_empty_paragraph(element):
                body.remove(element)

    next_anchor = first_existing_paragraph(body, info_paragraphs)
    if next_anchor is not None:
        remove_blank_paragraphs_between(body, last_title, next_anchor)
    else:
        children = list(body)
        last_index = children.index(last_title)
        last_index = children.index(last_title)
        while last_index + 1 < len(children) and is_empty_paragraph(children[last_index + 1]):
            body.remove(children[last_index + 1])
            children = list(body)
            last_index = children.index(last_title)

    before = make_blank_paragraph(options)
    first_index = list(body).index(first_title)
    body.insert(first_index, before)
    if next_anchor is None:
        after = make_blank_paragraph(options)
        last_index = list(body).index(last_title)
        body.insert(last_index + 1, after)


def standardize_info_bottom_blank_line(
    body: ET.Element,
    info_paragraphs: list[ET.Element],
    options: FormatOptions,
) -> None:
    last_info = last_existing_paragraph(body, info_paragraphs)
    if last_info is None:
        return

    children = list(body)
    last_index = children.index(last_info)
    while last_index + 1 < len(children) and is_empty_paragraph(children[last_index + 1]):
        body.remove(children[last_index + 1])
        children = list(body)
        last_index = children.index(last_info)

    body.insert(last_index + 1, make_blank_paragraph(options))


def remove_blank_lines_after_addressee(body: ET.Element, addressee_paragraphs: list[ET.Element]) -> None:
    last_addressee = last_existing_paragraph(body, addressee_paragraphs)
    if last_addressee is None:
        return

    children = list(body)
    last_index = children.index(last_addressee)
    while last_index + 1 < len(children) and is_empty_paragraph(children[last_index + 1]):
        body.remove(children[last_index + 1])
        children = list(body)
        last_index = children.index(last_addressee)


def remove_blank_lines_after_headings(body: ET.Element, heading_paragraphs: list[ET.Element]) -> None:
    for heading in heading_paragraphs:
        children = list(body)
        if heading not in children:
            continue
        heading_index = children.index(heading)
        while heading_index + 1 < len(children) and is_empty_paragraph(children[heading_index + 1]):
            body.remove(children[heading_index + 1])
            children = list(body)
            heading_index = children.index(heading)


def remove_blank_paragraphs_between(
    body: ET.Element,
    first_paragraph: ET.Element,
    second_paragraph: ET.Element,
) -> None:
    children = list(body)
    if first_paragraph not in children or second_paragraph not in children:
        return

    first_index = children.index(first_paragraph)
    second_index = children.index(second_paragraph)
    while second_index > first_index + 1:
        middle = children[first_index + 1]
        if is_empty_paragraph(middle):
            body.remove(middle)
            children = list(body)
            second_index = children.index(second_paragraph)
            continue
        break


def first_existing_paragraph(body: ET.Element, paragraphs: list[ET.Element]) -> ET.Element | None:
    children = list(body)
    for paragraph in paragraphs:
        if paragraph in children:
            return paragraph
    return None


def last_existing_paragraph(body: ET.Element, paragraphs: list[ET.Element]) -> ET.Element | None:
    children = list(body)
    for paragraph in reversed(paragraphs):
        if paragraph in children:
            return paragraph
    return None


def is_empty_paragraph(element: ET.Element) -> bool:
    return element.tag == qn("w:p") and not paragraph_text(element).strip()


def make_blank_paragraph(options: FormatOptions) -> ET.Element:
    paragraph = ET.Element(qn("w:p"))
    apply_blank_paragraph_style(paragraph, options)
    run = ET.SubElement(paragraph, qn("w:r"))
    style_run(run, "正文", options)
    ET.SubElement(run, qn("w:t")).text = ""
    return paragraph


def apply_paragraph_style(paragraph: ET.Element, role: str, options: FormatOptions) -> None:
    ppr = get_or_add(paragraph, "w:pPr", before="w:r")
    clear_all_children(ppr)

    spacing = ET.SubElement(ppr, qn("w:spacing"))
    spacing.set(qn("w:before"), "0")
    spacing.set(qn("w:after"), "0")
    spacing.set(qn("w:line"), points_to_twips(options.line_spacing_pt))
    spacing.set(qn("w:lineRule"), "exact")

    ind = ET.SubElement(ppr, qn("w:ind"))
    jc = ET.SubElement(ppr, qn("w:jc"))

    if role == "主标题":
        jc.set(qn("w:val"), "center")
        spacing.set(qn("w:line"), points_to_twips(options.title_line_spacing_pt))
    elif role in {"会议材料信息", "发言署名"}:
        jc.set(qn("w:val"), "center")
        spacing.set(qn("w:line"), points_to_twips(options.title_line_spacing_pt))
    elif role == "落款":
        jc.set(qn("w:val"), "right")
    elif role == "附件标注":
        jc.set(qn("w:val"), "left")
        ind.set(qn("w:leftChars"), "200")
    elif role == "附件首页标识":
        jc.set(qn("w:val"), "left")
    elif role == "主送单位":
        jc.set(qn("w:val"), "left")
    else:
        jc.set(qn("w:val"), "left")
        ind.set(qn("w:firstLineChars"), "200")


def apply_run_style(paragraph: ET.Element, role: str, options: FormatOptions) -> None:
    if role == "一级标题":
        normalize_first_level_heading_marker(paragraph, options)
    if role == "正文" and apply_leading_marker_style(paragraph, options):
        return

    for run in paragraph.findall("w:r", NS):
        style_run(run, role, options)


def apply_leading_marker_style(paragraph: ET.Element, options: FormatOptions) -> bool:
    text = paragraph_text(paragraph)
    parts = split_leading_marker(text)
    if parts is None:
        return False
    lead_text, body_text = parts

    ppr = paragraph.find("w:pPr", NS)
    for child in list(paragraph):
        if child is not ppr:
            paragraph.remove(child)

    add_text_run(paragraph, lead_text, "二级标题", options)
    if body_text:
        add_text_run(paragraph, body_text, "正文", options)
    return True


def split_leading_marker(text: str) -> tuple[str, str] | None:
    shi_sentence_match = re.match(r"^([一二三四五六七八九十]+是.+?。)(.*)$", text)
    if shi_sentence_match:
        return shi_sentence_match.group(1), shi_sentence_match.group(2)

    shi_match = re.match(r"^([一二三四五六七八九十]+是)(.+)$", text)
    if shi_match:
        return shi_match.group(1), shi_match.group(2)

    bracket_match = re.match(r"^(（[一二三四五六七八九十]+）.+?。)(.*)$", text)
    if bracket_match:
        return bracket_match.group(1), bracket_match.group(2)

    bracket_only_match = re.match(r"^(（[一二三四五六七八九十]+）)(.+)$", text)
    if bracket_only_match:
        return bracket_only_match.group(1), bracket_only_match.group(2)

    return None


def normalize_first_level_heading_marker(paragraph: ET.Element, options: FormatOptions) -> None:
    text = paragraph_text(paragraph)
    match = re.match(r"^(\d+)([、.．])(.*)$", text)
    if not match:
        return

    number = int(match.group(1))
    suffix = match.group(3)
    normalized = f"{number_to_chinese(number)}、{suffix.lstrip()}"

    ppr = paragraph.find("w:pPr", NS)
    for child in list(paragraph):
        if child is not ppr:
            paragraph.remove(child)

    add_text_run(paragraph, normalized, "一级标题", options)


def apply_missing_first_level_number(
    paragraph: ET.Element,
    number: int,
    options: FormatOptions,
) -> str:
    text = paragraph_text(paragraph).strip()
    normalized = f"{number_to_chinese(number)}、{text}"

    ppr = paragraph.find("w:pPr", NS)
    for child in list(paragraph):
        if child is not ppr:
            paragraph.remove(child)

    add_text_run(paragraph, normalized, "一级标题", options)
    return normalized


def add_text_run(paragraph: ET.Element, text: str, role: str, options: FormatOptions) -> None:
    run = ET.SubElement(paragraph, qn("w:r"))
    style_run(run, role, options)
    text_node = ET.SubElement(run, qn("w:t"))
    set_text(text_node, text)


def set_text(text_node: ET.Element, text: str) -> None:
    if " " in text or text.startswith(" ") or text.endswith(" "):
        text_node.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    text_node.text = text


def style_run(run: ET.Element, role: str, options: FormatOptions) -> None:
    rpr = get_or_add(run, "w:rPr", before="w:t")
    clear_children(rpr, {"w:rFonts", "w:sz", "w:szCs", "w:b", "w:bCs", "w:spacing"})

    fonts = ET.SubElement(rpr, qn("w:rFonts"))
    fonts.set(qn("w:ascii"), options.western_font)
    fonts.set(qn("w:hAnsi"), options.western_font)
    fonts.set(qn("w:eastAsia"), font_for(role, options))
    fonts.set(qn("w:cs"), options.western_font)

    char_spacing = ET.SubElement(rpr, qn("w:spacing"))
    char_spacing.set(qn("w:val"), points_to_twips(options.char_spacing_pt))

    if role == "三级标题":
        ET.SubElement(rpr, qn("w:b"))
        ET.SubElement(rpr, qn("w:bCs"))

    size = ET.SubElement(rpr, qn("w:sz"))
    size.set(qn("w:val"), size_for(role, options))
    size_cs = ET.SubElement(rpr, qn("w:szCs"))
    size_cs.set(qn("w:val"), size_for(role, options))


def font_for(role: str, options: FormatOptions) -> str:
    if role == "主标题":
        return options.title_font
    if role == "一级标题":
        return options.first_heading_font
    if role in {"二级标题", "会议材料信息", "发言署名"}:
        return options.second_heading_font
    return options.body_font


def size_for(role: str, options: FormatOptions) -> str:
    if role == "主标题":
        return points_to_half_points(options.title_size_pt)
    return points_to_half_points(options.body_size_pt)


def points_to_twips(points: float) -> str:
    return str(round(points * 20))


def points_to_half_points(points: float) -> str:
    return str(round(points * 2))


def set_page_setup(body: ET.Element) -> None:
    sect_pr = body.find("w:sectPr", NS)
    if sect_pr is None:
        sect_pr = ET.SubElement(body, qn("w:sectPr"))

    clear_children(sect_pr, {"w:pgSz", "w:pgMar"})

    pg_sz = ET.Element(qn("w:pgSz"))
    pg_sz.set(qn("w:w"), "11906")
    pg_sz.set(qn("w:h"), "16838")
    sect_pr.insert(0, pg_sz)

    pg_mar = ET.Element(qn("w:pgMar"))
    pg_mar.set(qn("w:top"), "2098")
    pg_mar.set(qn("w:right"), "1474")
    pg_mar.set(qn("w:bottom"), "1701")
    pg_mar.set(qn("w:left"), "1587")
    pg_mar.set(qn("w:header"), "851")
    pg_mar.set(qn("w:footer"), "1417")
    pg_mar.set(qn("w:gutter"), "0")
    sect_pr.insert(1, pg_mar)


def ensure_page_number_footer(tmp_path: Path, body: ET.Element, options: FormatOptions) -> None:
    word_dir = tmp_path / "word"
    rels_dir = word_dir / "_rels"
    rels_dir.mkdir(exist_ok=True)
    rels_path = rels_dir / "document.xml.rels"

    rels_tree = ET.parse(rels_path)
    rels_root = rels_tree.getroot()
    rel_ns = "http://schemas.openxmlformats.org/package/2006/relationships"
    ET.register_namespace("", rel_ns)

    existing_ids = [
        int(item.attrib.get("Id", "rId0").replace("rId", ""))
        for item in rels_root
        if item.attrib.get("Id", "").startswith("rId")
        and item.attrib.get("Id", "")[3:].isdigit()
    ]
    rel_id = f"rId{max(existing_ids, default=0) + 1}"

    footer_name = next_available_footer_name(word_dir)
    footer_path = word_dir / footer_name
    footer_path.write_text(build_footer_xml(options), encoding="utf-8")

    rel = ET.SubElement(rels_root, f"{{{rel_ns}}}Relationship")
    rel.set("Id", rel_id)
    rel.set("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer")
    rel.set("Target", footer_name)
    rels_tree.write(rels_path, encoding="UTF-8", xml_declaration=True)

    sect_pr = body.find("w:sectPr", NS)
    if sect_pr is None:
        sect_pr = ET.SubElement(body, qn("w:sectPr"))

    for child in list(sect_pr):
        if child.tag == qn("w:footerReference"):
            sect_pr.remove(child)

    footer_ref = ET.Element(qn("w:footerReference"))
    footer_ref.set(qn("w:type"), "default")
    footer_ref.set(qn("r:id"), rel_id)
    sect_pr.insert(0, footer_ref)

    update_content_types(tmp_path, footer_name)


def next_available_footer_name(word_dir: Path) -> str:
    index = 1
    while (word_dir / f"footer{index}.xml").exists():
        index += 1
    return f"footer{index}.xml"


def build_footer_xml(options: FormatOptions) -> str:
    western_font = escape_xml_attr(options.western_font)
    line_spacing = points_to_twips(options.line_spacing_pt)
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:xml="http://www.w3.org/XML/1998/namespace">
  <w:p>
    <w:pPr>
      <w:jc w:val="center"/>
      <w:spacing w:before="0" w:after="0" w:line="{line_spacing}" w:lineRule="exact"/>
    </w:pPr>
    <w:r>
      <w:rPr>
        <w:rFonts w:ascii="{western_font}" w:hAnsi="{western_font}" w:eastAsia="{western_font}"/>
        <w:sz w:val="28"/>
      </w:rPr>
      <w:t xml:space="preserve">— </w:t>
    </w:r>
    <w:r>
      <w:rPr>
        <w:rFonts w:ascii="{western_font}" w:hAnsi="{western_font}" w:eastAsia="{western_font}"/>
        <w:sz w:val="28"/>
      </w:rPr>
      <w:fldChar w:fldCharType="begin"/>
    </w:r>
    <w:r>
      <w:rPr>
        <w:rFonts w:ascii="{western_font}" w:hAnsi="{western_font}" w:eastAsia="{western_font}"/>
        <w:sz w:val="28"/>
      </w:rPr>
      <w:instrText xml:space="preserve">PAGE</w:instrText>
    </w:r>
    <w:r>
      <w:rPr>
        <w:rFonts w:ascii="{western_font}" w:hAnsi="{western_font}" w:eastAsia="{western_font}"/>
        <w:sz w:val="28"/>
      </w:rPr>
      <w:fldChar w:fldCharType="separate"/>
    </w:r>
    <w:r>
      <w:rPr>
        <w:rFonts w:ascii="{western_font}" w:hAnsi="{western_font}" w:eastAsia="{western_font}"/>
        <w:sz w:val="28"/>
      </w:rPr>
      <w:t>1</w:t>
    </w:r>
    <w:r>
      <w:rPr>
        <w:rFonts w:ascii="{western_font}" w:hAnsi="{western_font}" w:eastAsia="{western_font}"/>
        <w:sz w:val="28"/>
      </w:rPr>
      <w:fldChar w:fldCharType="end"/>
    </w:r>
    <w:r>
      <w:rPr>
        <w:rFonts w:ascii="{western_font}" w:hAnsi="{western_font}" w:eastAsia="{western_font}"/>
        <w:sz w:val="28"/>
      </w:rPr>
      <w:t xml:space="preserve"> —</w:t>
    </w:r>
  </w:p>
</w:ftr>
"""


def escape_xml_attr(value: str) -> str:
    return (
        value.replace("&", "&amp;")
        .replace('"', "&quot;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
    )


def update_content_types(tmp_path: Path, footer_name: str) -> None:
    content_types_path = tmp_path / "[Content_Types].xml"
    tree = ET.parse(content_types_path)
    root = tree.getroot()
    ct_ns = "http://schemas.openxmlformats.org/package/2006/content-types"
    ET.register_namespace("", ct_ns)
    part_name = f"/word/{footer_name}"

    for item in root:
        if item.attrib.get("PartName") == part_name:
            return

    override = ET.SubElement(root, f"{{{ct_ns}}}Override")
    override.set("PartName", part_name)
    override.set(
        "ContentType",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml",
    )
    tree.write(content_types_path, encoding="UTF-8", xml_declaration=True)
