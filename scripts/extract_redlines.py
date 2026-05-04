import sys
import zipfile
from pathlib import Path
from copy import deepcopy

from lxml import etree
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT


NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
}


def qn(tag):
    prefix, name = tag.split(":")
    return f"{{{NS[prefix]}}}{name}"


def get_text(element):
    texts = element.xpath(".//w:t/text() | .//w:delText/text()", namespaces=NS)
    return "".join(texts).strip()


def paragraph_text(paragraph):
    texts = paragraph.xpath(".//w:t/text() | .//w:delText/text()", namespaces=NS)
    return "".join(texts).strip()


def get_alignment(paragraph):
    jc = paragraph.xpath("./w:pPr/w:jc/@w:val", namespaces=NS)
    return jc[0] if jc else None


def detect_lesson(text, current_lesson):
    clean = text.strip()
    lowered = clean.lower()

    if lowered.startswith("lesson ") and ":" in clean:
        return clean

    if lowered.startswith("lesson:"):
        return clean.split(":", 1)[1].strip()

    if lowered.startswith("lesson name:"):
        return clean.split(":", 1)[1].strip()

    return current_lesson


def detect_template(text, current_template):
    clean = text.strip()
    lowered = clean.lower()

    known_templates = [
        "click and reveal template",
        "binary list template",
        "question and answer template",
        "text and image template",
        "video template",
        "hotspot image template",
    ]

    for template in known_templates:
        if lowered.startswith(template):
            return clean

    if lowered.startswith("select type"):
        return clean.replace("Select Type", "").replace(":", "").strip() or current_template

    if lowered.startswith("template:"):
        return clean.split(":", 1)[1].strip()

    if lowered.startswith("template used:"):
        return clean.split(":", 1)[1].strip()

    if lowered.startswith("template type:"):
        return clean.split(":", 1)[1].strip()

    return current_template


def is_red_run(run):
    color_values = run.xpath(".//w:rPr/w:color/@w:val", namespaces=NS)

    for color in color_values:
        color = color.lower()
        if color in ["ff0000", "red", "c00000", "e60000", "b00000", "a00000"]:
            return True

    return False


def is_strikethrough_run(run):
    return bool(run.xpath(".//w:rPr/w:strike | .//w:rPr/w:dstrike", namespaces=NS))


def is_underlined_run(run):
    return bool(run.xpath(".//w:rPr/w:u", namespaces=NS))


def is_visual_redline_run(run):
    return is_red_run(run) or is_strikethrough_run(run)


def extract_visual_redline_text(paragraph):
    changes = []
    current = []

    runs = paragraph.xpath(".//w:r", namespaces=NS)

    for run in runs:
        text = get_text(run)
        if not text:
            continue

        if is_visual_redline_run(run):
            current.append(text)
        else:
            if current:
                changes.append("".join(current).strip())
                current = []

    if current:
        changes.append("".join(current).strip())

    return [change for change in changes if change]


def extract_track_change_text(paragraph):
    changes = []

    for insertion in paragraph.xpath(".//w:ins", namespaces=NS):
        changed_text = get_text(insertion)
        if changed_text:
            changes.append(changed_text)

    for deletion in paragraph.xpath(".//w:del", namespaces=NS):
        changed_text = get_text(deletion)
        if changed_text:
            changes.append(changed_text)

    return changes


def paragraph_has_redline(paragraph):
    return bool(extract_track_change_text(paragraph) or extract_visual_redline_text(paragraph))


def add_run_with_style(docx_paragraph, text, red=False, strike=False, underline=False, bold=False):
    run = docx_paragraph.add_run(text)
    run.font.size = Pt(10)

    if red:
        run.font.color.rgb = RGBColor(192, 0, 0)

    if strike:
        run.font.strike = True

    if underline:
        run.font.underline = True

    if bold:
        run.bold = True

    return run


def add_xml_runs_to_docx_paragraph(source_paragraph, target_paragraph, redline_only=False):
    for run in source_paragraph.xpath(".//w:r", namespaces=NS):
        text = get_text(run)
        if not text:
            continue

        red = is_red_run(run)
        strike = is_strikethrough_run(run)
        underline = is_underlined_run(run)

        if redline_only and not (red or strike):
            continue

        add_run_with_style(
            target_paragraph,
            text,
            red=red or strike,
            strike=strike,
            underline=underline,
        )


def apply_alignment(docx_paragraph, alignment_value):
    if alignment_value == "center":
        docx_paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    elif alignment_value == "right":
        docx_paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
    elif alignment_value == "both":
        docx_paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
    else:
        docx_paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT


def set_cell_text(cell, label):
    paragraph = cell.paragraphs[0]
    run = paragraph.add_run(label)
    run.bold = True
    run.font.size = Pt(10)


def extract_redline_blocks(docx_path):
    results = []

    current_page = 1
    current_lesson = "Unknown"
    current_template = "Unknown"

    with zipfile.ZipFile(docx_path) as docx:
        xml = docx.read("word/document.xml")

    root = etree.fromstring(xml)
    paragraphs = root.xpath(".//w:p", namespaces=NS)

    for paragraph in paragraphs:
        page_breaks = paragraph.xpath(
            ".//w:lastRenderedPageBreak | .//w:br[@w:type='page']",
            namespaces=NS
        )

        full_text = paragraph_text(paragraph)

        if full_text:
            current_lesson = detect_lesson(full_text, current_lesson)
            current_template = detect_template(full_text, current_template)

        if paragraph_has_redline(paragraph):
            redline_parts = extract_track_change_text(paragraph) + extract_visual_redline_text(paragraph)

            results.append({
                "page": current_page,
                "lesson": current_lesson,
                "template": current_template,
                "original_paragraph": paragraph,
                "original_text": full_text,
                "redline_text": " ".join(redline_parts),
                "alignment": get_alignment(paragraph),
            })

        current_page += len(page_breaks)

    return results


def create_two_column_docx(results, output_file):
    doc = Document()

    section = doc.sections[0]
    section.left_margin = Inches(0.45)
    section.right_margin = Inches(0.45)
    section.top_margin = Inches(0.5)
    section.bottom_margin = Inches(0.5)

    title = doc.add_paragraph()
    title_run = title.add_run("Redline Review Output")
    title_run.bold = True
    title_run.font.size = Pt(16)

    intro = doc.add_paragraph()
    intro.add_run("Left column shows the original paragraph or sentence. Right column shows only the redlined content from that same block.")

    if not results:
        doc.add_paragraph("No redline changes found.")
        doc.save(output_file)
        return

    for index, item in enumerate(results, start=1):
        meta = doc.add_paragraph()
        meta_run = meta.add_run(
            f"Page Number: {item['page']} | Lesson: {item['lesson']} | Template Used: {item['template']}"
        )
        meta_run.bold = True
        meta_run.font.size = Pt(10)

        table = doc.add_table(rows=2, cols=2)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        table.style = "Table Grid"

        table.columns[0].width = Inches(3.75)
        table.columns[1].width = Inches(3.75)

        header_left = table.cell(0, 0)
        header_right = table.cell(0, 1)
        set_cell_text(header_left, "Original Content")
        set_cell_text(header_right, "Redline Content Only")

        original_cell = table.cell(1, 0)
        redline_cell = table.cell(1, 1)

        original_cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP
        redline_cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP

        original_p = original_cell.paragraphs[0]
        redline_p = redline_cell.paragraphs[0]

        apply_alignment(original_p, item["alignment"])
        apply_alignment(redline_p, item["alignment"])

        add_xml_runs_to_docx_paragraph(
            item["original_paragraph"],
            original_p,
            redline_only=False
        )

        add_xml_runs_to_docx_paragraph(
            item["original_paragraph"],
            redline_p,
            redline_only=True
        )

        doc.add_paragraph("")

    doc.save(output_file)


def main():
    if len(sys.argv) != 3:
        print("Usage: python scripts/extract_redlines.py <input_folder> <output_docx>")
        sys.exit(1)

    input_folder = Path(sys.argv[1])
    output_file = Path(sys.argv[2])
    output_file.parent.mkdir(parents=True, exist_ok=True)

    all_results = []

    for docx_file in sorted(input_folder.glob("**/*.docx")):
        all_results.extend(extract_redline_blocks(docx_file))

    create_two_column_docx(all_results, output_file)

    if not all_results:
        print("No redline changes found.")
    else:
        print(f"Created two-column DOCX with {len(all_results)} redline blocks.")


if __name__ == "__main__":
    main()
