import re
import sys
import zipfile
from pathlib import Path

from lxml import etree
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT


NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
}


RED = RGBColor(192, 0, 0)


def tag_name(element):
    return etree.QName(element).localname


def get_text(element):
    texts = element.xpath(".//w:t/text() | .//w:delText/text()", namespaces=NS)
    return "".join(texts)


def paragraph_text(paragraph):
    return get_text(paragraph).strip()


def is_red_run(run):
    colors = run.xpath(".//w:rPr/w:color/@w:val", namespaces=NS)
    for color in colors:
        color = color.lower()
        if color in ["ff0000", "c00000", "c0392b", "e60000", "b00000", "a00000", "red"]:
            return True
    return False


def is_strike_run(run):
    return bool(run.xpath(".//w:rPr/w:strike | .//w:rPr/w:dstrike", namespaces=NS))


def is_underline_run(run):
    return bool(run.xpath(".//w:rPr/w:u", namespaces=NS))


def detect_page(text, current_page):
    match = re.search(r"Page Number\s*:\s*(\d+)", text, re.IGNORECASE)
    if match:
        return int(match.group(1))
    return current_page


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


def add_docx_run(paragraph, text, red=False, strike=False, underline=False, bold=False):
    if not text:
        return

    run = paragraph.add_run(text)
    run.font.size = Pt(10)

    if red:
        run.font.color.rgb = RED

    if strike:
        run.font.strike = True

    if underline:
        run.font.underline = True

    if bold:
        run.bold = True


def collect_runs_from_element(element, inside_change=None):
    """
    Returns runs as dictionaries:
    {
      text,
      red,
      strike,
      underline,
      is_redline
    }
    """
    results = []
    local = tag_name(element)

    if local == "del":
        inside_change = "delete"
    elif local == "ins":
        inside_change = "insert"

    if local == "r":
        text = get_text(element)
        if text:
            visual_red = is_red_run(element)
            visual_strike = is_strike_run(element)
            visual_underline = is_underline_run(element)

            if inside_change == "delete":
                results.append({
                    "text": text,
                    "red": True,
                    "strike": True,
                    "underline": False,
                    "is_redline": True,
                })
            elif inside_change == "insert":
                results.append({
                    "text": text,
                    "red": True,
                    "strike": False,
                    "underline": True,
                    "is_redline": True,
                })
            elif visual_red or visual_strike:
                results.append({
                    "text": text,
                    "red": True,
                    "strike": visual_strike,
                    "underline": visual_underline,
                    "is_redline": True,
                })
            else:
                results.append({
                    "text": text,
                    "red": False,
                    "strike": False,
                    "underline": visual_underline,
                    "is_redline": False,
                })

        return results

    for child in element:
        results.extend(collect_runs_from_element(child, inside_change))

    return results


def paragraph_has_redline(paragraph):
    runs = collect_runs_from_element(paragraph)
    return any(run["is_redline"] for run in runs)


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
        text = paragraph_text(paragraph)

        if text:
            current_page = detect_page(text, current_page)
            current_lesson = detect_lesson(text, current_lesson)
            current_template = detect_template(text, current_template)

        if paragraph_has_redline(paragraph):
            runs = collect_runs_from_element(paragraph)

            results.append({
                "page": current_page,
                "lesson": current_lesson,
                "template": current_template,
                "runs": runs,
            })

    return results


def set_cell_header(cell, text):
    paragraph = cell.paragraphs[0]
    add_docx_run(paragraph, text, bold=True)


def create_two_column_docx(results, output_file):
    doc = Document()

    section = doc.sections[0]
    section.left_margin = Inches(0.45)
    section.right_margin = Inches(0.45)
    section.top_margin = Inches(0.5)
    section.bottom_margin = Inches(0.5)

    title = doc.add_paragraph()
    add_docx_run(title, "Redline Review Output", bold=True)

    intro = doc.add_paragraph()
    add_docx_run(
        intro,
        "Left column shows the original paragraph or sentence. Right column shows only the redlined content from that same block."
    )

    if not results:
        doc.add_paragraph("No redline changes found.")
        doc.save(output_file)
        return

    for item in results:
        meta = doc.add_paragraph()
        add_docx_run(
            meta,
            f"Page Number: {item['page']} | Lesson: {item['lesson']} | Template Used: {item['template']}",
            bold=True
        )

        table = doc.add_table(rows=2, cols=2)
        table.style = "Table Grid"
        table.alignment = WD_TABLE_ALIGNMENT.CENTER

        table.cell(0, 0).vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP
        table.cell(0, 1).vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP
        table.cell(1, 0).vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP
        table.cell(1, 1).vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP

        set_cell_header(table.cell(0, 0), "Original Content")
        set_cell_header(table.cell(0, 1), "Redline Content Only")

        original_paragraph = table.cell(1, 0).paragraphs[0]
        redline_paragraph = table.cell(1, 1).paragraphs[0]

        for run in item["runs"]:
            add_docx_run(
                original_paragraph,
                run["text"],
                red=run["red"],
                strike=run["strike"],
                underline=run["underline"],
            )

        for run in item["runs"]:
            if run["is_redline"]:
                add_docx_run(
                    redline_paragraph,
                    run["text"],
                    red=run["red"],
                    strike=run["strike"],
                    underline=run["underline"],
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

    docx_files = sorted(input_folder.glob("**/*.docx"))

    if not docx_files:
        print("ERROR: No .docx files found in the input folder.")
        sys.exit(2)

    all_results = []

    for docx_file in docx_files:
        print(f"Processing: {docx_file}")
        all_results.extend(extract_redline_blocks(docx_file))

    create_two_column_docx(all_results, output_file)

    print(f"Created report: {output_file}")
    print(f"Redline blocks found: {len(all_results)}")


if __name__ == "__main__":
    main()
