import sys
import zipfile
from pathlib import Path
from lxml import etree


NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
}


def text_from_element(element):
    texts = element.xpath(".//w:t/text()", namespaces=NS)
    return "".join(texts).strip()


def paragraph_text(paragraph):
    texts = paragraph.xpath(".//w:t/text()", namespaces=NS)
    return "".join(texts).strip()


def detect_lesson(text, current_lesson):
    lowered = text.lower()

    if lowered.startswith("lesson:"):
        return text.split(":", 1)[1].strip()

    if lowered.startswith("lesson name:"):
        return text.split(":", 1)[1].strip()

    return current_lesson


def detect_template(text, current_template):
    lowered = text.lower()

    if lowered.startswith("template:"):
        return text.split(":", 1)[1].strip()

    if lowered.startswith("template used:"):
        return text.split(":", 1)[1].strip()

    if lowered.startswith("template type:"):
        return text.split(":", 1)[1].strip()

    return current_template


def extract_redlines_from_docx(docx_path):
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

        insertions = paragraph.xpath(".//w:ins", namespaces=NS)
        deletions = paragraph.xpath(".//w:del", namespaces=NS)

        for insertion in insertions:
            changed_text = text_from_element(insertion)
            if changed_text:
                results.append({
                    "page": current_page,
                    "lesson": current_lesson,
                    "template": current_template,
                    "change": changed_text
                })

        for deletion in deletions:
            changed_text = text_from_element(deletion)
            if changed_text:
                results.append({
                    "page": current_page,
                    "lesson": current_lesson,
                    "template": current_template,
                    "change": changed_text
                })

        current_page += len(page_breaks)

    return results


def write_report(all_results, output_path):
    lines = []

    for result in all_results:
        lines.append(f"Page Number: {result['page']}")
        lines.append(f"Lesson: {result['lesson']}")
        lines.append(f"Template Used: {result['template']}")
        lines.append(f"Redline Change: {result['change']}")
        lines.append("")

    if lines:
        output_path.write_text("\n".join(lines).strip() + "\n", encoding="utf-8")
    else:
        output_path.write_text("No redline changes found.\n", encoding="utf-8")


def main():
    if len(sys.argv) != 3:
        print("Usage: python scripts/extract_redlines.py <input_folder> <output_file>")
        sys.exit(1)

    input_folder = Path(sys.argv[1])
    output_file = Path(sys.argv[2])
    output_file.parent.mkdir(parents=True, exist_ok=True)

    all_results = []

    for docx_file in sorted(input_folder.glob("**/*.docx")):
        all_results.extend(extract_redlines_from_docx(docx_file))

    write_report(all_results, output_file)

    if not all_results:
        print("No redline changes found.")
    else:
        print(f"Extracted {len(all_results)} redline changes.")


if __name__ == "__main__":
    main()
