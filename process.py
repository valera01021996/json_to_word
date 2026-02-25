import json
import sys
import time
import shutil
import logging
from pathlib import Path
from email import policy
from email.parser import BytesParser
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

SCRIPT_DIR = Path(__file__).parent
TEMPLATE_PATH = SCRIPT_DIR / "template.docx"
EML_WAIT_TIMEOUT = 60   # секунд ждать EML файл
EML_WAIT_INTERVAL = 2   # интервал проверки в секундах

# Используем именованный логгер — не трогаем конфиг рутового логгера.
# Когда process.py импортирует daemon.py, используется его basicConfig.
# Когда process.py запускается напрямую — вывод идёт в stderr (дефолт).
logger = logging.getLogger(__name__)


def human_size(num_bytes):
    for unit in ["B", "KB", "MB", "GB"]:
        if num_bytes < 1024 or unit == "GB":
            return f"{num_bytes:.0f} {unit}" if unit == "B" else f"{num_bytes:.2f} {unit}"
        num_bytes /= 1024


def parse_eml(eml_path):
    with open(eml_path, "rb") as f:
        msg = BytesParser(policy=policy.default).parse(f)

    text_plain_parts = []
    text_html_parts = []

    for part in msg.walk():
        if part.is_multipart():
            continue
        ctype = part.get_content_type()
        disp = part.get_content_disposition()
        filename = part.get_filename()

        if disp == "attachment" or filename:
            continue

        if ctype == "text/plain":
            try:
                text_plain_parts.append(part.get_content())
            except Exception:
                pass
        elif ctype == "text/html":
            try:
                text_html_parts.append(part.get_content())
            except Exception:
                pass

    body_text = ""
    if text_plain_parts:
        body_text = "\n\n".join(t.strip() for t in text_plain_parts if t and t.strip())
    elif text_html_parts:
        body_text = "\n\n".join(t.strip() for t in text_html_parts if t and t.strip())

    attachments = []
    for part in msg.walk():
        if part.is_multipart():
            continue
        disp = part.get_content_disposition()
        filename = part.get_filename()
        if disp != "attachment" and not filename:
            continue
        filename = filename or "attachment.bin"
        data = part.get_payload(decode=True)
        size = len(data) if data else 0
        attachments.append((filename, size))

    return body_text, attachments


def parse_json(json_path):
    with open(json_path, encoding="utf-8") as f:
        data = json.load(f)

    record = data["data"]["asdf"][0]
    meta = record["meta"]

    test3_raw = meta.get("test3", [])
    if isinstance(test3_raw, list):
        test3_val = "\n".join(str(v) for v in test3_raw)
    else:
        test3_val = str(test3_raw)

    eml_filename = record.get("filename") or meta.get("filename") or ""
    if isinstance(eml_filename, list):
        eml_filename = eml_filename[0] if eml_filename else ""

    return {
        "start_time": str(record.get("start_time", "")),
        "stop_time":  str(record.get("stop_time", "")),
        "test4":      str(meta.get("test4", "")),
        "test3":      test3_val,
        "eml_filename": eml_filename,
    }


def clear_cell(cell):
    for para in cell.paragraphs:
        for run in para.runs:
            run.text = ""
    for para in cell.paragraphs[1:]:
        para._element.getparent().remove(para._element)


def replace_cell_text(cell, new_text):
    lines = str(new_text).split("\n")
    clear_cell(cell)
    first_para = cell.paragraphs[0]
    if first_para.runs:
        first_para.runs[0].text = lines[0]
    else:
        first_para.add_run(lines[0])
    for line in lines[1:]:
        new_para = OxmlElement("w:p")
        cell._element.append(new_para)
        cell.paragraphs[-1].add_run(line)


def fill_table_cell(table, label, value):
    for row in table.rows:
        unique_cells = []
        seen = set()
        for cell in row.cells:
            if id(cell) not in seen:
                seen.add(id(cell))
                unique_cells.append(cell)
        for i, cell in enumerate(unique_cells):
            if cell.text.strip() == label and i + 1 < len(unique_cells):
                replace_cell_text(unique_cells[i + 1], value)
                return


def make_paragraph(text, bold=False):
    p = OxmlElement("w:p")
    if text:
        r = OxmlElement("w:r")
        if bold:
            rPr = OxmlElement("w:rPr")
            rPr.append(OxmlElement("w:b"))
            r.append(rPr)
        t = OxmlElement("w:t")
        t.text = text
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        r.append(t)
        p.append(r)
    return p


def insert_text_before_tables(doc, body_text, attachments):
    body = doc.element.body
    first_table = body.find(qn("w:tbl"))
    if first_table is None:
        return

    idx = list(body).index(first_table)
    inserts = []

    if body_text:
        inserts.append(make_paragraph("Текст сообщения:", bold=True))
        for line in body_text.split("\n"):
            inserts.append(make_paragraph(line))
        inserts.append(make_paragraph(""))

    if attachments:
        inserts.append(make_paragraph("Вложения:", bold=True))
        for name, size in attachments:
            inserts.append(make_paragraph(f"- {name} | {size} байт ({human_size(size)})"))
        inserts.append(make_paragraph(""))

    for para in reversed(inserts):
        body.insert(idx, para)


def process_json(json_file):
    json_path = Path(json_file)
    watch_dir = json_path.parent

    logger.info(f"Processing: {json_path}")

    try:
        data = parse_json(json_path)
    except Exception as e:
        logger.error(f"Failed to parse JSON {json_path}: {e}")
        return

    body_text = ""
    attachments = []
    eml_filename = data["eml_filename"]
    if eml_filename:
        eml_path = watch_dir / eml_filename
        if not eml_path.exists():
            logger.info(f"Waiting for EML: {eml_path}")
            elapsed = 0
            while not eml_path.exists() and elapsed < EML_WAIT_TIMEOUT:
                time.sleep(EML_WAIT_INTERVAL)
                elapsed += EML_WAIT_INTERVAL
        if eml_path.exists():
            try:
                body_text, attachments = parse_eml(eml_path)
            except Exception as e:
                logger.error(f"Failed to parse EML {eml_path}: {e}")
        else:
            logger.warning(f"EML not found after {EML_WAIT_TIMEOUT}s, skipping: {eml_path}")

    try:
        doc = Document(TEMPLATE_PATH)
    except Exception as e:
        logger.error(f"Failed to open template: {e}")
        return

    tables = doc.tables

    # Таблица 1: Test1, Test2, Test3, Test4
    fill_table_cell(tables[0], "Test1", data["start_time"])
    fill_table_cell(tables[0], "Test2", data["stop_time"])
    fill_table_cell(tables[0], "Test3", data["test4"])
    fill_table_cell(tables[0], "Test4", data["test3"])

    # Таблица 2: Test5, Test6, Test7, Test8
    fill_table_cell(tables[1], "Test5", data["start_time"])
    fill_table_cell(tables[1], "Test6", data["stop_time"])
    fill_table_cell(tables[1], "Test7", data["test4"])
    fill_table_cell(tables[1], "Test8", data["test3"])

    insert_text_before_tables(doc, body_text, attachments)

    output_path = watch_dir / (json_path.stem + ".docx")
    tmp_path = output_path.with_suffix(".tmp")
    try:
        doc.save(tmp_path)
        shutil.move(tmp_path, output_path)  # атомарная операция — либо целый файл либо ничего
        logger.info(f"Saved: {output_path}")
    except Exception as e:
        logger.error(f"Failed to save {output_path}: {e}")
        tmp_path.unlink(missing_ok=True)  # убрать мусор если что-то пошло не так


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python process.py file.json")
        sys.exit(1)
    # При запуске напрямую — вывод логов в консоль
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
    process_json(sys.argv[1])
