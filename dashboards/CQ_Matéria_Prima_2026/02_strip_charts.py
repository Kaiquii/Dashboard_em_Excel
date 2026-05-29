from pathlib import Path
import json
import posixpath
import tempfile
import zipfile
import xml.etree.ElementTree as ET

import openpyxl


SCRIPT_DIR = Path(__file__).resolve().parent
PROJECT_ROOT = SCRIPT_DIR.parents[1]
WORK_DIR = PROJECT_ROOT / ".tmp" / SCRIPT_DIR.name
CONFIG = json.loads((SCRIPT_DIR / "config.json").read_text(encoding="utf-8"))
INPUT = PROJECT_ROOT / CONFIG["entrada"]
SOURCE = WORK_DIR / "CQ Materia Prima - 2026 - Dashboard-openpyxl.xlsm"
OUTPUT = WORK_DIR / "CQ Materia Prima - 2026 - Dashboard-base-sem-graficos.xlsm"

NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
NS_CT = "http://schemas.openxmlformats.org/package/2006/content-types"
DRAWING_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing"
DRAWING_CT = "application/vnd.openxmlformats-officedocument.drawing+xml"


def read_parts(path):
    with zipfile.ZipFile(path, "r") as archive:
        return {item.filename: archive.read(item.filename) for item in archive.infolist()}


def write_parts(path, parts):
    with tempfile.NamedTemporaryFile(delete=False, suffix=path.suffix, dir=path.parent) as tmp:
        tmp_path = Path(tmp.name)

    with zipfile.ZipFile(tmp_path, "w", zipfile.ZIP_DEFLATED, allowZip64=True) as archive:
        for name, data in parts.items():
            archive.writestr(name, data)

    tmp_path.replace(path)


def rels_path(part_name):
    return posixpath.join(
        posixpath.dirname(part_name),
        "_rels",
        posixpath.basename(part_name) + ".rels",
    )


def resolve_target(source_part, target):
    if target.startswith("/"):
        return target.lstrip("/")
    return posixpath.normpath(posixpath.join(posixpath.dirname(source_part), target))


def relative_target(source_part, target_part):
    return posixpath.relpath(target_part, posixpath.dirname(source_part))


def local_name(tag):
    return tag.rsplit("}", 1)[-1]


def sheet_parts(parts):
    workbook = ET.fromstring(parts["xl/workbook.xml"])
    relationships = ET.fromstring(parts["xl/_rels/workbook.xml.rels"])
    rel_map = {rel.attrib["Id"]: rel.attrib["Target"] for rel in relationships}
    sheets = {}

    for sheet in workbook.findall(f".//{{{NS_MAIN}}}sheet"):
        rid = sheet.attrib[f"{{{NS_REL}}}id"]
        sheets[sheet.attrib["name"]] = resolve_target("xl/workbook.xml", rel_map[rid])

    return sheets


def parse_relationships(parts, path):
    if path in parts:
        return ET.fromstring(parts[path])
    return ET.Element(f"{{{NS_PKG_REL}}}Relationships")


def next_relationship_id(root):
    used = set()
    for rel in root:
        rid = rel.attrib.get("Id", "")
        if rid.startswith("rId") and rid[3:].isdigit():
            used.add(int(rid[3:]))

    next_id = 1
    while next_id in used:
        next_id += 1
    return f"rId{next_id}"


def insert_worksheet_drawing(sheet_root, drawing):
    # Excel expects drawings before tableParts/extLst inside worksheet XML.
    for index, child in enumerate(list(sheet_root)):
        if local_name(child.tag) in {"tableParts", "extLst"}:
            sheet_root.insert(index, drawing)
            return
    sheet_root.append(drawing)


def add_content_type(parts, part_name, content_type):
    root = ET.fromstring(parts["[Content_Types].xml"])
    existing = {
        child.attrib.get("PartName")
        for child in root
        if local_name(child.tag) == "Override"
    }
    normalized = "/" + part_name.lstrip("/")
    if normalized not in existing:
        ET.SubElement(
            root,
            f"{{{NS_CT}}}Override",
            {"PartName": normalized, "ContentType": content_type},
        )
        parts["[Content_Types].xml"] = ET.tostring(root, encoding="utf-8", xml_declaration=True)


def add_default_content_type(parts, extension, content_type):
    root = ET.fromstring(parts["[Content_Types].xml"])
    existing = {
        child.attrib.get("Extension")
        for child in root
        if local_name(child.tag) == "Default"
    }
    if extension not in existing:
        ET.SubElement(
            root,
            f"{{{NS_CT}}}Default",
            {"Extension": extension, "ContentType": content_type},
        )
        parts["[Content_Types].xml"] = ET.tostring(root, encoding="utf-8", xml_declaration=True)


def restore_original_drawings(target_path, original_path):
    original_parts = read_parts(original_path)
    target_parts = read_parts(target_path)
    original_sheets = sheet_parts(original_parts)
    target_sheets = sheet_parts(target_parts)
    restored = 0

    for media_name, data in original_parts.items():
        if media_name.startswith("xl/media/"):
            target_parts[media_name] = data
            extension = media_name.rsplit(".", 1)[-1].lower()
            if extension == "png":
                add_default_content_type(target_parts, "png", "image/png")
            elif extension in {"jpg", "jpeg"}:
                add_default_content_type(target_parts, extension, "image/jpeg")

    for sheet_name, original_sheet_part in original_sheets.items():
        if sheet_name not in target_sheets or sheet_name == "Dashboard2":
            continue

        original_rels_path = rels_path(original_sheet_part)
        if original_rels_path not in original_parts:
            continue

        original_rels = ET.fromstring(original_parts[original_rels_path])
        drawing_rels = [
            rel for rel in original_rels
            if rel.attrib.get("Type") == DRAWING_REL
        ]
        if not drawing_rels:
            continue

        target_sheet_part = target_sheets[sheet_name]
        target_rels_path = rels_path(target_sheet_part)
        target_rels = parse_relationships(target_parts, target_rels_path)
        target_sheet = ET.fromstring(target_parts[target_sheet_part])

        for existing in list(target_sheet.findall(f"{{{NS_MAIN}}}drawing")):
            target_sheet.remove(existing)

        for original_rel in drawing_rels:
            original_drawing_part = resolve_target(original_sheet_part, original_rel.attrib["Target"])
            if original_drawing_part not in original_parts:
                continue

            target_parts[original_drawing_part] = original_parts[original_drawing_part]
            original_drawing_rels_path = rels_path(original_drawing_part)
            if original_drawing_rels_path in original_parts:
                target_parts[original_drawing_rels_path] = original_parts[original_drawing_rels_path]

            rid = next_relationship_id(target_rels)
            ET.SubElement(
                target_rels,
                f"{{{NS_PKG_REL}}}Relationship",
                {
                    "Id": rid,
                    "Type": DRAWING_REL,
                    "Target": relative_target(target_sheet_part, original_drawing_part),
                },
            )
            drawing = ET.Element(f"{{{NS_MAIN}}}drawing", {f"{{{NS_REL}}}id": rid})
            insert_worksheet_drawing(target_sheet, drawing)
            add_content_type(target_parts, original_drawing_part, DRAWING_CT)
            restored += 1

        target_parts[target_rels_path] = ET.tostring(target_rels, encoding="utf-8", xml_declaration=True)
        target_parts[target_sheet_part] = ET.tostring(target_sheet, encoding="utf-8", xml_declaration=True)

    write_parts(target_path, target_parts)
    return restored


def main():
    wb = openpyxl.load_workbook(SOURCE, keep_vba=True)
    ws = wb["Dashboard2"]
    ws._charts = []
    ws.protection.sheet = False
    ws.protection.objects = False
    ws.protection.scenarios = False
    wb.save(OUTPUT)
    restored = restore_original_drawings(OUTPUT, INPUT)
    print(f"Etapa 2 concluida. Desenhos originais preservados: {restored}")


if __name__ == "__main__":
    main()
