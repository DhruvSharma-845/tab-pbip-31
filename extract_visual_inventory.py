import argparse
import csv
import json
import os
import zipfile
from pathlib import Path
from typing import Optional
from xml.etree import ElementTree as ET


def load_json(path: Path):
    with path.open("r", encoding="utf-8") as file:
        return json.load(file)


def strip_namespace(tag: str) -> str:
    if "}" in tag:
        return tag.split("}", 1)[1]
    return tag


def find_first_twb_in_twbx(twbx_path: Path) -> Optional[str]:
    with zipfile.ZipFile(twbx_path, "r") as zf:
        for name in zf.namelist():
            if name.lower().endswith(".twb"):
                return name
    return None


def read_twb_from_twbx(twbx_path: Path) -> str:
    twb_name = find_first_twb_in_twbx(twbx_path)
    if not twb_name:
        raise FileNotFoundError(f"No .twb found in {twbx_path}")
    with zipfile.ZipFile(twbx_path, "r") as zf:
        return zf.read(twb_name).decode("utf-8", errors="replace")


def format_field(field: dict) -> str:
    if not isinstance(field, dict):
        return json.dumps(field, ensure_ascii=True)

    if "Measure" in field:
        measure = field["Measure"]
        entity = (
            measure.get("Expression", {})
            .get("SourceRef", {})
            .get("Entity")
        )
        prop = measure.get("Property")
        return f"{entity}.{prop}" if entity and prop else prop or "Measure"

    if "Column" in field:
        column = field["Column"]
        entity = (
            column.get("Expression", {})
            .get("SourceRef", {})
            .get("Entity")
        )
        prop = column.get("Property")
        return f"{entity}.{prop}" if entity and prop else prop or "Column"

    if "Aggregation" in field:
        agg = field["Aggregation"]
        func = agg.get("Function") or "Aggregation"
        inner = agg.get("Expression") or {}
        return f"{func}({format_field(inner)})"

    if "Hierarchy" in field:
        hierarchy = field["Hierarchy"]
        return hierarchy.get("Name") or "Hierarchy"

    if "Level" in field:
        level = field["Level"]
        return level.get("Name") or "Level"

    return json.dumps(field, ensure_ascii=True)


def collect_projections(query_state: dict) -> list[dict]:
    projections = []

    def walk(obj):
        if isinstance(obj, dict):
            if "projections" in obj and isinstance(obj["projections"], list):
                for proj in obj["projections"]:
                    projections.append(proj)
            for value in obj.values():
                walk(value)
        elif isinstance(obj, list):
            for item in obj:
                walk(item)

    walk(query_state)
    return projections


def extract_pbip_visuals(report_dir: Path) -> dict:
    pages_path = report_dir / "definition" / "pages" / "pages.json"
    if not pages_path.exists():
        raise FileNotFoundError(f"PBIP pages.json not found at {pages_path}")

    pages_meta = load_json(pages_path)
    page_ids = pages_meta.get("pageOrder", [])
    pages = []
    visuals = []

    for page_id in page_ids:
        page_dir = report_dir / "definition" / "pages" / page_id
        page_json = load_json(page_dir / "page.json")
        page_info = {
            "page_id": page_id,
            "display_name": page_json.get("displayName"),
            "width": page_json.get("width"),
            "height": page_json.get("height"),
            "display_option": page_json.get("displayOption"),
        }
        pages.append(page_info)

        visuals_dir = page_dir / "visuals"
        if not visuals_dir.exists():
            continue

        for visual_id in sorted(os.listdir(visuals_dir)):
            visual_path = visuals_dir / visual_id / "visual.json"
            if not visual_path.exists():
                continue
            visual_json = load_json(visual_path)
            visual_def = visual_json.get("visual", {})
            position = visual_json.get("position", {})
            query_state = visual_def.get("query", {}).get("queryState", {})
            projections = collect_projections(query_state)
            fields = []
            for proj in projections:
                field = proj.get("field")
                if field:
                    fields.append(format_field(field))

            filter_fields = []
            for flt in visual_json.get("filterConfig", {}).get("filters", []):
                field = flt.get("field")
                if field:
                    filter_fields.append(format_field(field))

            visuals.append(
                {
                    "page_id": page_id,
                    "page_name": page_info["display_name"],
                    "visual_id": visual_json.get("name"),
                    "visual_type": visual_def.get("visualType"),
                    "x": position.get("x"),
                    "y": position.get("y"),
                    "width": position.get("width"),
                    "height": position.get("height"),
                    "fields": fields,
                    "filters": filter_fields,
                }
            )

    return {"pages": pages, "visuals": visuals}


def extract_tableau_inventory(twb_path: Path) -> dict:
    if twb_path.suffix.lower() == ".twbx":
        twb_xml = read_twb_from_twbx(twb_path)
    else:
        twb_xml = twb_path.read_text(encoding="utf-8", errors="replace")

    root = ET.fromstring(twb_xml)

    dashboards = []
    for dash in root.findall(".//dashboard"):
        dash_name = dash.get("name")
        zones = []
        for zone in dash.findall(".//zone"):
            zone_attrs = dict(zone.attrib)
            zone_type = zone_attrs.get("type") or zone_attrs.get("name")
            zones.append(
                {
                    "type": zone_type,
                    "attrs": zone_attrs,
                }
            )
        dashboards.append({"name": dash_name, "zones": zones})

    worksheets = []
    for ws in root.findall(".//worksheet"):
        ws_name = ws.get("name")
        marks = []
        for mark in ws.findall(".//mark"):
            mark_class = mark.get("class")
            if mark_class:
                marks.append(mark_class)
        worksheets.append(
            {
                "name": ws_name,
                "mark_types": sorted(set(marks)),
            }
        )

    return {"dashboards": dashboards, "worksheets": worksheets}


def write_csv(path: Path, rows: list[dict]):
    if not rows:
        return
    with path.open("w", newline="", encoding="utf-8") as file:
        writer = csv.DictWriter(file, fieldnames=list(rows[0].keys()))
        writer.writeheader()
        writer.writerows(rows)


def main():
    parser = argparse.ArgumentParser(
        description="Extract PBIP visual inventory and Tableau metadata."
    )
    parser.add_argument(
        "--pbip-report",
        default="withvisuals/Superstore.Report",
        help="Path to PBIP report folder.",
    )
    parser.add_argument(
        "--tableau",
        default="Superstore.twbx",
        help="Path to Tableau twb or twbx file.",
    )
    parser.add_argument(
        "--out-dir",
        default="out",
        help="Output directory for inventory JSON/CSV.",
    )
    args = parser.parse_args()

    pbip_report = Path(args.pbip_report)
    tableau_path = Path(args.tableau)
    out_dir = Path(args.out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    pbip_inventory = extract_pbip_visuals(pbip_report)
    tableau_inventory = extract_tableau_inventory(tableau_path)

    inventory = {
        "pbip_report": str(pbip_report),
        "tableau_workbook": str(tableau_path),
        "pbip": pbip_inventory,
        "tableau": tableau_inventory,
    }

    json_path = out_dir / "visual_inventory.json"
    with json_path.open("w", encoding="utf-8") as file:
        json.dump(inventory, file, indent=2, ensure_ascii=True)

    csv_path = out_dir / "pbip_visuals.csv"
    write_csv(csv_path, pbip_inventory["visuals"])

    print(f"Wrote {json_path}")
    print(f"Wrote {csv_path}")


if __name__ == "__main__":
    main()
