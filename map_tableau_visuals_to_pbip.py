import argparse
import json
import re
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET


SKIP_ZONE_TYPES = {
    "layout-basic",
    "layout-flow",
    "filter",
    "legend",
    "parameter",
    "storyboard",
}


def read_twb_xml(twb_or_twbx_path: Path) -> str:
    if twb_or_twbx_path.suffix.lower() == ".twbx":
        with zipfile.ZipFile(twb_or_twbx_path, "r") as zf:
            twb_name = None
            for name in zf.namelist():
                if name.lower().endswith(".twb"):
                    twb_name = name
                    break
            if not twb_name:
                raise FileNotFoundError(f"No .twb found in {twb_or_twbx_path}")
            return zf.read(twb_name).decode("utf-8", errors="replace")
    return twb_or_twbx_path.read_text(encoding="utf-8", errors="replace")


def slugify(value: str) -> str:
    value = value.lower()
    value = re.sub(r"[^a-z0-9]+", " ", value)
    value = re.sub(r"\s+", " ", value).strip()
    return value


def tokenize(value: str) -> list[str]:
    return [token for token in slugify(value).split(" ") if token]


def jaccard(a: set[str], b: set[str]) -> float:
    if not a or not b:
        return 0.0
    return len(a & b) / len(a | b)


def load_json(path: Path):
    return json.loads(path.read_text(encoding="utf-8"))


def read_tableau_dashboards(twb_or_twbx_path: Path) -> dict:
    twb_xml = read_twb_xml(twb_or_twbx_path)
    root = ET.fromstring(twb_xml)
    dashboards = {}
    for dash in root.findall(".//dashboard"):
        dash_name = dash.get("name")
        if not dash_name:
            continue
        worksheets = []
        for zone in dash.findall(".//zone"):
            name = zone.get("name")
            zone_type = zone.get("type-v2") or zone.get("type")
            if not name:
                continue
            if zone_type in SKIP_ZONE_TYPES:
                continue
            worksheets.append(
                {
                    "name": name,
                    "zone_type": zone_type,
                }
            )
        dashboards[dash_name] = worksheets
    return dashboards


def extract_text_runs(visual: dict) -> list[str]:
    texts = []
    objects = visual.get("objects", {})
    for section in objects.values():
        if not isinstance(section, list):
            continue
        for item in section:
            properties = item.get("properties", {})
            paragraphs = properties.get("paragraphs", [])
            for paragraph in paragraphs:
                for run in paragraph.get("textRuns", []):
                    value = run.get("value")
                    if value:
                        texts.append(value)
    return texts


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


def format_field(field: dict) -> str:
    if not isinstance(field, dict):
        return ""

    if "Measure" in field:
        measure = field["Measure"]
        prop = measure.get("Property")
        return prop or ""

    if "Column" in field:
        column = field["Column"]
        prop = column.get("Property")
        return prop or ""

    if "Aggregation" in field:
        agg = field["Aggregation"]
        func = agg.get("Function")
        inner = agg.get("Expression") or {}
        return f"{func} {format_field(inner)}".strip()

    if "Hierarchy" in field:
        hierarchy = field["Hierarchy"]
        return hierarchy.get("Name") or ""

    if "Level" in field:
        level = field["Level"]
        return level.get("Name") or ""

    return ""


def build_pbip_visual_index(report_dir: Path) -> dict:
    pages_path = report_dir / "definition" / "pages" / "pages.json"
    pages_meta = load_json(pages_path)
    index = {}

    for page_id in pages_meta.get("pageOrder", []):
        page_json = load_json(
            report_dir / "definition" / "pages" / page_id / "page.json"
        )
        display_name = page_json.get("displayName")
        if not display_name:
            continue

        visuals = []
        visuals_dir = report_dir / "definition" / "pages" / page_id / "visuals"
        if visuals_dir.exists():
            for visual_id in sorted(visuals_dir.iterdir()):
                visual_json_path = visual_id / "visual.json"
                if not visual_json_path.exists():
                    continue
                visual_json = load_json(visual_json_path)
                visual_def = visual_json.get("visual", {})
                query_state = visual_def.get("query", {}).get("queryState", {})
                projections = collect_projections(query_state)
                fields = []
                for proj in projections:
                    field = proj.get("field")
                    if field:
                        field_name = format_field(field)
                        if field_name:
                            fields.append(field_name)

                filter_fields = []
                for flt in visual_json.get("filterConfig", {}).get("filters", []):
                    field = flt.get("field")
                    if field:
                        field_name = format_field(field)
                        if field_name:
                            filter_fields.append(field_name)

                text_runs = []
                if visual_def.get("visualType") == "textbox":
                    text_runs = extract_text_runs(visual_def)

                visuals.append(
                    {
                        "visual_id": visual_json.get("name"),
                        "visual_type": visual_def.get("visualType"),
                        "fields": fields,
                        "filters": filter_fields,
                        "text": text_runs,
                        "path": str(visual_json_path),
                    }
                )

        index[display_name] = {
            "page_id": page_id,
            "visuals": visuals,
        }

    return index


def visual_tokens(visual: dict) -> set[str]:
    tokens = []
    if visual.get("visual_type"):
        tokens += tokenize(visual["visual_type"])
    for field in visual.get("fields", []):
        tokens += tokenize(field)
    for field in visual.get("filters", []):
        tokens += tokenize(field)
    for text in visual.get("text", []):
        tokens += tokenize(text)
    return set(tokens)


def match_worksheets_to_visuals(dashboard: str, worksheets: list, visuals: list) -> list:
    visual_token_map = []
    for visual in visuals:
        visual_token_map.append((visual, visual_tokens(visual)))

    results = []
    for ws in worksheets:
        ws_tokens = set(tokenize(ws["name"]))
        scored = []
        for visual, tokens in visual_token_map:
            score = jaccard(ws_tokens, tokens)
            if score > 0:
                scored.append((score, visual))
        scored.sort(key=lambda item: item[0], reverse=True)
        best = scored[:3]
        results.append(
            {
                "worksheet": ws["name"],
                "dashboard": dashboard,
                "best_matches": [
                    {
                        "visual_id": match["visual_id"],
                        "visual_type": match["visual_type"],
                        "score": round(score, 3),
                        "path": match["path"],
                    }
                    for score, match in best
                ],
            }
        )
    return results


def main():
    parser = argparse.ArgumentParser(
        description="Map Tableau worksheets to PBIP visuals by name/field tokens."
    )
    parser.add_argument(
        "--tableau-twbx",
        default="Superstore.twbx",
        help="Tableau TWB/TWBX file.",
    )
    parser.add_argument(
        "--pbip-report",
        default="SuperstorePBIP/Superstore.Report",
        help="PBIP report folder.",
    )
    parser.add_argument(
        "--out",
        default="out/tableau_to_pbip_visual_map.json",
        help="Output JSON path.",
    )
    args = parser.parse_args()

    dashboards = read_tableau_dashboards(Path(args.tableau_twbx))
    pbip_index = build_pbip_visual_index(Path(args.pbip_report))

    mappings = []
    for dash_name, worksheets in dashboards.items():
        page = pbip_index.get(dash_name)
        if not page:
            mappings.append(
                {
                    "dashboard": dash_name,
                    "pbip_page": None,
                    "note": "No PBIP page with matching displayName",
                    "worksheet_map": [],
                }
            )
            continue
        worksheet_map = match_worksheets_to_visuals(
            dash_name, worksheets, page["visuals"]
        )
        mappings.append(
            {
                "dashboard": dash_name,
                "pbip_page": dash_name,
                "pbip_page_id": page["page_id"],
                "worksheet_map": worksheet_map,
            }
        )

    out_path = Path(args.out)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_path.write_text(json.dumps(mappings, indent=2), encoding="utf-8")
    print(f"Wrote {out_path}")


if __name__ == "__main__":
    main()
