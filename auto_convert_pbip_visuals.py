import argparse
import json
import math
import re
from pathlib import Path
from typing import Dict, List, Optional, Tuple


TOP_TYPES = {"textbox", "card", "slicer"}
MIDDLE_TYPES = {
    "map",
    "filledMap",
    "shapeMap",
    "areaChart",
    "stackedAreaChart",
    "lineChart",
    "columnChart",
    "clusteredColumnChart",
    "barChart",
    "clusteredBarChart",
}
BOTTOM_TYPES = {"tableEx", "matrix", "scatterChart"}


def load_json(path: Path):
    return json.loads(path.read_text(encoding="utf-8"))


def write_json(path: Path, payload: dict):
    path.write_text(json.dumps(payload, indent=2), encoding="utf-8")


def slugify(value: str) -> str:
    value = value.lower()
    value = re.sub(r"[^a-z0-9]+", " ", value)
    value = re.sub(r"\s+", " ", value).strip()
    return value


def tokenize(value: str) -> List[str]:
    return [token for token in slugify(value).split(" ") if token]


def jaccard(a: set, b: set) -> float:
    if not a or not b:
        return 0.0
    return len(a & b) / len(a | b)


def read_snapshot_names(snapshots_dir: Path) -> List[str]:
    if not snapshots_dir.exists():
        return []
    return [path.stem for path in snapshots_dir.glob("*.png")]


def match_snapshot(page_name: str, snapshot_names: List[str]) -> Optional[str]:
    page_tokens = set(tokenize(page_name))
    best = None
    best_score = 0.0
    for snap in snapshot_names:
        score = jaccard(page_tokens, set(tokenize(snap)))
        if score > best_score:
            best = snap
            best_score = score
    return best


def extract_text_runs(visual: dict) -> List[str]:
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


def collect_projections(query_state: dict) -> List[dict]:
    projections = []

    def walk(obj):
        if isinstance(obj, dict):
            if "projections" in obj and isinstance(obj["projections"], list):
                projections.extend(obj["projections"])
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
        return field["Measure"].get("Property") or ""

    if "Column" in field:
        return field["Column"].get("Property") or ""

    if "Aggregation" in field:
        agg = field["Aggregation"]
        func = agg.get("Function")
        inner = agg.get("Expression") or {}
        return f"{func} {format_field(inner)}".strip()

    if "Hierarchy" in field:
        return field["Hierarchy"].get("Name") or ""

    if "Level" in field:
        return field["Level"].get("Name") or ""

    return ""


def collect_fields(visual_def: dict, filter_config: dict) -> List[str]:
    query_state = visual_def.get("query", {}).get("queryState", {})
    projections = collect_projections(query_state)
    fields = []
    for proj in projections:
        field = proj.get("field")
        if field:
            name = format_field(field)
            if name:
                fields.append(name)

    for flt in filter_config.get("filters", []):
        field = flt.get("field")
        if field:
            name = format_field(field)
            if name:
                fields.append(name)
    return fields


def has_any(fields: List[str], needles: List[str]) -> bool:
    field_text = " ".join(fields).lower()
    return any(needle.lower() in field_text for needle in needles)


def recommend_visual_type(page_name: str, visual: dict) -> Tuple[str, str]:
    visual_type = visual["visual_type"]
    fields = visual["fields"]
    text = " ".join(visual["text"]).lower()
    page = page_name.lower()

    if visual_type in {"textbox", "slicer"}:
        return visual_type, "Keep layout control visual"

    if visual_type == "tableEx":
        if has_any(fields, ["Order Year", "Order Month", "Order Date"]) and has_any(
            fields, ["Sales"]
        ):
            return "matrix", "Date-based table -> matrix with heatmap"
        if has_any(fields, ["Product Name"]) and has_any(fields, ["Sales", "Profit"]):
            return "scatterChart", "Product sales vs profit -> scatter/dot plot"
        if has_any(fields, ["Customer Name"]) and has_any(fields, ["Sales"]):
            return "barChart", "Customer ranking -> bar chart"
        if len(fields) <= 1:
            return "card", "Single measure -> card"
        return "tableEx", "Keep as table"

    if visual_type == "lineChart" and "forecast" in page:
        return "lineChart", "Forecast stays line chart"

    if visual_type == "columnChart" and "performance" in page:
        return "clusteredBarChart", "Performance vs target -> clustered bar"

    if visual_type == "map":
        return "filledMap", "Map with filled regions"

    if visual_type == "areaChart" and "shipping" in page:
        return "stackedAreaChart", "Shipping trend -> stacked area"

    if "sales and profit" in text and visual_type != "scatterChart":
        return "scatterChart", "Sales vs profit -> scatter"

    return visual_type, "No change"


def recommend_section(visual_type: str) -> str:
    if visual_type in TOP_TYPES:
        return "top"
    if visual_type in BOTTOM_TYPES:
        return "bottom"
    return "middle"


def distribute_section(
    visuals: List[dict],
    start_y: float,
    end_y: float,
    gap: float = 8.0,
):
    if not visuals:
        return
    visuals.sort(key=lambda v: v["position"]["y"])
    available = max(end_y - start_y, 1.0)
    total_height = sum(v["position"]["height"] for v in visuals)
    total_gap = gap * (len(visuals) - 1)
    scale = 1.0
    if total_height + total_gap > available:
        scale = max((available - total_gap) / total_height, 0.2)

    y_cursor = start_y
    for visual in visuals:
        old_height = visual["position"]["height"]
        new_height = round(old_height * scale, 2)
        visual["position"]["y"] = round(y_cursor, 2)
        visual["position"]["height"] = new_height
        y_cursor += new_height + gap


def process_report(
    report_dir: Path,
    snapshots_dir: Path,
    dry_run: bool,
) -> Dict:
    pages_path = report_dir / "definition" / "pages" / "pages.json"
    pages_meta = load_json(pages_path)
    snapshots = read_snapshot_names(snapshots_dir)

    report_changes = {"pages": []}

    for page_id in pages_meta.get("pageOrder", []):
        page_path = report_dir / "definition" / "pages" / page_id
        page_json = load_json(page_path / "page.json")
        page_name = page_json.get("displayName", page_id)
        page_height = float(page_json.get("height", 720))

        visuals_dir = page_path / "visuals"
        if not visuals_dir.exists():
            continue

        visuals = []
        for visual_dir in visuals_dir.iterdir():
            visual_path = visual_dir / "visual.json"
            if not visual_path.exists():
                continue
            visual_json = load_json(visual_path)
            visual_def = visual_json.get("visual", {})
            filter_config = visual_json.get("filterConfig", {})
            fields = collect_fields(visual_def, filter_config)
            text = []
            if visual_def.get("visualType") == "textbox":
                text = extract_text_runs(visual_def)
            visuals.append(
                {
                    "path": visual_path,
                    "json": visual_json,
                    "visual": visual_def,
                    "position": visual_json.get("position", {}),
                    "visual_type": visual_def.get("visualType"),
                    "fields": fields,
                    "text": text,
                }
            )

        page_changes = {
            "page_id": page_id,
            "page_name": page_name,
            "snapshot_match": match_snapshot(page_name, snapshots),
            "visual_changes": [],
        }

        for visual in visuals:
            old_type = visual["visual_type"]
            new_type, reason = recommend_visual_type(page_name, visual)
            if new_type != old_type:
                visual["visual"]["visualType"] = new_type
            visual["recommended_type"] = new_type
            visual["change_reason"] = reason

            if new_type != old_type:
                page_changes["visual_changes"].append(
                    {
                        "visual_path": str(visual["path"]),
                        "old_type": old_type,
                        "new_type": new_type,
                        "reason": reason,
                    }
                )

        top, middle, bottom = [], [], []
        for visual in visuals:
            section = recommend_section(visual["recommended_type"])
            if section == "top":
                top.append(visual)
            elif section == "bottom":
                bottom.append(visual)
            else:
                middle.append(visual)

        top_end = page_height * 0.2
        mid_end = page_height * 0.65
        distribute_section(top, 0, top_end)
        distribute_section(middle, top_end + 4, mid_end)
        distribute_section(bottom, mid_end + 4, page_height)

        for visual in visuals:
            position = visual["position"]
            visual["json"]["position"] = position
            if not dry_run:
                write_json(visual["path"], visual["json"])

        if page_changes["visual_changes"]:
            report_changes["pages"].append(page_changes)

    return report_changes


def main():
    parser = argparse.ArgumentParser(
        description="Auto-convert PBIP visual types and reposition by snapshot/page."
    )
    parser.add_argument(
        "--pbip-report",
        default="SuperstorePBIP/Superstore.Report",
        help="PBIP report folder.",
    )
    parser.add_argument(
        "--snapshots-dir",
        default="tableau snapshots",
        help="Directory with Tableau snapshot images.",
    )
    parser.add_argument(
        "--out-report",
        default="out/pbip_visual_conversion_report.json",
        help="Output report JSON path.",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Do not modify visual.json files.",
    )
    args = parser.parse_args()

    report_dir = Path(args.pbip_report)
    snapshots_dir = Path(args.snapshots_dir)

    report = process_report(report_dir, snapshots_dir, args.dry_run)
    out_path = Path(args.out_report)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_path.write_text(json.dumps(report, indent=2), encoding="utf-8")
    print(f"Wrote {out_path}")


if __name__ == "__main__":
    main()
