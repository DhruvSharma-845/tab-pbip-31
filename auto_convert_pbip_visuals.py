import argparse
import copy
import json
import re
import shutil
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


def extract_title_text(visual: dict) -> str:
    objects = visual.get("objects", {})
    title = objects.get("title", [])
    for item in title:
        props = item.get("properties", {})
        text = props.get("text", {}).get("expr", {}).get("Literal", {}).get("Value")
        if text:
            return text.strip("'")
    return ""


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


def replace_entity(obj, old_entity: str, new_entity: str):
    if isinstance(obj, dict):
        for key, value in obj.items():
            if key == "SourceRef" and isinstance(value, dict):
                if value.get("Entity") == old_entity:
                    value["Entity"] = new_entity
            if key == "queryRef" and isinstance(value, str):
                if value.startswith(f"{old_entity}."):
                    obj[key] = value.replace(old_entity, new_entity, 1)
            else:
                replace_entity(value, old_entity, new_entity)
    elif isinstance(obj, list):
        for item in obj:
            replace_entity(item, old_entity, new_entity)


def to_scatter_query(visual: dict, entity: str, detail_property: str):
    query = visual.get("query", {})
    query_state = query.get("queryState", {})
    sales_field = {
        "field": {
            "Aggregation": {
                "Expression": {
                    "Column": {
                        "Expression": {"SourceRef": {"Entity": entity}},
                        "Property": "Sales",
                    }
                },
                "Function": 0,
            }
        },
        "queryRef": f"Sum({entity}.Sales)",
        "nativeQueryRef": "Sum of Sales",
    }
    profit_field = {
        "field": {
            "Aggregation": {
                "Expression": {
                    "Column": {
                        "Expression": {"SourceRef": {"Entity": entity}},
                        "Property": "Profit",
                    }
                },
                "Function": 0,
            }
        },
        "queryRef": f"Sum({entity}.Profit)",
        "nativeQueryRef": "Sum of Profit",
    }
    detail_field = {
        "field": {
            "Column": {
                "Expression": {"SourceRef": {"Entity": entity}},
                "Property": detail_property,
            }
        },
        "queryRef": f"{entity}.{detail_property}",
        "nativeQueryRef": detail_property,
    }
    legend_field = {
        "field": {
            "Column": {
                "Expression": {"SourceRef": {"Entity": entity}},
                "Property": "Segment",
            }
        },
        "queryRef": f"{entity}.Segment",
        "nativeQueryRef": "Segment",
    }
    query_state.clear()
    query_state.update(
        {
            "X": {"projections": [sales_field]},
            "Y": {"projections": [profit_field]},
            "Details": {"projections": [detail_field]},
            "Legend": {"projections": [legend_field]},
        }
    )
    query["queryState"] = query_state
    visual["query"] = query


def to_customer_rank_query(visual: dict, entity: str):
    query = visual.get("query", {})
    query_state = query.get("queryState", {})
    query_state.clear()
    query_state.update(
        {
            "Category": {
                "projections": [
                    {
                        "field": {
                            "Column": {
                                "Expression": {"SourceRef": {"Entity": entity}},
                                "Property": "Customer Name",
                            }
                        },
                        "queryRef": f"{entity}.Customer Name",
                        "nativeQueryRef": "Customer Name",
                        "active": True,
                    }
                ]
            },
            "Y": {
                "projections": [
                    {
                        "field": {
                            "Aggregation": {
                                "Expression": {
                                    "Column": {
                                        "Expression": {"SourceRef": {"Entity": entity}},
                                        "Property": "Sales",
                                    }
                                },
                                "Function": 0,
                            }
                        },
                        "queryRef": f"Sum({entity}.Sales)",
                        "nativeQueryRef": "Sum of Sales",
                    }
                ]
            },
        }
    )
    query["queryState"] = query_state
    visual["query"] = query


def lock_visual_type(visual: dict):
    if "autoSelectVisualType" in visual:
        visual["autoSelectVisualType"] = False
    else:
        visual["autoSelectVisualType"] = False


def ensure_small_multiples(query_state: dict, field_name: str):
    if "SmallMultiples" in query_state:
        return
    series = query_state.get("Series")
    if series and isinstance(series, dict):
        query_state["SmallMultiples"] = series
        query_state.pop("Series", None)
        return
    legend = query_state.get("Legend")
    if legend and isinstance(legend, dict):
        query_state["SmallMultiples"] = legend


def build_segment_value_filter(entity: str, value: str, name_suffix: str) -> dict:
    return {
        "name": f"seg_value_{name_suffix}",
        "displayName": "Segment",
        "field": {
            "Column": {
                "Expression": {"SourceRef": {"Entity": entity}},
                "Property": "Segment",
            }
        },
        "type": "Categorical",
        "filter": {
            "Version": 2,
            "From": [{"Name": entity, "Entity": entity}],
            "Where": [
                {
                    "Condition": {
                        "Comparison": {
                            "ComparisonKind": 0,
                            "Left": {
                                "Column": {
                                    "Expression": {"SourceRef": {"Entity": entity}},
                                    "Property": "Segment",
                                }
                            },
                            "Right": {"Literal": {"Value": value}},
                        }
                    }
                }
            ],
        },
    }


def split_segment_visuals(
    visuals_dir: Path,
    visuals: List[dict],
    page_height: float,
) -> List[dict]:
    segment_visual = None
    for visual in visuals:
        name = str(visual.get("json", {}).get("name", ""))
        fields_text = " ".join(visual.get("fields", [])).lower()
        if name.startswith("seg_"):
            segment_visual = visual
            break
        if (
            visual.get("recommended_type") in {"areaChart", "stackedAreaChart"}
            and "segment" in fields_text
            and "order month" in fields_text
        ):
            segment_visual = visual
            break
    if not segment_visual:
        return visuals

    # Clean any existing segment split visuals
    for visual in list(visuals):
        name = str(visual.get("json", {}).get("name", ""))
        if name.startswith("seg_"):
            seg_dir = Path(visual["path"]).parent
            if seg_dir.exists():
                shutil.rmtree(seg_dir)
            visuals.remove(visual)

    # Remove the original segment visual folder before cloning
    original_dir = Path(segment_visual["path"]).parent
    if original_dir.exists():
        shutil.rmtree(original_dir)

    base_json = segment_visual["json"]
    base_position = segment_visual["position"]
    x = base_position.get("x", 20)
    width = base_position.get("width", 600)
    y = 390
    total_height = max(page_height - 410, 300)

    segments = [
        ("Consumer", "consumer"),
        ("Corporate", "corporate"),
        ("Home Office", "home_office"),
    ]
    gap = 6
    each_height = max((total_height - gap * (len(segments) - 1)) / len(segments), 80)

    new_visuals = []
    for idx, (segment_label, suffix) in enumerate(segments):
        visual_json = copy.deepcopy(base_json)
        visual_json["name"] = f"seg_{suffix}"
        visual_json["position"]["x"] = x
        visual_json["position"]["y"] = round(y + idx * (each_height + gap), 2)
        visual_json["position"]["width"] = width
        visual_json["position"]["height"] = round(each_height, 2)
        visual_json["visual"]["visualType"] = "stackedAreaChart"
        visual_json["visual"]["autoSelectVisualType"] = False
        query_state = visual_json["visual"].get("query", {}).get("queryState", {})
        query_state.pop("SmallMultiples", None)
        visual_json["visual"]["query"]["queryState"] = query_state

        filter_config = visual_json.get("filterConfig", {})
        filters = [
            f
            for f in filter_config.get("filters", [])
            if not (
                f.get("field", {})
                .get("Column", {})
                .get("Property", "")
                .lower()
                == "segment"
            )
        ]
        filters.append(build_segment_value_filter("Orders", segment_label, suffix))
        filter_config["filters"] = filters
        visual_json["filterConfig"] = filter_config

        visual_dir = visuals_dir / visual_json["name"]
        visual_dir.mkdir(parents=True, exist_ok=True)
        (visual_dir / "visual.json").write_text(
            json.dumps(visual_json, indent=2), encoding="utf-8"
        )
        new_visuals.append(
            {
                "path": visual_dir / "visual.json",
                "json": visual_json,
                "visual": visual_json["visual"],
                "position": visual_json["position"],
                "visual_type": visual_json["visual"]["visualType"],
                "fields": segment_visual.get("fields", []),
                "text": segment_visual.get("text", []),
                "title": segment_visual.get("title", ""),
                "recommended_type": "stackedAreaChart",
            }
        )

    # Remove original segment visual from list and add new visuals
    visuals = [v for v in visuals if v is not segment_visual]
    visuals.extend(new_visuals)
    return visuals


def recommend_visual_type(page_name: str, visual: dict) -> Tuple[str, str]:
    visual_type = visual["visual_type"]
    fields = visual["fields"]
    text = " ".join(visual["text"]).lower()
    page = page_name.lower()
    title = visual.get("title", "").lower()

    if visual_type in {"textbox", "slicer"}:
        return visual_type, "Keep layout control visual"

    if visual_type == "tableEx":
        if "customer" in title and "scatter" in title:
            return "scatterChart", "Customer scatter -> scatter plot"
        if "customer" in title and "rank" in title:
            return "barChart", "Customer rank -> bar chart"
        if "customer" in title and "overview" in title:
            return "matrix", "Customer overview -> matrix"
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

    if visual_type == "areaChart" and (
        "shipping" in page or has_any(fields, ["Segment", "Category"])
    ):
        return "stackedAreaChart", "Stacked area with segment/category"

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


def layout_profile(page_name: str, snapshot_name: Optional[str]) -> Optional[str]:
    name = (snapshot_name or page_name).lower()
    if "overview" in name:
        return "overview"
    if "product" in name:
        return "product"
    if "customers" in name:
        return "customers"
    if "order details" in name:
        return "order_details"
    if "shipping" in name:
        return "shipping"
    if "commission" in name:
        return "commission"
    if "performance" in name:
        return "performance"
    if "forecast" in name and "what if" in name:
        return "what_if_forecast"
    if "forecast" in name:
        return "forecast"
    return None


def apply_layout_overrides(profile: str, visuals: List[dict], page_height: float):
    if profile == "overview":
        title = [v for v in visuals if v["visual_type"] == "textbox"]
        cards = [v for v in visuals if v["recommended_type"] == "card"]
        maps = [
            v
            for v in visuals
            if v["recommended_type"] in {"map", "filledMap", "shapeMap"}
        ]
        areas = [
            v
            for v in visuals
            if v["recommended_type"] in {"areaChart", "stackedAreaChart"}
        ]
        segmented = [v for v in areas if str(v.get("json", {}).get("name", "")).startswith("seg_")]
        for visual in title:
            visual["position"]["y"] = 0
            visual["position"]["height"] = 40
            visual["position"]["x"] = 0
            visual["position"]["width"] = 1280
        for visual in cards:
            visual["position"]["y"] = 40
            visual["position"]["height"] = 90
        for visual in maps:
            visual["position"]["y"] = 140
            visual["position"]["height"] = 240
            visual["position"]["x"] = 20
            visual["position"]["width"] = 1240
            visual["visual"]["autoSelectVisualType"] = False
        for visual in areas:
            if visual in segmented:
                continue
            visual["position"]["y"] = 390
            visual["position"]["height"] = max(page_height - 410, 300)
            visual["visual"]["autoSelectVisualType"] = False
        return

    if profile == "product":
        for visual in visuals:
            title = visual.get("title", "").lower()
            fields = " ".join(visual.get("fields", [])).lower()
            if "sales by product category" in title or (
                visual["recommended_type"] == "matrix"
                and "order month" in fields
                and "order year" in fields
                and "category" in fields
            ):
                visual["position"]["x"] = 20
                visual["position"]["y"] = 50
                visual["position"]["width"] = 1240
                visual["position"]["height"] = 280
                visual["visual"]["autoSelectVisualType"] = False
            if "sales and profit" in title or visual["recommended_type"] == "scatterChart":
                visual["position"]["x"] = 20
                visual["position"]["y"] = 350
                visual["position"]["width"] = 1240
                visual["position"]["height"] = max(page_height - 370, 300)
                visual["visual"]["autoSelectVisualType"] = False
            if visual["visual_type"] == "textbox":
                visual["position"]["y"] = 0
                visual["position"]["height"] = 40
        return

    if profile == "customers":
        for visual in visuals:
            title = visual.get("title", "").lower()
            if "overview" in title:
                visual["position"]["y"] = 40
                visual["position"]["height"] = 120
            if "scatter" in title:
                visual["position"]["x"] = 20
                visual["position"]["y"] = 200
                visual["position"]["width"] = 600
                visual["position"]["height"] = max(page_height - 220, 400)
            if "rank" in title:
                visual["position"]["x"] = 640
                visual["position"]["y"] = 200
                visual["position"]["width"] = 620
                visual["position"]["height"] = max(page_height - 220, 400)
        return

    if profile == "order_details":
        slicers = [v for v in visuals if v["visual_type"] == "slicer"]
        tables = [v for v in visuals if v["visual_type"] in {"tableEx", "matrix"}]
        for visual in visuals:
            if visual["visual_type"] == "textbox":
                visual["position"]["y"] = 0
                visual["position"]["height"] = 40
        if slicers:
            slicers.sort(key=lambda v: v["position"]["x"])
            x_positions = [20, 240, 460, 680, 900, 1120]
            for idx, visual in enumerate(slicers):
                visual["position"]["y"] = 40
                visual["position"]["height"] = 40
                visual["position"]["width"] = 180
                visual["position"]["x"] = x_positions[idx % len(x_positions)]
        for visual in tables:
            visual["position"]["x"] = 20
            visual["position"]["y"] = 90
            visual["position"]["width"] = 1240
            visual["position"]["height"] = max(page_height - 110, 520)
            visual["visual"]["autoSelectVisualType"] = False
        return


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
            title = extract_title_text(visual_def)
            visuals.append(
                {
                    "path": visual_path,
                    "json": visual_json,
                    "visual": visual_def,
                    "position": visual_json.get("position", {}),
                    "visual_type": visual_def.get("visualType"),
                    "fields": fields,
                    "text": text,
                    "title": title,
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
                visual["visual"]["autoSelectVisualType"] = False
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

        if page_name.lower() == "order details":
            for visual in visuals:
                replace_entity(visual["json"], "Sales Target", "Sample - Superstore")

        if page_name.lower() == "product":
            for visual in visuals:
                if "sales and profit" in visual.get("title", "").lower():
                    to_scatter_query(visual["visual"], "Orders", "Product Name")
                    visual["visual"]["visualType"] = "scatterChart"
                    visual["visual"]["autoSelectVisualType"] = False
                    visual["recommended_type"] = "scatterChart"
                if "sales by product category" in visual.get("title", "").lower():
                    visual["visual"]["visualType"] = "matrix"
                    visual["visual"]["autoSelectVisualType"] = False
                    visual["recommended_type"] = "matrix"

        if page_name.lower() == "customers":
            for visual in visuals:
                title = visual.get("title", "").lower()
                if "scatter" in title:
                    to_scatter_query(visual["visual"], "Sample - Superstore", "Customer Name")
                    visual["visual"]["visualType"] = "scatterChart"
                    visual["visual"]["autoSelectVisualType"] = False
                    visual["recommended_type"] = "scatterChart"
                if "rank" in title:
                    to_customer_rank_query(visual["visual"], "Sample - Superstore")
                    visual["visual"]["visualType"] = "barChart"
                    visual["visual"]["autoSelectVisualType"] = False
                    visual["recommended_type"] = "barChart"

        if page_name.lower() == "overview":
            for visual in visuals:
                if visual["recommended_type"] not in {"areaChart", "stackedAreaChart"}:
                    continue
                fields_text = " ".join(visual.get("fields", [])).lower()
                query_state = visual["visual"].get("query", {}).get("queryState", {})
                if "segment" in fields_text:
                    ensure_small_multiples(query_state, "Segment")
                elif "category" in fields_text:
                    ensure_small_multiples(query_state, "Category")
                visual["visual"]["query"]["queryState"] = query_state
                lock_visual_type(visual["visual"])
            visuals = split_segment_visuals(visuals_dir, visuals, page_height)

        top, middle, bottom = [], [], []
        for visual in visuals:
            section = recommend_section(visual["recommended_type"])
            if section == "top":
                top.append(visual)
            elif section == "bottom":
                bottom.append(visual)
            else:
                middle.append(visual)

        snapshot_match = page_changes["snapshot_match"]
        profile = layout_profile(page_name, snapshot_match)
        if profile:
            apply_layout_overrides(profile, visuals, page_height)
            if profile in {"overview", "product", "order_details"}:
                for visual in visuals:
                    if visual["visual_type"] in {"textbox", "slicer", "card"}:
                        continue
                    lock_visual_type(visual["visual"])
        else:
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
