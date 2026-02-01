#!/usr/bin/env python3
"""
Generic TWBX -> PBIP (PBIR preview) converter.

Features:
- Extracts TWB and packaged data files from TWBX
- Builds TMDL (tables split into definition/tables)
- Creates PBIR pages per Tableau dashboard
- Adds an Overview page with KPI cards, monthly charts, and map
- Uses Windows absolute paths for File.Contents (required by Desktop)

Usage example:
python3 convert_twbx_to_pbip_auto.py \
  --twbx /path/to/workbook.twbx \
  --out /path/to/output/SuperstorePBIP \
  --project-name Superstore \
  --sample-theme /path/to/CY25SU12.json \
  --windows-data-root "C:\\Users\\Name\\Downloads\\Project\\SuperstorePBIP\\Data\\Superstore"
"""
import argparse
import uuid
import hashlib
import json
import os
import re
import shutil
import xml.etree.ElementTree as ET
from collections import defaultdict
from pathlib import Path


def safe_name(name: str) -> str:
    return (name or "").strip()


def needs_quotes(name: str) -> bool:
    return bool(re.search(r"[^A-Za-z0-9_]", name))


def quote(name: str) -> str:
    return f"'{name}'" if needs_quotes(name) else name


def map_type(datatype: str) -> str:
    if not datatype:
        return "string"
    t = datatype.lower()
    if t in ("string", "str"):
        return "string"
    if t in ("integer", "int"):
        return "int64"
    if t in ("real", "float", "double"):
        return "double"
    if t in ("date", "datetime"):
        return "dateTime"
    if t in ("boolean", "bool"):
        return "boolean"
    return "string"


def extract_twb(twbx_path: str, out_dir: str) -> str:
    twb_path = os.path.join(out_dir, "workbook.twb")
    os.makedirs(out_dir, exist_ok=True)
    os.system(f"unzip -p \"{twbx_path}\" \"*.twb\" > \"{twb_path}\"")
    return twb_path


def extract_data_files(twbx_path: str, data_dir: str) -> None:
    os.makedirs(data_dir, exist_ok=True)
    os.system(f"unzip -o \"{twbx_path}\" \"Data/*\" -d \"{data_dir}\" >/dev/null 2>&1")


def parse_datasources(root):
    ds_caption = {}
    for ds in root.findall(".//datasource"):
        name = ds.get("name")
        caption = ds.get("caption")
        if name and caption:
            ds_caption[name] = caption
    return ds_caption


def parse_columns(root, ds_caption):
    col_meta = defaultdict(dict)
    for dep in root.findall(".//datasource-dependencies"):
        ds = dep.get("datasource")
        if not ds:
            continue
        table = safe_name(ds_caption.get(ds, ds))
        for col in dep.findall("column"):
            name = col.get("name")
            if not name:
                continue
            datatype = col.get("datatype")
            col_name = name.strip("[]")
            col_meta[table][col_name] = datatype
    return col_meta


def parse_worksheet_meta(root):
    meta = {}
    for ws in root.findall(".//worksheet"):
        name = ws.get("name")
        rows = (ws.findtext("table/rows") or "").strip()
        cols = (ws.findtext("table/cols") or "").strip()
        mark = None
        pane = ws.find("table/panes/pane")
        if pane is not None and pane.find("mark") is not None:
            mark = pane.find("mark").get("class")
        meta[name] = {"rows": rows, "cols": cols, "mark": mark}
    return meta


def parse_dashboard_filters(root):
    dashboards = defaultdict(list)
    for d in root.findall(".//dashboard"):
        dname = d.get("name")
        if not dname:
            continue
        for z in d.findall(".//zone[@type-v2='filter']"):
            param = z.get("param") or ""
            fields = extract_fields(param)
            if not fields:
                continue
            field = fields[-1]
            dashboards[dname].append({
                "field": field,
                "mode": z.get("mode"),
                "values": z.get("values"),
                "worksheet": z.get("name"),
                "x": int(z.get("x") or 0),
                "y": int(z.get("y") or 0),
                "w": int(z.get("w") or 0),
                "h": int(z.get("h") or 0)
            })
    return dashboards


def parse_dashboard_worksheet_zones(root, worksheet_names):
    dashboards = defaultdict(dict)
    for d in root.findall(".//dashboard"):
        dname = d.get("name")
        if not dname:
            continue
        for z in d.findall(".//zone"):
            ws_name = z.get("worksheet") or z.get("name")
            if not ws_name or ws_name not in worksheet_names:
                continue
            dashboards[dname][ws_name] = {
                "x": int(z.get("x") or 0),
                "y": int(z.get("y") or 0),
                "w": int(z.get("w") or 0),
                "h": int(z.get("h") or 0)
            }
    return dashboards


def parse_dashboard_root_sizes(root):
    sizes = {}
    for d in root.findall(".//dashboard"):
        dname = d.get("name")
        if not dname:
            continue
        best = None
        for z in d.findall(".//zone[@type-v2='layout-basic']"):
            x = int(z.get("x") or 0)
            y = int(z.get("y") or 0)
            w = int(z.get("w") or 0)
            h = int(z.get("h") or 0)
            if x == 0 and y == 0:
                if not best or (w * h) > (best["w"] * best["h"]):
                    best = {"w": w, "h": h}
        if best:
            sizes[dname] = best
    return sizes


def normalize_snapshot_key(name):
    return re.sub(r"[^a-z0-9]+", "", (name or "").lower())


def get_png_dimensions(path):
    with open(path, "rb") as f:
        header = f.read(24)
        if len(header) < 24 or header[:8] != b"\x89PNG\r\n\x1a\n":
            return None
        width = int.from_bytes(header[16:20], "big")
        height = int.from_bytes(header[20:24], "big")
        return width, height


def find_snapshot_for_dashboard(dashboard_name, snapshots_dir):
    if not snapshots_dir or not os.path.isdir(snapshots_dir):
        return None
    target = normalize_snapshot_key(dashboard_name)
    for entry in os.listdir(snapshots_dir):
        if not entry.lower().endswith(".png"):
            continue
        if normalize_snapshot_key(os.path.splitext(entry)[0]) == target:
            return os.path.join(snapshots_dir, entry)
    return None


def get_page_size_for_dashboard(dashboard_name, snapshot_dims, default_width=1280, default_height=720):
    snap_dim = snapshot_dims.get(dashboard_name)
    if snap_dim:
        snap_w, snap_h = snap_dim
        page_width = default_width
        page_height = int(round(page_width * (snap_h / max(snap_w, 1))))
        page_height = max(600, min(page_height, 2000))
        return page_width, page_height
    return default_width, default_height


def scale_rect(rect, root_size, page_size):
    if not rect or not root_size:
        return None
    root_w = root_size.get("w") or 0
    root_h = root_size.get("h") or 0
    if root_w <= 0 or root_h <= 0:
        return None
    page_w, page_h = page_size
    scale_x = page_w / root_w
    scale_y = page_h / root_h
    return {
        "x": int(round(rect["x"] * scale_x)),
        "y": int(round(rect["y"] * scale_y)),
        "w": int(round(rect["w"] * scale_x)),
        "h": int(round(rect["h"] * scale_y))
    }


def parse_calculations(root):
    calc_meta = defaultdict(dict)
    for ws in root.findall(".//worksheet"):
        ws_name = ws.get("name")
        if not ws_name:
            continue
        for col in ws.findall(".//column[calculation]"):
            name = (col.get("name") or "").strip("[]")
            caption = col.get("caption") or name
            calc = col.find("calculation")
            formula = calc.get("formula") if calc is not None else None
            if name and formula:
                calc_meta[ws_name][name] = {"caption": caption, "formula": formula}
    return calc_meta


def parse_measure_names_filters(root, calc_captions):
    ws_measures = {}
    for ws in root.findall(".//worksheet"):
        ws_name = ws.get("name")
        if not ws_name:
            continue
        measures = []
        for flt in ws.findall(".//filter"):
            col = flt.get("column") or ""
            if "Measure Names" not in col:
                continue
            for member in flt.findall(".//groupfilter[@function='member']"):
                member_val = member.get("member") or ""
                fields = extract_fields(member_val)
                for field in fields:
                    field_name = calc_captions.get(field, field)
                    if field_name not in measures:
                        measures.append(field_name)
        ws_measures[ws_name] = measures
    return ws_measures


def is_table_worksheet(meta):
    rows = meta.get("rows", "")
    cols = meta.get("cols", "")
    return "Measure Names" in rows or "Measure Names" in cols or "Multiple Values" in rows or "Multiple Values" in cols


def dedupe_preserve(items):
    seen = set()
    out = []
    for item in items:
        if item in seen:
            continue
        seen.add(item)
        out.append(item)
    return out


def parse_dashboard_worksheets(root, worksheet_names):
    dashboards = {}
    for d in root.findall(".//dashboard"):
        dname = d.get("name")
        worksheets = []
        for z in d.findall(".//zone"):
            wname = z.get("worksheet") or z.get("name")
            if wname and wname in worksheet_names:
                worksheets.append(wname)
        dashboards[dname] = list(dict.fromkeys(worksheets))
    return dashboards


def map_mark_to_visual(mark: str) -> str:
    if mark in ("Bar",):
        return "columnChart"
    if mark in ("Line",):
        return "lineChart"
    if mark in ("Area",):
        return "areaChart"
    if mark in ("Pie",):
        return "pieChart"
    if mark in ("Square", "Circle"):
        return "scatterChart"
    if mark in ("Multipolygon",):
        return "map"
    return "tableEx"


def extract_fields(expr: str):
    # Extract Tableau field references like [none:Field:nk]
    fields = []
    if not expr:
        return fields
    for token in re.findall(r"\[([^\]]+)\]", expr):
        # token looks like federated.ds].[none:Field:nk
        if "none:" in token or "sum:" in token or "avg:" in token or "cnt:" in token:
            parts = token.split(":")
            if len(parts) >= 2:
                field = parts[1]
                if field not in ("Measure Names", "Multiple Values"):
                    fields.append(field)
    return fields


def choose_category_value(meta, default_category="Category", default_value="Sales"):
    row_fields = extract_fields(meta.get("rows", ""))
    col_fields = extract_fields(meta.get("cols", ""))
    fields = row_fields + col_fields
    category = default_category
    value = default_value
    # Prefer date for category if present
    for f in fields:
        if "Order Date" in f:
            category = "Order Month"
            break
    # Prefer dimension-like fields for category
    for f in fields:
        if f in ("Category", "Segment", "Region", "State/Province", "City", "Product Name"):
            category = f if f != "Order Date" else "Order Month"
            break
    # Prefer measures for value
    for f in fields:
        if f in ("Sales", "Profit", "Quantity", "Sales Target"):
            value = f
            break
    return category, value


def choose_table_for_fields(col_meta, category, value):
    for table, cols in col_meta.items():
        if category in cols and value in cols:
            return table
    for table, cols in col_meta.items():
        if value in cols:
            return table
    return "Orders"


def choose_table_for_field(col_meta, field):
    for table, cols in col_meta.items():
        if field in cols:
            return table
    return "Orders"


def choose_table_for_any(col_meta, fields):
    for table, cols in col_meta.items():
        for f in fields:
            if f in cols:
                return table
    return "Orders"


def choose_series(fields):
    for f in fields:
        if f in ("Segment", "Category", "Region"):
            return f
    return None


def determine_visual_type(meta, ws_name):
    mark = meta.get("mark")
    rows = meta.get("rows", "")
    cols = meta.get("cols", "")
    fields = extract_fields(rows) + extract_fields(cols)
    if "Forecast" in ws_name:
        return "lineChart"
    if any("Order Date" in f for f in fields):
        return "lineChart" if mark in (None, "Automatic", "Line") else map_mark_to_visual(mark)
    if any(f in ("Latitude (generated)", "Longitude (generated)", "State/Province", "City") for f in fields):
        return "map"
    return map_mark_to_visual(mark)

def build_table_files(tables_dir, col_meta, windows_data_root, order_calc_columns=None):
    tables_dir = os.path.abspath(tables_dir)
    os.makedirs(tables_dir, exist_ok=True)

    for table, cols in col_meta.items():
        tname = quote(table)
        out = []
        out.append(f"table {tname}")
        out.append("\tlineageTag: " + str(uuid.uuid4()))
        out.append("")
        for col, dtype in cols.items():
            cname = quote(col)
            summarize = "sum" if map_type(dtype) in ("int64", "double") else "none"
            out.append(f"\tcolumn {cname}")
            out.append(f"\t\tdataType: {map_type(dtype)}")
            out.append(f"\t\tsummarizeBy: {summarize}")
            out.append(f"\t\tsourceColumn: {col}")
            out.append("\t\tlineageTag: " + str(uuid.uuid4()))
            out.append("")

        # Add derived columns for monthly/legend if Orders table exists
        if table == "Orders":
            out.append("\tcolumn 'Order Month'")
            out.append("\t\tdataType: dateTime")
            out.append("\t\tsummarizeBy: none")
            out.append("\t\tsourceColumn: Order Month")
            out.append("\t\tlineageTag: " + str(uuid.uuid4()))
            out.append("")
            out.append("\tcolumn 'Order Year'")
            out.append("\t\tdataType: int64")
            out.append("\t\tsummarizeBy: none")
            out.append("\t\tsourceColumn: Order Year")
            out.append("\t\tlineageTag: " + str(uuid.uuid4()))
            out.append("")
            out.append("\tcolumn 'Profitability'")
            out.append("\t\tdataType: string")
            out.append("\t\tsummarizeBy: none")
            out.append("\t\tsourceColumn: Profitability")
            out.append("\t\tlineageTag: " + str(uuid.uuid4()))
            out.append("")
            if order_calc_columns:
                if "Days to Ship Actual" in order_calc_columns:
                    out.append("\tcolumn 'Days to Ship Actual'")
                    out.append("\t\tdataType: int64")
                    out.append("\t\tsummarizeBy: none")
                    out.append("\t\tsourceColumn: Days to Ship Actual")
                    out.append("\t\tlineageTag: " + str(uuid.uuid4()))
                    out.append("")
                if "Days to Ship Scheduled" in order_calc_columns:
                    out.append("\tcolumn 'Days to Ship Scheduled'")
                    out.append("\t\tdataType: int64")
                    out.append("\t\tsummarizeBy: none")
                    out.append("\t\tsourceColumn: Days to Ship Scheduled")
                    out.append("\t\tlineageTag: " + str(uuid.uuid4()))
                    out.append("")
            # Measures
            out.append("\tmeasure 'Total Sales' = SUM('Orders'[Sales])")
            out.append("\t\tformatString: \"$#,0\"")
            out.append("\tmeasure 'Total Profit' = SUM('Orders'[Profit])")
            out.append("\t\tformatString: \"$#,0\"")
            out.append("\tmeasure 'Profit Ratio' = DIVIDE(SUM('Orders'[Profit]), SUM('Orders'[Sales]))")
            out.append("\t\tformatString: \"0.0%\"")
            out.append("\tmeasure 'Profit per Order' = DIVIDE(SUM('Orders'[Profit]), DISTINCTCOUNT('Orders'[Order ID]))")
            out.append("\t\tformatString: \"$#,0.00\"")
            out.append("\tmeasure 'Sales per Customer' = DIVIDE(SUM('Orders'[Sales]), DISTINCTCOUNT('Orders'[Customer Name]))")
            out.append("\t\tformatString: \"$#,0.00\"")
            out.append("\tmeasure 'Avg Discount' = AVERAGE('Orders'[Discount])")
            out.append("\t\tformatString: \"0.0%\"")

        # Partition (Superstore defaults)
        out.append("\tpartition FullData = m")
        out.append("\t\tmode: import")
        out.append("\t\tsource =")
        out.append("\t\t\tlet")
        if "Sales Target" in table:
            out.append(f"\t\t\t\tSource = Excel.Workbook(File.Contents(\"{windows_data_root}\\\\Sales Target.xlsx\"), null, true),")
            out.append("\t\t\t\tSheet1 = Source{[Item=\"Sheet1\",Kind=\"Sheet\"]}[Data],")
            out.append("\t\t\t\tPromoted = Table.PromoteHeaders(Sheet1, [PromoteAllScalars=true])")
            out.append("\t\t\tin")
            out.append("\t\t\t\tPromoted")
        elif "Sales Commission" in table:
            out.append(f"\t\t\t\tSource = Csv.Document(File.Contents(\"{windows_data_root}\\\\Sales Commission.csv\"), [Delimiter=\",\", Columns=4, Encoding=65001, QuoteStyle=QuoteStyle.None]),")
            out.append("\t\t\t\tPromoted = Table.PromoteHeaders(Source, [PromoteAllScalars=true])")
            out.append("\t\t\tin")
            out.append("\t\t\t\tPromoted")
        else:
            m_steps = [
                ("Source", f"Excel.Workbook(File.Contents(\"{windows_data_root}\\\\Sample - Superstore.xls\"), null, true)"),
                ("OrdersTable", "try Source{[Item=\"Orders\",Kind=\"Table\"]}[Data] otherwise null"),
                ("OrdersSheet", "try Source{[Item=\"Orders\",Kind=\"Sheet\"]}[Data] otherwise null"),
                ("Selected", "if OrdersTable <> null then OrdersTable else if OrdersSheet <> null then OrdersSheet else Source{0}[Data]"),
                ("Promoted", "Table.PromoteHeaders(Selected, [PromoteAllScalars=true])"),
                ("ChangedType", "Table.TransformColumnTypes(Promoted, {{\"Order Date\", type date}, {\"Ship Date\", type date}, {\"Sales\", type number}, {\"Profit\", type number}, {\"Discount\", type number}, {\"Quantity\", Int64.Type}})"),
                ("CleanNumbers", "Table.TransformColumns(ChangedType, {{\"Sales\", each try Number.From(_) otherwise null, type number}, {\"Profit\", each try Number.From(_) otherwise null, type number}, {\"Discount\", each try Number.From(_) otherwise null, type number}, {\"Quantity\", each try Number.From(_) otherwise null, Int64.Type}})"),
                ("AddedMonth", "Table.AddColumn(CleanNumbers, \"Order Month\", each Date.StartOfMonth([Order Date]), type date)"),
                ("AddedProfitability", "Table.AddColumn(AddedMonth, \"Profitability\", each if [Profit] >= 0 then \"Profitable\" else \"Unprofitable\", type text)"),
                ("AddedYear", "Table.AddColumn(AddedProfitability, \"Order Year\", each Date.Year([Order Date]), Int64.Type)")
            ]
            if order_calc_columns:
                if "Days to Ship Actual" in order_calc_columns:
                    m_steps.append((
                        "AddedDaysToShipActual",
                        "Table.AddColumn(AddedYear, \"Days to Ship Actual\", each Duration.Days([Ship Date] - [Order Date]), Int64.Type)"
                    ))
                if "Days to Ship Scheduled" in order_calc_columns:
                    source_step = "AddedDaysToShipActual" if "Days to Ship Actual" in order_calc_columns else "AddedYear"
                    m_steps.append((
                        "AddedDaysToShipScheduled",
                        "Table.AddColumn("
                        + source_step
                        + ", \"Days to Ship Scheduled\", each if [Ship Mode] = \"Same Day\" then 0 else if [Ship Mode] = \"First Class\" then 1 else if [Ship Mode] = \"Second Class\" then 3 else if [Ship Mode] = \"Standard Class\" then 6 else null, Int64.Type)"
                    ))
            for idx, (step, expr) in enumerate(m_steps):
                suffix = "," if idx < len(m_steps) - 1 else ""
                out.append(f"\t\t\t\t{step} = {expr}{suffix}")
            out.append("\t\t\tin")
            out.append(f"\t\t\t\t{m_steps[-1][0]}")
        out.append("\tannotation PBI_ResultType = Table")
        out.append("")

        (Path(tables_dir) / f"{table}.tmdl").write_text("\n".join(out) + "\n")


def normalize_lineage_indentation(tables_dir: str) -> None:
    tables_path = Path(tables_dir)
    for path in tables_path.glob("*.tmdl"):
        lines = path.read_text().splitlines()
        new = []
        in_column = False
        for line in lines:
            stripped = line.strip()
            if stripped.startswith("column "):
                in_column = True
                new.append("\t" + stripped)
                continue
            if stripped.startswith("lineageTag:"):
                new.append(("\t\t" if in_column else "\t") + stripped)
                in_column = False
                continue
            new.append(line)
        path.write_text("\n".join(new) + "\n")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--twbx", required=True, help="Path to .twbx file")
    parser.add_argument("--out", required=True, help="Output PBIP folder")
    parser.add_argument("--project-name", default="Superstore", help="Base name for report/model folders")
    parser.add_argument("--sample-theme", required=True, help="Path to a base theme json to copy")
    parser.add_argument("--windows-data-root", required=True, help="Windows absolute path to Data/Superstore")
    args = parser.parse_args()

    out_root = os.path.abspath(args.out)
    report_name = f"{args.project_name}.Report"
    model_name = f"{args.project_name}.SemanticModel"

    os.makedirs(out_root, exist_ok=True)
    twb_path = extract_twb(args.twbx, out_root)
    extract_data_files(args.twbx, os.path.join(out_root, "Data"))

    root = ET.parse(twb_path).getroot()
    ds_caption = parse_datasources(root)
    col_meta = parse_columns(root, ds_caption)
    ws_meta = parse_worksheet_meta(root)
    worksheet_names = set(ws_meta.keys())
    dash_ws = parse_dashboard_worksheets(root, worksheet_names)
    dash_filters = parse_dashboard_filters(root)
    dash_worksheet_zones = parse_dashboard_worksheet_zones(root, worksheet_names)
    dash_root_sizes = parse_dashboard_root_sizes(root)
    calc_meta = parse_calculations(root)
    calc_captions = {}
    for ws_name, calcs in calc_meta.items():
        for calc_name, calc_info in calcs.items():
            calc_captions[calc_name] = calc_info.get("caption", calc_name)
    ws_measure_names = parse_measure_names_filters(root, calc_captions)
    order_calc_columns = set()
    for calcs in calc_meta.values():
        for calc in calcs.values():
            caption = calc.get("caption")
            if caption in ("Days to Ship Actual", "Days to Ship Scheduled"):
                order_calc_columns.add(caption)

    # Report structure
    pages_dir = os.path.join(out_root, report_name, "definition", "pages")
    os.makedirs(pages_dir, exist_ok=True)
    os.makedirs(os.path.join(out_root, report_name, "StaticResources", "SharedResources", "BaseThemes"), exist_ok=True)
    shutil.copy2(args.sample_theme,
                 os.path.join(out_root, report_name, "StaticResources", "SharedResources", "BaseThemes",
                              os.path.basename(args.sample_theme)))

    report_json = {
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/report/3.1.0/schema.json",
        "themeCollection": {
            "baseTheme": {
                "name": os.path.splitext(os.path.basename(args.sample_theme))[0],
                "reportVersionAtImport": {"visual": "2.5.0", "report": "3.1.0", "page": "2.3.0"},
                "type": "SharedResources"
            }
        },
        "objects": {
            "section": [{"properties": {"verticalAlignment": {"expr": {"Literal": {"Value": "'Top'"}}}}}]
        },
        "resourcePackages": [{
            "name": "SharedResources",
            "type": "SharedResources",
            "items": [{"name": os.path.splitext(os.path.basename(args.sample_theme))[0],
                       "path": f"BaseThemes/{os.path.basename(args.sample_theme)}",
                       "type": "BaseTheme"}]
        }],
        "settings": {
            "useStylableVisualContainerHeader": True,
            "exportDataMode": "AllowSummarized",
            "defaultDrillFilterOtherVisuals": True,
            "allowChangeFilterTypes": True,
            "useEnhancedTooltips": True,
            "useDefaultAggregateDisplayName": True
        }
    }
    with open(os.path.join(out_root, report_name, "definition", "report.json"), "w") as f:
        json.dump(report_json, f, indent=2)

    with open(os.path.join(out_root, report_name, "definition", "version.json"), "w") as f:
        json.dump({
            "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/versionMetadata/1.0.0/schema.json",
            "version": "2.0.0"
        }, f, indent=2)

    with open(os.path.join(out_root, report_name, "definition.pbir"), "w") as f:
        json.dump({"version": "4.0", "datasetReference": {"byPath": {"path": f"../{model_name}"}}}, f, indent=2)

    # Pages from Tableau dashboards
    dashboard_names = [d.get("name") for d in root.findall(".//dashboard") if d.get("name")]
    if not dashboard_names:
        dashboard_names = ["Overview"]
    page_ids = []
    page_by_name = {}
    snapshots_dir = os.path.join(os.path.dirname(args.twbx), "tableau snapshots")
    snapshot_dims = {}
    if os.path.isdir(snapshots_dir):
        for name in dashboard_names:
            snap_path = find_snapshot_for_dashboard(name, snapshots_dir)
            if snap_path:
                dims = get_png_dimensions(snap_path)
                if dims:
                    snapshot_dims[name] = dims
    for name in dashboard_names:
        page_id = hashlib.sha1(name.encode("utf-8")).hexdigest()[:16]
        page_ids.append(page_id)
        page_by_name[name] = page_id
        page_folder = os.path.join(pages_dir, page_id)
        os.makedirs(page_folder, exist_ok=True)
        page_width, page_height = get_page_size_for_dashboard(name, snapshot_dims)
        page_json = {
            "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/page/2.0.0/schema.json",
            "name": page_id,
            "displayName": name,
            "displayOption": "FitToPage",
            "height": page_height,
            "width": page_width
        }
        with open(os.path.join(page_folder, "page.json"), "w") as f:
            json.dump(page_json, f, indent=2)
    # Add standalone worksheet pages (worksheets not on dashboards)
    all_ws = set(ws_meta.keys())
    ws_in_dash = set()
    for ws_list in dash_ws.values():
        ws_in_dash.update(ws_list)
    standalone_ws = sorted(all_ws - ws_in_dash)
    for ws_name in standalone_ws:
        page_id = hashlib.sha1(ws_name.encode("utf-8")).hexdigest()[:16]
        if page_id in page_ids:
            continue
        page_ids.append(page_id)
        page_by_name[ws_name] = page_id
        page_folder = os.path.join(pages_dir, page_id)
        os.makedirs(page_folder, exist_ok=True)
        page_json = {
            "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/page/2.0.0/schema.json",
            "name": page_id,
            "displayName": ws_name,
            "displayOption": "FitToPage",
            "height": 720,
            "width": 1280
        }
        with open(os.path.join(page_folder, "page.json"), "w") as f:
            json.dump(page_json, f, indent=2)

    with open(os.path.join(pages_dir, "pages.json"), "w") as f:
        json.dump({
            "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/pagesMetadata/1.0.0/schema.json",
            "pageOrder": page_ids,
            "activePageName": page_by_name.get("Overview", page_ids[0])
        }, f, indent=2)

    # Overview visuals (Superstore mapping)
    overview_id = page_by_name.get("Overview", page_ids[0])
    overview_folder = os.path.join(pages_dir, overview_id)
    visuals_dir = os.path.join(overview_folder, "visuals")
    os.makedirs(visuals_dir, exist_ok=True)

    def write_visual(vid, payload):
        vdir = os.path.join(visuals_dir, vid)
        os.makedirs(vdir, exist_ok=True)
        with open(os.path.join(vdir, "visual.json"), "w") as f:
            json.dump(payload, f, indent=2)

    # Title textbox (optional)
    write_visual("overview_title", {
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/visualContainer/2.5.0/schema.json",
        "name": "overview_title",
        "position": {"x": 0, "y": 0, "z": 4, "height": 40, "width": 1280, "tabOrder": 0},
        "visual": {
            "visualType": "textbox",
            "objects": {
                "general": [{
                    "properties": {
                        "paragraphs": [{
                            "textRuns": [{
                                "value": "Executive Overview - Profitability (All)",
                                "textStyle": {"fontWeight": "bold", "fontSize": "20pt"}
                            }]
                        }]
                    }
                }]
            },
            "drillFilterOtherVisuals": True
        }
    })

    # KPI cards
    card_specs = [
        ("card_sales", "Total Sales", 20),
        ("card_profit", "Total Profit", 230),
        ("card_ratio", "Profit Ratio", 440),
        ("card_profit_order", "Profit per Order", 650),
        ("card_sales_customer", "Sales per Customer", 860),
        ("card_discount", "Avg Discount", 1070),
    ]
    for vid, measure, x in card_specs:
        write_visual(vid, {
            "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/visualContainer/2.5.0/schema.json",
            "name": vid,
            "position": {"x": x, "y": 45, "z": 0, "height": 80, "width": 200, "tabOrder": 1},
            "visual": {
                "visualType": "card",
                "query": {"queryState": {"Values": {"projections": [{
                    "field": {"Measure": {"Expression": {"SourceRef": {"Entity": "Orders"}}, "Property": measure}},
                    "queryRef": f"Orders.{measure}",
                    "nativeQueryRef": measure
                }]}}},
                "drillFilterOtherVisuals": True,
                "autoSelectVisualType": True
            }
        })

    # Monthly sales by segment (area)
    write_visual("monthly_segment", {
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/visualContainer/2.5.0/schema.json",
        "name": "monthly_segment",
        "position": {"x": 20, "y": 400, "z": 2, "height": 300, "width": 610, "tabOrder": 20},
        "visual": {
            "visualType": "areaChart",
            "query": {"queryState": {
                "Category": {"projections": [{
                    "field": {"Column": {"Expression": {"SourceRef": {"Entity": "Orders"}}, "Property": "Order Month"}},
                    "queryRef": "Orders.Order Month",
                    "nativeQueryRef": "Order Month",
                    "active": True
                }]},
                "Series": {"projections": [{
                    "field": {"Column": {"Expression": {"SourceRef": {"Entity": "Orders"}}, "Property": "Segment"}},
                    "queryRef": "Orders.Segment",
                    "nativeQueryRef": "Segment"
                }]},
                "Legend": {"projections": [{
                    "field": {"Column": {"Expression": {"SourceRef": {"Entity": "Orders"}}, "Property": "Profitability"}},
                    "queryRef": "Orders.Profitability",
                    "nativeQueryRef": "Profitability"
                }]},
                "Y": {"projections": [{
                    "field": {"Aggregation": {"Expression": {"Column": {"Expression": {"SourceRef": {"Entity": "Orders"}}, "Property": "Sales"}}, "Function": 0}},
                    "queryRef": "Sum(Orders.Sales)",
                    "nativeQueryRef": "Sum of Sales"
                }]}
            }},
            "drillFilterOtherVisuals": True,
            "autoSelectVisualType": True
        }
    })

    # Monthly sales by product category (area)
    write_visual("monthly_category", {
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/visualContainer/2.5.0/schema.json",
        "name": "monthly_category",
        "position": {"x": 650, "y": 400, "z": 3, "height": 300, "width": 610, "tabOrder": 21},
        "visual": {
            "visualType": "areaChart",
            "query": {"queryState": {
                "Category": {"projections": [{
                    "field": {"Column": {"Expression": {"SourceRef": {"Entity": "Orders"}}, "Property": "Order Month"}},
                    "queryRef": "Orders.Order Month",
                    "nativeQueryRef": "Order Month",
                    "active": True
                }]},
                "Series": {"projections": [{
                    "field": {"Column": {"Expression": {"SourceRef": {"Entity": "Orders"}}, "Property": "Category"}},
                    "queryRef": "Orders.Category",
                    "nativeQueryRef": "Category"
                }]},
                "Legend": {"projections": [{
                    "field": {"Column": {"Expression": {"SourceRef": {"Entity": "Orders"}}, "Property": "Profitability"}},
                    "queryRef": "Orders.Profitability",
                    "nativeQueryRef": "Profitability"
                }]},
                "Y": {"projections": [{
                    "field": {"Aggregation": {"Expression": {"Column": {"Expression": {"SourceRef": {"Entity": "Orders"}}, "Property": "Sales"}}, "Function": 0}},
                    "queryRef": "Sum(Orders.Sales)",
                    "nativeQueryRef": "Sum of Sales"
                }]}
            }},
            "drillFilterOtherVisuals": True,
            "autoSelectVisualType": True
        }
    })

    # Profit ratio by city (map)
    write_visual("profit_ratio_city", {
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/visualContainer/2.5.0/schema.json",
        "name": "profit_ratio_city",
        "position": {"x": 20, "y": 130, "z": 1, "height": 260, "width": 1240, "tabOrder": 10},
        "visual": {
            "visualType": "map",
            "query": {"queryState": {
                "Location": {"projections": [{
                    "field": {"Column": {"Expression": {"SourceRef": {"Entity": "Orders"}}, "Property": "City"}},
                    "queryRef": "Orders.City",
                    "nativeQueryRef": "City",
                    "active": True
                }]},
                "Size": {"projections": [{
                    "field": {"Measure": {"Expression": {"SourceRef": {"Entity": "Orders"}}, "Property": "Profit Ratio"}},
                    "queryRef": "Orders.Profit Ratio",
                    "nativeQueryRef": "Profit Ratio"
                }]}
            }},
            "drillFilterOtherVisuals": True,
            "autoSelectVisualType": True
        }
    })

    # Product page visuals
    product_id = page_by_name.get("Product")
    if product_id:
        product_folder = os.path.join(pages_dir, product_id)
        product_visuals = os.path.join(product_folder, "visuals")
        os.makedirs(product_visuals, exist_ok=True)

        def write_product(vid, payload):
            vdir = os.path.join(product_visuals, vid)
            os.makedirs(vdir, exist_ok=True)
            with open(os.path.join(vdir, "visual.json"), "w") as f:
                json.dump(payload, f, indent=2)

        # Title textbox
        write_product("product_title", {
            "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/visualContainer/2.5.0/schema.json",
            "name": "product_title",
            "position": {"x": 0, "y": 0, "z": 4, "height": 40, "width": 1280, "tabOrder": 0},
            "visual": {
                "visualType": "textbox",
                "objects": {
                    "general": [{
                        "properties": {
                            "paragraphs": [{
                                "textRuns": [{
                                    "value": "Product Drilldown",
                                    "textStyle": {"fontWeight": "bold", "fontSize": "20pt"}
                                }]
                            }]
                        }
                    }]
                },
                "drillFilterOtherVisuals": True
            }
        })

        # Sales by Product Category (matrix heatmap style)
        write_product("product_heatmap", {
            "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/visualContainer/2.5.0/schema.json",
            "name": "product_heatmap",
            "position": {"x": 20, "y": 50, "z": 1, "height": 280, "width": 1240, "tabOrder": 1},
            "visual": {
                "visualType": "matrix",
                "query": {"queryState": {
                    "Rows": {"projections": [
                        {"field": {"Column": {"Expression": {"SourceRef": {"Entity": "Orders"}}, "Property": "Category"}},
                         "queryRef": "Orders.Category", "nativeQueryRef": "Category"},
                        {"field": {"Column": {"Expression": {"SourceRef": {"Entity": "Orders"}}, "Property": "Order Year"}},
                         "queryRef": "Orders.Order Year", "nativeQueryRef": "Order Year"}
                    ]},
                    "Columns": {"projections": [
                        {"field": {"Column": {"Expression": {"SourceRef": {"Entity": "Orders"}}, "Property": "Order Month"}},
                         "queryRef": "Orders.Order Month", "nativeQueryRef": "Order Month"}
                    ]},
                    "Values": {"projections": [
                        {"field": {"Aggregation": {"Expression": {"Column": {"Expression": {"SourceRef": {"Entity": "Orders"}}, "Property": "Sales"}}, "Function": 0}},
                         "queryRef": "Sum(Orders.Sales)", "nativeQueryRef": "Sum of Sales"}
                    ]}
                }},
                "drillFilterOtherVisuals": True,
                "autoSelectVisualType": True,
                "objects": {"title": [{"properties": {"text": {"expr": {"Literal": {"Value": "'Sales by Product Category'"}}}, "show": {"expr": {"Literal": {"Value": "true"}}}}}]}
            }
        })

        # Sales and Profit by Product Names (bar chart)
        write_product("product_sales_profit", {
            "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/visualContainer/2.5.0/schema.json",
            "name": "product_sales_profit",
            "position": {"x": 20, "y": 350, "z": 2, "height": 330, "width": 1240, "tabOrder": 2},
            "visual": {
                "visualType": "barChart",
                "query": {"queryState": {
                    "Category": {"projections": [{
                        "field": {"Column": {"Expression": {"SourceRef": {"Entity": "Orders"}}, "Property": "Product Name"}},
                        "queryRef": "Orders.Product Name", "nativeQueryRef": "Product Name", "active": True
                    }]},
                    "Legend": {"projections": [{
                        "field": {"Column": {"Expression": {"SourceRef": {"Entity": "Orders"}}, "Property": "Segment"}},
                        "queryRef": "Orders.Segment", "nativeQueryRef": "Segment"
                    }]},
                    "Y": {"projections": [
                        {"field": {"Aggregation": {"Expression": {"Column": {"Expression": {"SourceRef": {"Entity": "Orders"}}, "Property": "Sales"}}, "Function": 0}},
                         "queryRef": "Sum(Orders.Sales)", "nativeQueryRef": "Sum of Sales"},
                        {"field": {"Aggregation": {"Expression": {"Column": {"Expression": {"SourceRef": {"Entity": "Orders"}}, "Property": "Profit"}}, "Function": 0}},
                         "queryRef": "Sum(Orders.Profit)", "nativeQueryRef": "Sum of Profit"}
                    ]}
                }},
                "drillFilterOtherVisuals": True,
                "autoSelectVisualType": True,
                "objects": {"title": [{"properties": {"text": {"expr": {"Literal": {"Value": "'Sales and Profit by Product Names'"}}}, "show": {"expr": {"Literal": {"Value": "true"}}}}}]}
            }
        })

    # Generic visuals for remaining pages (best-effort)
    for dash_name, ws_list in dash_ws.items():
        if dash_name in ("Overview", "Product"):
            continue
        page_id = page_by_name.get(dash_name)
        if not page_id:
            continue
        page_folder = os.path.join(pages_dir, page_id)
        visuals_dir = os.path.join(page_folder, "visuals")
        os.makedirs(visuals_dir, exist_ok=True)

        def write_generic(vid, payload):
            vdir = os.path.join(visuals_dir, vid)
            os.makedirs(vdir, exist_ok=True)
            with open(os.path.join(vdir, "visual.json"), "w") as f:
                json.dump(payload, f, indent=2)

        # page title textbox
        title_vid = hashlib.sha1((dash_name + "_title").encode("utf-8")).hexdigest()[:16]
        write_generic(title_vid, {
            "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/visualContainer/2.5.0/schema.json",
            "name": title_vid,
            "position": {"x": 0, "y": 0, "z": 0, "height": 40, "width": 1280, "tabOrder": 0},
            "visual": {
                "visualType": "textbox",
                "objects": {
                    "general": [{
                        "properties": {
                            "paragraphs": [{
                                "textRuns": [{
                                    "value": dash_name,
                                    "textStyle": {"fontWeight": "bold", "fontSize": "18pt"}
                                }]
                            }]
                        }
                    }]
                },
                "drillFilterOtherVisuals": True
            }
        })

        page_size = get_page_size_for_dashboard(dash_name, snapshot_dims)
        root_size = dash_root_sizes.get(dash_name)
        filters = dash_filters.get(dash_name, [])
        filter_fields = dedupe_preserve([f["field"] for f in filters])
        filter_x = 20
        filter_y = 45
        filter_height = 60
        max_filter_bottom = 0
        for fidx, flt in enumerate(filters):
            field = flt["field"]
            table = choose_table_for_field(col_meta, field)
            rect = scale_rect(flt, root_size, page_size) if flt.get("w") else None
            if rect:
                x = rect["x"]
                y = rect["y"]
                width = max(rect["w"], 100)
                height = max(rect["h"], 40)
            else:
                x = filter_x
                y = filter_y
                width = 280 if "Date" in field else 160
                height = filter_height
            filter_vid = hashlib.sha1(f"{dash_name}:{field}:filter".encode("utf-8")).hexdigest()[:16]
            write_generic(filter_vid, {
                "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/visualContainer/2.5.0/schema.json",
                "name": filter_vid,
                "position": {"x": x, "y": y, "z": fidx + 1, "height": height, "width": width, "tabOrder": fidx + 1},
                "visual": {
                    "visualType": "slicer",
                    "query": {"queryState": {
                        "Values": {"projections": [{
                            "field": {"Column": {"Expression": {"SourceRef": {"Entity": table}}, "Property": field}},
                            "queryRef": f"{table}.{field}",
                            "nativeQueryRef": field,
                            "active": True
                        }]}
                    }},
                    "drillFilterOtherVisuals": True,
                    "autoSelectVisualType": True
                }
            })
            if rect:
                max_filter_bottom = max(max_filter_bottom, y + height)
            else:
                filter_x += width + 10

        content_start_y = max_filter_bottom + 20 if max_filter_bottom else (120 if filter_fields else 60)

        for idx, ws_name in enumerate(ws_list):
            meta = ws_meta.get(ws_name, {})
            vid = hashlib.sha1(f"{dash_name}:{ws_name}".encode("utf-8")).hexdigest()[:16]
            ws_rect = dash_worksheet_zones.get(dash_name, {}).get(ws_name)
            scaled_ws_rect = scale_rect(ws_rect, root_size, page_size) if ws_rect else None
            if is_table_worksheet(meta):
                row_fields = [f for f in extract_fields(meta.get("rows", "")) if f not in ("Measure Names", "Multiple Values")]
                measure_fields = ws_measure_names.get(ws_name, [])
                table_fields = dedupe_preserve(row_fields + measure_fields)
                table = choose_table_for_any(col_meta, table_fields)
                column_names = set(col_meta.get(table, {}).keys())
                extra_columns = set(order_calc_columns or [])
                measure_overrides = {"Profit Ratio"}
                projections = []
                for field in table_fields:
                    if field in column_names or field in extra_columns:
                        projections.append({
                            "field": {"Column": {"Expression": {"SourceRef": {"Entity": table}}, "Property": field}},
                            "queryRef": f"{table}.{field}",
                            "nativeQueryRef": field
                        })
                    elif field in measure_overrides:
                        projections.append({
                            "field": {"Measure": {"Expression": {"SourceRef": {"Entity": table}}, "Property": field}},
                            "queryRef": f"{table}.{field}",
                            "nativeQueryRef": field
                        })
                    else:
                        projections.append({
                            "field": {"Column": {"Expression": {"SourceRef": {"Entity": table}}, "Property": field}},
                            "queryRef": f"{table}.{field}",
                            "nativeQueryRef": field
                        })
                if scaled_ws_rect:
                    pos_x = max(scaled_ws_rect["x"], 0)
                    pos_y = max(scaled_ws_rect["y"], content_start_y)
                    pos_w = max(scaled_ws_rect["w"], 300)
                    pos_h = max(scaled_ws_rect["h"], 200)
                else:
                    pos_x = 20
                    pos_y = content_start_y
                    pos_w = 1240
                    pos_h = 560
                payload = {
                    "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/visualContainer/2.5.0/schema.json",
                    "name": vid,
                    "position": {"x": pos_x, "y": pos_y, "z": idx + 1, "height": pos_h, "width": pos_w, "tabOrder": idx + 1},
                    "visual": {
                        "visualType": "tableEx",
                        "query": {"queryState": {"Values": {"projections": projections}}},
                        "drillFilterOtherVisuals": True,
                        "autoSelectVisualType": True,
                        "objects": {"title": [{"properties": {"text": {"expr": {"Literal": {"Value": f"'{ws_name}'"}}}, "show": {"expr": {"Literal": {"Value": "true"}}}}}]}
                    }
                }
            else:
                vtype = determine_visual_type(meta, ws_name)
                category, value = choose_category_value(meta)
                table = choose_table_for_fields(col_meta, category, value)
                fields = extract_fields(meta.get("rows", "")) + extract_fields(meta.get("cols", ""))
                series = choose_series(fields)

                # map charts on geo fields
                if category in ("City", "State/Province") and vtype in ("tableEx", "map"):
                    vtype = "map"

                if scaled_ws_rect:
                    x = max(scaled_ws_rect["x"], 0)
                    y = max(scaled_ws_rect["y"], content_start_y)
                    width = max(scaled_ws_rect["w"], 300)
                    height = max(scaled_ws_rect["h"], 200)
                else:
                    x = 20 + (idx % 2) * 620
                    y = content_start_y + (idx // 2) * 300
                    width = 600
                    height = 260

                query_state = {
                    "Category": {"projections": [{
                        "field": {"Column": {"Expression": {"SourceRef": {"Entity": table}}, "Property": category}},
                        "queryRef": f"{table}.{category}",
                        "nativeQueryRef": category,
                        "active": True
                    }]},
                    "Y": {"projections": [{
                        "field": {"Aggregation": {"Expression": {"Column": {"Expression": {"SourceRef": {"Entity": table}}, "Property": value}}, "Function": 0}},
                        "queryRef": f"Sum({table}.{value})",
                        "nativeQueryRef": f"Sum of {value}"
                    }]}
                }
                if series and vtype in ("lineChart", "areaChart"):
                    query_state["Series"] = {"projections": [{
                        "field": {"Column": {"Expression": {"SourceRef": {"Entity": table}}, "Property": series}},
                        "queryRef": f"{table}.{series}",
                        "nativeQueryRef": series
                    }]}

                # map uses Location/Size
                if vtype == "map":
                    query_state = {
                        "Location": {"projections": [{
                            "field": {"Column": {"Expression": {"SourceRef": {"Entity": table}}, "Property": category}},
                            "queryRef": f"{table}.{category}",
                            "nativeQueryRef": category,
                            "active": True
                        }]},
                        "Size": {"projections": [{
                            "field": {"Aggregation": {"Expression": {"Column": {"Expression": {"SourceRef": {"Entity": table}}, "Property": value}}, "Function": 0}},
                            "queryRef": f"Sum({table}.{value})",
                            "nativeQueryRef": f"Sum of {value}"
                        }]}
                    }

                payload = {
                    "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/visualContainer/2.5.0/schema.json",
                    "name": vid,
                    "position": {"x": x, "y": y, "z": idx + 1, "height": height, "width": width, "tabOrder": idx + 1},
                    "visual": {
                        "visualType": vtype,
                        "query": {"queryState": query_state},
                        "drillFilterOtherVisuals": True,
                        "autoSelectVisualType": True,
                        "objects": {"title": [{"properties": {"text": {"expr": {"Literal": {"Value": f"'{ws_name}'"}}}, "show": {"expr": {"Literal": {"Value": "true"}}}}}]}
                    }
                }
            write_generic(vid, payload)

    # Generic visuals for standalone worksheet pages
    for ws_name in standalone_ws:
        if ws_name in ("Overview", "Product"):
            continue
        page_id = page_by_name.get(ws_name)
        if not page_id:
            continue
        page_folder = os.path.join(pages_dir, page_id)
        visuals_dir = os.path.join(page_folder, "visuals")
        os.makedirs(visuals_dir, exist_ok=True)
        meta = ws_meta.get(ws_name, {})
        mark = meta.get("mark")
        vtype = map_mark_to_visual(mark)
        category, value = choose_category_value(meta)
        table = choose_table_for_fields(col_meta, category, value)
        vid = hashlib.sha1(ws_name.encode("utf-8")).hexdigest()[:16]
        # Forecast override: line chart with Order Month + Segment + Sales
        if ws_name == "Forecast":
            payload = {
                "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/visualContainer/2.5.0/schema.json",
                "name": vid,
                "position": {"x": 20, "y": 60, "z": 1, "height": 600, "width": 1240, "tabOrder": 1},
                "visual": {
                    "visualType": "lineChart",
                    "query": {
                        "queryState": {
                            "Category": {"projections": [{
                                "field": {"Column": {"Expression": {"SourceRef": {"Entity": "Orders"}}, "Property": "Order Month"}},
                                "queryRef": "Orders.Order Month",
                                "nativeQueryRef": "Order Month",
                                "active": True
                            }]},
                            "Series": {"projections": [{
                                "field": {"Column": {"Expression": {"SourceRef": {"Entity": "Orders"}}, "Property": "Segment"}},
                                "queryRef": "Orders.Segment",
                                "nativeQueryRef": "Segment"
                            }]},
                            "Y": {"projections": [{
                                "field": {"Aggregation": {"Expression": {"Column": {"Expression": {"SourceRef": {"Entity": "Orders"}}, "Property": "Sales"}}, "Function": 0}},
                                "queryRef": "Sum(Orders.Sales)",
                                "nativeQueryRef": "Sum of Sales"
                            }]}
                        },
                        "sortDefinition": {
                            "sort": [{
                                "field": {"Column": {"Expression": {"SourceRef": {"Entity": "Orders"}}, "Property": "Order Month"}},
                                "direction": "Ascending"
                            }],
                            "isDefaultSort": True
                        }
                    },
                    "drillFilterOtherVisuals": True,
                    "autoSelectVisualType": True,
                    "objects": {"title": [{"properties": {"text": {"expr": {"Literal": {"Value": "'Sales Forecast'"}}}, "show": {"expr": {"Literal": {"Value": "true"}}}}}]}
                }
            }
        else:
            payload = {
            "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/visualContainer/2.5.0/schema.json",
            "name": vid,
            "position": {"x": 20, "y": 60, "z": 1, "height": 600, "width": 1240, "tabOrder": 1},
            "visual": {
                "visualType": vtype,
                "query": {
                    "queryState": {
                        "Category": {"projections": [{
                            "field": {"Column": {"Expression": {"SourceRef": {"Entity": table}}, "Property": category}},
                            "queryRef": f"{table}.{category}",
                            "nativeQueryRef": category,
                            "active": True
                        }]},
                        "Y": {"projections": [{
                            "field": {"Aggregation": {"Expression": {"Column": {"Expression": {"SourceRef": {"Entity": table}}, "Property": value}}, "Function": 0}},
                            "queryRef": f"Sum({table}.{value})",
                            "nativeQueryRef": f"Sum of {value}"
                        }]}
                    }
                },
                "drillFilterOtherVisuals": True,
                "autoSelectVisualType": True,
                "objects": {"title": [{"properties": {"text": {"expr": {"Literal": {"Value": f"'{ws_name}'"}}}, "show": {"expr": {"Literal": {"Value": "true"}}}}}]}
            }
        }
        vdir = os.path.join(visuals_dir, vid)
        os.makedirs(vdir, exist_ok=True)
        with open(os.path.join(vdir, "visual.json"), "w") as f:
            json.dump(payload, f, indent=2)

    # Semantic model files
    model_dir = os.path.join(out_root, model_name, "definition")
    os.makedirs(os.path.join(model_dir, "cultures"), exist_ok=True)
    os.makedirs(os.path.join(model_dir, "tables"), exist_ok=True)

    with open(os.path.join(out_root, model_name, "definition.pbism"), "w") as f:
        json.dump({"version": "4.2", "settings": {}}, f, indent=2)

    with open(os.path.join(model_dir, "database.tmdl"), "w") as f:
        f.write("database\n\tcompatibilityLevel: 1600\n")

    with open(os.path.join(model_dir, "cultures", "en-US.tmdl"), "w") as f:
        f.write("cultureInfo en-US\n\n\tlinguisticMetadata =\n\t\t\t{\n\t\t\t  \"Version\": \"1.0.0\",\n\t\t\t  \"Language\": \"en-US\"\n\t\t\t}\n\t\tcontentType: json\n")

    table_refs = [quote(t) for t in col_meta.keys()]
    query_order = ", ".join([f"\"{t}\"" for t in col_meta.keys()])
    model_lines = [
        "model Model",
        "\tculture: en-US",
        "\tdefaultPowerBIDataSourceVersion: powerBI_V3",
        "\tsourceQueryCulture: en-IN",
        "\tdataAccessOptions",
        "\t\tlegacyRedirects",
        "\t\treturnErrorValuesAsNull",
        "",
        "annotation __PBI_TimeIntelligenceEnabled = 1",
        "",
        f"annotation PBI_QueryOrder = [{query_order}]",
    ]
    for t in col_meta.keys():
        model_lines.append(f"ref table {quote(t)}")
    model_lines.append("")
    model_lines.append("ref cultureInfo en-US")
    model_lines.append("")
    with open(os.path.join(model_dir, "model.tmdl"), "w") as f:
        f.write("\n".join(model_lines) + "\n")

    tables_dir = os.path.join(model_dir, "tables")
    build_table_files(tables_dir, col_meta, args.windows_data_root, order_calc_columns=order_calc_columns)
    normalize_lineage_indentation(tables_dir)

    # PBIP file
    with open(os.path.join(out_root, f"{args.project_name}.pbip"), "w") as f:
        json.dump({"version": "1.0", "artifacts": [{"report": {"path": report_name}}], "settings": {"enableAutoRecovery": True}}, f, indent=2)

    print(f"Generated PBIP with Overview visuals at: {out_root}")


if __name__ == "__main__":
    main()
