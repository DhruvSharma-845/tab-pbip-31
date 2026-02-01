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


def parse_dashboard_worksheets(root, worksheet_names):
    dashboards = {}
    for d in root.findall(".//dashboard"):
        dname = d.get("name")
        worksheets = []
        for z in d.findall(".//zone"):
            wname = z.get("name") or z.get("worksheet")
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

def build_table_files(tables_dir, col_meta, windows_data_root):
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
            out.append(f"\t\t\t\tSource = Excel.Workbook(File.Contents(\"{windows_data_root}\\\\Sample - Superstore.xls\"), null, true),")
            out.append("\t\t\t\tOrdersTable = try Source{[Item=\"Orders\",Kind=\"Table\"]}[Data] otherwise null,")
            out.append("\t\t\t\tOrdersSheet = try Source{[Item=\"Orders\",Kind=\"Sheet\"]}[Data] otherwise null,")
            out.append("\t\t\t\tSelected = if OrdersTable <> null then OrdersTable else if OrdersSheet <> null then OrdersSheet else Source{0}[Data],")
            out.append("\t\t\t\tPromoted = Table.PromoteHeaders(Selected, [PromoteAllScalars=true]),")
            out.append("\t\t\t\tChangedType = Table.TransformColumnTypes(Promoted, {{\"Order Date\", type date}, {\"Ship Date\", type date}, {\"Sales\", type number}, {\"Profit\", type number}, {\"Discount\", type number}, {\"Quantity\", Int64.Type}}),")
            out.append("\t\t\t\tCleanNumbers = Table.TransformColumns(ChangedType, {{\"Sales\", each try Number.From(_) otherwise null, type number}, {\"Profit\", each try Number.From(_) otherwise null, type number}, {\"Discount\", each try Number.From(_) otherwise null, type number}, {\"Quantity\", each try Number.From(_) otherwise null, Int64.Type}}),")
            out.append("\t\t\t\tAddedMonth = Table.AddColumn(CleanNumbers, \"Order Month\", each Date.StartOfMonth([Order Date]), type date),")
            out.append("\t\t\t\tAddedProfitability = Table.AddColumn(AddedMonth, \"Profitability\", each if [Profit] >= 0 then \"Profitable\" else \"Unprofitable\", type text),")
            out.append("\t\t\t\tAddedYear = Table.AddColumn(AddedProfitability, \"Order Year\", each Date.Year([Order Date]), Int64.Type)")
            out.append("\t\t\tin")
            out.append("\t\t\t\tAddedYear")
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
    for name in dashboard_names:
        page_id = hashlib.sha1(name.encode("utf-8")).hexdigest()[:16]
        page_ids.append(page_id)
        page_by_name[name] = page_id
        page_folder = os.path.join(pages_dir, page_id)
        os.makedirs(page_folder, exist_ok=True)
        page_json = {
            "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/page/2.0.0/schema.json",
            "name": page_id,
            "displayName": name,
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

        for idx, ws_name in enumerate(ws_list):
            meta = ws_meta.get(ws_name, {})
            mark = meta.get("mark")
            vtype = map_mark_to_visual(mark)
            category, value = choose_category_value(meta)
            table = choose_table_for_fields(col_meta, category, value)

            # map charts on geo fields
            if category in ("City", "State/Province") and vtype in ("tableEx", "map"):
                vtype = "map"

            vid = hashlib.sha1(f"{dash_name}:{ws_name}".encode("utf-8")).hexdigest()[:16]
            x = 20 + (idx % 2) * 620
            y = 60 + (idx // 2) * 300

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
                "position": {"x": x, "y": y, "z": idx + 1, "height": 260, "width": 600, "tabOrder": idx + 1},
                "visual": {
                    "visualType": vtype,
                    "query": {"queryState": query_state},
                    "drillFilterOtherVisuals": True,
                    "autoSelectVisualType": True,
                    "objects": {"title": [{"properties": {"text": {"expr": {"Literal": {"Value": f"'{ws_name}'"}}}, "show": {"expr": {"Literal": {"Value": "true"}}}}}]}
                }
            }
            write_generic(vid, payload)

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
    build_table_files(tables_dir, col_meta, args.windows_data_root)
    normalize_lineage_indentation(tables_dir)

    # PBIP file
    with open(os.path.join(out_root, f"{args.project_name}.pbip"), "w") as f:
        json.dump({"version": "1.0", "artifacts": [{"report": {"path": report_name}}], "settings": {"enableAutoRecovery": True}}, f, indent=2)

    print(f"Generated PBIP with Overview visuals at: {out_root}")


if __name__ == "__main__":
    main()
