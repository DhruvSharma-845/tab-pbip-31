#!/usr/bin/env python3
"""
PBIR preview converter (best-effort) that builds a PBIP with visuals.

Important:
- Requires PBIR preview enabled in Power BI Desktop.
- Uses boundVisual-style visual.json schema (visualContainer/2.5.0).
- Partitions use Windows absolute paths (required by Desktop).
"""
import argparse
import hashlib
import json
import os
import re
import shutil
import xml.etree.ElementTree as ET
from collections import defaultdict


def safe_table_name(name: str) -> str:
    return name.strip()


def needs_quotes(name: str) -> bool:
    return bool(re.search(r"[^A-Za-z0-9_]", name))


def quote_name(name: str) -> str:
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
    # best-effort extract common Tableau packaged data files
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
        table = safe_table_name(ds_caption.get(ds, ds))
        for col in dep.findall("column"):
            name = col.get("name")
            if not name:
                continue
            datatype = col.get("datatype")
            col_name = name.strip("[]")
            col_meta[table][col_name] = datatype
    return col_meta


def pick_visual_fields(col_meta):
    # pick first table with at least one string + numeric column
    for table, cols in col_meta.items():
        str_cols = [c for c, t in cols.items() if map_type(t) == "string"]
        num_cols = [c for c, t in cols.items() if map_type(t) in ("int64", "double")]
        if str_cols and num_cols:
            return table, str_cols[0], num_cols[0]
    # fallback
    for table, cols in col_meta.items():
        cols_list = list(cols.keys())
        if cols_list:
            return table, cols_list[0], cols_list[0]
    return "Table", "Category", "Value"


def build_table_files(tables_dir, col_meta, windows_data_root):
    tables_dir = os.path.abspath(tables_dir)
    os.makedirs(tables_dir, exist_ok=True)

    for table, cols in col_meta.items():
        tname = quote_name(table)
        out = []
        out.append(f"table {tname}")
        out.append("\tlineageTag: 00000000-0000-0000-0000-000000000000")
        out.append("")
        for col, dtype in cols.items():
            cname = quote_name(col)
            summarize = "sum" if map_type(dtype) in ("int64", "double") else "none"
            out.append(f"\tcolumn {cname}")
            out.append(f"\t\tdataType: {map_type(dtype)}")
            out.append(f"\t\tsummarizeBy: {summarize}")
            out.append(f"\t\tsourceColumn: {col}")
            out.append(f"\t\tlineageTag: 00000000-0000-0000-0000-000000000001")
            out.append("")

        # Partition: best-effort guessing of data file names
        out.append("\tpartition FullData = m")
        out.append("\t\tmode: import")
        out.append("\t\tsource =")
        out.append("\t\t\tlet")
        # These are Superstore-specific defaults; override with correct files after generation.
        if "Sales Target" in table:
            out.append(f"\t\t\t\tSource = Excel.Workbook(File.Contents(\"{windows_data_root}\\\\Sales Target.xlsx\"), null, true),")
            out.append("\t\t\t\tSheet1 = Source{[Item=\"Sheet1\",Kind=\"Sheet\"]}[Data],")
            out.append("\t\t\t\tPromoted = Table.PromoteHeaders(Sheet1, [PromoteAllScalars=true])")
        elif "Sales Commission" in table:
            out.append(f"\t\t\t\tSource = Csv.Document(File.Contents(\"{windows_data_root}\\\\Sales Commission.csv\"), [Delimiter=\",\", Columns=4, Encoding=65001, QuoteStyle=QuoteStyle.None]),")
            out.append("\t\t\t\tPromoted = Table.PromoteHeaders(Source, [PromoteAllScalars=true])")
        else:
            out.append(f"\t\t\t\tSource = Excel.Workbook(File.Contents(\"{windows_data_root}\\\\Sample - Superstore.xls\"), null, true),")
            out.append("\t\t\t\tOrdersTable = try Source{[Item=\"Orders\",Kind=\"Table\"]}[Data] otherwise null,")
            out.append("\t\t\t\tOrdersSheet = try Source{[Item=\"Orders\",Kind=\"Sheet\"]}[Data] otherwise null,")
            out.append("\t\t\t\tSelected = if OrdersTable <> null then OrdersTable else if OrdersSheet <> null then OrdersSheet else Source{0}[Data],")
            out.append("\t\t\t\tPromoted = Table.PromoteHeaders(Selected, [PromoteAllScalars=true])")
        out.append("\t\t\tin")
        out.append("\t\t\t\tPromoted")
        out.append("\tannotation PBI_ResultType = Table")
        out.append("")

        (Path(tables_dir) / f"{table}.tmdl").write_text("\n".join(out) + "\n")


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

    data_dir = os.path.join(out_root, "Data")
    extract_data_files(args.twbx, data_dir)

    root = ET.parse(twb_path).getroot()
    ds_caption = parse_datasources(root)
    col_meta = parse_columns(root, ds_caption)

    # Create report structure
    pages_dir = os.path.join(out_root, report_name, "definition", "pages")
    os.makedirs(pages_dir, exist_ok=True)
    os.makedirs(os.path.join(out_root, report_name, "StaticResources", "SharedResources", "BaseThemes"), exist_ok=True)
    shutil.copy2(args.sample_theme,
                 os.path.join(out_root, report_name, "StaticResources", "SharedResources", "BaseThemes",
                              os.path.basename(args.sample_theme)))

    # report.json
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
            "section": [
                {"properties": {"verticalAlignment": {"expr": {"Literal": {"Value": "'Top'"}}}}}
            ]
        },
        "resourcePackages": [
            {
                "name": "SharedResources",
                "type": "SharedResources",
                "items": [{"name": os.path.splitext(os.path.basename(args.sample_theme))[0],
                           "path": f"BaseThemes/{os.path.basename(args.sample_theme)}",
                           "type": "BaseTheme"}]
            }
        ],
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

    # version.json
    with open(os.path.join(out_root, report_name, "definition", "version.json"), "w") as f:
        json.dump({
            "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/versionMetadata/1.0.0/schema.json",
            "version": "2.0.0"
        }, f, indent=2)

    # definition.pbir
    with open(os.path.join(out_root, report_name, "definition.pbir"), "w") as f:
        json.dump({"version": "4.0", "datasetReference": {"byPath": {"path": f"../{model_name}"}}}, f, indent=2)

    # pages.json and first page
    page_id = hashlib.sha1("Page 1".encode("utf-8")).hexdigest()[:16]
    page_folder = os.path.join(pages_dir, page_id)
    os.makedirs(page_folder, exist_ok=True)
    page_json = {
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/page/2.0.0/schema.json",
        "name": page_id,
        "displayName": "Page 1",
        "displayOption": "FitToPage",
        "height": 720,
        "width": 1280
    }
    with open(os.path.join(page_folder, "page.json"), "w") as f:
        json.dump(page_json, f, indent=2)
    with open(os.path.join(pages_dir, "pages.json"), "w") as f:
        json.dump({
            "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/pagesMetadata/1.0.0/schema.json",
            "pageOrder": [page_id],
            "activePageName": page_id
        }, f, indent=2)

    # Visual
    table, category, value = pick_visual_fields(col_meta)
    visual_id = hashlib.sha1(f"{table}:{category}:{value}".encode("utf-8")).hexdigest()[:16]
    visual_dir = os.path.join(page_folder, "visuals", visual_id)
    os.makedirs(visual_dir, exist_ok=True)
    visual_json = {
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/visualContainer/2.5.0/schema.json",
        "name": visual_id,
        "position": {
            "x": 60,
            "y": 60,
            "z": 0,
            "height": 300,
            "width": 520,
            "tabOrder": 0
        },
        "visual": {
            "visualType": "columnChart",
            "query": {
                "queryState": {
                    "Category": {
                        "projections": [
                            {
                                "field": {
                                    "Column": {
                                        "Expression": {"SourceRef": {"Entity": table}},
                                        "Property": category
                                    }
                                },
                                "queryRef": f"{table}.{category}",
                                "nativeQueryRef": category,
                                "active": True
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
                                                "Expression": {"SourceRef": {"Entity": table}},
                                                "Property": value
                                            }
                                        },
                                        "Function": 0
                                    }
                                },
                                "queryRef": f"Sum({table}.{value})",
                                "nativeQueryRef": f"Sum of {value}"
                            }
                        ]
                    }
                }
            },
            "drillFilterOtherVisuals": True,
            "autoSelectVisualType": True
        }
    }
    with open(os.path.join(visual_dir, "visual.json"), "w") as f:
        json.dump(visual_json, f, indent=2)

    # Semantic model structure
    model_dir = os.path.join(out_root, model_name, "definition")
    os.makedirs(os.path.join(model_dir, "cultures"), exist_ok=True)
    os.makedirs(os.path.join(model_dir, "tables"), exist_ok=True)

    with open(os.path.join(out_root, model_name, "definition.pbism"), "w") as f:
        json.dump({"version": "4.2", "settings": {}}, f, indent=2)

    with open(os.path.join(model_dir, "database.tmdl"), "w") as f:
        f.write("database\n\tcompatibilityLevel: 1600\n")

    with open(os.path.join(model_dir, "cultures", "en-US.tmdl"), "w") as f:
        f.write("cultureInfo en-US\n\n\tlinguisticMetadata =\n\t\t\t{\n\t\t\t  \"Version\": \"1.0.0\",\n\t\t\t  \"Language\": \"en-US\"\n\t\t\t}\n\t\tcontentType: json\n")

    # model.tmdl with refs
    table_refs = [quote_name(t) for t in col_meta.keys()]
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
        f"annotation PBI_QueryOrder = [{', '.join([f'\"{t}\"' for t in col_meta.keys()])}]",
    ]
    for t in col_meta.keys():
        model_lines.append(f"ref table {quote_name(t)}")
    model_lines.append("")
    model_lines.append("ref cultureInfo en-US")
    model_lines.append("")
    with open(os.path.join(model_dir, "model.tmdl"), "w") as f:
        f.write("\n".join(model_lines) + "\n")

    build_table_files(os.path.join(model_dir, "tables"), col_meta, args.windows_data_root)

    # PBIP file
    with open(os.path.join(out_root, f"{args.project_name}.pbip"), "w") as f:
        json.dump({"version": "1.0", "artifacts": [{"report": {"path": report_name}}], "settings": {"enableAutoRecovery": True}}, f, indent=2)

    print(f"Generated PBIP with one bound visual at: {out_root}")


if __name__ == "__main__":
    main()
