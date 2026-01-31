#!/usr/bin/env python3
"""
Stable PBIP generator from a Tableau TWBX workbook.
- Outputs PBIP compatible with official stable schemas.
- No visualContainers (stable page schema 2.0.0) so pages are blank but valid.
- Adds FullData import partitions with empty typed tables.
"""
import argparse
import json
import os
import re
import shutil
import hashlib
import xml.etree.ElementTree as ET
from collections import defaultdict


def ds_to_table(ds_caption: str) -> str:
    caption_table_map = {
        'Sample - Superstore': 'Orders',
        'Sales Commission': 'Sales Commission',
        'Sales Target': 'Sales Target'
    }
    return caption_table_map.get(ds_caption, ds_caption)


def map_type(t: str) -> str:
    if not t:
        return 'string'
    t = t.lower()
    if t in ('string', 'str'):
        return 'string'
    if t in ('integer', 'int'):
        return 'int64'
    if t in ('real', 'float', 'double'):
        return 'double'
    if t in ('date', 'datetime'):
        return 'dateTime'
    if t in ('boolean', 'bool'):
        return 'boolean'
    return 'string'


def build_model_tmdl(model_path, tables_meta):
    m_type_map = {
        'string': 'text',
        'int64': 'number',
        'double': 'number',
        'dateTime': 'datetime',
        'boolean': 'logical'
    }

    lines = [
        'model Model',
        '\tculture: en-US',
        '\tdefaultPowerBIDataSourceVersion: powerBI_V3',
        '\tsourceQueryCulture: en-IN',
        '\tdataAccessOptions',
        '\t\tlegacyRedirects',
        '\t\treturnErrorValuesAsNull',
        '',
        'annotation __PBI_TimeIntelligenceEnabled = 1',
        '',
        'annotation PBI_ProTooling = ["DevMode"]',
        '',
        'ref cultureInfo en-US',
        ''
    ]

    for table, data in tables_meta.items():
        lines.append(f"table '{table}'")
        cols = data.get('columns', {})
        for col_name, dtype in cols.items():
            lines.append(f"\tcolumn '{col_name}'")
            lines.append(f"\t\tdataType: {map_type(dtype)}")
        measures = data.get('measures', {})
        for measure_name, expr in measures.items():
            lines.append(f"\tmeasure '{measure_name}' = {expr}")

        # FullData partition (empty typed table) for stable PBIP
        m_fields = []
        for col_name, dtype in cols.items():
            m_fields.append(f"{col_name} = {m_type_map.get(map_type(dtype), 'text')}")
        m_schema = ', '.join(m_fields) if m_fields else ''
        m_source = "let\n            Source = #table(type table [" + m_schema + "], {})\n        in\n            Source"

        lines.append("\tpartition 'FullData'")
        lines.append("\t\tmode: import")
        lines.append("\t\tsource =")
        for line in m_source.splitlines():
            lines.append("\t\t\t" + line)
        lines.append('')

    with open(model_path, 'w') as f:
        f.write('\n'.join(lines) + '\n')


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--twbx', required=True, help='Path to .twbx file')
    parser.add_argument('--out', required=True, help='Output PBIP folder')
    parser.add_argument('--project-name', default='Superstore', help='Base name for report/model folders')
    parser.add_argument('--sample-theme', required=True, help='Path to a base theme json to copy')
    args = parser.parse_args()

    twbx_path = os.path.abspath(args.twbx)
    out_root = os.path.abspath(args.out)
    project_name = args.project_name

    # Extract TWB
    twb_path = os.path.join(out_root, project_name + '.twb')
    if not os.path.exists(out_root):
        os.makedirs(out_root, exist_ok=True)
    os.system(f"unzip -p '{twbx_path}' '*.twb' > '{twb_path}'")

    root = ET.parse(twb_path).getroot()

    # datasource name -> caption
    ds_caption = {}
    for ds in root.findall('.//datasource'):
        name = ds.get('name')
        caption = ds.get('caption')
        if name and caption:
            ds_caption[name] = caption

    # Collect column metadata
    col_meta = defaultdict(dict)
    for dep in root.findall('.//datasource-dependencies'):
        ds = dep.get('datasource')
        if not ds:
            continue
        table = ds_to_table(ds_caption.get(ds, ds))
        for col in dep.findall('column'):
            name = col.get('name')
            if not name:
                continue
            datatype = col.get('datatype')
            col_name = name.strip('[]')
            col_meta[table][col_name] = datatype

    # Build PBIP structure
    report_name = f"{project_name}.Report"
    model_name = f"{project_name}.SemanticModel"

    os.makedirs(out_root, exist_ok=True)
    os.makedirs(os.path.join(out_root, report_name, 'definition', 'pages'), exist_ok=True)
    os.makedirs(os.path.join(out_root, report_name, 'StaticResources', 'SharedResources', 'BaseThemes'), exist_ok=True)
    os.makedirs(os.path.join(out_root, model_name, 'definition', 'cultures'), exist_ok=True)

    # copy theme
    shutil.copy2(args.sample_theme,
                 os.path.join(out_root, report_name, 'StaticResources', 'SharedResources', 'BaseThemes', os.path.basename(args.sample_theme)))

    # pbip file
    with open(os.path.join(out_root, f'{project_name}.pbip'), 'w') as f:
        json.dump({
            'version': '1.0',
            'artifacts': [{'report': {'path': report_name}}],
            'settings': {'enableAutoRecovery': True}
        }, f, indent=2)

    # definition.pbir
    with open(os.path.join(out_root, report_name, 'definition.pbir'), 'w') as f:
        json.dump({'version': '4.0', 'datasetReference': {'byPath': {'path': f'../{model_name}'}}}, f, indent=2)

    # report.json
    report_json = {
        '$schema': 'https://developer.microsoft.com/json-schemas/fabric/item/report/definition/report/3.1.0/schema.json',
        'themeCollection': {
            'baseTheme': {
                'name': os.path.splitext(os.path.basename(args.sample_theme))[0],
                'reportVersionAtImport': {'visual': '2.5.0', 'report': '3.1.0', 'page': '2.0.0'},
                'type': 'SharedResources'
            }
        },
        'objects': {
            'section': [
                {'properties': {'verticalAlignment': {'expr': {'Literal': {'Value': "'Top'"}}}}}
            ]
        },
        'resourcePackages': [
            {
                'name': 'SharedResources',
                'type': 'SharedResources',
                'items': [{'name': os.path.splitext(os.path.basename(args.sample_theme))[0],
                           'path': f"BaseThemes/{os.path.basename(args.sample_theme)}",
                           'type': 'BaseTheme'}]
            }
        ],
        'settings': {
            'useStylableVisualContainerHeader': True,
            'exportDataMode': 'AllowSummarized',
            'defaultDrillFilterOtherVisuals': True,
            'allowChangeFilterTypes': True,
            'useEnhancedTooltips': True,
            'useDefaultAggregateDisplayName': True
        }
    }
    with open(os.path.join(out_root, report_name, 'definition', 'report.json'), 'w') as f:
        json.dump(report_json, f, indent=2)

    # version.json
    with open(os.path.join(out_root, report_name, 'definition', 'version.json'), 'w') as f:
        json.dump({
            '$schema': 'https://developer.microsoft.com/json-schemas/fabric/item/report/definition/versionMetadata/1.0.0/schema.json',
            'version': '2.0.0'
        }, f, indent=2)

    # pages.json (names only, no visualContainers for stable schema)
    dashboards = root.findall('.//dashboard')
    page_ids = []
    for dash in dashboards:
        dash_name = dash.get('name')
        page_id = hashlib.sha1(dash_name.encode('utf-8')).hexdigest()[:16]
        page_ids.append(page_id)
        page_folder = os.path.join(out_root, report_name, 'definition', 'pages', page_id)
        os.makedirs(page_folder, exist_ok=True)
        page_json = {
            '$schema': 'https://developer.microsoft.com/json-schemas/fabric/item/report/definition/page/2.0.0/schema.json',
            'name': page_id,
            'displayName': dash_name,
            'displayOption': 'FitToPage',
            'height': 720,
            'width': 1280
        }
        with open(os.path.join(page_folder, 'page.json'), 'w') as f:
            json.dump(page_json, f, indent=2)

    with open(os.path.join(out_root, report_name, 'definition', 'pages', 'pages.json'), 'w') as f:
        json.dump({
            '$schema': 'https://developer.microsoft.com/json-schemas/fabric/item/report/definition/pagesMetadata/1.0.0/schema.json',
            'pageOrder': page_ids,
            'activePageName': page_ids[0] if page_ids else ''
        }, f, indent=2)

    # semantic model files
    with open(os.path.join(out_root, model_name, 'definition.pbism'), 'w') as f:
        json.dump({'version': '4.2', 'settings': {}}, f, indent=2)

    with open(os.path.join(out_root, model_name, 'definition', 'database.tmdl'), 'w') as f:
        f.write('database\n\tcompatibilityLevel: 1600\n')

    with open(os.path.join(out_root, model_name, 'definition', 'cultures', 'en-US.tmdl'), 'w') as f:
        f.write('cultureInfo en-US\n\n\tlinguisticMetadata =\n\t\t\t{\n\t\t\t  "Version": "1.0.0",\n\t\t\t  "Language": "en-US"\n\t\t\t}\n\t\tcontentType: json\n')

    # model.tmdl
    tables_meta = defaultdict(lambda: {'columns': {}, 'measures': {}})
    for table, cols in col_meta.items():
        for col_name, dtype in cols.items():
            tables_meta[table]['columns'][col_name] = dtype

    # Example measures to avoid blank model - use unique names per table
    for table in tables_meta.keys():
        if 'Sales' in tables_meta[table]['columns']:
            tables_meta[table]['measures'][f"{table} Sum of Sales"] = f"SUM('{table}'[Sales])"

    build_model_tmdl(os.path.join(out_root, model_name, 'definition', 'model.tmdl'), tables_meta)

    print(f"Generated stable PBIP at: {out_root}")


if __name__ == '__main__':
    main()
