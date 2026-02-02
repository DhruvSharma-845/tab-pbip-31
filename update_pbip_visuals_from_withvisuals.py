import argparse
import json
import shutil
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET


def load_json(path: Path):
    with path.open("r", encoding="utf-8") as file:
        return json.load(file)


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


def read_dashboard_names(twb_or_twbx_path: Path) -> list[str]:
    twb_xml = read_twb_xml(twb_or_twbx_path)
    root = ET.fromstring(twb_xml)
    names = []
    for dash in root.findall(".//dashboard"):
        name = dash.get("name")
        if name:
            names.append(name)
    return sorted(set(names))


def build_page_map(report_dir: Path) -> dict[str, str]:
    pages_path = report_dir / "definition" / "pages" / "pages.json"
    pages_meta = load_json(pages_path)
    page_map = {}
    for page_id in pages_meta.get("pageOrder", []):
        page_json = load_json(
            report_dir / "definition" / "pages" / page_id / "page.json"
        )
        display_name = page_json.get("displayName")
        if display_name:
            page_map[display_name] = page_id
    return page_map


def replace_visuals(source_page: Path, target_page: Path):
    source_visuals = source_page / "visuals"
    target_visuals = target_page / "visuals"
    if not source_visuals.exists():
        return False
    if target_visuals.exists():
        shutil.rmtree(target_visuals)
    shutil.copytree(source_visuals, target_visuals)
    return True


def main():
    parser = argparse.ArgumentParser(
        description="Copy visual.jsons from source PBIP into target PBIP."
    )
    parser.add_argument(
        "--source-report",
        default="withvisuals/Superstore.Report",
        help="Source PBIP report folder with visuals.",
    )
    parser.add_argument(
        "--target-report",
        default="SuperstorePBIP/Superstore.Report",
        help="Target PBIP report folder to update.",
    )
    parser.add_argument(
        "--tableau-twbx",
        default="Superstore.twbx",
        help="Tableau TWB/TWBX file used for dashboard-to-page mapping.",
    )
    parser.add_argument(
        "--include-non-dashboard-pages",
        action="store_true",
        help="Also update pages not found in Tableau dashboards.",
    )
    args = parser.parse_args()

    source_report = Path(args.source_report)
    target_report = Path(args.target_report)
    twb_path = Path(args.tableau_twbx)

    source_pages = build_page_map(source_report)
    target_pages = build_page_map(target_report)
    dashboards = read_dashboard_names(twb_path)

    updated = []
    skipped = []
    missing = []

    names_to_update = list(dashboards)
    if args.include_non_dashboard_pages:
        names_to_update = sorted(set(names_to_update + list(target_pages.keys())))

    for name in names_to_update:
        source_id = source_pages.get(name)
        target_id = target_pages.get(name)
        if not source_id or not target_id:
            missing.append(name)
            continue
        source_page = source_report / "definition" / "pages" / source_id
        target_page = target_report / "definition" / "pages" / target_id
        if replace_visuals(source_page, target_page):
            updated.append(name)
        else:
            skipped.append(name)

    print("Updated pages:")
    for name in updated:
        print(f"- {name}")

    if skipped:
        print("\nSkipped (no visuals in source):")
        for name in skipped:
            print(f"- {name}")

    if missing:
        print("\nMissing page in source/target:")
        for name in missing:
            print(f"- {name}")


if __name__ == "__main__":
    main()
