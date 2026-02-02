import argparse
import json
import re
import zipfile
from pathlib import Path
from typing import Optional, Tuple
from xml.etree import ElementTree as ET


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


def read_dashboard_names(twb_or_twbx_path: Path) -> list[str]:
    twb_xml = read_twb_xml(twb_or_twbx_path)
    root = ET.fromstring(twb_xml)
    names = []
    for dash in root.findall(".//dashboard"):
        name = dash.get("name")
        if name:
            names.append(name)
    return sorted(set(names))


def read_pbip_pages(report_dir: Path) -> dict[str, str]:
    pages_path = report_dir / "definition" / "pages" / "pages.json"
    pages_meta = json.loads(pages_path.read_text(encoding="utf-8"))
    pages = {}
    for page_id in pages_meta.get("pageOrder", []):
        page_json = json.loads(
            (report_dir / "definition" / "pages" / page_id / "page.json").read_text(
                encoding="utf-8"
            )
        )
        display_name = page_json.get("displayName")
        if display_name:
            pages[display_name] = page_id
    return pages


def read_snapshot_names(snapshots_dir: Path) -> list[str]:
    names = []
    for path in sorted(snapshots_dir.glob("*.png")):
        names.append(path.stem)
    return names


def build_candidates(names: list[str]) -> dict[str, set[str]]:
    return {name: set(tokenize(name)) for name in names}


def best_match(name: str, candidates: dict[str, set[str]]) -> Tuple[Optional[str], float]:
    tokens = set(tokenize(name))
    best = None
    best_score = 0.0
    for candidate, cand_tokens in candidates.items():
        score = jaccard(tokens, cand_tokens)
        if score > best_score:
            best = candidate
            best_score = score
    return best, best_score


def main():
    parser = argparse.ArgumentParser(
        description="Map Tableau snapshot names to Tableau dashboards and PBIP pages."
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
        "--snapshots-dir",
        default="tableau snapshots",
        help="Directory with Tableau snapshot PNGs.",
    )
    parser.add_argument(
        "--out",
        default="out/snapshot_to_pbip_map.json",
        help="Output JSON path.",
    )
    parser.add_argument(
        "--min-score",
        type=float,
        default=0.1,
        help="Minimum Jaccard score to accept a match.",
    )
    args = parser.parse_args()

    dashboards = read_dashboard_names(Path(args.tableau_twbx))
    pbip_pages = read_pbip_pages(Path(args.pbip_report))
    snapshots = read_snapshot_names(Path(args.snapshots_dir))

    dashboard_tokens = build_candidates(dashboards)
    pbip_tokens = build_candidates(list(pbip_pages.keys()))

    results = []
    for snapshot in snapshots:
        dash_match, dash_score = best_match(snapshot, dashboard_tokens)
        page_match, page_score = best_match(snapshot, pbip_tokens)
        results.append(
            {
                "snapshot": snapshot,
                "dashboard_match": dash_match if dash_score >= args.min_score else None,
                "dashboard_score": round(dash_score, 3),
                "pbip_page_match": page_match if page_score >= args.min_score else None,
                "pbip_page_id": pbip_pages.get(page_match) if page_score >= args.min_score else None,
                "pbip_score": round(page_score, 3),
            }
        )

    output_path = Path(args.out)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(json.dumps(results, indent=2), encoding="utf-8")
    print(f"Wrote {output_path}")


if __name__ == "__main__":
    main()
