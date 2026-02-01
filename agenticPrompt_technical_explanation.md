# TWBX to PBIP Conversion: Technical Explanation

## Executive Summary
This document explains the technical approach used to convert Tableau TWBX workbooks into Power BI PBIP projects with high fidelity. The conversion process reads the TWB XML inside a TWBX, reconstructs the semantic model (TMDL), and generates report visuals (PBIR JSON) that match Tableau dashboards pixel‑for‑pixel where possible. It is designed to be robust, schema‑compliant, and safe from PBIP corruption.

## What the Approach Produces
- A PBIP folder with a Power BI report and semantic model.
- All table relationships, measures, and calculated fields mapped from Tableau.
- Visuals recreated from Tableau dashboards, aligned to snapshot imagery.
- A repeatable process that enforces schema rules to prevent breakage.

## High‑Level Architecture
```
Inputs
  ├─ Tableau TWBX
  │   ├─ TWB XML (metadata + calculations + layout)
  │   └─ Packaged data (extractable files)
  ├─ Empty PBIP template
  └─ Tableau snapshot images
              │
              ▼
Parsing + Mapping Engine
  ├─ XML parser (datasources, worksheets, dashboards)
  ├─ Calculation translator (Tableau → DAX)
  ├─ Relationship builder
  ├─ Visual mapper + layout engine
  └─ Schema validator + corruption guardrails
              │
              ▼
Outputs
  ├─ PBIP item shortcut (.pbip)
  ├─ Report definition (PBIR JSON)
  └─ Semantic model (TMDL)
```

## Detailed Data Flow
```
TWBX
  └─ unzip → TWB XML
          ├─ Datasources + tables → TMDL tables
          ├─ Joins/relationships → model.tmdl
          ├─ Calculations → DAX measures/columns
          ├─ Worksheets → visual mappings
          ├─ Dashboards → report pages
          └─ Filters/parameters → slicers/interactions
Snapshots
  └─ layout + style references → visual positions + formatting
```

## Architecture Diagram: TWB XML to DAX + PBIP
```
TWBX (zip)
  ├─ workbook.twb (XML)
  │    ├─ datasources → tables + columns → TMDL tables/*.tmdl
  │    ├─ relations/join graph → model.tmdl relationships
  │    ├─ calculations → DAX measures/columns
  │    ├─ worksheets → visual spec + fields + mark type
  │    ├─ dashboards → PBIR pages + layout coordinates
  │    └─ filters/parameters → PBIR filters + slicers + interactions
  └─ Data/* (extractable files)
            └─ Power Query source mapping (File.Contents)

Snapshots (images)
  └─ layout + style → pixel alignment + formatting
```

## Architecture Diagram: PBIP Assembly
```
Semantic Model (PBISM + TMDL)
  ├─ definition.pbism (version + settings)
  ├─ database.tmdl (compatibilityLevel)
  ├─ model.tmdl (relationships + metadata + annotations)
  └─ tables/*.tmdl (columns + measures + partitions)
                 │
                 ▼
Report (PBIR JSON)
  ├─ definition.pbir (datasetReference)
  ├─ report.json (theme + settings)
  ├─ pages/pages.json (order + active)
  └─ pages/<pageId>/
       ├─ page.json (size + displayName)
       └─ visuals/<visualId>/visual.json (queries + positions)
```

## Core Components
### 1) TWB XML Parser
Extracts all metadata needed to reproduce the model and report:
- Datasources and table schemas
- Columns with data types
- Worksheet definitions and mark types
- Dashboard layout and zones
- Filters, parameters, and global settings

### 2) Calculation Translator (Tableau → DAX)
- Generates exact DAX for calculations when an equivalent exists.
- If the Tableau expression is non‑portable, the translator applies a closest‑match rule and records it in an assumptions log.
- Supports:
  - Aggregations and row‑level calcs
  - FIXED/INCLUDE/EXCLUDE patterns when mappable
  - Parameter‑driven logic
  - KPI switching via calculation groups

### 3) Semantic Model Builder (TMDL)
Produces:
- `tables/*.tmdl` for each datasource or logical table
- `model.tmdl` relationships and metadata
- `database.tmdl` compatibility settings

### 4) Report Generator (PBIR JSON)
Produces:
- Pages for each Tableau dashboard
- Visuals for each worksheet
- KPI cards with complete `visual.json` definitions
- Page order and active page metadata

### 5) Validation + Corruption Guardrails
Strict schema rules prevent invalid PBIP output:
- All JSON files adhere to official schemas and versions
- All IDs are unique and referenced correctly
- All visuals have valid positions and queries
- Paths are relative and OS‑agnostic

## Schema Differences: Tableau TWB vs Power BI PBIP
### Tableau TWB (XML) logic
- A single XML document that encodes datasources, relationships, worksheets, dashboards, filters, and calculations.
- Calculations are stored as Tableau expressions with context‑dependent evaluation (LOD, table calcs, parameters).
- Worksheets describe visual intent (mark type, rows/cols shelves) rather than a fully materialized rendering definition.
- Dashboards define layout through zones and nested containers.

### Power BI PBIP logic
- PBIP is a folder structure with separate schemas:
  - **Semantic model** in TMDL (`*.tmdl`) with explicit columns, measures, relationships, and partitions.
  - **Report** in PBIR JSON with explicit visuals, positions, and queries.
- DAX is required for all calculations and is evaluated in a different filter context model than Tableau.
- Visuals are fully described with query projections and container metadata.

### Key Differences in Logic
- **Context evaluation:** Tableau’s LOD/table calcs are context‑driven; Power BI uses filter context and DAX semantics.
- **Visual definition:** Tableau worksheets define intent; PBIR requires explicit visual JSON with queries and positions.
- **Relationships:** Tableau can use logical/physical layers; Power BI requires explicit relationships in model.tmdl.
- **Parameters and filters:** Tableau parameters are global by default; Power BI often uses slicers or calculation groups.

### How the Conversion Bridges the Gap
- Tableau calculations are translated to DAX with explicit context handling.
- Worksheet intent is translated to concrete visual JSON (visualType + query + position).
- Dashboard zones are mapped to page coordinates, ensuring pixel alignment.
- Relationships are generated explicitly in model.tmdl to match logical data model behavior.

## Folder Output (PBIP)
```
<Project>.pbip
<Project>.Report/
  definition.pbir
  definition/
    report.json
    version.json
    pages/pages.json
    pages/<pageId>/page.json
    pages/<pageId>/visuals/<visualId>/visual.json
  StaticResources/SharedResources/BaseThemes/<theme>.json (optional)
<Project>.SemanticModel/
  definition.pbism
  definition/
    database.tmdl
    model.tmdl
    cultures/en-US.tmdl
    tables/*.tmdl
  diagramLayout.json (optional)
```

## Accuracy Strategy
Accuracy is achieved through layered validation and direct metadata use:
1. **Direct XML extraction** ensures no loss of information from the TWBX.
2. **Exact DAX translation** targets functional parity with Tableau logic.
3. **Schema‑first output** reduces risk of Power BI load failure.
4. **Snapshot‑aligned layout** ensures pixel‑level visual similarity.

### Accuracy Considerations
- **Best case:** 1:1 mapping with identical numbers, visual layout, and interactions.
- **Edge cases:** Certain Tableau features can be approximated but may require assumptions:
  - Complex LOD expressions
  - Custom marks or proprietary Tableau features
  - Nested table calcs with context dependency

These exceptions are explicitly logged and surfaced in an “Assumptions” section in the final output.

## KPI Switching via Calculation Groups
KPI switching is implemented via a calculation group that dynamically changes the selected measure for a KPI visual. This enables a single card or chart to swap between metrics based on a slicer/parameter. Each KPI card visual is fully defined in its own `visual.json`.

## Why This Does Not Break PBIP
- Schema versions and URLs are fixed and validated.
- Each JSON file strictly matches its official definition.
- No additional files or properties are injected.
- Names, IDs, and references are consistent across report metadata.

## Leadership Summary
This approach provides:
- **Speed:** Automated conversion from Tableau to Power BI.
- **Precision:** Model and visuals closely replicate the source.
- **Reliability:** Schema‑safe PBIP output that opens cleanly in Desktop.
- **Traceability:** Clear assumptions logged for any non‑portable features.

The result is an enterprise‑grade conversion pipeline that preserves business logic and visual fidelity while minimizing the risk of PBIP corruption or report failure.
