# Technical Explanation and Architecture

This document provides a detailed technical overview of the Tableau TWBX to Power BI PBIP conversion system. This is reference documentation for understanding the system architecture and is separate from the actionable agent prompt.

## System Overview

This conversion agent operates as a multi-modal pipeline that transforms Tableau TWBX workbooks into Power BI PBIP (Power BI Project) format. The architecture leverages both structured data extraction (XML parsing) and visual understanding (vision models) to achieve high-fidelity conversion.

## Architecture Components

```
┌─────────────────────────────────────────────────────────────────────────────┐
│                           INPUT LAYER                                       │
├─────────────────┬─────────────────────┬─────────────────────────────────────┤
│  TWBX Package   │  Dashboard Images   │  Sample PBIP Template               │
│  (.twbx/.twb)   │  (PNG/SVG snapshots)│  (samplepbipfolder)                 │
└────────┬────────┴──────────┬──────────┴──────────────────┬──────────────────┘
         │                   │                              │
         ▼                   ▼                              ▼
┌─────────────────────────────────────────────────────────────────────────────┐
│                        EXTRACTION LAYER                                     │
├─────────────────────────────────────────────────────────────────────────────┤
│  ┌─────────────────────┐  ┌─────────────────────┐  ┌─────────────────────┐  │
│  │  TWB XML Parser     │  │  Vision Model       │  │  Template Loader    │  │
│  │  - Datasources      │  │  - Chart detection  │  │  - Schema refs      │  │
│  │  - Calculations     │  │  - Layout analysis  │  │  - Visual templates │  │
│  │  - Worksheets       │  │  - Color extraction │  │  - TMDL patterns    │  │
│  │  - Relationships    │  │  - Label parsing    │  │                     │  │
│  └─────────────────────┘  └─────────────────────┘  └─────────────────────┘  │
└─────────────────────────────────────────────────────────────────────────────┘
         │                   │                              │
         ▼                   ▼                              ▼
┌─────────────────────────────────────────────────────────────────────────────┐
│                       TRANSFORMATION LAYER                                  │
├─────────────────────────────────────────────────────────────────────────────┤
│  ┌───────────────────────────────────────────────────────────────────────┐  │
│  │                    SEMANTIC MODEL GENERATOR                           │  │
│  │  Tableau Datasource → TMDL Tables                                     │  │
│  │  Tableau Calculations → DAX Measures/Calculated Columns               │  │
│  │  Tableau Relationships → Model Relationships                          │  │
│  │  Tableau Parameters → What-If Parameters / Calculation Groups         │  │
│  └───────────────────────────────────────────────────────────────────────┘  │
│  ┌───────────────────────────────────────────────────────────────────────┐  │
│  │                      REPORT GENERATOR                                 │  │
│  │  Tableau Dashboards → PBIR Pages                                      │  │
│  │  Tableau Worksheets → Visual Containers (visual.json)                 │  │
│  │  Tableau Filters → Slicers / Report Filters / Visual Filters          │  │
│  │  Tableau Actions → Drill-through / Bookmarks                          │  │
│  └───────────────────────────────────────────────────────────────────────┘  │
└─────────────────────────────────────────────────────────────────────────────┘
         │
         ▼
┌─────────────────────────────────────────────────────────────────────────────┐
│                          OUTPUT LAYER                                       │
├─────────────────────────────────────────────────────────────────────────────┤
│  <Project>.pbip                                                             │
│  <Project>.Report/                                                          │
│    ├── definition.pbir                                                      │
│    └── definition/                                                          │
│        ├── report.json, version.json, pages/...                             │
│  <Project>.SemanticModel/                                                   │
│    ├── definition.pbism                                                     │
│    └── definition/                                                          │
│        ├── database.tmdl, model.tmdl, tables/*.tmdl, cultures/*.tmdl        │
└─────────────────────────────────────────────────────────────────────────────┘
```

## Data Flow Pipeline

### 1. Ingestion Phase
- TWBX is unzipped to extract the embedded TWB (XML) and any packaged data extracts (.hyper/.tde)
- Dashboard snapshots (PNG/SVG) are loaded for visual analysis
- Sample PBIP template provides schema validation and structure reference

### 2. Parsing Phase
- **TWB XML Parser**: Extracts datasources, columns, calculated fields, worksheets, dashboards, filters, parameters, and relationships from the XML structure
- **Vision Model**: Analyzes dashboard snapshots to detect chart types, layout positions, color palettes, label placements, and visual hierarchies
- **Conflict Resolution**: TWB is authoritative for data/logic; snapshots are authoritative for visual layout and styling

### 3. Transformation Phase
- **Semantic Model**: Converts Tableau datasources to TMDL table definitions with proper data types, converts Tableau calculations to DAX expressions, builds model relationships
- **Report Definition**: Maps Tableau worksheets to Power BI visual containers, positions visuals based on snapshot analysis, applies styling and themes

### 4. Generation Phase
- Outputs complete PBIP folder structure with all required JSON schemas
- Validates all file references and field mappings
- Ensures Power BI Desktop compatibility

## Key Technical Mappings

| Tableau Concept | Power BI Equivalent | Output Location |
|-----------------|---------------------|-----------------|
| Datasource | Table | `tables/*.tmdl` |
| Dimension | Column | `tables/*.tmdl` |
| Measure | DAX Measure | `tables/*.tmdl` |
| Calculated Field | Calculated Column/Measure | `tables/*.tmdl` |
| Parameter | What-If Parameter | `tables/*.tmdl` |
| Relationship | Model Relationship | `model.tmdl` |
| Dashboard | Report Page | `pages/<id>/page.json` |
| Worksheet | Visual Container | `visuals/<id>/visual.json` |
| Quick Filter | Slicer Visual | `visuals/<id>/visual.json` |
| Action Filter | Cross-filter / Drill-through | `visual.json` interactions |

## Vision Model Integration

The vision model performs critical functions:
- **Chart Type Detection**: Identifies bar, line, scatter, map, and other chart types from snapshots
- **Layout Extraction**: Determines exact pixel positions and dimensions for visual placement
- **Color Palette Extraction**: Captures hex values for category colors, backgrounds, and accents
- **Text Extraction**: Reads titles, axis labels, legend entries for validation against TWB
- **Encoding Detection**: Identifies how data is encoded (color, size, position, shape)

## Error Prevention Architecture

The system includes multiple validation layers:
1. **Schema Validation**: All JSON files validated against official Microsoft schemas
2. **Reference Integrity**: All visual queries reference only existing model fields
3. **Type Validation**: Data types are consistent between columns and measure references
4. **Position Validation**: All position values are non-negative numbers
5. **ID Uniqueness**: Visual and page IDs are guaranteed unique

## File Format Details

### TWBX Structure (Input)
```
Workbook.twbx (ZIP archive)
├── Workbook.twb (XML - main workbook definition)
├── Data/
│   └── Extract files (.hyper, .tde)
└── Images/ (optional embedded images)
```

### PBIP Structure (Output)
```
<Project>.pbip (JSON shortcut)
<Project>.Report/
├── definition.pbir (report metadata)
└── definition/
    ├── report.json (report settings)
    ├── version.json (version metadata)
    └── pages/
        ├── pages.json (page order)
        └── <pageId>/
            ├── page.json (page settings)
            └── visuals/
                └── <visualId>/
                    └── visual.json (visual definition)
<Project>.SemanticModel/
├── definition.pbism (model metadata)
└── definition/
    ├── database.tmdl (database settings)
    ├── model.tmdl (model + relationships)
    ├── cultures/
    │   └── en-US.tmdl (culture settings)
    └── tables/
        └── <Table>.tmdl (table + columns + measures)
```

## Calculation Translation Examples

### Tableau to DAX Mapping

| Tableau Function | DAX Equivalent |
|------------------|----------------|
| `SUM([Sales])` | `SUM(Table[Sales])` |
| `AVG([Profit])` | `AVERAGE(Table[Profit])` |
| `COUNTD([Customer ID])` | `DISTINCTCOUNT(Table[Customer ID])` |
| `IF [Profit] > 0 THEN "Positive" ELSE "Negative" END` | `IF(Table[Profit] > 0, "Positive", "Negative")` |
| `DATEPART('year', [Order Date])` | `YEAR(Table[Order Date])` |
| `DATEDIFF('day', [Order Date], [Ship Date])` | `DATEDIFF(Table[Order Date], Table[Ship Date], DAY)` |
| `ZN([Sales])` | `IF(ISBLANK(Table[Sales]), 0, Table[Sales])` |
| `CONTAINS([Category], "Tech")` | `CONTAINSSTRING(Table[Category], "Tech")` |

### Table Calculation Considerations
Tableau table calculations (RUNNING_SUM, WINDOW_AVG, INDEX, etc.) require careful translation as they depend on the visual's partition and addressing. These typically become:
- DAX window functions (SUMX with filters)
- Calculation groups for dynamic contexts
- Separate measures with explicit filter contexts
