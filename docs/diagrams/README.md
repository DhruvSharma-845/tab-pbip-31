# Architecture Diagrams

This folder contains architecture diagrams for the Tableau TWBX to Power BI PBIP conversion system.

## Diagram Files

| File | Description |
|------|-------------|
| `system_architecture.mmd` | High-level system architecture showing input, extraction, transformation, and output layers |
| `data_flow_pipeline.mmd` | Four-phase data flow: Ingestion → Parsing → Transformation → Generation |
| `schema_mapping.mmd` | Mapping between Tableau TWB XML elements and Power BI PBIP components |
| `visual_type_mapping.mmd` | Decision tree for mapping Tableau mark types to Power BI visual types |
| `file_structure.mmd` | File structure comparison between TWBX input and PBIP output |
| `validation_layers.mmd` | Five-layer validation architecture for error prevention |

## Viewing the Diagrams

### Option 1: GitHub / GitLab
Mermaid diagrams render automatically in GitHub and GitLab markdown files.

### Option 2: VS Code
Install the "Mermaid Preview" or "Markdown Preview Mermaid Support" extension.

### Option 3: Mermaid Live Editor
Copy the `.mmd` file contents to [mermaid.live](https://mermaid.live) for interactive editing and export.

### Option 4: CLI Export
Use the Mermaid CLI to export to PNG/SVG:
```bash
npm install -g @mermaid-js/mermaid-cli
mmdc -i system_architecture.mmd -o system_architecture.png
```

## Existing Image Files

| File | Description |
|------|-------------|
| `twbx_to_pbip_architecture.jpg` | Original architecture diagram |
| `pbip_assembly_architecture.jpg` | PBIP assembly process |
| `twb_vs_pbip_schema_logic.jpg` | Schema comparison diagram |

## Updating Diagrams

The `.mmd` files use [Mermaid](https://mermaid.js.org/) syntax. Key diagram types used:

- **flowchart TB/LR** - Top-to-bottom or left-to-right flow diagrams
- **subgraph** - Grouped sections
- **classDef** - Custom styling for nodes

When updating, ensure:
1. Node IDs are unique within the diagram
2. Connections use proper syntax (`-->`, `==>`, `-.->`)
3. Labels with special characters are quoted
