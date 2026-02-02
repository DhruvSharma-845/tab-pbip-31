## PBIP Converter

Stable PBIP generator from a Tableau TWBX workbook. This output is compatible with official stable Power BI schemas.

### Usage

```
python3 /Users/dhrsharm/tab-pbip-31/convert_twbx_to_pbip.py \
  --twbx /Users/dhrsharm/tab-pbip-31/Superstore.twbx \
  --out /Users/dhrsharm/tab-pbip-31/SuperstorePBIP \
  --project-name Superstore \
  --sample-theme /Users/dhrsharm/tab-pbip-31/samplepbipfolder/Sample.Report/StaticResources/SharedResources/BaseThemes/CY25SU12.json
```

### Usage (PBIR Preview With Visuals)

```
python3 /Users/dhrsharm/tab-pbip-31/convert_twbx_to_pbip_with_visuals.py \
  --twbx /Users/dhrsharm/tab-pbip-31/Superstore.twbx \
  --out /Users/dhrsharm/tab-pbip-31/SuperstorePBIP \
  --project-name Superstore \
  --sample-theme /Users/dhrsharm/tab-pbip-31/samplepbipfolder/Sample.Report/StaticResources/SharedResources/BaseThemes/CY25SU12.json \
  --windows-data-root "C:\Users\Pooja Satija\Downloads\tab-pbip-31-code-new-visuals\tab-pbip-31-code-new-visuals\SuperstorePBIP\Data\Superstore"
```

### Usage (Auto Overview + Product)

```
python3 /Users/dhrsharm/tab-pbip-31/convert_twbx_to_pbip_auto.py \
  --twbx /Users/dhrsharm/tab-pbip-31/Superstore.twbx \
  --out /Users/dhrsharm/tab-pbip-31/SuperstorePBIP \
  --project-name Superstore \
  --sample-theme /Users/dhrsharm/tab-pbip-31/samplepbipfolder/Sample.Report/StaticResources/SharedResources/BaseThemes/CY25SU12.json \
  --windows-data-root "C:\Users\Pooja Satija\Downloads\tab-pbip-31-code-new-visuals\tab-pbip-31-code-new-visuals\SuperstorePBIP\Data\Superstore"
```

Command used:
```
python3 /Users/dhrsharm/tab-pbip-31/convert_twbx_to_pbip_auto.py \
  --twbx /Users/dhrsharm/tab-pbip-31/Superstore.twbx \
  --out /Users/dhrsharm/tab-pbip-31/SuperstorePBIP \
  --project-name Superstore \
  --sample-theme /Users/dhrsharm/tab-pbip-31/samplepbipfolder/Sample.Report/StaticResources/SharedResources/BaseThemes/CY25SU12.json \
  --windows-data-root "C:\Users\Pooja Satija\Downloads\tab-pbip-31-code-new-visuals\tab-pbip-31-code-new-visuals\SuperstorePBIP\Data\Superstore"
```

### Usage (Auto Convert PBIP Visuals)

```
python3 /Users/dhrsharm/tab-pbip-31/auto_convert_pbip_visuals.py \
  --pbip-report /Users/dhrsharm/tab-pbip-31/SuperstorePBIP/Superstore.Report \
  --snapshots-dir "/Users/dhrsharm/tab-pbip-31/tableau snapshots" \
  --out-report /Users/dhrsharm/tab-pbip-31/out/pbip_visual_conversion_report.json
```
