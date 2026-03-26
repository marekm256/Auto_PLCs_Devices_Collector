# Auto_PLCs_Devices_Collector

Minimal guide for Siemens device inventory pipeline.

## Data Pipeline
1. `ProjectsScannerCopier` scans input folder and copies PLC projects into structured collection folders.
2. `Tia15-21_DevicesExporter` exports found devices in JSONs from Siemens TIA projects (`ap15`..`ap21`).
3. `Step7_ConfigsExporter` (Python, GUI automation) opens STEP7 projects and exports `.cfg` files.
4. `Step7_DevicesExporter` parses STEP7 `.cfg` files and exports devices in JSONs.
5. `Devices_Processor` from gathered JSONs creates formatted Excel files and enriches data with Siemens metadata (image, lifecycle, description).

Current scope:
- Siemens projects and Siemens devices.
- Planned extension: Rockwell and additional vendors.

## Apps Overview

### ProjectsScannerCopier
- Purpose: discover and copy project files from source archive into standardized folders.
- Output: organized project folders for next pipeline steps.

### Tia15-21_DevicesExporter
- Purpose: adapter for TIA versions 15..21.
- Input: collected TIA project folders.
- Output: JSON files with `Device`, `DeviceItem`, `OrderNumber`, `Firmware`.

### Step7_ConfigsExporter (Python)
- Purpose: automated GUI export of STEP7 hardware configs.
- Input: STEP7 projects.
- Output: `.cfg` files in `Step7_ConfigsExporter/Configs`.

### Step7_DevicesExporter
- Purpose: parse exported STEP7 `.cfg` files and extract device inventory.
- Input: `.cfg` files from Step7_ConfigsExporter.
- Output: JSON files with standardized device rows.

### Devices_Processor
- Purpose: final processing of JSON inventory.
- Input: `collected_components`.
- Output:
  - formatted Excel files (`collected_excels`)
  - downloaded image cache (`downloaded_images`)
- Metadata source: Siemens product pages (Lifecycle, Description, product image).
