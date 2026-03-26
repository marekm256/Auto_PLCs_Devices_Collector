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
- Purpose: scan archive folders, detect PLC project types, and copy each matched project folder into a normalized structure.
- Input:
  - `source_root` (folder to scan recursively)
  - `target_root` (where copied projects are saved)
  - optional ignored folders: `"IgnoreA;IgnoreB"` (semicolon-separated)
- Searches for project files:
  - STEP7: `*.s7p`
  - TIA: `*.ap10`, `*.ap11`, `*.ap12`, `*.ap13`, `*.ap14`, `*.ap15`, `*.ap15_1`, `*.ap16`, `*.ap17`, `*.ap18`, `*.ap19`, `*.ap20`
  - Rockwell: `*.acd`
- Excludes from search:
  - any path containing `$`
  - folder name `Source Codes Archive` (default)
  - any custom folder names passed in optional ignore argument
- Project folder naming (short logic):
  - uses source folder path segments and looks for machine code format: `^[A-Z0-9]{7}$`
  - first match in segment `4`, then `5`
  - if no match: uses first non-empty from segment `6`, `5`, `4`, source folder name, else `UNKNOWN`
  - if same name already exists in target: adds version suffix (`_v2`, `_v3`, ...)
- Output tree:
```text
target_root/
+- step7/
|  +- MACHINE01/
|  +- MACHINE01_v2/
+- tia10/
+- tia11/
+- tia12/
+- tia13/
+- tia14/
+- tia15/
+- tia15_1/
+- tia16/
+- tia17/
+- tia18/
+- tia19/
+- tia20/
+- rockwell/
```
- Simple example:
```powershell
ProjectsScannerCopier.exe "D:\Archive\Projects" "D:\Collected\collected_projects" "Backup;Old"
```
Scans `D:\Archive\Projects`, skips folders containing `Backup`, `Old`, `Source Codes Archive`, and copies found project folders into typed subfolders under `D:\Collected\collected_projects`.

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

