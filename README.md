# Power BI Report Export, PBIX→PBIP Conversion & Similarity Analyzer

## 1. Overview
This project automates three stages of Power BI report processing:
1. Export PBIX reports from a specified Power BI workspace.
2. Convert downloaded PBIX files to PBIP (project) format using UI automation.
3. Compute similarity metrics between PBIP report projects to identify duplicates and master versions.

Outputs:
- Downloaded PBIX files: `./download_pbix_files`
- Converted PBIP project folders: `./converted_pbip_files`
- Similarity results Excel: `<OUTPUT_PATH>/Report_similarity_matrix.xlsx`

## 2. Repository Structure
```
DeDuplication_Report_Similarity.py   # Python similarity & deduplication logic
export_reports.ps1                   # Script to export PBIX reports
PBIXtoPBIP_PBITConversion.psm1       # PowerShell module for PBIX→PBIP conversion (UI automation)
run_exporter_console.ps1             # Interactive console wrapper (export & optional conversion)
setup_dependencies.ps1               # One-time setup script to install all required PowerShell modules
converted_pbip_files/                # Generated PBIP project folders
download_pbix_files/                 # Exported PBIX files
requirements.txt                     # Python package requirements
```

## 3. Prerequisites
### PowerShell
- Windows PowerShell 5.1 (default) or newer
- Sufficient permissions to access the target Power BI workspace (Viewer or higher)
- PowerShell Modules (auto-installed by setup script):
  - NuGet package provider (v2.8.5.201+)
  - PSGallery repository (trusted)
  - MicrosoftPowerBIMgmt (with sub-modules)
  - PS2EXE (optional, for building .exe)

### Power BI Desktop
- Installed locally (required for PBIX → PBIP UI automation conversion)

### Python
- Python 3.9+ (recommended)
- Packages: `pandas`, `openpyxl`

Install Python dependencies:
```powershell
pip install -r requirements.txt
```
(If you do not use the `requirements.txt`, run: `pip install pandas openpyxl`)

## 4. Step 1: Install Required PowerShell Modules
**IMPORTANT: Run this first before using any export or conversion scripts.**

The repository includes `setup_dependencies.ps1` which automatically installs all required PowerShell modules and configures your environment.

Run the setup script:
```powershell
# From repository root
.\setup_dependencies.ps1
```

This script will:
- Set execution policy for the current session
- Install NuGet package provider
- Configure PSGallery as a trusted repository
- Install MicrosoftPowerBIMgmt and related modules (with AcceptLicense)
- Optionally install PS2EXE for building executables
- Verify all installations

**Note:** You only need to run this once per machine, unless modules need updating.

## 5. Step 2: Export PBIX Reports (Optional Immediate Conversion)
Use the console wrapper script to:
- Prompt for Workspace Id
- Export reports to `./download_pbix_files`
- Optionally convert them to PBIP into `./converted_pbip_files`

Run:
```powershell
# From repository root
.\run_exporter_console.ps1
```
Follow on-screen prompts.

### Behind the Scenes
- The script automatically sets execution policy for the current session
- `export_reports.ps1` handles downloading PBIX via Power BI REST/API calls (requires proper authentication context)
- Power BI modules are automatically imported if already installed (run `setup_dependencies.ps1` first if not)
- If conversion selected, the module `PBIXtoPBIP_PBITConversion.psm1` is imported and automates Power BI Desktop to produce PBIP projects

## 6. (Optional) Build PowerShell Exporter as .EXE
Install PS2EXE once per environment:
```powershell
Install-Module PS2EXE -Scope CurrentUser
```
Package the script:
```powershell
Invoke-PS2EXE -InputFile .\run_exporter_console.ps1 -OutputFile .\ReportExportTool.exe -Title "Power BI Report Exporter" -Company "" -Product "" -Copyright "" -IconFile "" -NoConsole:$false -STA
```
Then run the generated `ReportExportTool.exe` directly.

## 7. Step 3: Configure & Run Similarity Analysis
Open `DeDuplication_Report_Similarity.py` in VS Code.

Two ways to set paths:
1. Interactive: Uncomment the lines:
```python
# REPORTS_ROOT = get_reports_root()
# OUTPUT_PATH = get_output_path()
```
2. Static (current default): Replace the hardcoded values:
```python
REPORTS_ROOT = r"<AbsolutePathToConvertedPBIPFolders>"
OUTPUT_PATH = r"<AbsolutePathWhereExcelShouldBeSaved>"
```

Run analysis:
```powershell
python .\DeDuplication_Report_Similarity.py
```
Result Excel: `<OUTPUT_PATH>\Report_similarity_matrix.xlsx`
Console output includes:
- Pairwise similarity matrix (written to Excel)
- Groups above thresholds (70%, 80%, 90%, 95%, 100%)
- Candidate masters (reports covering others)
- Reports to keep vs. eligible for elimination

### Key Tunable Variables
Inside `DeDuplication_Report_Similarity.py`:
- `VISUAL_MATCH_THRESHOLD` (default 0.9): Jaccard similarity needed to consider two visuals a match.
- `GROUP_THRESHOLDS = [0.7, 0.8, 0.9, 0.95, 1.0]`: Clustering similarity bands.
- `MASTER_THRESHOLD = 0.95`: Visual match threshold when evaluating master coverage.

Adjust as needed for stricter or looser matching.

## 8. (Optional) Build Python Similarity Script as .EXE
```powershell
pip install pyinstaller
pyinstaller --onefile .\DeDuplication_Report_Similarity.py
```
Generated binary appears under `dist/DeDuplication_Report_Similarity.exe`.
Distribute along with instructions for providing proper folder paths.

## 9. Recommended Workflow
1. **First-time setup**: Run `.\setup_dependencies.ps1` to install all required PowerShell modules.
2. Run `.\run_exporter_console.ps1` → export PBIXs.
3. Choose conversion → produce PBIP folders in `converted_pbip_files`.
4. Review converted structure; ensure all report folders end with `.Report` (or consistent naming).
5. Set `REPORTS_ROOT` to `converted_pbip_files` absolute path.
6. Choose/prepare an output folder and set `OUTPUT_PATH`.
7. Run similarity script: `python .\DeDuplication_Report_Similarity.py`
8. Open Excel matrix → identify high-similarity pairs/groups.
9. Use console output lists of masters vs. eliminations to decide archival/deletion.

## 10. Troubleshooting
- **Module installation fails**: Run `setup_dependencies.ps1` with administrator privileges if needed, or check internet connectivity for downloading from PSGallery.
- **Execution policy errors**: The scripts set execution policy to Bypass for the current session automatically. If issues persist, run: `Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force`
- **Power BI authentication fails**: Ensure you have appropriate permissions in the target workspace and valid credentials.
- **Excel not generated**: Verify `OUTPUT_PATH` exists and you have write permissions.
- **Missing PBIP folders**: Ensure Power BI Desktop was installed and UI automation not blocked (no modal dialogs).
- **Similarity all zeros**: Check that PBIP folders contain JSON visual definitions; conversion may have failed.
- **Import-Module fails**: Confirm full absolute path and file is not blocked (Unblock-File if downloaded).
- **Pandas Excel write error**: Install `openpyxl` (`pip install openpyxl`).

## 11. Security & Access Notes
- Workspace export requires appropriate Power BI service permissions.
- Avoid storing credentials in plain text scripts.
- Run scripts from a trusted environment; UI automation can be sensitive to focus changes.

## 12. Cleaning Up
- Delete temporary `build/` and `dist/` folders after PyInstaller if space is a concern.
- Archive original PBIX files after successful PBIP conversion to reduce redundancy.

## 13. Requirements Summary
PowerShell Modules (installed via `setup_dependencies.ps1`):
- NuGet package provider (v2.8.5.201+)
- MicrosoftPowerBIMgmt (with AcceptLicense)
- MicrosoftPowerBIMgmt.Reports
- MicrosoftPowerBIMgmt.Workspaces
- MicrosoftPowerBIMgmt.Profile
- PS2EXE (optional, for building executables)

Local PowerShell Modules:
- `PBIXtoPBIP_PBITConversion.psm1` (included in repository)

Python Packages:
- `pandas`
- `openpyxl` (Excel writer engine)

## 14. License / Attribution
Internal use. Do not redistribute outside authorized organization contexts unless explicitly permitted.

## 15. Next Enhancements (Ideas)
- Add argparse CLI for similarity script (pass paths & thresholds without editing code).
- Add REST auth & token management wrapper for export script.
- Produce HTML summary report with charts of similarity distribution.
- Integrate automatic duplicate archiving logic.

---
Feel free to update paths or thresholds to match your environment. Let me know if you’d like a CLI version or additional automation.
