# Spreadsheet_Comparator
R-AM-PJ

# Spreadsheet version comparator

## Features

- Compare multiple spreadsheet files at once.
- Detect:
  - Changed cell values
  - Missing cells
  - Missing rows
  - Newly added rows
- Highlight differences directly in the output workbook.
- Support parallel processing with up to 4 workers.
- Display real-time comparison progress.
- Save application settings automatically.
- Automatically open the output folder after comparison.


## How to Use

[Download Xlsx Comparator](https://rak.box.com/s/m9u5e1k53893m2rj3pk2mj54ay4n25rn)

1. Run `ExcelComparator.exe`
2. Select the **Old Version Folder** and **New Version Folder**.
3. Select a **Pairing Method**:
   - **Filename**: Match files based on their filenames.
   - **Folder Order**: Match files according to their order in each folder.
4. Select whether to enable **Parallel Processing**.
5. Click **Compare**.
6. The application compares each matched pair and creates the results in the output folder.


## Requirements

- Pair of Old/New version spreadsheet - stored in separated folder
- Network

No Python or any other tools required.


## Output

Default output path is where your New Version Folder is.


### Color Summary

| Color | Meaning |
|---|---|
| Normal | No difference |
| Light Red (`#F4CCCC`) | Changed cell |
| Light Red (`#F4CCCC`) | Missing cell |
| Light Red (`#F4CCCC`) | Missing Old Version row |
| Light Blue (`#CFE2F3`) | Added New Version row |

The colors are only used to visually indicate differences. The New Version remains the primary source of the output data.


## Notes

- Only `.xlsx` and `.csv` files are supported.
- Files inside subfolders are not included.
- Unsupported files are ignored. No need to clean.


## Version

**Xlsx Comparator v1.0**


## Key Logic


<img src="./images/mermaid-diagram.png" alt="Diff Detecting Algorithm" width="700"/>

### Difference Rules
| Old Version | New Version | Result |
|---|---|---|
| Same value | Same value | No difference |
| Value | Different value | Changed cell |
| Value | Empty | Missing cell |
| Empty | Value | Ignored |
| Empty | Empty | Ignored |
| Row exists | Row exists | Compare cells |
| Row exists | Row missing | Missing row |
| Row missing | Row exists | Added row |

## Authors

Sana(Sua Oh)