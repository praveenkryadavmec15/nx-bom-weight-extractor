# Siemens NX BOM & Weight Extractor — User & Engineering Guide

This comprehensive guide covers practical workflows, engineering case studies, NX integration, and troubleshooting for [`Extract_BOM_Weight.py`](Extract_BOM_Weight.py).

---

## Table of Contents
1. [Workflow Overview & Architecture](#1-workflow-overview--architecture)
2. [Step-by-Step Usage Guide](#2-step-by-step-usage-guide)
3. [Setting Up a 1-Click Ribbon Button in NX](#3-setting-up-a-1-click-ribbon-button-in-nx)
4. [Real-World Case Studies](#4-real-world-case-studies)
   - [Case 1: EV Battery Pack & Cells (Direct Weight Override)](#case-1-ev-battery-pack--cells-direct-weight-override)
   - [Case 2: Custom Materials & Library Auto-Harvesting](#case-2-custom-materials--library-auto-harvesting)
   - [Case 3: Fasteners & Standard Hardware](#case-3-fasteners--standard-hardware)
   - [Case 4: Design Iterations & Part Additions/Deletions](#case-4-design-iterations--part-additionsdeletions)
5. [How Overrides & Re-Runs Work (The Sync Engine)](#5-how-overrides--re-runs-work-the-sync-engine)
6. [Teamcenter vs. Native NX Attribute Resolution](#6-teamcenter-vs-native-nx-attribute-resolution)
7. [Reference Sets & Geometry Measurement Rules](#7-reference-sets--geometry-measurement-rules)
8. [Excel Column & Formula Reference](#8-excel-column--formula-reference)
9. [Troubleshooting & Frequently Asked Questions](#9-troubleshooting--frequently-asked-questions)

---

## 1. Workflow Overview & Architecture

Engineering organizations often enforce strict IT environments where corporate laptops cannot run `pip install` or download external binary packages. Furthermore, native CAD BOM exports are typically one-way: once exported, any manual corrections you make in Excel are discarded whenever you re-export.

`Extract_BOM_Weight.py` addresses these limitations using a **pure native architecture**:

```
[Siemens NX Assembly]
       │
       ▼ (Reads Geometry, Ref Sets, Mass Properties, TC Attributes)
[Extract_BOM_Weight.py (NX Journal)]
       │
       ├──► Reads Existing Desktop\NX_BOM_Staircase_Master.xls
       │    └──► Extracts previous user overrides & custom materials
       │    └──► Archives old file with timestamp (YYYYMMDD_HHMMSS)
       │
       └──► Generates Fresh Interactive Excel (.xls XML 2003)
            ├──► Worksheet "BOM": 22 columns with staircase indentation
            └──► Worksheet "Materials Library": Dropdown table & lookup
```

Because the output is generated in **Microsoft Excel XML Spreadsheet 2003** format:
- It opens natively in Microsoft Excel, LibreOffice Calc, and Google Sheets without security warnings.
- It supports embedded formulas (`VLOOKUP`, conditional logic, cell arithmetic).
- It embeds native Excel data validation (dropdown list in column 17 referencing the Materials Library).
- It requires **zero third-party Python packages** (uses Python's built-in `xml.etree.ElementTree` and `os`).

---

## 2. Step-by-Step Usage Guide

### First-Time Run
1. Launch **Siemens NX** and load your top-level assembly.
2. Ensure the assembly structure is fully loaded or set to load lightweight components on demand.
3. In NX, press **`Alt + F8`** (or go to **Menu** $\rightarrow$ **Tools** $\rightarrow$ **Journal** $\rightarrow$ **Play...**).
4. Browse to the directory containing [`Extract_BOM_Weight.py`](Extract_BOM_Weight.py) and click **Run**.
5. Watch the NX **Information Window / Listing Window** for execution progress:
   - Identifies active work part.
   - Traverses assembly tree recursively.
   - Resolves reference sets and calculates solid bodies mass/volume.
   - Writes `NX_BOM_Staircase_Master.xls` to your Desktop.
6. Open the file on your Desktop. You will see:
   - The top banner with root Part Number and Revision.
   - The staircase hierarchy showing assembly nesting levels.
   - Automatically extracted quantities, materials, densities, and weights.

---

## 3. Setting Up a 1-Click Ribbon Button in NX

Instead of opening the Journal dialog every time, you can create a dedicated button in your NX Ribbon:

1. In Siemens NX, press **`Ctrl + 1`** (or right-click the empty gray area of any ribbon tab and select **Customize...**).
2. In the Customize dialog, switch to the **Commands** tab.
3. In the Categories list on the left, scroll down and click **New Item**.
4. Drag **New User Command** from the Commands list onto any Ribbon tab (e.g., *Assemblies* or *Home*).
5. Right-click the newly placed button on the ribbon:
   - **Edit Action**: Change Type to **Journal File**, then browse and select `Extract_BOM_Weight.py`.
   - **Edit Name**: Change the display text to `Export BOM & Weight`.
   - **Edit Button Display**: Choose *Image and Text*.
   - **Change Icon**: Pick an icon (such as an Excel sheet, scale, or list icon).
6. Click **Close** on the Customize dialog.

Now you can generate and update your BOM spreadsheet with a single click at any time!

---

## 4. Real-World Case Studies

### Case 1: EV Battery Pack & Cells (Direct Weight Override)
**The Problem**:
In an EV battery pack CAD assembly, each individual lithium-ion cell (e.g., 21700 or prismatic format) is typically modeled as an exterior aluminum or steel enclosure. Modeling the internal electrode winding (anode, cathode, separator) and liquid electrolyte in CAD would balloon the file size and degrade graphics performance. Consequently, CAD-computed weight for a cell might register as only `0.015 kg` (just the can), whereas the real physical cell weighs `0.069 kg`.

**The Solution**:
1. Run the script to generate `NX_BOM_Staircase_Master.xls`.
2. Locate the cell row in the BOM sheet.
3. In **Column 20 (`Override Weight/Part (kg)`)**, enter the exact physical cell weight: `0.069`.
4. Excel immediately recalculates:
   - **Column 21 (`Weight/Part`)** reflects `0.069`.
   - **Column 22 (`Total Weight`)** multiplies `0.069 * Qty` (e.g., `0.069 * 440 = 30.360 kg`).
5. Next week, when CAD updates the cell count or bracket geometry, **re-run the journal**. The script remembers that you assigned `0.069 kg` to this cell part number and keeps your override intact!

---

### Case 2: Custom Materials & Library Auto-Harvesting
**The Problem**:
You have specialized materials (e.g., thermal interface materials, potting polyurethane, busbar copper alloys) that might not be formally assigned in the NX material library.

**The Solution**:
- **Option A: Select from Dropdown**:
  In **Column 17 (`Override Material`)**, click any cell. A dropdown arrow appears containing all materials from the `Materials Library` sheet. Selecting a material automatically updates **Column 19 (`Active Density`)** via `VLOOKUP`.
- **Option B: Enter a Direct Density Override**:
  In **Column 18 (`Override Density`)**, type any density in $\text{kg/mm}^3$ (e.g., `0.00000210`). The `Active Density` column prioritizes this number immediately.
- **Option C: Add to Materials Library Sheet**:
  Switch to the `Materials Library` worksheet in Excel, add a new row at the bottom (e.g., `SILICONE GAP PAD` and `0.00000180`). When you re-run the journal, the script reads this custom entry back into its internal library and preserves it in future runs!

---

### Case 3: Fasteners & Standard Hardware
**The Problem**:
Standard fasteners (bolts, nuts, washers, rivnuts) often clutter BOMs and frequently have generic densities or simplified thread representations.

**The Solution**:
- The script automatically inspects component names for keywords (`BOLT`, `SCREW`, `NUT`, `WASHER`, `INSERT`, `RIVET`, `FASTENER`, `STUD`).
- It flags their code as **`H`** (Hardware) and category as **`Mech.`**.
- If no material density is assigned in NX, the script checks its default table for standard steel/stainless steel density.

---

### Case 4: Design Iterations & Part Additions/Deletions
**The Problem**:
You spend hours annotating an exported BOM with supplier notes and cell weights. Then your design team adds 4 new cooling pipe brackets and removes an old connector. A typical CAD export would wipe out your entire spreadsheet.

**The Solution**:
With this tool:
1. Save and **close** the existing Excel file.
2. Run `Extract_BOM_Weight.py` in NX.
3. The script:
   - Creates a timestamped backup copy (`NX_BOM_Staircase_Master_20260909_211500.xls`).
   - Parses the existing overrides for every part instance.
   - Extracts the newly added brackets and removes the deleted connector.
   - Re-applies your cell weight overrides and material choices to the matching parts.
   - Writes the new updated master spreadsheet.

---

## 5. How Overrides & Re-Runs Work (The Sync Engine)

The sync engine uses an **Instance Key** tracking strategy:

$$\text{Instance Key} = \text{PartNumber}\_\text{OccurrenceIndex}$$

For example, if part `BATT-MOD-01` appears multiple times across different subassemblies:
- First appearance: `BATT-MOD-01_1`
- Second appearance: `BATT-MOD-01_2`

When the journal executes:
1. `backup_and_read_existing(filepath)` opens the current `NX_BOM_Staircase_Master.xls` using `xml.etree.ElementTree`.
2. It parses the previous rows, storing:
   - Part Number & occurrence sequence.
   - Column 17 (`Override Material`).
   - Column 18 (`Override Density`).
   - Column 20 (`Override Weight/Part`).
   - All custom rows in the `Materials Library` sheet.
3. When the new NX structure is traversed, matching instance keys receive their saved overrides.
4. If new components were added in CAD, they are appended to their appropriate structural position with default values ready for review.

---

## 6. Teamcenter vs. Native NX Attribute Resolution

The script handles both Teamcenter-managed environments and standalone native NX files seamlessly:

### Teamcenter Attributes (Priority 1)
| Data Field | Teamcenter User Attribute | Fallback if Missing |
|:---|:---|:---|
| **Part Number** | `DB_PART_NO` | Part leaf name (e.g. `001234/01` $\rightarrow$ `001234`) |
| **Revision** | `DB_PART_REV` | Extracted from slash suffix or last 2 digits of leaf |
| **Part Title / Name** | `DB_PART_NAME` | Suffix of leaf name after dash `-` or `DisplayName` |
| **Material** | `NXOpen.UF.Sf.LocateMaterial` | `Material` or `DB_MATERIAL` attribute |

### Native Unmanaged NX (Priority 2)
If running outside Teamcenter, the script automatically parses the NX part leaf string:
- `1004567/A` $\rightarrow$ Part Number: `1004567`, Revision: `A`
- `1004567_rev01` $\rightarrow$ Extracts item and revision safely.

---

## 7. Reference Sets & Geometry Measurement Rules

Reference Sets are an NX feature used to control which geometry represents a part in an assembly (e.g., solid bodies, sheet bodies, or lightweight faceting).

The script evaluates geometry in the following order:

1. **Target Reference Set Search**:
   It inspects the reference sets defined in `TARGET_REFSETS`:
   ```python
   TARGET_REFSETS = ["Ahead Model", "AHEAD_MODEL", "AHEAD MODEL", "MODEL"]
   ```
   *(Matches are case-insensitive and ignore delimiter variations).*

2. **Body Filtering**:
   Only bodies that satisfy all of the following conditions are measured:
   - `body.IsSolidBody == True` (excludes sheet bodies, surfaces, sketches, and datums).
   - `body.IsBlanked == False` (excludes hidden/suppressed reference geometry).

3. **Fallback to All Bodies**:
   If no specified reference set exists in the part, the script automatically inspects all unblanked solid bodies in `part_obj.Bodies`.

4. **Mass Properties Measurement**:
   Measurements use `work_part.MeasureManager.NewMassProperties` with metric units:
   - Volume: $\text{mm}^3$
   - Mass: $\text{kg}$
   - Calculated CAD Density: $\rho = \frac{\text{Mass}}{\text{Volume}}$ ($\text{kg/mm}^3$)

---

## 8. Excel Column & Formula Reference

The table below describes how formulas in the output spreadsheet interact:

| Column | Name | Formula / Behavior |
|:---:|:---|:---|
| **Col 14** | Total Volume | `=RC[-2]*RC[-1]` (Volume/Part $\times$ Qty) |
| **Col 17** | Override Material | Data validation dropdown referencing `=MatDropdownList` |
| **Col 18** | Override Density | Direct numeric input cell |
| **Col 19** | Active Density | `=IF(ISNUMBER(RC18), IF(RC18>0, RC18, RC16), IF(ISNA(VLOOKUP(RC17, MatLookupTable, 2, FALSE)), RC16, VLOOKUP(RC17, MatLookupTable, 2, FALSE)))` |
| **Col 20** | Override Weight/Part | Direct numeric input cell (e.g., physical cell weight) |
| **Col 21** | Weight/Part | `=IF(ISNUMBER(RC20), IF(RC20>0, RC20, RC19*RC13), RC19*RC13)` |
| **Col 22** | Total Weight | `=RC21*RC12` (Weight/Part $\times$ Qty) |

### Formula Logic Explained:
- **Active Density (Col 19)**:
  1. If you entered a number in **Override Density (Col 18)** $> 0$, use it.
  2. Otherwise, check if an **Override Material (Col 17)** was picked: lookup its density in `MatLookupTable`.
  3. If no override was entered, fall back to **NX CAD Density (Col 16)**.
- **Weight/Part (Col 21)**:
  1. If you entered a number in **Override Weight/Part (Col 20)** $> 0$, use it directly.
  2. Otherwise, calculate: $\text{Active Density} \times \text{Volume/Part}$.

---

## 9. Troubleshooting & Frequently Asked Questions

### Q: I ran the script, but I received an error: `[Errno 13] Permission denied`
**Cause**: The master spreadsheet (`NX_BOM_Staircase_Master.xls`) is currently open in Microsoft Excel. Windows locks open Excel files against renaming or overwriting.  
**Fix**: Save and close the file in Excel, then re-run the journal in NX.

### Q: Some components report 0.00 volume and 0.00 weight. Why?
**Cause**:
1. The component is unloaded or lightweight.
2. The reference set in the component does not match any name in `TARGET_REFSETS` and has no active solid body.
3. The component is an empty container assembly (which correctly has zero direct volume).  
**Fix**:
- Ensure components are loaded in NX (*Assembly Navigator* $\rightarrow$ right-click $\rightarrow$ *Open* or *Load*).
- Verify the solid body belongs to the `MODEL` or `Ahead Model` reference set.

### Q: How do I change the output path from Desktop to a shared network drive?
Open `Extract_BOM_Weight.py` in any text editor and edit line 313:
```python
# Default:
# desktop = os.path.join(os.path.expanduser("~"), "Desktop")

# Custom path:
desktop = r"P:\Engineering\Battery_Projects\BOM_Exports"
```

### Q: Can I run this script in batch / headless mode?
Yes. You can invoke the NX journal from the command prompt or an automated script using `run_journal.exe`:
```powershell
"C:\Program Files\Siemens\NX\NXBIN\run_journal.exe" "Extract_BOM_Weight.py"
```

---

*Need help or found a bug? Please open an issue on GitHub.*

