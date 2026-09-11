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
            ├──► Worksheet "BOM": 23 columns with staircase indentation
            └──► Worksheet "Materials Library": Dropdown table & lookup
```

Because the output is generated in **Microsoft Excel XML Spreadsheet 2003** format:
- It opens natively in Microsoft Excel, LibreOffice Calc, and Google Sheets without security warnings.
- It supports embedded formulas (`VLOOKUP`, conditional logic, cell arithmetic).
- It embeds native Excel data validation (dropdown list in column 18 referencing the Materials Library).
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
   - Deep scans for materials via NX API, text attributes, and custom aliases (`MATERIAL1`).
   - Traverses assembly tree recursively with mutually exclusive classification.
   - Resolves reference sets and calculates solid bodies mass/volume.
   - Writes `NX_BOM_Staircase_Master.xls` to your Desktop.
   ```text
   Processing BOM with 'MATERIAL1' extraction...
   --------------------------------------------------
   SUCCESS! Excel BOM generated cleanly.
   File updated: C:\Users\<Username>\Desktop\NX_BOM_Staircase_Master.xls
   --------------------------------------------------
   ```
6. Open the file on your Desktop. You will see:
   - The top banner with root Part Number and Revision.
   - The staircase hierarchy showing assembly nesting levels.
   - Automatically classified item types (`A`, `C`, `H`, `P`) and disciplines (`Mech`, `EEE`).
   - Extracted revisions, quantities, materials, densities, and weights.

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
3. In **Column 21 (`Override Weight/Part (kg)`)**, enter the exact physical cell weight: `0.069`.
4. Excel immediately recalculates:
   - **Column 22 (`Weight/Part`)** reflects `0.069`.
   - **Column 23 (`Total Weight`)** multiplies `0.069 * Qty` (e.g., `0.069 * 440 = 30.360 kg`).
5. Next week, when CAD updates the cell count or bracket geometry, **re-run the journal**. The script remembers that you assigned `0.069 kg` to this cell part number and keeps your override intact!

---

### Case 2: Custom Materials & Library Auto-Harvesting
**The Problem**:
You have specialized materials (e.g., thermal interface materials, potting polyurethane, busbar copper alloys) that might not be formally assigned in the NX material library.

**The Solution**:
- **Option A: Select from Dropdown**:
  In **Column 18 (`Override Material`)**, click any cell. A dropdown arrow appears containing all materials from the `Materials Library` sheet. Selecting a material automatically updates **Column 20 (`Active Density`)** via `VLOOKUP`.
- **Option B: Enter a Direct Density Override**:
  In **Column 19 (`Override Density`)**, type any density in $\text{kg/mm}^3$ (e.g., `0.00000210`). The `Active Density` column prioritizes this number immediately.
- **Option C: Add to Materials Library Sheet**:
  Switch to the `Materials Library` worksheet in Excel, add a new row at the bottom (e.g., `SILICONE GAP PAD` and `0.00000180`). When you re-run the journal, the script reads this custom entry back into its internal library and preserves it in future runs!

---

### Case 3: Hardware, Fasteners & Proprietary Parts Classification
**The Problem**:
Engineers need clear, mutually exclusive distinction between fabricated parts, standard catalog hardware, off-the-shelf proprietary items, and assembly nodes, along with proper discipline grouping (`Mech` vs `EEE`). Standard CAD exports often group everything together as generic parts or fail to recognize proprietary electrical bought-out items.

**The Solution**:
The script employs an updated, two-tier classification pipeline:

1. **Assembly / Child Part Classification (`Assly/Child Part`)**:
   - **`A` (Assembly)**: Any component that contains children or sub-assemblies.
   - **`H` (Hardware & Fasteners)**: Components whose names contain standard fastener terms:
     `BOLT`, `SCREW`, `NUT`, `WASHER`, `INSERT`, `RIVET`, `FASTENER`, `STUD`, `PIN`, `CLIP`, `STANDOFF`, `SPACER`.
   - **`P` (Proprietary & Bought-Out Parts)**: Off-the-shelf and electrical items:
     `CELL`, `FUSE`, `CONNECTOR`, `CONN `, `CONN_`, `SPLICE`, `RELAY`, `SENSOR`, `SWITCH`, `BMS`, `PCB`, `CONTACTOR`, `PLUG`, `BREAKER`, `DIODE`.
   - **`C` (Child Part)**: All other fabricated or machined detail piece-parts.

2. **Discipline Categorization (`Category`: `Mech` vs `EEE`)**:
   - **Mechanical Structural Precedence**:
     If the component name contains mechanical/enclosure keywords (`HOUSING`, `BRACKET`, `HOLDER`, `COVER`, `TAPE`, `TRAY`, `BASE`, `GASKET`, `SEAL`, `FOAM`, `CARRIER PLATE`) and no electrical override term (such as `BUSBAR`, `PCB`, `CABLE`, `WIRE`, `TERMINAL`, `CELL`, `FUSE`, `BMS`, `SHUNT`), it is classified as **`Mech`**.
   - **Battery Assembly Protection**:
     Top-level and sub-assembly nodes such as `INSTALL BATTERY`, `BATTERY ASSEMBLY`, `ASSEMBLY BATTERY`, or `CELL PACK ASSEMBLY` are explicitly categorized as **`Mech`**, preventing them from being mistakenly tagged as electrical components simply because they contain the word "BATTERY".
   - **Electrical Classification**:
     Items containing electrical keywords (`BATTERY`, `CELL`, `BUSBAR`, `WIRE`, `CABLE`, `CONN`, `SPLICE`, `TERMINAL`, `PCB`, `HARNESS`, `FUSE`, `BMS`, `RELAY`, `CONTACTOR`, `SHUNT`, `ELECTRICAL`, `TAB CELL`) are classified as **`EEE`**.
   - All other items default to **`Mech`**.

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
2. It dynamically parses header labels and previous rows, storing:
   - Part Number & occurrence sequence (Col 10).
   - Column 18 (`Override Material`).
   - Column 19 (`Override Density`).
   - Column 21 (`Override Weight/Part`).
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
| **Material** | `NXOpen.UF.Sf.LocateMaterial` | Deep alias search across bodies, prototypes & components: `MATERIAL1`, `Material`, `DB_MATERIAL`, `DB_MATERIAL_NAME`, `Material Name`, `NX_Material`, `MASSPROP_MATERIAL`, `DB_PART_MATERIAL`, `Matl` |

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

The table below describes how formulas in the 23-column output spreadsheet interact:

| Column | Name | Formula / Behavior |
|:---:|:---|:---|
| **Col 1–6** | Assembly Level | Visual staircase indicator (Levels 1 to 5+) |
| **Col 7** | Item Name | `DB_PART_NAME` or component leaf suffix |
| **Col 8** | Assly / Child Part | Classification code: `A` (Assly), `C` (Child), `H` (Hardware), `P` (Proprietary) |
| **Col 9** | Category | Discipline category: `Mech` or `EEE` |
| **Col 10** | Part Number | `DB_PART_NO` or cleaned part number |
| **Col 11** | Parent Part | Parent subassembly part number |
| **Col 12** | Qty | Instance occurrence count within parent |
| **Col 13** | Rev | Part revision (`DB_PART_REV` or parsed from leaf) |
| **Col 14** | Volume/Part (mm³) | Solid volume measured from target reference sets |
| **Col 15** | Total Volume (mm³) | `=RC[-3]*RC[-1]` (Volume/Part $\times$ Qty) |
| **Col 16** | NX Material | Extracted via native NX material tag, `MATERIAL1`, or text attributes |
| **Col 17** | NX Density (kg/mm³) | CAD density or auto-synchronized library density |
| **Col 18** | Override Material | Data validation dropdown referencing `=MatDropdownList` |
| **Col 19** | Override Density | Direct numeric input cell in $\text{kg/mm}^3$ |
| **Col 20** | Active Density | `=IF(ISNUMBER(RC19), IF(RC19>0, RC19, RC17), IF(ISNA(VLOOKUP(RC18, MatLookupTable, 2, FALSE)), RC17, VLOOKUP(RC18, MatLookupTable, 2, FALSE)))` |
| **Col 21** | Override Weight/Part | Direct numeric input cell in $\text{kg}$ (e.g., physical battery cell weight) |
| **Col 22** | Weight/Part | `=IF(ISNUMBER(RC21), IF(RC21>0, RC21, RC20*RC14), RC20*RC14)` |
| **Col 23** | Total Weight | `=RC22*RC12` (Weight/Part $\times$ Qty) |

### Formula Logic Explained:
- **Active Density (Col 20)**:
  1. If you entered a number in **Override Density (Col 19)** $> 0$, use it.
  2. Otherwise, check if an **Override Material (Col 18)** was picked from the dropdown: lookup its density in `MatLookupTable`.
  3. If no override was entered, fall back to **NX CAD Density (Col 17)**.
- **Weight/Part (Col 22)**:
  1. If you entered a number in **Override Weight/Part (Col 21)** $> 0$, use it directly.
  2. Otherwise, calculate: $\text{Active Density (Col 20)} \times \text{Volume/Part (Col 14)}$.
- **Total Weight (Col 23)**:
  Multiplies $\text{Weight/Part (Col 22)} \times \text{Qty (Col 12)}$.

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
Open `Extract_BOM_Weight.py` in any text editor and edit around line 405:
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

