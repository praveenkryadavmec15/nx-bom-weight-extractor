"""
Siemens NX & Teamcenter BOM and Weight Extractor (NX Journal)
============================================================
Automated Bill of Materials (BOM) hierarchy extraction and mass/volume calculator
with two-way persistent Excel overrides and zero external dependencies.

Compatible with Siemens NX (NX 11 through NX 2406+) and Teamcenter / Native NX.
Runs on native Python bundled with Siemens NX.
"""

import os
import datetime
import xml.etree.ElementTree as ET
import NXOpen
import NXOpen.UF

XML_EXCEL_NAME = "NX_BOM_Staircase_Master.xls"
INCLUDE_REVISION_IN_TITLE = True

# Updated to include all case and delimiter variations of your reference sets
TARGET_REFSETS = ["Ahead Model", "AHEAD_MODEL", "AHEAD MODEL", "MODEL"]

DEFAULT_MATERIALS = {
    "STEEL (MILD / CARBON)": 0.00000785,
    "STAINLESS STEEL 304": 0.00000800,
    "ALUMINUM 6061": 0.00000270,
    "BRASS / FREE CUTTING BRASS": 0.00000847,
    "COPPER": 0.00000896,
    "ABS PLASTIC": 0.00000105,
    "NYLON (PA6)": 0.00000114,
    "POLYCARBONATE (PC)": 0.00000120,
    "POM (DELRIN / ACETAL)": 0.00000142,
    "RUBBER / EPDM": 0.00000115
}

def clean_material_name(raw_name):
    if not raw_name: return ""
    parts = [p.strip() for p in raw_name.split(":")]
    return parts[0]

def backup_and_read_existing(filepath):
    saved_order = []
    overrides = {}
    pno_counts = {}
    existing_mat_lib = {}

    if not os.path.exists(filepath):
        return saved_order, overrides, existing_mat_lib

    try:
        with open(filepath, 'r', encoding='utf-8', errors='ignore') as f:
            content = f.read()

        if not content.startswith("PK"):
            root = ET.fromstring(content)
            for ws in root:
                ws_name = ws.attrib.get('{urn:schemas-microsoft-com:office:spreadsheet}Name', '')
                
                if ws_name == 'Materials Library':
                    for child in ws:
                        if child.tag.endswith('Table'):
                            for row in child:
                                if not row.tag.endswith('Row'): continue
                                c_idx = 1
                                m_name, m_den = "", ""
                                for cell in row:
                                    if not cell.tag.endswith('Cell'): continue
                                    idx_attr = next((v for k,v in cell.attrib.items() if k.endswith('Index')), None)
                                    if idx_attr: c_idx = int(idx_attr)
                                    val = ""
                                    for data in cell:
                                        if data.tag.endswith('Data'):
                                            val = data.text if data.text else ""
                                            break
                                    if c_idx == 1: m_name = val.strip()
                                    if c_idx == 2: m_den = val.strip()
                                    c_idx += 1
                                if m_name and m_name != "Material Name":
                                    try:
                                        existing_mat_lib[m_name] = float(m_den)
                                    except:
                                        pass

                if ws_name == 'BOM':
                    for child in ws:
                        if child.tag.endswith('Table'):
                            for row in child:
                                if not row.tag.endswith('Row'): continue
                                col_idx = 1
                                pno, mat_ov, den_ov, wt_ov = "", "", "", ""
                                for cell in row:
                                    if not cell.tag.endswith('Cell'): continue
                                    idx_attr = next((v for k,v in cell.attrib.items() if k.endswith('Index')), None)
                                    if idx_attr: col_idx = int(idx_attr)
                                    val = ""
                                    for data in cell:
                                        if data.tag.endswith('Data'):
                                            val = data.text if data.text else ""
                                            break
                                    if col_idx == 10: pno = val.strip()
                                    if col_idx == 17: mat_ov = val.strip()
                                    if col_idx == 18: den_ov = val.strip()
                                    if col_idx == 20: wt_ov = val.strip()
                                    col_idx += 1

                                if pno and pno != "Part Number":
                                    pno_counts[pno] = pno_counts.get(pno, 0) + 1
                                    instance_key = f"{pno}_{pno_counts[pno]}"
                                    saved_order.append(instance_key)
                                    overrides[instance_key] = {'mat': mat_ov, 'density': den_ov, 'weight': wt_ov}
    except Exception:
        pass 

    try:
        ctime = os.path.getctime(filepath)
        ts_str = datetime.datetime.fromtimestamp(ctime).strftime("%Y%m%d_%H%M%S")
        dir_name = os.path.dirname(filepath)
        base_name, ext = os.path.splitext(os.path.basename(filepath))
        archive_name = f"{base_name}_{ts_str}{ext}"
        archive_path = os.path.join(dir_name, archive_name)
        if os.path.exists(archive_path):
            archive_path = os.path.join(dir_name, f"{base_name}_{ts_str}_{int(datetime.datetime.now().timestamp())}{ext}")
        os.rename(filepath, archive_path)
    except:
        pass

    return saved_order, overrides, existing_mat_lib

def get_item_and_rev(part_obj):
    if not part_obj: return "", ""
    pno, rev = "", ""
    try:
        pno = part_obj.GetUserAttributeAsString("DB_PART_NO", NXOpen.NXObject.AttributeType.String, -1).strip()
    except: pass
    try:
        rev = part_obj.GetUserAttributeAsString("DB_PART_REV", NXOpen.NXObject.AttributeType.String, -1).strip()
    except: pass
    
    if not pno:
        leaf = part_obj.Leaf.strip()
        if "/" in leaf:
            pno, rev = leaf.split("/")[0].strip(), leaf.split("/")[1].strip()
        elif len(leaf) >= 10 and leaf[-2:].isdigit():
            pno, rev = leaf[:-2], leaf[-2:]
        else:
            pno = leaf
    return pno, rev

def get_cleaned_comp_number(child):
    proto = child.Prototype
    if proto:
        pno, _ = get_item_and_rev(proto)
        if pno: return pno
    raw = child.DisplayName
    return raw.split("/")[0].strip() if "/" in raw else raw.strip()

def get_target_bodies(part_obj, session):
    """
    Checks reference sets supporting AHEAD_MODEL, AHEAD MODEL, and MODEL.
    Falls back to unblanked solid bodies if no reference sets match.
    """
    bodies = []
    matched_rs = None
    
    for target in TARGET_REFSETS:
        for rs in part_obj.GetAllReferenceSets():
            if rs.Name.strip().lower() == target.lower():
                matched_rs = rs
                break
        if matched_rs: break

    if matched_rs:
        for obj in matched_rs.AskMembersInReferenceSet():
            body = None
            if isinstance(obj, NXOpen.Body):
                body = obj
            else:
                try:
                    tagged = session.GetObjectManager().GetTaggedObject(obj.Tag)
                    if isinstance(tagged, NXOpen.Body):
                        body = tagged
                except:
                    pass
            
            if body and body.IsSolidBody and not body.IsBlanked:
                if body not in bodies:
                    bodies.append(body)

    if not bodies:
        for b in part_obj.Bodies:
            if b.IsSolidBody and not b.IsBlanked:
                if b not in bodies:
                    bodies.append(b)
            
    return bodies

def get_material(uf_session, session, obj):
    try:
        mat_tag = uf_session.Sf.LocateMaterial(obj.Tag)
        if mat_tag != NXOpen.Tag.Null:
            mat_obj = session.GetObjectManager().GetTaggedObject(mat_tag)
            if hasattr(mat_obj, "Name"): return mat_obj.Name
    except: pass
    return ""

def classify_component(name, is_assembly, has_geometry):
    uname = name.upper()
    if is_assembly and not has_geometry: return "A", "Mech."
    if any(h in uname for h in ["BOLT", "SCREW", "NUT", "WASHER", "INSERT", "RIVET", "FASTENER", "STUD"]): return "H", "Mech."
    if any(e in uname for e in ["BATTERY", "CELL", "BUSBAR", "WIRE", "CABLE", "CONN", "SPLICE", "TERMINAL", "PCB", "HARNESS", "EEE"]): return "P", "EEE"
    if is_assembly: return "A", "Mech."
    return "C", "Mech."

def measure_geometry(child, work_part, session, uf_session):
    proto = child.Prototype
    if not proto or not proto.IsFullyLoaded:
        try: uf_session.Part.OpenQuiet(child.DisplayName)
        except: pass
        proto = child.Prototype

    comp_name, part_no = "", ""
    if proto:
        part_no, _ = get_item_and_rev(proto)
        try: comp_name = proto.GetUserAttributeAsString("DB_PART_NAME", NXOpen.NXObject.AttributeType.String, -1).strip()
        except: pass

    raw_leaf = child.DisplayName
    if not part_no: part_no = raw_leaf.split("/")[0].strip() if "/" in raw_leaf else raw_leaf.strip()
    if not comp_name: comp_name = raw_leaf.split("-")[-1].strip() if "-" in raw_leaf else raw_leaf.strip()

    vol, mass, den, m_name, has_g = 0.0, 0.0, 0.0, "", False
    dmats = []
    
    if proto:
        bodies = get_target_bodies(proto, session)
        if bodies:
            has_g = True
            units = [work_part.UnitCollection.FindObject("MilliMeter"), None, None, None, None]
            try: units[2] = work_part.UnitCollection.FindObject("Kilogram")
            except: units[2] = work_part.UnitCollection.GetBase("Mass")

            try:
                mass_props = work_part.MeasureManager.NewMassProperties(units, 0.999, bodies)
                if mass_props:
                    vol = float(mass_props.Volume)
                    mass = float(mass_props.Mass)
                    mass_props.Dispose()
            except Exception:
                pass
                
            if vol > 0: den = mass / vol
            
            for b in bodies:
                m = get_material(uf_session, session, b)
                if m:
                    m = clean_material_name(m)
                    if m not in dmats: dmats.append(m)
        
        if not dmats:
            m = get_material(uf_session, session, proto)
            if m:
                m = clean_material_name(m)
                if m not in dmats: dmats.append(m)
            
        if not dmats:
            for attr in ["Material", "DB_MATERIAL"]:
                try:
                    m = proto.GetUserAttributeAsString(attr, NXOpen.NXObject.AttributeType.String, -1).strip()
                    if m:
                        m = clean_material_name(m)
                        if m not in dmats: dmats.append(m)
                except: pass

        m_name = "/".join(dmats) if dmats else ""

    return {"name": comp_name, "part_no": part_no, "has_geometry": has_g, "volume": vol, "mass": mass, "density": den, "material": m_name}

def process_level(parent_comp, level, parent_id, rows, work_part, session, uf_session):
    children = parent_comp.GetChildren()
    if not children: return

    sorted_children = sorted(children, key=lambda c: c.DisplayName.upper() if c.DisplayName else "")
    child_groups, group_order = {}, []
    
    for c in sorted_children:
        if c.IsSuppressed: continue
        key = get_cleaned_comp_number(c)
        if key not in child_groups:
            child_groups[key] = []
            group_order.append(key)
        child_groups[key].append(c)

    for key in group_order:
        group = child_groups[key]
        rep = group[0]
        sub_children = rep.GetChildren()
        is_assembly = (sub_children is not None and len(sub_children) > 0)
        geom = measure_geometry(rep, work_part, session, uf_session)
        assly_type, cat = classify_component(geom["name"], is_assembly, geom["has_geometry"])

        rows.append({
            "level": level, "name": geom["name"], "part_no": geom["part_no"],
            "parent_id": parent_id, "assly_code": assly_type, "category": cat,
            "qty": len(group), "volume": geom["volume"], "nx_material": geom["material"], "nx_density": geom["density"]
        })

        if is_assembly:
            process_level(rep, level + 1, geom["part_no"], rows, work_part, session, uf_session)

def main():
    session = NXOpen.Session.GetSession()
    uf_session = NXOpen.UF.UFSession.GetUFSession()
    work_part = session.Parts.Work

    if not work_part:
        session.ListingWindow.Open()
        session.ListingWindow.WriteLine("No active NX assembly loaded.")
        return

    session.ListingWindow.Open()
    session.ListingWindow.WriteLine("Processing BOM and extracting Geometry...")

    desktop = os.path.join(os.path.expanduser("~"), "Desktop")
    excel_path = os.path.join(desktop, XML_EXCEL_NAME)

    saved_order, existing_overrides, existing_mat_lib = backup_and_read_existing(excel_path)

    root_comp = work_part.ComponentAssembly.RootComponent
    root_pno, root_rev = get_item_and_rev(work_part)
    banner_title = f"{root_pno}/{root_rev} - BOM STRUCTURE" if (INCLUDE_REVISION_IN_TITLE and root_rev) else f"{root_pno} - BOM STRUCTURE"

    rows = []
    if root_comp:
        process_level(root_comp, 1, root_pno, rows, work_part, session, uf_session)
    else:
        session.ListingWindow.WriteLine("Error: Root component not found.")
        return

    mat_lib = {k.upper(): v for k, v in DEFAULT_MATERIALS.items()}
    for m, d in existing_mat_lib.items():
        if d > 0:
            mat_lib[m.upper()] = d

    for r in rows:
        if r["nx_density"] > 0 and r["nx_material"] and "/" not in r["nx_material"]:
            mat_lib[r["nx_material"].upper()] = r["nx_density"]

    for r in rows:
        if r["nx_density"] == 0.0 and r["nx_material"]:
            mats = r["nx_material"].split("/")
            if len(mats) == 1:
                mat_key = mats[0].upper()
                if mat_key in mat_lib:
                    r["nx_density"] = mat_lib[mat_key]

    if saved_order:
        row_pool = {}
        for r in rows:
            pno = r["part_no"]
            if pno not in row_pool: row_pool[pno] = []
            row_pool[pno].append(r)

        reordered_rows = []
        for instance_key in saved_order:
            pno = instance_key.rsplit("_", 1)[0]
            if pno in row_pool and len(row_pool[pno]) > 0:
                row = row_pool[pno].pop(0)
                row["instance_key"] = instance_key
                reordered_rows.append(row)
                
        for pno, remaining_rows in row_pool.items():
            for r in remaining_rows:
                reordered_rows.append(r)
                
        rows = reordered_rows

    for r in rows:
        mat_name = r["nx_material"].strip()
        if mat_name and "/" not in mat_name:
            mat_key = mat_name.upper()
            if mat_key not in mat_lib or mat_lib[mat_key] == 0.0:
                if r["nx_density"] > 0:
                    mat_lib[mat_key] = r["nx_density"]

    xml = [
        '<?xml version="1.0"?>',
        '<?mso-application progid="Excel.Sheet"?>',
        '<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"',
        ' xmlns:o="urn:schemas-microsoft-com:office:office"',
        ' xmlns:x="urn:schemas-microsoft-com:office:excel"',
        ' xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet">',
        ' <Styles>',
        '  <Style ss:ID="Default" ss:Name="Normal"><Font ss:FontName="Arial" ss:Size="10"/></Style>',
        '  <Style ss:ID="sBanner"><Alignment ss:Horizontal="Center" ss:Vertical="Center"/><Font ss:FontName="Arial" ss:Size="12" ss:Bold="1"/><Interior ss:Color="#D9E1F2" ss:Pattern="Solid"/><Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/></Borders></Style>',
        '  <Style ss:ID="sNavyHeader"><Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/><Font ss:FontName="Arial" ss:Size="10" ss:Bold="1" ss:Color="#FFFFFF"/><Interior ss:Color="#002060" ss:Pattern="Solid"/><Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/></Borders></Style>',
        '  <Style ss:ID="sMatHeader"><Alignment ss:Horizontal="Center" ss:Vertical="Center"/><Font ss:FontName="Arial" ss:Size="10" ss:Bold="1" ss:Color="#FFFFFF"/><Interior ss:Color="#002060" ss:Pattern="Solid"/></Style>',
        '  <Style ss:ID="sYellowLevel"><Alignment ss:Horizontal="Center" ss:Vertical="Center"/><Font ss:FontName="Arial" ss:Size="10" ss:Bold="1"/><Interior ss:Color="#FFFF00" ss:Pattern="Solid"/><Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/><Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/><Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/></Borders></Style>',
        '  <Style ss:ID="sBlankCell"><Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/><Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/><Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/></Borders></Style>',
        '  <Style ss:ID="sCenter"><Alignment ss:Horizontal="Center" ss:Vertical="Center"/><Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/><Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/><Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/></Borders></Style>',
        '  <Style ss:ID="sLeftBold"><Alignment ss:Horizontal="Left" ss:Vertical="Center"/><Font ss:FontName="Arial" ss:Size="10" ss:Bold="1"/><Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/><Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/><Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/></Borders></Style>',
        '  <Style ss:ID="sLeft"><Alignment ss:Horizontal="Left" ss:Vertical="Center"/><Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/><Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/><Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/></Borders></Style>',
        '  <Style ss:ID="sDecimal"><Alignment ss:Horizontal="Right" ss:Vertical="Center"/><NumberFormat ss:Format="#,##0.00"/><Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/><Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/><Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/></Borders></Style>',
        '  <Style ss:ID="sDensity"><Alignment ss:Horizontal="Right" ss:Vertical="Center"/><NumberFormat ss:Format="0.00000000"/><Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/><Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/><Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/></Borders></Style>',
        '  <Style ss:ID="sWeight"><Alignment ss:Horizontal="Right" ss:Vertical="Center"/><NumberFormat ss:Format="#,##0.000"/><Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/><Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/><Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/></Borders></Style>',
        ' </Styles>',
        ' <Names>',
        '  <NamedRange ss:Name="MatDropdownList" ss:RefersTo="=\'Materials Library\'!R2C1:R200C1"/>',
        '  <NamedRange ss:Name="MatLookupTable" ss:RefersTo="=\'Materials Library\'!R2C1:R200C2"/>',
        ' </Names>',
        ' <Worksheet ss:Name="BOM">',
        '  <Table ss:DefaultRowHeight="18">',
        '   <Column ss:Width="25" ss:Span="5"/>',
        '   <Column ss:Index="7" ss:Width="190"/>',
        '   <Column ss:Index="8" ss:Width="65"/>',
        '   <Column ss:Index="9" ss:Width="65"/>',
        '   <Column ss:Index="10" ss:Width="100"/>',
        '   <Column ss:Index="11" ss:Width="100"/>',
        '   <Column ss:Index="12" ss:Width="45"/>',
        '   <Column ss:Index="13" ss:Width="95"/>',
        '   <Column ss:Index="14" ss:Width="95"/>',
        '   <Column ss:Index="15" ss:Width="110"/>',
        '   <Column ss:Index="16" ss:Width="95"/>',
        '   <Column ss:Index="17" ss:Width="110"/>',
        '   <Column ss:Index="18" ss:Width="95"/>',
        '   <Column ss:Index="19" ss:Width="95"/>',
        '   <Column ss:Index="20" ss:Width="100"/>',
        '   <Column ss:Index="21" ss:Width="85"/>',
        '   <Column ss:Index="22" ss:Width="85"/>',
        f'   <Row ss:Height="24"><Cell ss:Index="1" ss:MergeAcross="21" ss:StyleID="sBanner"><Data ss:Type="String">{banner_title}</Data></Cell></Row>',
        '   <Row ss:Height="28">',
        '    <Cell ss:Index="1" ss:MergeAcross="5" ss:StyleID="sNavyHeader"><Data ss:Type="String">Assembly level</Data></Cell>',
        '    <Cell ss:Index="7" ss:StyleID="sNavyHeader"><Data ss:Type="String">Item Name</Data></Cell>',
        '    <Cell ss:Index="8" ss:StyleID="sNavyHeader"><Data ss:Type="String">Assly/&#10;Child Part</Data></Cell>',
        '    <Cell ss:Index="9" ss:StyleID="sNavyHeader"><Data ss:Type="String">Category</Data></Cell>',
        '    <Cell ss:Index="10" ss:StyleID="sNavyHeader"><Data ss:Type="String">Part Number</Data></Cell>',
        '    <Cell ss:Index="11" ss:StyleID="sNavyHeader"><Data ss:Type="String">Parent Part</Data></Cell>',
        '    <Cell ss:Index="12" ss:StyleID="sNavyHeader"><Data ss:Type="String">Qty</Data></Cell>',
        '    <Cell ss:Index="13" ss:StyleID="sNavyHeader"><Data ss:Type="String">Volume/Part&#10;(mm3)</Data></Cell>',
        '    <Cell ss:Index="14" ss:StyleID="sNavyHeader"><Data ss:Type="String">Total Volume&#10;(mm3)</Data></Cell>',
        '    <Cell ss:Index="15" ss:StyleID="sNavyHeader"><Data ss:Type="String">NX Material</Data></Cell>',
        '    <Cell ss:Index="16" ss:StyleID="sNavyHeader"><Data ss:Type="String">NX Density&#10;(kg/mm3)</Data></Cell>',
        '    <Cell ss:Index="17" ss:StyleID="sNavyHeader"><Data ss:Type="String">Override Material</Data></Cell>',
        '    <Cell ss:Index="18" ss:StyleID="sNavyHeader"><Data ss:Type="String">Override Density&#10;(kg/mm3)</Data></Cell>',
        '    <Cell ss:Index="19" ss:StyleID="sNavyHeader"><Data ss:Type="String">Active Density&#10;(kg/mm3)</Data></Cell>',
        '    <Cell ss:Index="20" ss:StyleID="sNavyHeader"><Data ss:Type="String">Override Weight/&#10;Part (kg)</Data></Cell>',
        '    <Cell ss:Index="21" ss:StyleID="sNavyHeader"><Data ss:Type="String">Weight/Part&#10;(kg)</Data></Cell>',
        '    <Cell ss:Index="22" ss:StyleID="sNavyHeader"><Data ss:Type="String">Total Weight&#10;(kg)</Data></Cell>',
        '   </Row>'
    ]

    for idx, r in enumerate(rows):
        clamped_lvl = min(r["level"], 5)
        row_cells = []
        for c_idx in range(6):
            if c_idx == clamped_lvl:
                row_cells.append(f'<Cell ss:StyleID="sYellowLevel"><Data ss:Type="Number">{r["level"]}</Data></Cell>')
            else: row_cells.append('<Cell ss:StyleID="sBlankCell"/>')

        name_style = "sLeftBold" if r["assly_code"] == "A" else "sLeft"
        row_cells.append(f'<Cell ss:StyleID="{name_style}"><Data ss:Type="String">{r["name"]}</Data></Cell>')
        row_cells.append(f'<Cell ss:StyleID="sCenter"><Data ss:Type="String">{r["assly_code"]}</Data></Cell>')
        row_cells.append(f'<Cell ss:StyleID="sCenter"><Data ss:Type="String">{r["category"]}</Data></Cell>')
        row_cells.append(f'<Cell ss:StyleID="sCenter"><Data ss:Type="String">{r["part_no"]}</Data></Cell>')
        row_cells.append(f'<Cell ss:StyleID="sCenter"><Data ss:Type="String">{r["parent_id"]}</Data></Cell>')
        row_cells.append(f'<Cell ss:StyleID="sCenter"><Data ss:Type="Number">{r["qty"]}</Data></Cell>')
        row_cells.append(f'<Cell ss:StyleID="sDecimal"><Data ss:Type="Number">{r["volume"]:.4f}</Data></Cell>')
        
        row_cells.append('<Cell ss:StyleID="sDecimal" ss:Formula="=RC[-2]*RC[-1]"/>')
        row_cells.append(f'<Cell ss:StyleID="sLeft"><Data ss:Type="String">{r["nx_material"]}</Data></Cell>')
        row_cells.append(f'<Cell ss:StyleID="sDensity"><Data ss:Type="Number">{r["nx_density"]:.8f}</Data></Cell>')

        inst_key = r.get("instance_key", "")
        prior = existing_overrides.get(inst_key, {})
        ov_mat = prior.get("mat", "")
        ov_den = prior.get("density", "")
        ov_wt = prior.get("weight", "")

        row_cells.append(f'<Cell ss:StyleID="sLeft"><Data ss:Type="String">{ov_mat}</Data></Cell>' if ov_mat else '<Cell ss:StyleID="sLeft"/>')
        row_cells.append(f'<Cell ss:StyleID="sDensity"><Data ss:Type="Number">{ov_den}</Data></Cell>' if ov_den else '<Cell ss:StyleID="sDensity"/>')
        row_cells.append('<Cell ss:StyleID="sDensity" ss:Formula="=IF(ISNUMBER(RC18),IF(RC18&gt;0,RC18,RC16),IF(ISNA(VLOOKUP(RC17,MatLookupTable,2,FALSE)),RC16,VLOOKUP(RC17,MatLookupTable,2,FALSE)))"/>')
        row_cells.append(f'<Cell ss:StyleID="sWeight"><Data ss:Type="Number">{ov_wt}</Data></Cell>' if ov_wt else '<Cell ss:StyleID="sWeight"/>')
        row_cells.append('<Cell ss:StyleID="sWeight" ss:Formula="=IF(ISNUMBER(RC20),IF(RC20&gt;0,RC20,RC19*RC13),RC19*RC13)"/>')
        row_cells.append('<Cell ss:StyleID="sWeight" ss:Formula="=RC21*RC12"/>')
        xml.append('   <Row>' + "".join(row_cells) + '</Row>')

    xml.extend([
        '  </Table>',
        '  <DataValidation xmlns="urn:schemas-microsoft-com:office:excel">',
        '   <Range>R5C17:R10000C17</Range>',
        '   <Type>List</Type>',
        '   <Value>=MatDropdownList</Value>',
        '  </DataValidation>',
        ' </Worksheet>',
        ' <Worksheet ss:Name="Materials Library">',
        '  <Table ss:DefaultRowHeight="18">',
        '   <Column ss:Index="1" ss:Width="200"/>',
        '   <Column ss:Index="2" ss:Width="120"/>',
        '   <Row>',
        '    <Cell ss:StyleID="sMatHeader"><Data ss:Type="String">Material Name</Data></Cell>',
        '    <Cell ss:StyleID="sMatHeader"><Data ss:Type="String">Density (kg/mm3)</Data></Cell>',
        '   </Row>'
    ])

    for m_name, m_den in sorted(mat_lib.items()):
        xml.append(f'   <Row><Cell><Data ss:Type="String">{m_name}</Data></Cell><Cell ss:StyleID="sDensity"><Data ss:Type="Number">{m_den:.8f}</Data></Cell></Row>')

    xml.extend(['  </Table>', ' </Worksheet>', '</Workbook>'])

    with open(excel_path, "w", encoding="utf-8") as f:
        f.write("\n".join(xml))

    session.ListingWindow.WriteLine("--------------------------------------------------")
    session.ListingWindow.WriteLine("SUCCESS! Excel BOM generated with all reference set variations.")
    session.ListingWindow.WriteLine(f"File updated: {excel_path}")
    session.ListingWindow.WriteLine("--------------------------------------------------")

if __name__ == "__main__":
    main()