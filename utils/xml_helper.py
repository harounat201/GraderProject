import xml.etree.ElementTree as ET
import os

# Excel XML Namespaces
NS = {
    'main': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'c': 'http://schemas.openxmlformats.org/drawingml/2006/chart',
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'xdr': 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing',
    'x14': 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main',
    'xm': 'http://schemas.microsoft.com/office/excel/2006/main'
}


def parse_chart_xml(chart_path):
    """
    Parses a chart XML and returns a dict of formatting details.
    """
    if not os.path.exists(chart_path):
        return {}
        
    try:
        tree = ET.parse(chart_path)
        root = tree.getroot()
        
        chart_info = {
            "types": [],
            "title": "",
            "title_formula": "",
            "axes": {},
            "series": [],
            "legend_pos": ""
        }
        
        # 1. Detect Chart Types (Collection for combination)
        plot_area = root.find('.//c:plotArea', NS)
        if plot_area is not None:
             type_map = {
                 'barChart': 'Bar',
                 'lineChart': 'Line',
                 'pieChart': 'Pie',
                 'scatterChart': 'XY Scatter',
                 'bubbleChart': 'Bubble',
                 'areaChart': 'Area',
                 'radarChart': 'Radar',
                 'stockChart': 'Stock'
             }
             for t, label in type_map.items():
                 if plot_area.find(f'.//c:{t}', NS) is not None:
                     chart_info["types"].append(label)
             
             # Compatibility for old field
             if chart_info["types"]:
                 chart_info["type"] = chart_info["types"][0]

        # 2. Chart Title & Formula
        title_elem = root.find('.//c:title', NS)
        if title_elem is not None:
             chart_info["title"] = title_elem.findtext('.//c:v', default="", namespaces=NS)
             chart_info["title_formula"] = title_elem.findtext('.//c:f', default="", namespaces=NS)

        # 3. Extract Axis Scales & Titles
        for ax_tag in ['c:valAx', 'c:catAx', 'c:dateAx']:
            for ax in root.findall(f'.//{ax_tag}', NS):
                ax_id_elem = ax.find('c:axId', NS)
                ax_id = ax_id_elem.get('val') if ax_id_elem is not None else "Unknown"
                
                scaling = ax.find('c:scaling', NS)
                ax_data = {"tag": ax_tag}
                if scaling is not None:
                    ax_data.update({
                        "min": scaling.find('c:min', NS).get('val') if scaling.find('c:min', NS) is not None else "auto",
                        "max": scaling.find('c:max', NS).get('val') if scaling.find('c:max', NS) is not None else "auto",
                        "orientation": scaling.find('c:orientation', NS).get('val') if scaling.find('c:orientation', NS) is not None else "minMax"
                    })
                
                major_unit_elem = ax.find('c:majorUnit', NS)
                ax_data["major_unit"] = major_unit_elem.get('val') if major_unit_elem is not None else "auto"
                
                # Axis Title
                ax_title = ax.find('c:title', NS)
                if ax_title is not None:
                     ax_data["title"] = ax_title.findtext('.//c:v', default="", namespaces=NS)
                     ax_data["title_formula"] = ax_title.findtext('.//c:f', default="", namespaces=NS)

                chart_info["axes"][ax_id] = ax_data

        # 4. Extract Series Data
        # For combination charts, we need to know WHICH chart type block each series belongs to
        for t_tag, label in type_map.items():
            for plot_type_block in plot_area.findall(f'c:{t_tag}', NS):
                for ser in plot_type_block.findall('c:ser', NS):
                    # Try to get series name
                    ser_name = ser.findtext('.//c:tx//c:v', default=None, namespaces=NS)
                    if not ser_name:
                        ser_name = ser.findtext('.//c:tx//c:f', default="Series", namespaces=NS)
                    
                    # Traditional charts (bar, line, etc.)
                    val_ref = ser.findtext('.//c:val//c:f', default="", namespaces=NS)
                    cat_ref = ser.findtext('.//c:cat//c:f', default="", namespaces=NS)
                    
                    # Scatter/Bubble charts
                    if not val_ref:
                        val_ref = ser.findtext('.//c:yVal//c:f', default="", namespaces=NS)
                    if not cat_ref:
                        cat_ref = ser.findtext('.//c:xVal//c:f', default="", namespaces=NS)
                    
                    chart_info["series"].append({
                        "name": ser_name,
                        "values": val_ref,
                        "categories": cat_ref,
                        "type": label  # Per-series type for combination charts
                    })

        # 5. Legend Position
        legend = root.find('.//c:legend', NS)
        if legend is not None:
            lp_elem = legend.find('c:legendPos', NS)
            chart_info["legend_pos"] = lp_elem.get('val') if lp_elem is not None else "r"

        return chart_info
    except Exception as e:
        return {"error": f"Failed to parse chart XML: {str(e)}"}

def parse_pivot_table_xml(pivot_path):
    """
    Parses a Pivot Table XML to extract its structure (rows, columns, data fields).
    """
    if not os.path.exists(pivot_path):
        return {}
        
    try:
        tree = ET.parse(pivot_path)
        root = tree.getroot()
        
        info = {
            "name": root.get('name'),
            "location": root.find('main:location', NS).get('ref') if root.find('main:location', NS) is not None else "",
            "rowFields": [],
            "colFields": [],
            "dataFields": []
        }
        
        # Fields mapping
        field_names = []
        cache_def = root.find('main:cacheDefinition', NS) # This might needs cache file parsing for real names
        # For now, we trust the labels in pivotFields if they exist
        
        for field in root.findall('.//main:pivotField', NS):
             field_names.append(field.get('name') or "Unknown Field")

        # Rows
        row_fields = root.find('main:rowFields', NS)
        if row_fields is not None:
            for f in row_fields.findall('main:field', NS):
                idx = int(f.get('x'))
                if idx >= 0 and idx < len(field_names):
                    info["rowFields"].append(field_names[idx])

        # Columns
        col_fields = root.find('main:colFields', NS)
        if col_fields is not None:
            for f in col_fields.findall('main:field', NS):
                idx = int(f.get('x'))
                if idx >= 0 and idx < len(field_names):
                    info["colFields"].append(field_names[idx])

        # Data
        data_fields = root.find('main:dataFields', NS)
        if data_fields is not None:
            for f in data_fields.findall('main:dataField', NS):
                info["dataFields"].append({
                    "name": f.get('name'),
                    "fld": f.get('fld'),
                    "subtotal": f.get('subtotal')
                })

        return info
    except Exception as e:
        return {"error": f"Failed to parse pivot table: {str(e)}"}

def parse_drawing_xml(drawing_path, unzip_dir):
    """
    Parses a drawing XML to find shapes, connectors, and charts.
    """
    if not os.path.exists(drawing_path):
        return {"objects": []}
        
    try:
        tree = ET.parse(drawing_path)
        root = tree.getroot()
        results = {"objects": []}
        
        # 1. Look for Shapes (sp)
        for sp in root.findall('.//xdr:sp', NS):
             nvsp = sp.find('xdr:nvSpPr', NS)
             name = nvsp.find('xdr:cNvPr', NS).get('name') if nvsp is not None else "Shape"
             results["objects"].append({"type": "shape", "name": name})

        # 2. Look for Connectors (cxnSp) - Precedence Arrows!
        for cxn in root.findall('.//xdr:cxnSp', NS):
             nvcxn = cxn.find('xdr:nvCxnSpPr', NS)
             name = nvcxn.find('xdr:cNvPr', NS).get('name') if nvcxn is not None else "Connector"
             results["objects"].append({"type": "connector", "name": name})

        # 3. Look for Charts
        drawing_filename = os.path.basename(drawing_path)
        chart_rels = parse_drawing_rels(unzip_dir, drawing_filename)
        
        for graphic in root.findall('.//a:graphic', NS):
             chart_ref = graphic.find('.//c:chart', NS)
             if chart_ref is not None:
                 rid = chart_ref.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
                 chart_target = chart_rels.get(rid)
                 if chart_target:
                     chart_full_path = os.path.join(unzip_dir, 'xl', chart_target)
                     chart_details = parse_chart_xml(chart_full_path)
                     results["objects"].append({
                         "type": "chart",
                         "rId": rid,
                         "details": chart_details
                     })
                     
        return results
    except Exception as e:
        return {"error": f"Failed to parse drawing XML: {str(e)}", "objects": []}

def parse_drawing_rels(unzip_dir, drawing_filename):
    """
    Parses xl/drawings/_rels/drawing[N].xml.rels to find chart targets.
    """
    rels_path = os.path.join(unzip_dir, 'xl', 'drawings', '_rels', f'{drawing_filename}.rels')
    if not os.path.exists(rels_path):
        return {}
        
    try:
        tree = ET.parse(rels_path)
        root = tree.getroot()
        
        rels = {}
        for rel in root.findall('{http://schemas.openxmlformats.org/package/2006/relationships}Relationship'):
            rId = rel.get('Id')
            target = rel.get('Target')
            if target.startswith('../charts/'):
                target = target.replace('../charts/', 'charts/')
            rels[rId] = target
        return rels
    except:
        return {}

def parse_styles_xml(unzip_dir):
    """
    Parses styles.xml and returns a dictionary mapping style index to human-readable formatting.
    """
    styles_path = os.path.join(unzip_dir, 'xl', 'styles.xml')
    if not os.path.exists(styles_path):
        return {}

    try:
        tree = ET.parse(styles_path)
        root = tree.getroot()

        # 1. Number Formats (Custom)
        # Built-in IDs below 164 are usually: 0=General, 1=Decimal, 2=Fixed, 3=Comma, 4=Percentage, etc.
        num_fmts = {}
        for nf in root.findall('.//main:numFmt', NS):
            num_fmts[nf.get('numFmtId')] = nf.get('formatCode')

        # 2. Fonts
        fonts = []
        for font in root.findall('.//main:font', NS):
            fonts.append({
                "bold": font.find('main:b', NS) is not None,
                "italic": font.find('main:i', NS) is not None,
                "name": font.findtext('main:name', default="", namespaces=NS)
            })

        # 3. Fills (Shading)
        fills = []
        for fill in root.findall('.//main:fill', NS):
            pattern = fill.find('main:patternFill', NS)
            fg_color = ""
            if pattern is not None:
                fg = pattern.find('main:fgColor', NS)
                if fg is not None:
                    fg_color = fg.get('rgb') or fg.get('theme') or ""
            fills.append({
                "pattern": pattern.get('patternType') if pattern is not None else "none",
                "color": fg_color
            })

        # 4. Borders
        borders = []
        for border in root.findall('.//main:border', NS):
            border_info = {}
            for side in ['left', 'right', 'top', 'bottom']:
                side_elem = border.find(f'main:{side}', NS)
                if side_elem is not None and side_elem.get('style'):
                    border_info[side] = side_elem.get('style')
            borders.append(border_info)

        # 5. Cell Formats (XFs) - This is what style_idx refers to
        cell_styles = {}
        xfs = root.find('main:cellXfs', NS)
        if xfs is not None:
            for idx, xf in enumerate(xfs.findall('main:xf', NS)):
                style = {
                    "num_fmt": num_fmts.get(xf.get('numFmtId'), xf.get('numFmtId')),
                    "font": fonts[int(xf.get('fontId'))] if xf.get('fontId') else {},
                    "fill": fills[int(xf.get('fillId'))] if xf.get('fillId') else {},
                    "border": borders[int(xf.get('borderId'))] if xf.get('borderId') else {},
                    "alignment": {}
                }
                
                align = xf.find('main:alignment', NS)
                if align is not None:
                    style["alignment"] = {
                        "horizontal": align.get('horizontal'),
                        "vertical": align.get('vertical'),
                        "wrap": align.get('wrapText') == '1'
                    }
                
                cell_styles[str(idx)] = style

        return cell_styles
    except Exception as e:
        return {"error": f"Failed to parse styles XML: {str(e)}"}

def parse_workbook_rels(unzip_dir):
    """
    Parses workbook.xml.rels to map rIds to file paths (like worksheets).
    Useful if workbook.xml uses rIds to reference sheets.
    """
    rels_path = os.path.join(unzip_dir, 'xl', '_rels', 'workbook.xml.rels')
    if not os.path.exists(rels_path):
        return {}
    
    tree = ET.parse(rels_path)
    root = tree.getroot()
    # Namespace for Relationships usually: http://schemas.openxmlformats.org/package/2006/relationships
    # But usually parsing without specific NS for attributes is easier or using the default one
    
    rels = {}
    for rel in root.findall('{http://schemas.openxmlformats.org/package/2006/relationships}Relationship'):
        rId = rel.get('Id')
        target = rel.get('Target')
        rels[rId] = target
    return rels

def get_sheet_map(unzip_dir):
    """
    Returns a dict mapping Sheet Name -> Absolute Path to sheet XML.
    Example: { "Data": "/tmp/xl/worksheets/sheet1.xml" }
    """
    workbook_path = os.path.join(unzip_dir, 'xl', 'workbook.xml')
    if not os.path.exists(workbook_path):
        raise FileNotFoundError("workbook.xml not found")
        
    tree = ET.parse(workbook_path)
    root = tree.getroot()
    
    sheets = {}
    rels = parse_workbook_rels(unzip_dir)
    
    # <sheets> <sheet name="Data" sheetId="1" r:id="rId1"/> ...
    for sheet in root.findall('.//main:sheet', NS):
        name = sheet.get('name')
        rId = sheet.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
        
        # Resolve target path
        if rId in rels:
            # target is usually relative to xl/ like "worksheets/sheet1.xml"
            target = rels[rId]
            # normalized path calculation
            full_path = os.path.normpath(os.path.join(unzip_dir, 'xl', target))
            sheets[name] = full_path
            
    return sheets

def get_shared_strings(unzip_dir):
    """
    Parses sharedStrings.xml and returns a list of strings.
    """
    ss_path = os.path.join(unzip_dir, 'xl', 'sharedStrings.xml')
    strings = []
    
    if not os.path.exists(ss_path):
        return strings # Return empty if no shared strings
        
    tree = ET.parse(ss_path)
    root = tree.getroot()
    
    # <si> <t>Value</t> </si>
    for si in root.findall('main:si', NS):
        # Text can be in <t> directly or inside <r><t> (rich text)
        text_nodes = si.findall('.//main:t', NS)
        text_val = "".join([t.text for t in text_nodes if t.text])
        strings.append(text_val)
        
    return strings

def parse_sheet_data(sheet_xml_path, shared_strings):
    """
    Parses a sheet XML and returns a dict of cell data.
    structure: { "A1": { "value": "100", "formula": "SUM(B1:B2)", "style": "1" }, ... }
    """
    if not os.path.exists(sheet_xml_path):
        return {}
        
    tree = ET.parse(sheet_xml_path)
    root = tree.getroot()
    
    cells = {}
    
    for c in root.findall('.//main:c', NS):
        coord = c.get('r') # e.g., "A1"
        sBox = c.get('s') # valid style index
        t = c.get('t') # type: 's' for shared string, 'str' for formula string, etc.
        
        formula_elem = c.find('main:f', NS)
        val_elem = c.find('main:v', NS)
        
        formula = formula_elem.text if formula_elem is not None else None
        
        raw_val = val_elem.text if val_elem is not None else ""
        final_val = raw_val
        
        if t == 's': # Shared String lookup
            try:
                idx = int(raw_val)
                final_val = shared_strings[idx]
            except (ValueError, IndexError):
                final_val = "ERROR_STRING_LOOKUP"
        
        cells[coord] = {
            "value": final_val,
            "formula": formula,
            "style_idx": sBox,
            # we could store more info here if needed
        }
        
    return cells

def parse_sheet_rels(unzip_dir, sheet_filename):
    """
    Parses the .rels file for a specific sheet to find related objects like PivotTables.
    sheet_filename e.g. "sheet1.xml"
    """
    rels_path = os.path.join(unzip_dir, 'xl', 'worksheets', '_rels', f'{sheet_filename}.rels')
    if not os.path.exists(rels_path):
        return {}
        
    tree = ET.parse(rels_path)
    root = tree.getroot()
    
    rels = {}
    for rel in root.findall('{http://schemas.openxmlformats.org/package/2006/relationships}Relationship'):
        rId = rel.get('Id')
        target = rel.get('Target')
        type_uri = rel.get('Type')
        rels[rId] = {"target": target, "type": type_uri}
    return rels

def parse_sheet_full(sheet_xml_path, shared_strings, unzip_dir=None, sheet_filename=None):
    """
    Parses a sheet XML and returns (cells, metadata).
    metadata includes validations, conditional formatting, and drawing refs.
    If unzip_dir and sheet_filename are provided, checks .rels for Pivot tables.
    """
    if not os.path.exists(sheet_xml_path):
        return {}, {}
        
    tree = ET.parse(sheet_xml_path)
    root = tree.getroot()
    
    cells = {}
    metadata = {
        "validations": [],
        "conditional_formatting": [],
        "drawings": [],
        "sparklines": [],
        "merge_cells": [],
        "view_settings": {}
    }
    
    # 1. Parse Cells
    for c in root.findall('.//main:c', NS):
        coord = c.get('r') # e.g., "A1"
        sBox = c.get('s') 
        t = c.get('t') 
        
        formula_elem = c.find('main:f', NS)
        val_elem = c.find('main:v', NS)
        
        formula = formula_elem.text if formula_elem is not None else None
        formula_type = formula_elem.get('t') if formula_elem is not None else "normal"
        formula_ref = formula_elem.get('ref') if formula_elem is not None else ""
        
        raw_val = val_elem.text if val_elem is not None else ""
        final_val = raw_val
        
        if t == 's': # Shared String lookup
            try:
                idx = int(raw_val)
                final_val = shared_strings[idx]
            except (ValueError, IndexError):
                final_val = "ERROR_STRING_LOOKUP"
        
        cells[coord] = {
            "value": final_val,
            "formula": formula,
            "formula_type": formula_type,
            "formula_ref": formula_ref,
            "style_idx": sBox,
        }
    
    # 2. Extract Data Validations
    # <dataValidations> <dataValidation type="list" ...> ...
    dvs = root.find('main:dataValidations', NS)
    if dvs is not None:
        for dv in dvs.findall('main:dataValidation', NS):
             dv_info = {
                 "type": dv.get('type'),
                 "sqref": dv.get('sqref'),
                 "formula1": dv.findtext('main:formula1', default="", namespaces=NS),
                 "formula2": dv.findtext('main:formula2', default="", namespaces=NS),
             }
             metadata['validations'].append(dv_info)

    # 3. Extract Conditional Formatting
    # <conditionalFormatting sqref="E3:E14"> <cfRule type="colorScale" ...>
    for cf in root.findall('main:conditionalFormatting', NS):
        sqref = cf.get('sqref')
        for rule in cf.findall('main:cfRule', NS):
            metadata['conditional_formatting'].append({
                "sqref": sqref,
                "type": rule.get('type'),
                "dxfId": rule.get('dxfId'),
                "priority": rule.get('priority'),
                "formula": rule.findtext('main:formula', default="", namespaces=NS)
            })

    # 4. Extract Drawings (Charts, Shapes, Connectors)
    drawing = root.find('main:drawing', NS)
    if drawing is not None:
        dr_rid = drawing.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
        if unzip_dir and sheet_filename:
            sheet_rels = parse_sheet_rels(unzip_dir, sheet_filename)
            dr_info = sheet_rels.get(dr_rid)
            if dr_info:
                 dr_path_rel = dr_info.get("target")
                 # Correct path resolution
                 if dr_path_rel.startswith('../'):
                     dr_full_path = os.path.normpath(os.path.join(unzip_dir, 'xl', 'worksheets', dr_path_rel))
                 else:
                     dr_full_path = os.path.join(unzip_dir, 'xl', 'drawings', os.path.basename(dr_path_rel))
                     
                 drawing_data = parse_drawing_xml(dr_full_path, unzip_dir)
                 for obj in drawing_data.get("objects", []):
                     metadata['drawings'].append(obj)
        else:
            metadata['drawings'].append(f"Drawing Reference rId={dr_rid}")
        
    legacy = root.find('main:legacyDrawing', NS)
    if legacy is not None:
        rid = legacy.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
        metadata['drawings'].append(f"Legacy Drawing Reference rId={rid}")

    # 5. Check Rels for Pivot Tables (if context provided)
    if unzip_dir and sheet_filename:
        rels = parse_sheet_rels(unzip_dir, sheet_filename)
        for rid, info in rels.items():
            if "pivotTable" in info.get("type", ""):
                 target = info.get("target") # e.g. ../pivotTables/pivotTable1.xml
                 pivot_path = os.path.normpath(os.path.join(unzip_dir, 'xl', 'worksheets', target))
                 pivot_details = parse_pivot_table_xml(pivot_path)
                 metadata['drawings'].append({
                     "type": "pivotTable",
                     "rId": rid,
                     "details": pivot_details
                 })

    # 6. Extract Sparklines
    # Sparklines are in an extension list or sparklineGroups element
    sgs = root.find('main:extLst/main:ext/{http://schemas.microsoft.com/office/spreadsheetml/2009/9/main}sparklineGroups', NS)
    if sgs is None:
        # Some versions might have it elsewhere or without extension list
        sgs = root.find('.//x14:sparklineGroups', NS)
        
    if sgs is not None:
        for group in sgs.findall('x14:sparklineGroup', NS):
            g_info = {
                "type": group.get('type', 'line'),
                "sparklines": []
            }
            for sl in group.findall('x14:sparklines/x14:sparkline', NS):
                g_info["sparklines"].append({
                    "data_range": sl.findtext('xm:f', namespaces=NS) or sl.get('f'),   # Formula referencing data
                    "location": sl.findtext('xm:sqref', namespaces=NS) or sl.get('sqref')  # Cell where sparkline is placed
                })
            metadata['sparklines'].append(g_info)

    # 7. Extract Merged Cells
    mcs = root.find('main:mergeCells', NS)
    if mcs is not None:
        for mc in mcs.findall('main:mergeCell', NS):
            metadata['merge_cells'].append(mc.get('ref'))

    # 8. Extract Sheet Views (Active Tab)
    svs = root.find('main:sheetViews', NS)
    if svs is not None:
        view = svs.find('main:sheetView', NS)
        if view is not None:
            metadata['view_settings'] = {
                "tabSelected": view.get('tabSelected') == "1",
                "showGridLines": view.get('showGridLines') != "0",
                "zoomScale": view.get('zoomScale')
            }

    return cells, metadata

def parse_workbook_to_json(unzip_dir):
    """
    Parses an entire workbook (all sheets) into a large JSON-friendly dict.
    Returns:
    {
       "sheets": {
           "Sheet1": { "cells": {...}, "metadata": {...} },
           ...
       }
    }
    """
    sheet_map = get_sheet_map(unzip_dir)
    shared_strings = get_shared_strings(unzip_dir)
    styles = parse_styles_xml(unzip_dir)
    
    workbook_data = {
        "sheets": {},
        "workbook_metadata": {
            "definedNames": {}
        }
    }

    # Extract Defined Names (Named Ranges, Solver Settings)
    workbook_xml_path = os.path.join(unzip_dir, 'xl', 'workbook.xml')
    if os.path.exists(workbook_xml_path):
        wb_tree = ET.parse(workbook_xml_path)
        wb_root = wb_tree.getroot()
        dns = wb_root.find('main:definedNames', NS)
        if dns is not None:
             for dn in dns.findall('main:definedNames/main:definedName', NS) or dns.findall('main:definedName', NS):
                 name = dn.get('name')
                 workbook_data["workbook_metadata"]["definedNames"][name] = dn.text
    
    for name, path in sheet_map.items():
        filename = os.path.basename(path)
        cells, metadata = parse_sheet_full(path, shared_strings, unzip_dir, filename)
        
        # Optimization: Only include cells with content OR special formatting (borders/shading)
        clean_cells = {}
        for coord, data in cells.items():
            # Resolve Style
            s_idx = data.get('style_idx')
            compact_style = {}
            if s_idx in styles:
                full_style = styles[s_idx]
                # Prune default styles to save tokens
                
                # 1. Fill (Shading)
                if full_style.get('fill', {}).get('pattern') != 'none':
                    compact_style['fill'] = full_style['fill']
                
                # 2. Border (Key for "Separator Lines")
                if full_style.get('border'):
                    compact_style['border'] = full_style['border']
                    
                # 3. Number Format
                num_fmt = full_style.get('num_fmt')
                if num_fmt and num_fmt not in ['0', 'General']:
                    compact_style['num_fmt'] = num_fmt
                    
                # 4. Font (Bold/Italic)
                font = full_style.get('font', {})
                if font.get('bold') or font.get('italic'):
                    compact_style['font'] = {k: v for k, v in font.items() if v}

            has_content = bool(data.get('value') or data.get('formula'))
            has_formatting = bool(compact_style)
            
            # Keep cell if it has either content OR non-default formatting
            if has_content or has_formatting:
                if compact_style:
                    data['style'] = compact_style
                
                # Cleanup internal fields
                if 'style_idx' in data:
                    del data['style_idx']
                clean_cells[coord] = data
                
        # Skip empty sheets (unless they have metadata like drawings)
        has_content = len(clean_cells) > 0
        has_meta = len(metadata.get('validations', [])) > 0 or \
                   len(metadata.get('conditional_formatting', [])) > 0 or \
                   len(metadata.get('drawings', [])) > 0
                   
        if has_content or has_meta:
            workbook_data["sheets"][name] = {
                "cells": clean_cells,
                "metadata": metadata
            }
        
    return workbook_data
