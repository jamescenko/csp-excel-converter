from flask import Flask, request, send_file
import openpyxl
from openpyxl.cell.cell import MergedCell
from openpyxl.styles import numbers
from copy import copy
import io
import base64

app = Flask(__name__)

def safe_write_cell(ws, cell_ref, value, is_currency=False):
    """Safely write to a cell, handling merged cells and formatting"""
    try:
        cell = ws[cell_ref]
        if isinstance(cell, MergedCell):
            for merged_range in ws.merged_cells.ranges:
                if cell.coordinate in merged_range:
                    master_cell = ws.cell(merged_range.min_row, merged_range.min_col)
                    master_cell.value = value
                    if is_currency:
                        master_cell.number_format = '$#,##0.00'
                    return
        else:
            cell.value = value
            if is_currency:
                cell.number_format = '$#,##0.00'
    except Exception as e:
        print(f"Warning: Could not write to {cell_ref}: {e}")

def copy_sheet(source_sheet, target_sheet):
    """Copy sheet structure including formatting"""
    for row in source_sheet.iter_rows():
        for cell in row:
            new_cell = target_sheet[cell.coordinate]
            if cell.has_style:
                new_cell.font = copy(cell.font)
                new_cell.border = copy(cell.border)
                new_cell.fill = copy(cell.fill)
                new_cell.number_format = copy(cell.number_format)
                new_cell.protection = copy(cell.protection)
                new_cell.alignment = copy(cell.alignment)
            if cell.value:
                new_cell.value = cell.value
    
    # Copy merged cells
    for merged_cell_range in source_sheet.merged_cells.ranges:
        target_sheet.merge_cells(str(merged_cell_range))
    
    # Copy column widths
    for col in source_sheet.column_dimensions:
        target_sheet.column_dimensions[col].width = source_sheet.column_dimensions[col].width
    
    # Copy row heights
    for row in source_sheet.row_dimensions:
        target_sheet.row_dimensions[row].height = source_sheet.row_dimensions[row].height

@app.route('/convert', methods=['POST'])
def fill_template():
    try:
        data = request.json
        
        meta = data.get("meta", {})
        summary = data.get("summary", {})
        regions = data.get("regions", [])
        
        exchange_rate = meta.get("exchangeRate", 1.39)
        
        # Handle both string and dict formats for template
        template_data = data.get("template", "")
        if not template_data:
            return {"error": "No template provided"}, 400
        
        if isinstance(template_data, dict):
            template_b64 = template_data.get("data", "")
            print("Received template as dict, extracted data property")
        else:
            template_b64 = template_data
            print("Received template as string")
            
        if not template_b64:
            return {"error": "No template data found"}, 400
            
        template_bytes = base64.b64decode(template_b64)
        wb = openpyxl.load_workbook(io.BytesIO(template_bytes))
        
        def safe_float(val):
            try:
                return float(val)
            except:
                return 0.0
        
        # === SUMMARY SHEET ===
        ws = wb["Summary"]
        
        # Exchange Rate (G4) - NO dollar sign for exchange rate
        safe_write_cell(ws, "G4", float(exchange_rate), is_currency=False)
        
        # Support Summary Table
        # Food Distribution (row 20)
        safe_write_cell(ws, "C20", summary.get("totalChildren", 0), is_currency=False)  # Number of children
        safe_write_cell(ws, "D20", safe_float(summary.get("foodDistCAD", 0)), is_currency=True)
        safe_write_cell(ws, "E20", safe_float(summary.get("foodDistUSD", 0)), is_currency=True)
        
        # Salary case-worker (row 22)
        safe_write_cell(ws, "D22", safe_float(summary.get("salaryCAD", 0)), is_currency=True)
        safe_write_cell(ws, "E22", safe_float(summary.get("salaryUSD", 0)), is_currency=True)
        
        # Subtotal Regular Support (row 25)
        safe_write_cell(ws, "D25", safe_float(summary.get("foodDistCAD", 0)) + safe_float(summary.get("salaryCAD", 0)), is_currency=True)
        safe_write_cell(ws, "E25", safe_float(summary.get("foodDistUSD", 0)) + safe_float(summary.get("salaryUSD", 0)), is_currency=True)
        
        # Additional Support
        # Family gifts (row 28)
        safe_write_cell(ws, "D28", safe_float(summary.get("familyCADTotal", 0)), is_currency=True)
        safe_write_cell(ws, "E28", safe_float(summary.get("familyUSDTotal", 0)), is_currency=True)
        
        # Medical gifts (row 29)
        safe_write_cell(ws, "D29", safe_float(summary.get("medicalCADTotal", 0)), is_currency=True)
        safe_write_cell(ws, "E29", safe_float(summary.get("medicalUSDTotal", 0)), is_currency=True)
        
        # Subtotal Additional (row 30)
        safe_write_cell(ws, "D30", safe_float(summary.get("familyCADTotal", 0)) + safe_float(summary.get("medicalCADTotal", 0)), is_currency=True)
        safe_write_cell(ws, "E30", safe_float(summary.get("familyUSDTotal", 0)) + safe_float(summary.get("medicalUSDTotal", 0)), is_currency=True)
        
        # GRAND TOTAL (row 32)
        safe_write_cell(ws, "D32", safe_float(summary.get("totalCAD", 0)), is_currency=True)
        safe_write_cell(ws, "E32", safe_float(summary.get("totalUSD", 0)), is_currency=True)
        
        # === REGION SHEETS ===
        if "Region" not in wb.sheetnames:
            print("Warning: No 'Region' template sheet found")
        else:
            region_template = wb["Region"]
            
            # Process each region
            for region_data in regions:
                region_name = region_data.get("region", "")
                
                if "GRAND COUNT" in region_name or not region_name:
                    continue
                
                # Sanitize sheet name
                sheet_name = region_name.replace(":", "-").replace("/", "-").replace("\\", "-")
                sheet_name = sheet_name.replace("?", "").replace("*", "").replace("[", "(").replace("]", ")")
                sheet_name = sheet_name[:31]
                
                print(f"Processing region: {sheet_name}")
                
                # Create new sheet from template
                if sheet_name in wb.sheetnames:
                    ws_r = wb[sheet_name]
                else:
                    ws_r = wb.create_sheet(title=sheet_name)
                    copy_sheet(region_template, ws_r)
                
                # Exchange Rate (G4)
                safe_write_cell(ws_r, "G4", float(exchange_rate), is_currency=False)
                
                # CRITICAL: Region sheets use DIFFERENT ROWS than Summary!
                # Looking at your images, Region sheet structure is different
                
                # Support Summary Table - REGIONS USE ROW 20 for Food Distribution
                safe_write_cell(ws_r, "D20", region_data.get("children", 0), is_currency=False)  # Number of children
                safe_write_cell(ws_r, "E20", safe_float(region_data.get("foodDistCAD", 0)), is_currency=True)
                safe_write_cell(ws_r, "F20", safe_float(region_data.get("foodDistUSD", 0)), is_currency=True)
                
                # Salary (row 22)
                safe_write_cell(ws_r, "E22", safe_float(region_data.get("salaryCAD", 0)), is_currency=True)
                safe_write_cell(ws_r, "F22", safe_float(region_data.get("salaryUSD", 0)), is_currency=True)
                
                # Subtotal Regular (row 25)
                safe_write_cell(ws_r, "E25", safe_float(region_data.get("foodDistCAD", 0)) + safe_float(region_data.get("salaryCAD", 0)), is_currency=True)
                safe_write_cell(ws_r, "F25", safe_float(region_data.get("foodDistUSD", 0)) + safe_float(region_data.get("salaryUSD", 0)), is_currency=True)
                
                # Additional Support
                # Family gifts (row 28)
                safe_write_cell(ws_r, "E28", safe_float(region_data.get("familyCAD", 0)), is_currency=True)
                safe_write_cell(ws_r, "F28", safe_float(region_data.get("familyUSD", 0)), is_currency=True)
                
                # Medical gifts (row 29)
                safe_write_cell(ws_r, "E29", safe_float(region_data.get("medicalCAD", 0)), is_currency=True)
                safe_write_cell(ws_r, "F29", safe_float(region_data.get("medicalUSD", 0)), is_currency=True)
                
                # Subtotal Additional (row 30)
                safe_write_cell(ws_r, "E30", safe_float(region_data.get("familyCAD", 0)) + safe_float(region_data.get("medicalCAD", 0)), is_currency=True)
                safe_write_cell(ws_r, "F30", safe_float(region_data.get("familyUSD", 0)) + safe_float(region_data.get("medicalUSD", 0)), is_currency=True)
                
                # GRAND TOTAL (row 32)
                safe_write_cell(ws_r, "E32", safe_float(region_data.get("totalCAD", 0)), is_currency=True)
                safe_write_cell(ws_r, "F32", safe_float(region_data.get("totalUSD", 0)), is_currency=True)
        
        # Save and return
        output = io.BytesIO()
        wb.save(output)
        output.seek(0)
        
        print("SUCCESS: File generated successfully")
        
        return send_file(
            output,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name="CSP-Report-Filled.xlsx"
        )
        
    except Exception as e:
        print(f"ERROR: {str(e)}")
        import traceback
        traceback.print_exc()
        return {"error": str(e), "details": type(e).__name__}, 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=10000)