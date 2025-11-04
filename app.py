from flask import Flask, request, send_file, jsonify
from openpyxl import load_workbook
from datetime import datetime
import io
import os

app = Flask(__name__)

@app.route('/health', methods=['GET'])
def health():
    """Health check endpoint"""
    return jsonify({
        "status": "healthy", 
        "message": "CSP Excel Service Running",
        "timestamp": datetime.now().isoformat()
    })

@app.route('/populate-excel', methods=['POST'])
def populate_excel():
    """
    Receives calculation data from n8n and returns populated Excel
    """
    try:
        data = request.json
        if not data:
            return jsonify({"error": "No data received"}), 400
        
        # Extract main sections
        summary_data = data.get('summary', {})
        regions_data = data.get('regions', [])
        exchange_rate = float(data.get('exchangeRate', 1.39))
        period_from = data.get('reportPeriodFrom', '')
        period_to = data.get('reportPeriodTo', '')
        
        print(f"📥 Processing {len(regions_data)} regions with {summary_data.get('totalChildren', 0)} children")
        
        # Load template
        template_path = 'CSP_Automate_Template.xlsx'
        if not os.path.exists(template_path):
            return jsonify({"error": "Template not found"}), 500
        
        wb = load_workbook(template_path)
        
        results = {
            "summary_updated": False,
            "regions_processed": 0,
            "regions_skipped": [],
            "children_written": 0
        }
        
        # ============================================
        # POPULATE SUMMARY SHEET
        # ============================================
        try:
            ws = wb['Summary']
            
            # Period & Exchange Rate (Top section)
            ws['B4'] = period_from
            ws['B5'] = period_to
            ws['H4'] = exchange_rate
            
            # Regular Support Section (Rows 20-25)
            ws['C20'] = summary_data.get('totalChildren', 0)
            ws['D20'] = summary_data.get('foodDistCAD', 0)
            ws['E20'] = summary_data.get('foodDistUSD', 0)
            ws['D22'] = summary_data.get('salaryCAD', 0)
            ws['E22'] = summary_data.get('salaryUSD', 0)
            ws['D24'] = summary_data.get('incentiveCAD', 0)
            ws['E24'] = summary_data.get('incentiveUSD', 0)
            # D25 and E25 are SUBTOTALS (Regular Support)
            ws['D25'] = summary_data.get('foodDistCAD', 0) + summary_data.get('salaryCAD', 0) + summary_data.get('incentiveCAD', 0)
            ws['E25'] = summary_data.get('foodDistUSD', 0) + summary_data.get('salaryUSD', 0) + summary_data.get('incentiveUSD', 0)
            
            # Additional Support (Rows 28-30)
            ws['D28'] = summary_data.get('familyCAD', 0)
            ws['E28'] = summary_data.get('familyUSD', 0)
            ws['D29'] = summary_data.get('medicalCAD', 0)
            ws['E29'] = summary_data.get('medicalUSD', 0)
            # D30 and E30 are SUBTOTALS (Additional Support/After Admin Fee)
            ws['D30'] = summary_data.get('familyCAD', 0) + summary_data.get('medicalCAD', 0)
            ws['E30'] = summary_data.get('familyUSD', 0) + summary_data.get('medicalUSD', 0)
            
            # D32 and E32 are GRAND TOTALS
            ws['D32'] = ws['D25'].value + ws['D30'].value
            ws['E32'] = ws['E25'].value + ws['E30'].value
            
            # Cross Check Section (Rows 36-43)
            ws['C36'] = summary_data.get('totalChildren', 0)  # # of active child
            ws['C37'] = summary_data.get('newChildrenCount', 0)  # # of New Child
            ws['C40'] = summary_data.get('foodDistCAD', 0)  # Food CAD
            ws['D40'] = summary_data.get('foodDistUSD', 0)  # Food USD
            ws['C41'] = summary_data.get('salaryCAD', 0)  # Salary CAD
            ws['D41'] = summary_data.get('salaryUSD', 0)  # Salary USD
            ws['C42'] = summary_data.get('incentiveCAD', 0)  # Incentive CAD
            ws['D42'] = summary_data.get('incentiveUSD', 0)  # Incentive USD
            # C43 and D43 are Subtotals
            ws['C43'] = ws['C40'].value + ws['C41'].value + ws['C42'].value
            ws['D43'] = ws['D40'].value + ws['D41'].value + ws['D42'].value
            
            # Admin Fee & Gifts Section (Rows 47-51)
            ws['C47'] = summary_data.get('familyCAD', 0)
            ws['D47'] = summary_data.get('familyUSD', 0)
            ws['C48'] = summary_data.get('medicalCAD', 0)
            ws['D48'] = summary_data.get('medicalUSD', 0)
            # C49 and D49 are Subtotals
            ws['C49'] = ws['C47'].value + ws['C48'].value
            ws['D49'] = ws['D47'].value + ws['D48'].value
            # C51 and D51 are Grand Totals (after admin fee)
            ws['C51'] = ws['C43'].value + ws['C49'].value
            ws['D51'] = ws['D43'].value + ws['D49'].value
            
            # Summary Statistics Section (Rows 55-58)
            # Left column values (B or C), Right column values (G)
            ws['C55'] = summary_data.get('totalChildren', 0)  # Total Active Children
            ws['G55'] = exchange_rate  # Exchange Rate
            
            ws['C56'] = summary_data.get('newChildrenCount', 0)  # New Children
            total_usd = summary_data.get('totalUSD', 0)
            total_children = summary_data.get('totalChildren', 0)
            avg_per_child = round(total_usd / total_children, 2) if total_children > 0 else 0
            ws['G56'] = avg_per_child  # Avg per Child (USD)
            
            ws['C57'] = summary_data.get('totalChildren', 0)  # Food Recipients
            ws['G57'] = summary_data.get('totalCAD', 0)  # Total CAD Amount
            
            ws['C58'] = len(regions_data)  # Salary Support Recipients
            ws['G58'] = summary_data.get('totalUSD', 0)  # Total USD Amount
            
            results['summary_updated'] = True
            print("✅ Summary sheet populated")
            
        except Exception as e:
            print(f"❌ Summary error: {str(e)}")
        
        # ============================================
        # POPULATE REGION SHEETS
        # ============================================
        for region in regions_data:
            region_code = region.get('code', '').strip()
            
            if not region_code or region_code not in wb.sheetnames:
                results['regions_skipped'].append(region_code or 'UNKNOWN')
                print(f"⚠️ Sheet '{region_code}' not found")
                continue
            
            try:
                ws = wb[region_code]
                
                # Period & Exchange Rate
                ws['B4'] = period_from
                ws['B5'] = period_to
                ws['H4'] = exchange_rate  # CORRECT: H4 not G4
                
                # Wire & Location Info
                ws['B10'] = str(region.get('wireId', ''))  # Partner ID
                ws['B11'] = region.get('caseworker', 'Unknown')  # Caseworker
                ws['B13'] = region.get('beneficiary', '')  # Beneficiary
                ws['B14'] = region.get('region', region_code)  # Region
                ws['B15'] = region.get('city', '')  # City
                
                # Regular Support Section (Rows 20-24)
                ws['C20'] = region.get('children', 0)
                ws['D20'] = region.get('foodDistCAD', 0)
                ws['E20'] = region.get('foodDistUSD', 0)
                ws['D22'] = region.get('salaryCAD', 0)
                ws['E22'] = region.get('salaryUSD', 0)
                ws['D24'] = region.get('incentiveCAD', 0)
                ws['E24'] = region.get('incentiveUSD', 0)
                
                # D25 and E25 are SUBTOTALS (Regular Support)
                ws['D25'] = ws['D20'].value + ws['D22'].value + ws['D24'].value
                ws['E25'] = ws['E20'].value + ws['E22'].value + ws['E24'].value
                
                # Additional Support (Rows 28-30)
                ws['D28'] = region.get('familyCAD', 0)
                ws['E28'] = region.get('familyUSD', 0)
                ws['D29'] = region.get('medicalCAD', 0)
                ws['E29'] = region.get('medicalUSD', 0)
                
                # D30 and E30 are SUBTOTALS (Additional Support)
                ws['D30'] = ws['D28'].value + ws['D29'].value
                ws['E30'] = ws['E28'].value + ws['E29'].value
                
                # D32 and E32 are GRAND TOTALS
                ws['D32'] = ws['D25'].value + ws['D30'].value
                ws['E32'] = ws['E25'].value + ws['E30'].value
                
                # ============================================
                # POPULATE CHILDREN DETAILS TABLE (Row 36+)
                # ============================================
                child_details = region.get('childDetails', [])
                
                # Clear old data (rows 36-200)
                for row in range(36, 201):
                    for col in ['A', 'B', 'C', 'D', 'E', 'F', 'G']:
                        ws[f'{col}{row}'] = None
                
                # Write new children data starting at row 36
                if child_details:
                    start_row = 36
                    for idx, child in enumerate(child_details):
                        row = start_row + idx
                        
                        ws[f'A{row}'] = child.get('cspId', '')
                        ws[f'B{row}'] = child.get('childName', child.get('name', ''))
                        
                        # Financial data
                        food_usd = child.get('foodDistUSD', child.get('foodAmount', 0))
                        medical = child.get('medicalGifts', 0)
                        family = child.get('familyGifts', 0)
                        
                        ws[f'C{row}'] = food_usd
                        ws[f'D{row}'] = medical
                        ws[f'E{row}'] = family
                        ws[f'F{row}'] = round(food_usd + medical + family, 2)
                        ws[f'G{row}'] = ''  # Signature
                        
                        results['children_written'] += 1
                
                results['regions_processed'] += 1
                print(f"✅ {region_code}: {len(child_details)} children")
                
            except Exception as e:
                results['regions_skipped'].append(region_code)
                print(f"❌ {region_code}: {str(e)}")
        
        # ============================================
        # SAVE AND RETURN
        # ============================================
        output = io.BytesIO()
        wb.save(output)
        output.seek(0)
        
        print(f"\n✅ COMPLETE: {results['regions_processed']} regions, {results['children_written']} children")
        
        return send_file(
            output,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=f'CSP_Wiring_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx'
        )
    
    except Exception as e:
        print(f"❌ CRITICAL: {str(e)}")
        return jsonify({"error": str(e)}), 500

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 8080))
    app.run(host='0.0.0.0', port=port)