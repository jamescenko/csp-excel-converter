from flask import Flask, request, send_file, jsonify
from openpyxl import load_workbook
from datetime import datetime
import io
import os
import traceback

app = Flask(__name__)

@app.route('/health', methods=['GET'])
def health():
    return jsonify({
        "status": "healthy",
        "message": "CSP Excel Service Running",
        "timestamp": datetime.now().isoformat()
    })

@app.route('/populate-excel', methods=['POST'])
def populate_excel():
    try:
        data = request.json
        if not data:
            return jsonify({"error": "No data received"}), 400
        
        print(f"\n{'='*60}")
        print(f"📥 RECEIVED DATA:")
        print(f"Top-level keys: {list(data.keys())}")
        
        # Extract data with multiple fallback paths
        summary_data = data.get('summary', {})
        exchange_rate = float(data.get('exchangeRate', 1.39))
        period_from = data.get('reportPeriodFrom', '01-01-2025')
        period_to = data.get('reportPeriodTo', '31-03-2025')
        
        # Get regions - try multiple paths
        regions_data = data.get('regions', [])
        if not regions_data:
            regions_data = summary_data.get('regions', [])
        
        print(f"Summary keys: {list(summary_data.keys())}")
        print(f"Regions found: {len(regions_data)}")
        print(f"Exchange rate: {exchange_rate}")
        print(f"{'='*60}\n")
        
        # Load template
        template_path = 'CSP_Automate_Template.xlsx'
        if not os.path.exists(template_path):
            return jsonify({"error": "Template not found"}), 500
        
        wb = load_workbook(template_path)
        print(f"✓ Template loaded. Available sheets: {len(wb.sheetnames)}")
        
        results = {
            "summary_updated": False,
            "regions_processed": 0,
            "regions_skipped": [],
            "children_written": 0,
            "errors": []
        }
        
        # ============================================
        # POPULATE SUMMARY SHEET
        # ============================================
        try:
            ws = wb['Summary']
            print(f"\n📊 POPULATING SUMMARY SHEET...")
            
            # Period & Exchange Rate
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
            
            # D25 and E25 are SUBTOTALS
            d25 = (summary_data.get('foodDistCAD', 0) + 
                  summary_data.get('salaryCAD', 0) + 
                  summary_data.get('incentiveCAD', 0))
            e25 = (summary_data.get('foodDistUSD', 0) + 
                  summary_data.get('salaryUSD', 0) + 
                  summary_data.get('incentiveUSD', 0))
            ws['D25'] = d25
            ws['E25'] = e25
            
            # Additional Support (Rows 28-30)
            ws['D28'] = summary_data.get('familyCAD', 0)
            ws['E28'] = summary_data.get('familyUSD', 0)
            ws['D29'] = summary_data.get('medicalCAD', 0)
            ws['E29'] = summary_data.get('medicalUSD', 0)
            
            # D30 and E30 are SUBTOTALS
            d30 = summary_data.get('familyCAD', 0) + summary_data.get('medicalCAD', 0)
            e30 = summary_data.get('familyUSD', 0) + summary_data.get('medicalUSD', 0)
            ws['D30'] = d30
            ws['E30'] = e30
            
            # D32 and E32 are GRAND TOTALS
            ws['D32'] = d25 + d30
            ws['E32'] = e25 + e30
            
            # Cross Check Section (Rows 36-43)
            ws['C36'] = summary_data.get('totalChildren', 0)
            ws['C37'] = summary_data.get('newChildrenCount', 0)
            ws['C40'] = summary_data.get('foodDistCAD', 0)
            ws['D40'] = summary_data.get('foodDistUSD', 0)
            ws['C41'] = summary_data.get('salaryCAD', 0)
            ws['D41'] = summary_data.get('salaryUSD', 0)
            ws['C42'] = summary_data.get('incentiveCAD', 0)
            ws['D42'] = summary_data.get('incentiveUSD', 0)
            
            # C43 and D43 are Subtotals
            ws['C43'] = summary_data.get('foodDistCAD', 0) + summary_data.get('salaryCAD', 0) + summary_data.get('incentiveCAD', 0)
            ws['D43'] = summary_data.get('foodDistUSD', 0) + summary_data.get('salaryUSD', 0) + summary_data.get('incentiveUSD', 0)
            
            # Admin Fee & Gifts Section (Rows 47-51)
            ws['C47'] = summary_data.get('familyCAD', 0)
            ws['D47'] = summary_data.get('familyUSD', 0)
            ws['C48'] = summary_data.get('medicalCAD', 0)
            ws['D48'] = summary_data.get('medicalUSD', 0)
            ws['C49'] = d30
            ws['D49'] = e30
            ws['C51'] = ws['C43'].value + d30
            ws['D51'] = ws['D43'].value + e30
            
            # Summary Statistics Section (Rows 55-58)
            ws['C55'] = summary_data.get('totalChildren', 0)
            ws['G55'] = exchange_rate
            ws['C56'] = summary_data.get('newChildrenCount', 0)
            
            total_usd = summary_data.get('totalUSD', 0)
            total_children = summary_data.get('totalChildren', 1)
            ws['G56'] = round(total_usd / total_children, 2) if total_children > 0 else 0
            
            ws['C57'] = summary_data.get('totalChildren', 0)
            ws['G57'] = summary_data.get('totalCAD', 0)
            ws['C58'] = len(regions_data)
            ws['G58'] = summary_data.get('totalUSD', 0)
            
            results['summary_updated'] = True
            print(f"   ✅ Summary populated with {summary_data.get('totalChildren', 0)} children")
            
        except Exception as e:
            error_msg = f"Summary error: {str(e)}"
            results['errors'].append(error_msg)
            print(f"   ❌ {error_msg}")
            print(traceback.format_exc())
        
        # ============================================
        # POPULATE REGION SHEETS
        # ============================================
        print(f"\n📍 PROCESSING {len(regions_data)} REGIONS...")
        
        if not regions_data:
            print(f"   ⚠️ WARNING: No regions data found!")
            print(f"   Check if regions are at data['regions'] or data['summary']['regions']")
        
        for idx, region in enumerate(regions_data):
            region_code = str(region.get('code', '')).strip().upper()
            
            print(f"\n[{idx+1}/{len(regions_data)}] Region: {region_code}")
            print(f"   Keys in region: {list(region.keys())}")
            
            if not region_code:
                results['regions_skipped'].append('UNKNOWN')
                print(f"   ⚠️ Skipped: No code")
                continue
            
            # Check if sheet exists (case-insensitive)
            sheet_found = None
            for sheet_name in wb.sheetnames:
                if sheet_name.upper() == region_code:
                    sheet_found = sheet_name
                    break
            
            if not sheet_found:
                results['regions_skipped'].append(f"{region_code} - not found")
                print(f"   ⚠️ Skipped: Sheet not in workbook")
                print(f"   Available: {', '.join(wb.sheetnames[:5])}...")
                continue
            
            try:
                ws = wb[sheet_found]
                
                # Period & Exchange Rate
                ws['B4'] = period_from
                ws['B5'] = period_to
                ws['H4'] = exchange_rate
                
                # Wire & Location Info
                ws['B10'] = str(region.get('wireId', ''))
                ws['B11'] = region.get('caseworker', 'Unknown')
                ws['B13'] = region.get('beneficiary', '')
                ws['B14'] = region.get('region', region_code)
                ws['B15'] = region.get('city', '')
                
                # Regular Support Section (Rows 20-24)
                ws['C20'] = region.get('children', 0)
                ws['D20'] = region.get('foodDistCAD', 0)
                ws['E20'] = region.get('foodDistUSD', 0)
                ws['D22'] = region.get('salaryCAD', 0)
                ws['E22'] = region.get('salaryUSD', 0)
                ws['D24'] = region.get('incentiveCAD', 0)
                ws['E24'] = region.get('incentiveUSD', 0)
                
                # D25 and E25 are SUBTOTALS
                d25 = region.get('foodDistCAD', 0) + region.get('salaryCAD', 0) + region.get('incentiveCAD', 0)
                e25 = region.get('foodDistUSD', 0) + region.get('salaryUSD', 0) + region.get('incentiveUSD', 0)
                ws['D25'] = d25
                ws['E25'] = e25
                
                # Additional Support (Rows 28-30)
                ws['D28'] = region.get('familyCAD', 0)
                ws['E28'] = region.get('familyUSD', 0)
                ws['D29'] = region.get('medicalCAD', 0)
                ws['E29'] = region.get('medicalUSD', 0)
                
                # D30 and E30 are SUBTOTALS
                d30 = region.get('familyCAD', 0) + region.get('medicalCAD', 0)
                e30 = region.get('familyUSD', 0) + region.get('medicalUSD', 0)
                ws['D30'] = d30
                ws['E30'] = e30
                
                # D32 and E32 are GRAND TOTALS
                ws['D32'] = d25 + d30
                ws['E32'] = e25 + e30
                
                # ============================================
                # POPULATE CHILDREN DETAILS TABLE (Row 36+)
                # ============================================
                child_details = region.get('childDetails', [])
                print(f"   👶 {len(child_details)} children")
                
                # Clear old data (rows 36-200)
                for row in range(36, 201):
                    for col in ['A', 'B', 'C', 'D', 'E', 'F', 'G']:
                        ws[f'{col}{row}'] = None
                
                # Write new children data starting at row 36
                if child_details:
                    start_row = 36
                    for child_idx, child in enumerate(child_details):
                        row = start_row + child_idx
                        
                        ws[f'A{row}'] = child.get('cspId', '')
                        ws[f'B{row}'] = child.get('childName', child.get('name', ''))
                        
                        # Handle foodAmount or foodDistUSD
                        food_usd = child.get('foodDistUSD', child.get('foodAmount', 0))
                        if food_usd == 0 and 'foodAmount' in child:
                            food_usd = child['foodAmount']
                        
                        medical = child.get('medicalGifts', 0)
                        family = child.get('familyGifts', 0)
                        
                        ws[f'C{row}'] = food_usd
                        ws[f'D{row}'] = medical
                        ws[f'E{row}'] = family
                        ws[f'F{row}'] = round(food_usd + medical + family, 2)
                        ws[f'G{row}'] = ''
                        
                        results['children_written'] += 1
                
                results['regions_processed'] += 1
                print(f"   ✅ Success")
                
            except Exception as e:
                error_msg = f"{region_code}: {str(e)}"
                results['errors'].append(error_msg)
                results['regions_skipped'].append(region_code)
                print(f"   ❌ Error: {str(e)}")
                print(traceback.format_exc())
        
        # ============================================
        # SAVE AND RETURN
        # ============================================
        output = io.BytesIO()
        wb.save(output)
        output.seek(0)
        
        print(f"\n{'='*60}")
        print(f"✅ COMPLETION:")
        print(f"   Summary: {'✓' if results['summary_updated'] else '✗'}")
        print(f"   Regions: {results['regions_processed']}/{len(regions_data)}")
        print(f"   Children: {results['children_written']}")
        print(f"   Skipped: {len(results['regions_skipped'])}")
        if results['regions_skipped']:
            print(f"   Skipped list: {', '.join(results['regions_skipped'][:10])}")
        print(f"   Errors: {len(results['errors'])}")
        print(f"{'='*60}\n")
        
        return send_file(
            output,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=f'CSP_Wiring_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx'
        )
    
    except Exception as e:
        error_msg = f"CRITICAL: {str(e)}"
        print(f"❌ {error_msg}")
        print(traceback.format_exc())
        return jsonify({"error": error_msg}), 500

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 8080))
    app.run(host='0.0.0.0', port=port)
