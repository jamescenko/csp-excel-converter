from flask import Flask, request, send_file, jsonify
from openpyxl import load_workbook
from openpyxl.styles import numbers
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
        
        print(f"\n{'='*80}")
        print(f"📥 RAW DATA RECEIVED:")
        print(f"Top-level keys: {list(data.keys())}")
        
        # Extract data
        summary_data = data.get('summary', {})
        exchange_rate = float(data.get('exchangeRate', 1.39))
        period_from = data.get('reportPeriodFrom', '01-01-2025')
        period_to = data.get('reportPeriodTo', '31-03-2025')
        
        # CRITICAL: Get regions from BOTH possible paths
        regions_data = data.get('regions', [])
        if not regions_data:
            regions_data = summary_data.get('regions', [])
        
        print(f"\n🔍 DATA ANALYSIS:")
        print(f"   Summary keys: {list(summary_data.keys())}")
        print(f"   Regions found: {len(regions_data)}")
        if regions_data:
            print(f"   First region: {regions_data[0].get('code', 'NO CODE')} - {len(regions_data[0].get('childDetails', []))} children")
        print(f"   Exchange rate: {exchange_rate}")
        print(f"{'='*80}\n")
        
        # Load template
        template_path = 'CSP_Automate_Template.xlsx'
        if not os.path.exists(template_path):
            return jsonify({"error": "Template not found"}), 500
        
        wb = load_workbook(template_path)
        print(f"✓ Template loaded: {len(wb.sheetnames)} sheets")
        
        results = {
            "summary_updated": False,
            "regions_processed": 0,
            "regions_skipped": [],
            "children_written": 0,
            "errors": []
        }
        
        # Currency format
        currency_format = '#,##0.00'
        
        # ============================================
        # POPULATE SUMMARY SHEET
        # ============================================
        try:
            ws = wb['Summary']
            print(f"\n📊 POPULATING SUMMARY...")
            
            # Period & Exchange Rate
            ws['B4'] = period_from
            ws['B5'] = period_to
            ws['H4'] = exchange_rate
            ws['H4'].number_format = '#,##0.00'
            
            # Regular Support Section (Rows 20-25)
            ws['C20'] = summary_data.get('totalChildren', 0)
            ws['D20'] = summary_data.get('foodDistCAD', 0)
            ws['D20'].number_format = currency_format
            ws['E20'] = summary_data.get('foodDistUSD', 0)
            ws['E20'].number_format = currency_format
            
            ws['D22'] = summary_data.get('salaryCAD', 0)
            ws['D22'].number_format = currency_format
            ws['E22'] = summary_data.get('salaryUSD', 0)
            ws['E22'].number_format = currency_format
            
            ws['D24'] = summary_data.get('incentiveCAD', 0)
            ws['D24'].number_format = currency_format
            ws['E24'] = summary_data.get('incentiveUSD', 0)
            ws['E24'].number_format = currency_format
            
            # D25 and E25 SUBTOTALS
            d25 = summary_data.get('foodDistCAD', 0) + summary_data.get('salaryCAD', 0) + summary_data.get('incentiveCAD', 0)
            e25 = summary_data.get('foodDistUSD', 0) + summary_data.get('salaryUSD', 0) + summary_data.get('incentiveUSD', 0)
            ws['D25'] = d25
            ws['D25'].number_format = currency_format
            ws['E25'] = e25
            ws['E25'].number_format = currency_format
            
            # Additional Support (Rows 28-30)
            ws['D28'] = summary_data.get('familyCAD', 0)
            ws['D28'].number_format = currency_format
            ws['E28'] = summary_data.get('familyUSD', 0)
            ws['E28'].number_format = currency_format
            
            ws['D29'] = summary_data.get('medicalCAD', 0)
            ws['D29'].number_format = currency_format
            ws['E29'] = summary_data.get('medicalUSD', 0)
            ws['E29'].number_format = currency_format
            
            # D30 and E30 SUBTOTALS
            d30 = summary_data.get('familyCAD', 0) + summary_data.get('medicalCAD', 0)
            e30 = summary_data.get('familyUSD', 0) + summary_data.get('medicalUSD', 0)
            ws['D30'] = d30
            ws['D30'].number_format = currency_format
            ws['E30'] = e30
            ws['E30'].number_format = currency_format
            
            # D32 and E32 GRAND TOTALS
            ws['D32'] = d25 + d30
            ws['D32'].number_format = currency_format
            ws['E32'] = e25 + e30
            ws['E32'].number_format = currency_format
            
            # Cross Check Section (Rows 36-43)
            ws['C36'] = summary_data.get('totalChildren', 0)
            ws['C37'] = summary_data.get('newChildrenCount', 0)
            
            ws['C40'] = summary_data.get('foodDistCAD', 0)
            ws['C40'].number_format = currency_format
            ws['D40'] = summary_data.get('foodDistUSD', 0)
            ws['D40'].number_format = currency_format
            
            ws['C41'] = summary_data.get('salaryCAD', 0)
            ws['C41'].number_format = currency_format
            ws['D41'] = summary_data.get('salaryUSD', 0)
            ws['D41'].number_format = currency_format
            
            ws['C42'] = summary_data.get('incentiveCAD', 0)
            ws['C42'].number_format = currency_format
            ws['D42'] = summary_data.get('incentiveUSD', 0)
            ws['D42'].number_format = currency_format
            
            # C43 and D43 Subtotals
            c43 = summary_data.get('foodDistCAD', 0) + summary_data.get('salaryCAD', 0) + summary_data.get('incentiveCAD', 0)
            d43 = summary_data.get('foodDistUSD', 0) + summary_data.get('salaryUSD', 0) + summary_data.get('incentiveUSD', 0)
            ws['C43'] = c43
            ws['C43'].number_format = currency_format
            ws['D43'] = d43
            ws['D43'].number_format = currency_format
            
            # Admin Fee & Gifts (Rows 47-51)
            ws['C47'] = summary_data.get('familyCAD', 0)
            ws['C47'].number_format = currency_format
            ws['D47'] = summary_data.get('familyUSD', 0)
            ws['D47'].number_format = currency_format
            
            ws['C48'] = summary_data.get('medicalCAD', 0)
            ws['C48'].number_format = currency_format
            ws['D48'] = summary_data.get('medicalUSD', 0)
            ws['D48'].number_format = currency_format
            
            ws['C49'] = d30
            ws['C49'].number_format = currency_format
            ws['D49'] = e30
            ws['D49'].number_format = currency_format
            
            ws['C51'] = c43 + d30
            ws['C51'].number_format = currency_format
            ws['D51'] = d43 + e30
            ws['D51'].number_format = currency_format
            
            # Summary Statistics (Rows 55-58) - COLUMN B NOT C!!!
            ws['B55'] = summary_data.get('totalChildren', 0)
            ws['G55'] = exchange_rate
            ws['G55'].number_format = '#,##0.00'
            
            ws['B56'] = summary_data.get('newChildrenCount', 0)
            total_usd = summary_data.get('totalUSD', 0)
            total_children = summary_data.get('totalChildren', 1)
            ws['G56'] = round(total_usd / total_children, 2) if total_children > 0 else 0
            ws['G56'].number_format = currency_format
            
            ws['B57'] = summary_data.get('totalChildren', 0)
            ws['G57'] = summary_data.get('totalCAD', 0)
            ws['G57'].number_format = currency_format
            
            ws['B58'] = len(regions_data)
            ws['G58'] = summary_data.get('totalUSD', 0)
            ws['G58'].number_format = currency_format
            
            results['summary_updated'] = True
            print(f"   ✅ Summary done: {summary_data.get('totalChildren', 0)} children, ${summary_data.get('totalCAD', 0):,.2f} CAD")
            
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
            print(f"   ⚠️ CRITICAL: NO REGIONS DATA!")
            print(f"   Checked: data['regions'] = {data.get('regions', 'NOT FOUND')}")
            print(f"   Checked: data['summary']['regions'] = {summary_data.get('regions', 'NOT FOUND')}")
            print(f"   Full data keys: {list(data.keys())}")
        
        for idx, region in enumerate(regions_data):
            region_code = str(region.get('code', '')).strip().upper()
            
            print(f"\n[{idx+1}/{len(regions_data)}] {region_code}")
            
            if not region_code:
                results['regions_skipped'].append('UNKNOWN')
                print(f"   ⚠️ Skip: No code")
                continue
            
            # Find sheet (case-insensitive)
            sheet_found = None
            for sheet_name in wb.sheetnames:
                if sheet_name.upper() == region_code:
                    sheet_found = sheet_name
                    break
            
            if not sheet_found:
                results['regions_skipped'].append(f"{region_code}")
                print(f"   ⚠️ Skip: Not in template")
                continue
            
            try:
                ws = wb[sheet_found]
                
                # Period & Exchange Rate
                ws['B4'] = period_from
                ws['B5'] = period_to
                ws['H4'] = exchange_rate
                ws['H4'].number_format = '#,##0.00'
                
                # Wire & Location
                ws['B10'] = str(region.get('wireId', ''))
                ws['B11'] = region.get('caseworker', 'Unknown')
                ws['B13'] = region.get('beneficiary', '')
                ws['B14'] = region.get('region', region_code)
                ws['B15'] = region.get('city', '')
                
                # Regular Support (Rows 20-24)
                ws['C20'] = region.get('children', 0)
                ws['D20'] = region.get('foodDistCAD', 0)
                ws['D20'].number_format = currency_format
                ws['E20'] = region.get('foodDistUSD', 0)
                ws['E20'].number_format = currency_format
                
                ws['D22'] = region.get('salaryCAD', 0)
                ws['D22'].number_format = currency_format
                ws['E22'] = region.get('salaryUSD', 0)
                ws['E22'].number_format = currency_format
                
                ws['D24'] = region.get('incentiveCAD', 0)
                ws['D24'].number_format = currency_format
                ws['E24'] = region.get('incentiveUSD', 0)
                ws['E24'].number_format = currency_format
                
                # D25/E25 SUBTOTALS
                d25 = region.get('foodDistCAD', 0) + region.get('salaryCAD', 0) + region.get('incentiveCAD', 0)
                e25 = region.get('foodDistUSD', 0) + region.get('salaryUSD', 0) + region.get('incentiveUSD', 0)
                ws['D25'] = d25
                ws['D25'].number_format = currency_format
                ws['E25'] = e25
                ws['E25'].number_format = currency_format
                
                # Additional Support (Rows 28-30)
                ws['D28'] = region.get('familyCAD', 0)
                ws['D28'].number_format = currency_format
                ws['E28'] = region.get('familyUSD', 0)
                ws['E28'].number_format = currency_format
                
                ws['D29'] = region.get('medicalCAD', 0)
                ws['D29'].number_format = currency_format
                ws['E29'] = region.get('medicalUSD', 0)
                ws['E29'].number_format = currency_format
                
                # D30/E30 SUBTOTALS
                d30 = region.get('familyCAD', 0) + region.get('medicalCAD', 0)
                e30 = region.get('familyUSD', 0) + region.get('medicalUSD', 0)
                ws['D30'] = d30
                ws['D30'].number_format = currency_format
                ws['E30'] = e30
                ws['E30'].number_format = currency_format
                
                # D32/E32 GRAND TOTALS
                ws['D32'] = d25 + d30
                ws['D32'].number_format = currency_format
                ws['E32'] = e25 + e30
                ws['E32'].number_format = currency_format
                
                # Children Table (Row 36+)
                child_details = region.get('childDetails', [])
                print(f"   {len(child_details)} kids")
                
                # Clear old data
                for row in range(36, 201):
                    for col in ['A', 'B', 'C', 'D', 'E', 'F', 'G']:
                        ws[f'{col}{row}'] = None
                
                # Write children
                if child_details:
                    for child_idx, child in enumerate(child_details):
                        row = 36 + child_idx
                        
                        ws[f'A{row}'] = child.get('cspId', '')
                        ws[f'B{row}'] = child.get('childName', child.get('name', ''))
                        
                        food_usd = child.get('foodDistUSD', child.get('foodAmount', 0))
                        medical = child.get('medicalGifts', 0)
                        family_gift = child.get('familyGifts', 0)
                        
                        ws[f'C{row}'] = food_usd
                        ws[f'C{row}'].number_format = currency_format
                        ws[f'D{row}'] = medical
                        ws[f'D{row}'].number_format = currency_format
                        ws[f'E{row}'] = family_gift
                        ws[f'E{row}'].number_format = currency_format
                        ws[f'F{row}'] = round(food_usd + medical + family_gift, 2)
                        ws[f'F{row}'].number_format = currency_format
                        ws[f'G{row}'] = ''
                        
                        results['children_written'] += 1
                
                results['regions_processed'] += 1
                print(f"   ✅ Done")
                
            except Exception as e:
                error_msg = f"{region_code}: {str(e)}"
                results['errors'].append(error_msg)
                results['regions_skipped'].append(region_code)
                print(f"   ❌ {str(e)}")
        
        # SAVE
        output = io.BytesIO()
        wb.save(output)
        output.seek(0)
        
        print(f"\n{'='*80}")
        print(f"✅ FINAL RESULTS:")
        print(f"   Summary: {'✓' if results['summary_updated'] else '✗'}")
        print(f"   Regions: {results['regions_processed']}/{len(regions_data)}")
        print(f"   Children: {results['children_written']}")
        print(f"   Skipped: {', '.join(results['regions_skipped'][:5]) if results['regions_skipped'] else 'none'}")
        print(f"{'='*80}\n")
        
        return send_file(
            output,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=f'CSP_Wiring_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx'
        )
    
    except Exception as e:
        print(f"❌ CRITICAL: {str(e)}")
        print(traceback.format_exc())
        return jsonify({"error": str(e)}), 500

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 8080))
    app.run(host='0.0.0.0', port=port)
