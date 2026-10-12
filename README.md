# Area_Velocity_Fully_Dynamic.xlsx
import openpyxl

def generate_fully_dynamic_excel(output_filename="Area_Velocity_Fully_Dynamic.xlsx"):
    wb = openpyxl.Workbook()
    sheet = wb.active
    sheet.title = "Area-Velocity"
    
    # Ensure grid lines are visible
    sheet.views.sheetView[0].showGridLines = True
    
    # Title & Header
    sheet['A1'] = "AREA–VELOCITY METHOD | Editable Calculation Sheet"
    sheet['A3'] = "Formula"
    sheet['B3'] = "Q = A × V"
    sheet['D3'] = "Blue cells = editable input values; calculated cells contain formulas."
    
    sheet['A4'] = "Discharge (Q)"
    sheet['B4'] = "m³/s"
    
    sheet['A6'] = "PARAMETERS"
    sheet['A7'] = "Date"
    sheet['B7'] = "08/10/2026"
    sheet['C7'] = "Start MSL"
    sheet['D7'] = 0
    sheet['E7'] = "End MSL"
    sheet['F7'] = 0
    sheet['G7'] = "Mean MSL"
    sheet['H7'] = "=ROUND(AVERAGE(D7,F7),3)"
    
    sheet['A8'] = "Start RD"
    sheet['B8'] = 103  # Default Start RD
    sheet['C8'] = "End RD"
    sheet['D8'] = 120  # Default End RD
    sheet['E8'] = "Slope (S)"
    sheet['F8'] = 0.0005
    sheet['G8'] = "Velocity factor"
    sheet['H8'] = 0.89
    
    # Table Headers
    headers = ["#", "RD", "DEPTH", "WIDTH", "D/D", "WET CORR", "AREA", "REVS", "VEL", "NEW VEL (× factor)", "Q"]
    for col_idx, h in enumerate(headers, start=1):
        sheet.cell(10, col_idx).value = h
        
    # Row 11: Left Bank (W/E @ L/B)
    sheet.cell(11, 1).value = 1
    sheet.cell(11, 2).value = "=$B$8+0.2"
    sheet.cell(11, 3).value = "W/E @ L/B"
    
    # Rows 12 to 91 (80 dynamic calculation steps supporting up to 40m span at 0.5m intervals)
    max_steps = 80
    start_data_row = 12
    end_data_row = start_data_row + max_steps - 1
    
    for i in range(1, max_steps + 1):
        r = start_data_row + i - 1
        
        if i == 1:
            rd_formula = f'=IF(($B$8 + 0.5) < $D$8, $B$8 + 0.5, IF(($B$8 + 0.5) >= $D$8, $D$8, ""))'
        else:
            prev_r = r - 1
            rd_formula = f'=IF(B{prev_r}="", "", IF(B{prev_r} < $D$8, IF(($B$8 + {i}*0.5) < $D$8, $B$8 + {i}*0.5, $D$8), ""))'
            
        sheet.cell(r, 1).value = f'=IF(B{r}<>"", {i+1}, "")'
        sheet.cell(r, 2).value = rd_formula
        sheet.cell(r, 3).value = f'=IF(B{r}="", "", IF(B{r}=$D$8, "W/E @ R/B", 0.5))'
        sheet.cell(r, 4).value = f'=IF(OR(B{r}="", C{r}="W/E @ R/B"), "", 0.5)'
        sheet.cell(r, 5).value = f'=IF(OR(C{r}="", C{r}="W/E @ R/B", C{r-1}="W/E @ R/B"), "", ABS(C{r}-C{r-1}))'
        sheet.cell(r, 6).value = f'=IF(OR(D{r}="", E{r}=""), "", SQRT(D{r}^2+E{r}^2)-D{r})'
        sheet.cell(r, 7).value = f'=IF(OR(C{r}="", D{r}="", C{r}="W/E @ R/B"), "", C{r}*D{r})'
        sheet.cell(r, 8).value = f'=IF(OR(B{r}="", C{r}="W/E @ R/B"), "", 10)'
        sheet.cell(r, 9).value = f'=IF(H{r}="", "", H{r}*0.02)'
        sheet.cell(r, 10).value = f'=IF(I{r}="", "", I{r}*$H$8)'
        sheet.cell(r, 11).value = f'=IF(OR(G{r}="", J{r}=""), "", G{r}*J{r})'

    # Results Section
    res_start = end_data_row + 3
    
    sheet.cell(res_start, 1).value = "RESULTS"
    
    results_mapping = [
        ("Top Width (m)", f'=LOOKUP(2, 1/(B{start_data_row}:B{end_data_row}<>""), B{start_data_row}:B{end_data_row}) - B11', "End RD - first RD"),
        ("Total Wet Corr (m)", f'=SUM(F{start_data_row}:F{end_data_row})', "Σ wet correction"),
        ("Wetted Perimeter (m)", f'=B{res_start + 1}+B{res_start + 2}', "Top width + wet correction"),
        ("Total Area (m²)", f'=SUM(G{start_data_row}:G{end_data_row})', "Σ(width × depth)"),
        ("Total Discharge Q (m³/s)", f'=SUM(K{start_data_row}:K{end_data_row})', "Σ(area × adjusted velocity)"),
        ("Max Point Velocity (m/s)", f'=IFERROR(MAX(J{start_data_row}:J{end_data_row}),0)', "Maximum adjusted velocity"),
        ("Hydraulic Radius R (m)", f'=IFERROR(B{res_start + 4}/B{res_start + 3},0)', "Area / wetted perimeter"),
        ("Hydraulic Depth (m)", f'=IFERROR(B{res_start + 4}/B{res_start + 1},0)', "Area / top width"),
        ("Mean Velocity (m/s)", f'=IFERROR(B{res_start + 5}/B{res_start + 4},0)', "Q / area"),
        ("Chezy C", f'=IFERROR(B{res_start + 9}/SQRT(B{res_start + 7}*$F$8),0)', "V / √(R × S)"),
        ("Manning n", f'=IFERROR(B{res_start + 7}^(1/6)/B{res_start + 10},0)', "R^(1/6) / C")
    ]
    
    for idx, (label, formula, note) in enumerate(results_mapping, start=1):
        r = res_start + idx
        sheet.cell(r, 1).value = label
        sheet.cell(r, 2).value = formula
        sheet.cell(r, 3).value = note

    wb.save(output_filename)
    print(f"Successfully generated {output_filename}")

generate_fully_dynamic_excel()
