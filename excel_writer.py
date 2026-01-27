import pandas as pd
import numpy as np
import xlsxwriter
import os
import subprocess
from datetime import datetime

def write_excel_report(master, rules_cannabis, rules_accessory, report_days, log_func):
    """
    Extracts Excel formatting logic from hh_logic.py into a dedicated module.
    Generates a formatted report with Control Panel and Order Builder sheets.
    """
    output_filename = f'Order_Rec_{datetime.now().strftime("%Y-%m-%d_%H%M")}.xlsx'
    log_func(f"Writing Excel: {output_filename}...")
    
    cols_static = ['SKU', 'Product Name', 'Category', 'Brand', 'Case_Size', 'Case_Cost', 'New_SKU_This_Week', 'Available_Cases']
    
    writer = pd.ExcelWriter(output_filename, engine='xlsxwriter')
    workbook = writer.book
    
    # Define formats
    # === DEFINE FORMATS WITH PREMIUM COLOR THEME ===
    border_col = '#E2E8F0'
    fmt_header = workbook.add_format({'bold': True, 'bg_color': '#1E293B', 'font_color': '#FFFFFF', 'border': 1, 'border_color': border_col, 'align': 'center', 'valign': 'vcenter'})
    fmt_text = workbook.add_format({'bg_color': '#FFFFFF', 'border': 1, 'border_color': border_col, 'valign': 'vcenter'})
    fmt_num = workbook.add_format({'bg_color': '#FFFFFF', 'border': 1, 'border_color': border_col, 'align': 'center', 'valign': 'vcenter'})
    fmt_dec = workbook.add_format({'num_format': '0.00', 'bg_color': '#FFFFFF', 'border': 1, 'border_color': border_col, 'align': 'center', 'valign': 'vcenter'})
    fmt_pct = workbook.add_format({'num_format': '0.0%', 'bg_color': '#FFFFFF', 'border': 1, 'border_color': border_col, 'align': 'center', 'valign': 'vcenter'})
    fmt_curr = workbook.add_format({'num_format': '$#,##0.00', 'bg_color': '#FFFFFF', 'border': 1, 'border_color': border_col, 'valign': 'vcenter'})
    fmt_inc = workbook.add_format({'bg_color': '#F8FAFC', 'border': 1, 'border_color': border_col, 'align': 'center', 'valign': 'vcenter'})
    fmt_soq = workbook.add_format({'bold': True, 'bg_color': '#FFEB3B', 'font_color': '#1F1F1F', 'border': 1, 'border_color': border_col, 'align': 'center', 'valign': 'vcenter'})

    
    # Control Panel formats - PREMIUM THEME
    fmt_control = workbook.add_format({'bold': True, 'bg_color': '#1E293B', 'font_color': '#FFFFFF', 'border': 1, 'align': 'center', 'valign': 'vcenter', 'font_size': 14})
    fmt_section_title = workbook.add_format({'bold': True, 'bg_color': '#F8FAFC', 'font_color': '#334155', 'border': 1, 'align': 'left', 'valign': 'vcenter', 'font_size': 11})
    fmt_label = workbook.add_format({'bg_color': '#F1F5F9', 'font_color': '#64748B', 'border': 1, 'align': 'right', 'valign': 'vcenter'})
    fmt_input = workbook.add_format({'bg_color': '#FFFFFF', 'font_color': '#0F172A', 'border': 1, 'align': 'center', 'valign': 'vcenter', 'num_format': '0.00'})
    fmt_readonly = workbook.add_format({'bg_color': '#F8FAFC', 'font_color': '#475569', 'border': 1, 'align': 'center', 'valign': 'vcenter', 'num_format': '0.00'})
    fmt_curr_input = workbook.add_format({'bg_color': '#FFFFFF', 'font_color': '#0F172A', 'border': 1, 'align': 'center', 'valign': 'vcenter', 'num_format': '$#,##0.00'})
    fmt_curr_readonly = workbook.add_format({'bg_color': '#FEF9C3', 'font_color': '#854D0E', 'border': 1, 'align': 'center', 'valign': 'vcenter', 'num_format': '$#,##0.00', 'bold': True})

    
    # Location-specific header formats (Modernized Palette)
    loc_configs = [
        ('Hill', workbook.add_format({'bold': True, 'bg_color': '#DBEAFE', 'font_color': '#1E40AF', 'border': 1, 'border_color': border_col})),
        ('Valley', workbook.add_format({'bold': True, 'bg_color': '#FEF3C7', 'font_color': '#92400E', 'border': 1, 'border_color': border_col})),
        ('Jasper', workbook.add_format({'bold': True, 'bg_color': '#D1FAE5', 'font_color': '#065F46', 'border': 1, 'border_color': border_col}))
    ]

    
    # Status-specific formats (Contemporary Muted Tints)
    fmt_status_hot = workbook.add_format({'bg_color': '#FEE2E2', 'font_color': '#991B1B', 'bold': True, 'border': 1, 'border_color': border_col})
    fmt_status_reorder = workbook.add_format({'bg_color': '#FFEDD5', 'font_color': '#9A3412', 'bold': True, 'border': 1, 'border_color': border_col})
    fmt_status_good = workbook.add_format({'bg_color': '#DCFCE7', 'font_color': '#166534', 'border': 1, 'border_color': border_col})
    fmt_status_filler = workbook.add_format({'bg_color': '#FEF9C3', 'font_color': '#854D0E', 'border': 1, 'border_color': border_col})
    fmt_status_dead = workbook.add_format({'bg_color': '#F1F5F9', 'font_color': '#475569', 'border': 1, 'border_color': border_col})
    fmt_status_new = workbook.add_format({'bg_color': '#F5F3FF', 'font_color': '#5B21B6', 'border': 1, 'border_color': border_col})
    fmt_status_cold = workbook.add_format({'bg_color': '#ECFEFF', 'font_color': '#155E75', 'border': 1, 'border_color': border_col})

    
    # === CREATE CONTROL PANEL SHEET ===
    control_sheet = workbook.add_worksheet("Control Panel")
    writer.sheets['Control Panel'] = control_sheet
    
    # Global styles
    control_sheet.set_default_row(22)
    
    # Control Panel Header
    control_sheet.merge_range(0, 0, 0, 8, '🎯 LOGIC CONTROL CENTER', fmt_control)
    control_sheet.set_row(0, 30)
    
    # Row 2: Global Settings
    row_idx = 2
    control_sheet.write(row_idx, 0, '⚙️ GLOBAL SETTINGS', fmt_section_title)
    control_sheet.write(row_idx, 1, 'Analysis Period:', fmt_label)
    control_sheet.write(row_idx, 2, report_days, fmt_input)  # C3
    control_sheet.write(row_idx, 3, 'Weeks Factor:', fmt_label)
    control_sheet.write_formula(row_idx, 4, '=C3/7', fmt_readonly)  # E3
    control_sheet.write(row_idx, 8, report_days, fmt_readonly)  # I3 - hidden reference
    # Fill gaps
    for c in [5, 6, 7]: control_sheet.write(row_idx, c, "", fmt_label)
    
    # Row 3: Cannabis Business Rules
    row_idx = 3
    control_sheet.write(row_idx, 0, '🌿 CANNABIS RULES', fmt_section_title)
    control_sheet.write(row_idx, 1, 'Hot Vel:', fmt_label)
    control_sheet.write(row_idx, 2, rules_cannabis['hot_velocity'], fmt_input) # C4
    control_sheet.write(row_idx, 3, 'Reorder:', fmt_label)
    control_sheet.write(row_idx, 4, rules_cannabis['reorder_point'], fmt_input) # E4
    control_sheet.write(row_idx, 5, 'Target:', fmt_label)
    control_sheet.write(row_idx, 6, rules_cannabis['target_wos'], fmt_input) # G4
    control_sheet.write(row_idx, 7, 'Dead:', fmt_label)
    control_sheet.write(row_idx, 8, rules_cannabis['dead_wos'], fmt_input) # I4
    
    # Row 4: Accessory Business Rules
    row_idx = 4
    control_sheet.write(row_idx, 0, '📦 ACCESSORY RULES', fmt_section_title)
    control_sheet.write(row_idx, 1, 'Hot Vel:', fmt_label)
    control_sheet.write(row_idx, 2, rules_accessory['hot_velocity'], fmt_input) # C5
    control_sheet.write(row_idx, 3, 'Reorder:', fmt_label)
    control_sheet.write(row_idx, 4, rules_accessory['reorder_point'], fmt_input) # E5
    control_sheet.write(row_idx, 5, 'Target:', fmt_label)
    control_sheet.write(row_idx, 6, rules_accessory['target_wos'], fmt_input) # G5
    control_sheet.write(row_idx, 7, 'Dead:', fmt_label)
    control_sheet.write(row_idx, 8, rules_accessory['dead_wos'], fmt_input) # I5

    # Row 6: Financial Totals
    row_idx = 6
    control_sheet.write(row_idx, 0, '💰 FINANCIAL SUMMARY', fmt_section_title)
    control_sheet.write(row_idx, 1, 'Hill Total:', fmt_label)
    control_sheet.write(row_idx, 3, 'Valley Total:', fmt_label)
    control_sheet.write(row_idx, 5, 'Jasper Total:', fmt_label)
    control_sheet.write(row_idx, 7, 'Grand Total:', fmt_label)
    # Buffers
    control_sheet.write(row_idx, 2, 0, fmt_curr_readonly)
    control_sheet.write(row_idx, 4, 0, fmt_curr_readonly)
    control_sheet.write(row_idx, 6, 0, fmt_curr_readonly)
    control_sheet.write(row_idx, 8, 0, fmt_curr_readonly)



    
    # Set column widths
    control_sheet.set_column(0, 0, 35) # Section titles
    for col_i in range(1, 9):
        control_sheet.set_column(col_i, col_i, 20) # Balanced pairs


    
    # === CREATE ORDER BUILDER SHEET ===
    worksheet = workbook.add_worksheet("Order Builder")
    writer.sheets['Order Builder'] = worksheet
    
    # Global styles
    worksheet.set_default_row(18)
    
    header_row = 0
    worksheet.set_row(header_row, 25) # Main header height
    current_col = 0

    
    key_group_start = current_col
    col_display_names = {
        'SKU': 'SKU', 'Product Name': 'Product Name', 'Category': 'Category', 
        'Brand': 'Brand', 'Case_Size': 'Case Size', 'Case_Cost': 'Case Cost', 
        'New_SKU_This_Week': 'New SKU This Week', 'Available_Cases': 'Available Cases'
    }
    worksheet.write(header_row, current_col, "Key", fmt_header)
    current_col += 1
    for col in cols_static:
        display_name = col_display_names.get(col, col.replace('_', ' '))
        worksheet.write(header_row, current_col, display_name, fmt_header)
        current_col += 1
    key_group_end = current_col - 1
    
    # Metrics per location: 9 columns (Stock now includes incoming display)
    metrics = [
        ('Status', 'Status', fmt_text), 
        ('Stock', 'StockDisplay', fmt_text),   # Combined: "5 + 12 🚚"
        ('Buy', 'SOQ', fmt_soq), 
        ('Sold', 'Sold', fmt_num), 
        ('Gross', 'Gross', fmt_curr), 
        ('Profit', 'Profit', fmt_curr), 
        ('Mrg%', 'Mrg', fmt_pct), 
        ('Vel', 'Vel', fmt_dec), 
        ('WOS', 'WOS', fmt_dec)
    ]
    
    # Location emojis for headers
    loc_emojis = {'Hill': '🔵', 'Valley': '🟠', 'Jasper': '🟢'}
    
    loc_groups = []
    for loc_name, fmt in loc_configs:
        loc_start = current_col
        emoji = loc_emojis.get(loc_name, '')
        for i, (title, _, _) in enumerate(metrics):
            if i == 0:
                # First column shows location name with emoji
                worksheet.write(header_row, current_col, f"{emoji} {loc_name}", fmt)
            else:
                # Other columns just show metric name
                worksheet.write(header_row, current_col, title, fmt)
            current_col += 1
        loc_end = current_col - 1
        loc_groups.append((loc_name, loc_start, loc_end))

    
    # Sort data
    master_sorted = master.sort_values(
        by=['Hill_Status', 'Valley_Status', 'Jasper_Status', 'SKU'],
        ascending=[True, True, True, True]
    ).reset_index(drop=True)
    
    data_end_row = header_row + len(master_sorted)
    
    def col_letter(col_idx):
        result = ""
        col_idx += 1
        while col_idx > 0:
            col_idx -= 1
            result = chr(65 + (col_idx % 26)) + result
            col_idx //= 26
        return result
    
    key_col_letter = col_letter(key_group_start)
    case_size_col_idx = cols_static.index('Case_Size') + key_group_start + 1
    case_cost_col_idx = cols_static.index('Case_Cost') + key_group_start + 1
    case_size_col_letter = col_letter(case_size_col_idx)
    case_cost_col_letter = col_letter(case_cost_col_idx)
    
    hill_buy_col = key_group_end + 1 + 2
    valley_buy_col = hill_buy_col + len(metrics)
    jasper_buy_col = valley_buy_col + len(metrics)
    
    hill_buy_col_letter = col_letter(hill_buy_col)
    valley_buy_col_letter = col_letter(valley_buy_col)
    jasper_buy_col_letter = col_letter(jasper_buy_col)
    
    first_data_row = header_row + 2
    last_data_row = data_end_row + 1
    
    # Budget tracking formulas
    r = 6
    for loc_char, col_i in [("H", 2), ("V", 4), ("J", 6)]:
        formula = (
            f'=SUMPRODUCT((LEN(\'Order Builder\'!{key_col_letter}${first_data_row}:{key_col_letter}${last_data_row}) - '
            f'LEN(SUBSTITUTE(UPPER(\'Order Builder\'!{key_col_letter}${first_data_row}:{key_col_letter}${last_data_row}), "{loc_char}", "")))*'
            f'(\'Order Builder\'!{case_cost_col_letter}${first_data_row}:{case_cost_col_letter}${last_data_row}))'
        )
        control_sheet.write_formula(r, col_i, formula, fmt_curr_readonly)
    
    # Grand Total
    control_sheet.write_formula(6, 8, '=C7+E7+G7', fmt_curr_readonly)


    
    # Row 8: Order Counts
    row_idx = 8
    control_sheet.write(row_idx, 0, '📊 ORDER COUNTS', fmt_section_title)
    control_sheet.write(row_idx, 1, 'Total SKUs:', fmt_label)
    control_sheet.write_formula(row_idx, 2, f'=COUNTA(\'Order Builder\'!{col_letter(key_group_start)}{first_data_row}:\'Order Builder\'!{col_letter(key_group_start)}{last_data_row})', fmt_readonly)
    
    control_sheet.write(row_idx, 3, 'Hill SKUs:', fmt_label)
    control_sheet.write_formula(row_idx, 4, f'=COUNTIF(\'Order Builder\'!{hill_buy_col_letter}{first_data_row}:\'Order Builder\'!{hill_buy_col_letter}{last_data_row},"<>-")', fmt_readonly)
    
    control_sheet.write(row_idx, 5, 'Valley SKUs:', fmt_label)
    control_sheet.write_formula(row_idx, 6, f'=COUNTIF(\'Order Builder\'!{valley_buy_col_letter}{first_data_row}:\'Order Builder\'!{valley_buy_col_letter}{last_data_row},"<>-")', fmt_readonly)
    
    control_sheet.write(row_idx, 7, 'Jasper SKUs:', fmt_label)
    control_sheet.write_formula(row_idx, 8, f'=COUNTIF(\'Order Builder\'!{jasper_buy_col_letter}{first_data_row}:\'Order Builder\'!{jasper_buy_col_letter}{last_data_row},"<>-")', fmt_readonly)


    
    # Summary Tables
    sku_col_letter = col_letter(key_group_start + 1)
    summary_start_row = 11
    r = summary_start_row
    labels = ['SKU', 'CASE COUNT', 'SKU', 'Quantity', 'SKU', 'Quantity', 'SKU', 'Quantity', 'COST']
    
    # Export Header Styling
    fmt_export_header = workbook.add_format({'bold': True, 'bg_color': '#334155', 'font_color': '#FFFFFF', 'border': 1, 'align': 'center'})
    for i, lbl in enumerate(labels):
        control_sheet.write(r, i, lbl, fmt_export_header)
    
    r = summary_start_row + 1
    fmt_loc_mark = workbook.add_format({'bold': True, 'bg_color': '#CBD5E1', 'font_color': '#1E293B', 'border': 1, 'align': 'center'})
    control_sheet.write(r, 2, 'Items to Order (Hill)', fmt_loc_mark)
    control_sheet.write(r, 4, 'Items to Order (Valley)', fmt_loc_mark)
    control_sheet.write(r, 6, 'Items to Order (Jasper)', fmt_loc_mark)
    # Fill remaining cells in this header row
    for c in [0, 1, 3, 5, 7, 8]: control_sheet.write(r, c, "", fmt_loc_mark)

    
    data_start_row_summary = summary_start_row + 2
    max_summary_rows = 500
    
    for i in range(max_summary_rows):
        ds_r = data_start_row_summary + i
        # SKU master list
        sku_formula = (
            f'=IFERROR(INDEX(\'Order Builder\'!{sku_col_letter}${first_data_row}:{sku_col_letter}${last_data_row},'
            f'SMALL(IF((\'Order Builder\'!{key_col_letter}${first_data_row}:{key_col_letter}${last_data_row}<>"")*'
            f'((ISNUMBER(SEARCH("H",UPPER(\'Order Builder\'!{key_col_letter}${first_data_row}:{key_col_letter}${last_data_row}))))+'
            f'(ISNUMBER(SEARCH("V",UPPER(\'Order Builder\'!{key_col_letter}${first_data_row}:{key_col_letter}${last_data_row}))))+'
            f'(ISNUMBER(SEARCH("J",UPPER(\'Order Builder\'!{key_col_letter}${first_data_row}:{key_col_letter}${last_data_row}))))>0),'
            f'ROW(\'Order Builder\'!{key_col_letter}${first_data_row}:{key_col_letter}${last_data_row})-ROW(\'Order Builder\'!{key_col_letter}${first_data_row})+1),{i+1})),"")'
        )
        control_sheet.write_array_formula(ds_r, 0, ds_r, 0, sku_formula, fmt_readonly)
        
        cur_sku_ref = f'{col_letter(0)}{ds_r+1}'
        # Case count
        control_sheet.write_formula(ds_r, 1, f'=IF({cur_sku_ref}="","",IFERROR(LEN(TRIM(INDEX(\'Order Builder\'!{key_col_letter}${first_data_row}:{key_col_letter}${last_data_row},MATCH({cur_sku_ref},\'Order Builder\'!{sku_col_letter}${first_data_row}:{sku_col_letter}${last_data_row},0)))),0))', fmt_readonly)
        
        # Loc specific SKUs and Quantities
        for loc_c, sku_col, qty_col in [("H", 2, 3), ("V", 4, 5), ("J", 6, 7)]:
            sku_loc_f = f'=IF({cur_sku_ref}="","",IF(ISNUMBER(SEARCH("{loc_c}",UPPER(TRIM(IFERROR(INDEX(\'Order Builder\'!{key_col_letter}${first_data_row}:{key_col_letter}${last_data_row},MATCH({cur_sku_ref},\'Order Builder\'!{sku_col_letter}${first_data_row}:{sku_col_letter}${last_data_row},0)),""))))),{cur_sku_ref},""))'
            control_sheet.write_formula(ds_r, sku_col, sku_loc_f, fmt_readonly)
            
            qty_loc_f = (
                f'=IF({col_letter(sku_col)}{ds_r+1}="","",IFERROR(INDEX(\'Order Builder\'!{case_size_col_letter}${first_data_row}:{case_size_col_letter}${last_data_row},MATCH({cur_sku_ref},\'Order Builder\'!{sku_col_letter}${first_data_row}:{sku_col_letter}${last_data_row},0))*'
                f'(LEN(TRIM(INDEX(\'Order Builder\'!{key_col_letter}${first_data_row}:{key_col_letter}${last_data_row},MATCH({cur_sku_ref},\'Order Builder\'!{sku_col_letter}${first_data_row}:{sku_col_letter}${last_data_row},0))))-'
                f'LEN(SUBSTITUTE(UPPER(TRIM(INDEX(\'Order Builder\'!{key_col_letter}${first_data_row}:{key_col_letter}${last_data_row},MATCH({cur_sku_ref},\'Order Builder\'!{sku_col_letter}${first_data_row}:{sku_col_letter}${last_data_row},0)))),"{loc_c}",""))),""))'
            )
            control_sheet.write_formula(ds_r, qty_col, qty_loc_f, fmt_readonly)
            
        # Cost = Case Count × Case Cost
        case_count_ref = f'B{ds_r+1}'
        control_sheet.write_formula(ds_r, 8, 
            f'=IF({cur_sku_ref}="","",{case_count_ref}*IFERROR(INDEX(\'Order Builder\'!{case_cost_col_letter}${first_data_row}:{case_cost_col_letter}${last_data_row},MATCH({cur_sku_ref},\'Order Builder\'!{sku_col_letter}${first_data_row}:{sku_col_letter}${last_data_row},0)),0))', 
            fmt_curr_readonly)


    # Write Data Rows
    fmt_key = workbook.add_format({'bg_color': '#FFFACD', 'font_color': '#1F1F1F', 'bold': True, 'border': 1, 'align': 'center'})
    row_bg_even = '#FFFFFF'
    row_bg_odd = '#F9F9F9'
    
    for r_idx, (_, row) in enumerate(master_sorted.iterrows()):
        xls_r = header_row + 1 + r_idx
        row_bg = row_bg_even if r_idx % 2 == 0 else row_bg_odd
        c_idx = 0
        worksheet.write(xls_r, c_idx, "", fmt_key)
        c_idx += 1
        
        for col in cols_static:
            val = row.get(col, "")
            if pd.isna(val): val = 0 if col == 'Available_Cases' else ""
            
            if col == 'Case_Cost':
                f = workbook.add_format({'num_format': '$#,##0.00', 'bg_color': row_bg, 'border': 1, 'border_color': border_col, 'valign': 'vcenter'})
            elif col == 'Available_Cases':
                f = workbook.add_format({'num_format': '0', 'bg_color': row_bg, 'border': 1, 'border_color': border_col, 'valign': 'vcenter', 'align': 'center'})
            elif col == 'New_SKU_This_Week' and str(val).upper() in ['TRUE', 'YES', '1', 'X']:
                f = workbook.add_format({'bg_color': '#F5F3FF', 'font_color': '#5B21B6', 'bold': True, 'border': 1, 'border_color': border_col, 'valign': 'vcenter', 'align': 'center'})
            else:
                f = workbook.add_format({'bg_color': row_bg, 'border': 1, 'border_color': border_col, 'valign': 'vcenter'})
            worksheet.write(xls_r, c_idx, val, f)
            c_idx += 1

            
        loc_col_info = {}
        for loc_name, _ in loc_configs:
            # Status
            worksheet.write(xls_r, c_idx, row.get(f"{loc_name}_Status", ""), workbook.add_format({'bg_color': row_bg, 'border': 1, 'border_color': border_col, 'align': 'center', 'valign': 'vcenter'}))
            c_idx += 1
            
            # Stock Display (combined: "5 + 12 🚚")
            worksheet.write(xls_r, c_idx, row.get(f"{loc_name}_StockDisplay", "0"), workbook.add_format({'bg_color': row_bg, 'border': 1, 'border_color': border_col, 'align': 'center', 'valign': 'vcenter'}))
            c_idx += 1

            
            # SOQ placeholder
            soq_c = c_idx
            worksheet.write(xls_r, soq_c, "-", fmt_soq)
            c_idx += 1
            
            # Scaled Period Formulas
            orig_p = "'Control Panel'!I3"
            curr_p = "'Control Panel'!C3"

            
            # Sold
            worksheet.write_formula(xls_r, c_idx, f'={row.get(f"{loc_name}_Sold", 0)}*({curr_p}/{orig_p})', workbook.add_format({'bg_color': row_bg, 'border': 1, 'border_color': border_col, 'valign': 'vcenter', 'align': 'center'}))
            sold_l = col_letter(c_idx)
            c_idx += 1
            
            # Gross
            worksheet.write_formula(xls_r, c_idx, f'={row.get(f"{loc_name}_Gross", 0)}*({curr_p}/{orig_p})', workbook.add_format({'num_format': '$#,##0.00', 'bg_color': row_bg, 'border': 1, 'border_color': border_col, 'valign': 'vcenter'}))
            c_idx += 1
            
            # Profit
            worksheet.write_formula(xls_r, c_idx, f'={row.get(f"{loc_name}_Profit", 0)}*({curr_p}/{orig_p})', workbook.add_format({'num_format': '$#,##0.00', 'bg_color': row_bg, 'border': 1, 'border_color': border_col, 'valign': 'vcenter'}))
            prof_l = col_letter(c_idx)
            c_idx += 1
            
            # Net Sales hidden ref for Margin
            net_sales_ref_c = key_group_end + len(metrics) * 3 + loc_configs.index((loc_name, _)) + 1
            worksheet.write_formula(xls_r, net_sales_ref_c, f'={row.get(f"{loc_name}_Net", 0)}*({curr_p}/{orig_p})', fmt_curr)
            net_s_l = col_letter(net_sales_ref_c)
            
            # Margin
            worksheet.write_formula(xls_r, c_idx, f'=IF({net_s_l}{xls_r+1}<>0, {prof_l}{xls_r+1}/{net_s_l}{xls_r+1}, 0)', workbook.add_format({'num_format': '0.0%', 'bg_color': row_bg, 'border': 1, 'border_color': border_col, 'valign': 'vcenter', 'align': 'center'}))
            c_idx += 1
            
            # Velocity
            w_fact = "'Control Panel'!E3"
            worksheet.write_formula(xls_r, c_idx, f'=IF({sold_l}{xls_r+1}>0, {sold_l}{xls_r+1}/{w_fact}, 0)', workbook.add_format({'num_format': '0.00', 'bg_color': row_bg, 'border': 1, 'border_color': border_col, 'valign': 'vcenter', 'align': 'center'}))
            vel_l = col_letter(c_idx)
            c_idx += 1
            
            # WOS - use numeric stock value for calculation
            stock_val = row.get(f"{loc_name}_Stock", 0)
            inc_val = row.get(f"{loc_name}_Inc_Num", 0)
            total_stock = stock_val + inc_val
            worksheet.write_formula(xls_r, c_idx, f'=IF({vel_l}{xls_r+1}>0, {total_stock}/{vel_l}{xls_r+1}, IF({total_stock}>0, 999, 0))', workbook.add_format({'num_format': '0.00', 'bg_color': row_bg, 'border': 1, 'border_color': border_col, 'valign': 'vcenter', 'align': 'center'}))
            c_idx += 1

            
            loc_col_info[loc_name] = {'soq_c': soq_c, 'vel_l': vel_l, 'stock_val': stock_val, 'inc_val': inc_val}

        # Backfill SOQ formulas
        sku_l = col_letter(key_group_start + 1)
        cs_l = col_letter(key_group_start + cols_static.index('Case_Size') + 1)
        target_w = "'Control Panel'!G4"


        for loc_name, info in loc_col_info.items():
            stock_val = info['stock_val']
            inc_val = info['inc_val']
            f = (f'=IF(AND(LEFT({sku_l}{xls_r+1},4)="CNB-",{info["vel_l"]}{xls_r+1}>0, '
                 f'({info["vel_l"]}{xls_r+1}*{target_w} - {stock_val} - {inc_val})>0), '
                 f'CEILING(MAX(({info["vel_l"]}{xls_r+1}*{target_w} - {stock_val} - {inc_val}), 0) / MAX({cs_l}{xls_r+1}, 1), 1), "-")')
            worksheet.write_formula(xls_r, info['soq_c'], f, fmt_soq)

    # Hide hidden columns
    for i in range(3):
        worksheet.set_column(key_group_end + len(metrics)*3 + i + 1, key_group_end + len(metrics)*3 + i + 1, None, None, {'hidden': True})

    # Conditional Formatting
    for _, loc_start, _ in loc_groups:
        st_c = loc_start
        # Status formatting (Landing removed)
        for val, fmt in [('🔥', fmt_status_hot), ('🚨', fmt_status_reorder),
                         ('✅', fmt_status_good), ('📦', fmt_status_filler), ('💀', fmt_status_dead), 
                         ('✨', fmt_status_new), ('❄️', fmt_status_cold)]:
            worksheet.conditional_format(1, st_c, data_end_row, st_c, {'type': 'text', 'criteria': 'containing', 'value': val, 'format': fmt})
        
        # Stock (col 1) - highlight when has incoming
        worksheet.conditional_format(1, st_c+1, data_end_row, st_c+1, {'type': 'text', 'criteria': 'containing', 'value': '🚚', 'format': workbook.add_format({'bg_color': '#DBEAFE', 'border': 1, 'border_color': border_col})})
        
        # WOS (col 8) - highlight low/high
        worksheet.conditional_format(1, st_c+8, data_end_row, st_c+8, {'type': 'cell', 'criteria': '<', 'value': 2.5, 'format': workbook.add_format({'bg_color': '#FEE2E2', 'border': 1, 'border_color': border_col})})
        worksheet.conditional_format(1, st_c+8, data_end_row, st_c+8, {'type': 'cell', 'criteria': '>', 'value': 26, 'format': workbook.add_format({'bg_color': '#FEF3C7', 'border': 1, 'border_color': border_col})})
        
        # Margin (col 6) - highlight good/bad margins
        m_l = col_letter(st_c+6)
        worksheet.conditional_format(1, st_c+6, data_end_row, st_c+6, {'type': 'formula', 'criteria': f'AND({m_l}2>=0.25,{m_l}2<=0.45)', 'format': workbook.add_format({'bg_color': '#DCFCE7', 'border': 1, 'border_color': border_col})})
        worksheet.conditional_format(1, st_c+6, data_end_row, st_c+6, {'type': 'cell', 'criteria': '<', 'value': 0.25, 'format': workbook.add_format({'bg_color': '#FEE2E2', 'border': 1, 'border_color': border_col})})
        worksheet.conditional_format(1, st_c+6, data_end_row, st_c+6, {'type': 'cell', 'criteria': '>', 'value': 0.45, 'format': workbook.add_format({'bg_color': '#FEF3C7', 'border': 1, 'border_color': border_col})})


    # Set column widths for Order Builder
    worksheet.set_column(0, 0, 8)  # Key
    worksheet.set_column(1, 1, 15) # SKU
    worksheet.set_column(2, 2, 35) # Product Name
    worksheet.set_column(3, 4, 15) # Cat/Brand
    worksheet.set_column(5, 8, 12) # Pricing/Stock metrics
    
    # Grid Polish
    worksheet.hide_gridlines(2)
    
    # Grouping and Freeze

    for _, l_s, l_e in loc_groups:
        worksheet.set_column(l_s, l_s, None, None, {'level': 1})
        if l_e > l_s: worksheet.set_column(l_s+1, l_e, None, None, {'level': 2})
    worksheet.freeze_panes(1, key_group_end + 1)
    
    writer.close()
    log_func("✅ Success! Report Generated.")
    
    try:
        os.startfile(output_filename)
    except:
        subprocess.call(['open', output_filename])
    
    return output_filename
