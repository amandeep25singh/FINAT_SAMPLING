import pandas as pd
import math
import os
import shutil

def clean_string(val):
    if pd.isna(val):
        return ""
    return str(val).replace('\xa0', ' ').strip().lower()

def get_bracket_and_percentage(voutype, dh_desc, amount):
    v_type = clean_string(voutype)
    desc = clean_string(dh_desc)
    
    if "ac bill" in v_type or "dc bill" in v_type: return "AC_DC_All", 1.00 
    elif "grant" in v_type or "grant" in desc: return "GIA_All", 1.00 
    elif "medical" in desc or "medical" in v_type:
        if amount >= 10000: return "Med_>=10k", 0.25
    elif "ltc" in desc or "leave travel" in desc or "ltc" in v_type:
        if amount >= 25000: return "LTC_>=25k", 0.25
    elif "ta bill" in v_type or "ta voucher" in v_type or "travel" in desc:
        if amount >= 15000: return "TA_>=15k", 0.25
    elif "telephone" in desc or "mobile" in desc:
        if amount >= 5000: return "Tel_>=5k", 0.50
    elif "establishment" in v_type or "salary" in desc:
        if amount < 500000: return "Est_<5L", 0.01
        elif amount <= 1000000: return "Est_5-10L", 0.01
        elif amount <= 1500000: return "Est_10-15L", 0.01
        else: return "Est_>15L", 0.02
    elif "contingent" in v_type or "non salary" in v_type or "works" in desc:
        prefix = "Works" if "works" in desc else "Cont"
        if amount < 100000: return f"{prefix}_<1L", 0.05
        elif amount < 300000: return f"{prefix}_1-3L", 0.10
        elif amount < 500000: return f"{prefix}_3-5L", 0.25
        else: return f"{prefix}_>=5L", 1.00
            
    return "None", 0.0

BRACKET_DEFS = {
    "Est_<5L": ("Establishment Vouchers", "Below Rs 5 lakh", 0.01),
    "Est_5-10L": ("", "Rs 5 lakh to 10 lakh", 0.01),
    "Est_10-15L": ("", "Rs 10 lakh to 15 lakh", 0.01),
    "Est_>15L": ("", "Above Rs 15 lakh", 0.02),
    "Med_>=10k": ("Medical Bills", "Rs 10,000 and above", 0.25),
    "LTC_>=25k": ("Leave Travel Concession", "Rs 25,000 and above", 0.25),
    "TA_>=15k": ("Travel Expenses / TA", "Rs 15,000 and above", 0.25),
    "Cont_<1L": ("Contingent Bills", "Below Rs 1 lakh", 0.05),
    "Cont_1-3L": ("", "Rs 1 lakh to 3 lakh", 0.10),
    "Cont_3-5L": ("", "Rs 3 lakh to 5 lakh", 0.25),
    "Cont_>=5L": ("", "Rs 5 lakh and above", 1.00),
    "Works_<1L": ("Works", "Below Rs 1 lakh", 0.05),
    "Works_1-3L": ("", "Rs 1 lakh to 3 lakh", 0.10),
    "Works_3-5L": ("", "Rs 3 lakh to 5 lakh", 0.25),
    "Works_>=5L": ("", "Rs 5 lakh and above", 1.00),
    "Tel_>=5k": ("Telephone / Mobile", "Rs 5,000 and above", 0.50),
    "AC_DC_All": ("AC / DC Bills", "All Amounts", 1.00),
    "GIA_All": ("Grant In Aid", "All Amounts", 1.00),
}

def process_audit_data():
    file_path = input("Enter the full path to the Master Excel or CSV file: ").strip()
    
    if file_path.startswith(('"', "'")) and file_path.endswith(('"', "'")):
        file_path = file_path[1:-1]
        
    if not os.path.exists(file_path):
        print(f"Error: File not found at {file_path}")
        return

    print("Loading data...")
    try:
        try:
            df_master = pd.read_csv(file_path)
            is_csv = True
            print("Successfully loaded as CSV.")
        except Exception:
            xls = pd.ExcelFile(file_path)
            master_sheet = [s for s in xls.sheet_names if 'MASTER' in s.upper()]
            master_sheet = master_sheet[0] if master_sheet else xls.sheet_names[0]
            
            df_raw = pd.read_excel(file_path, sheet_name=master_sheet, header=None)
            header_row_index = 0
            for i in range(min(150, len(df_raw))):
                row_str = ' '.join(df_raw.iloc[i].astype(str).str.lower().fillna(''))
                if 'amount' in row_str and ('desc' in row_str or 'dh' in row_str or 'dept' in row_str):
                    header_row_index = i
                    break
            df_master = pd.read_excel(file_path, sheet_name=master_sheet, header=header_row_index)
            is_csv = False
            print(f"Successfully loaded Excel sheet: '{master_sheet}'.")
            
    except Exception as e:
        print(f"Failed to read file. Error: {e}")
        return

    df_master.columns = [str(c).replace('\xa0', ' ').strip().upper() for c in df_master.columns]

    dept_col = next((col for col in df_master.columns if 'DEPT' in col), None)
    amount_col = next((col for col in df_master.columns if 'AMOUNT' in col), None)
    voutype_col = next((col for col in df_master.columns if 'VOUTYPE' in col), None)
    dh_desc_col = next((col for col in df_master.columns if 'DH DESC' in col), None)

    if not all([dept_col, amount_col, voutype_col, dh_desc_col]):
        print("\nError: Could not find required columns.")
        return

    df_master[amount_col] = pd.to_numeric(df_master[amount_col], errors='coerce').fillna(0)
    df_master = df_master.sort_values(by=amount_col, ascending=False)

    all_selected_vouchers = []
    dept_overall_dict = {}
    dept_sampled_dict = {}
    summary_counts = {bracket: {} for bracket in BRACKET_DEFS.keys()}

    print("Applying sampling criteria...")
    departments = sorted(df_master[dept_col].dropna().unique()) 

    for dept in departments:
        df_dept = df_master[df_master[dept_col] == dept].copy()
        
        df_dept[['Bracket', 'SamplePct']] = df_dept.apply(
            lambda row: pd.Series(get_bracket_and_percentage(row[voutype_col], row[dh_desc_col], row[amount_col])), axis=1
        )
        
        df_dept['Selected_For_Audit'] = 'No'
        selected_for_dept = []
        
        grouped = df_dept[df_dept['Bracket'] != "None"].groupby('Bracket', sort=False)
        
        for bracket, group in grouped:
            total_count = len(group)
            pct = group['SamplePct'].iloc[0]
            sample_size = math.ceil(total_count * pct)
            
            if bracket in summary_counts:
                summary_counts[bracket][dept] = {"total": total_count, "selected": sample_size}
            
            if sample_size > 0:
                sampled_indices = group.head(sample_size).index
                df_dept.loc[sampled_indices, 'Selected_For_Audit'] = 'Yes'
                selected_for_dept.append(df_dept.loc[sampled_indices])
                
        df_dept_clean = df_dept.drop(columns=['Bracket', 'SamplePct'])
        dept_overall_dict[dept] = df_dept_clean
        
        if selected_for_dept:
            df_sampled_combined = pd.concat(selected_for_dept).drop(columns=['Bracket', 'SamplePct', 'Selected_For_Audit'])
            df_sampled_combined = df_sampled_combined.sort_values(by=amount_col, ascending=False)
            dept_sampled_dict[dept] = df_sampled_combined
            all_selected_vouchers.append(df_sampled_combined)

    # --- THE FIX: Building a flat dataframe to bypass the MultiIndex error ---
    row1_header = ["Voutype / Dh Desc", "Amount involved in the voucher", "Percentage of vouchers to be checked in case of "]
    row2_header = ["", "", "Departments covered in local audit plan"]

    for dept in departments:
        row1_header.extend([dept, ""])
        row2_header.extend(["Number of Vouchers", "Selected vouchers"])

    matrix_data = [row1_header, row2_header]
    
    for bracket, (cat_name, amount_desc, pct_val) in BRACKET_DEFS.items():
        row = [cat_name, amount_desc, pct_val]
        for dept in departments:
            stats = summary_counts[bracket].get(dept, {"total": 0, "selected": 0})
            row.extend([stats["total"], stats["selected"]])
        matrix_data.append(row)
        
    df_summary_matrix = pd.DataFrame(matrix_data)

    if all_selected_vouchers:
        df_all_selected = pd.concat(all_selected_vouchers).sort_values(by=amount_col, ascending=False)
    else:
        df_all_selected = pd.DataFrame(columns=df_master.columns)

    import datetime
    timestamp = datetime.datetime.now().strftime("%H%M%S")
    output_path = os.path.join(os.path.dirname(file_path), f"Sampled_Output_Audit_{timestamp}.xlsx")
    
    if not is_csv:
        print(f"Creating output file and preserving original formatting...")
        shutil.copyfile(file_path, output_path)
        write_mode = 'a'
        if_exists = 'replace'
    else:
        print(f"Original was a CSV. Creating fresh Excel output...")
        write_mode = 'w'
        if_exists = None

    print(f"Writing calculated sheets to: {output_path}")
    
    with pd.ExcelWriter(output_path, engine='openpyxl', mode=write_mode, if_sheet_exists=if_exists) as writer:
        
        if write_mode == 'w':
            df_master.to_excel(writer, sheet_name='MASTER', index=False)

        # Write the matrix with header=False to strictly use our custom flat rows
        if not df_summary_matrix.empty:
            df_summary_matrix.to_excel(writer, sheet_name='Sampling Criterial', header=False, index=False)

        df_all_selected.to_excel(writer, sheet_name='Selected Vouchers all deptts', index=False)
        
        for dept_name in departments:
            safe_name = str(dept_name).replace('/', '_').replace('\\', '_')[:31] 
            
            if dept_name in dept_overall_dict:
                dept_overall_dict[dept_name].to_excel(writer, sheet_name=safe_name, index=False)
            
            if dept_name in dept_sampled_dict:
                sampled_sheet_name = f"Sampled_{safe_name}"[:31]
                dept_sampled_dict[dept_name].to_excel(writer, sheet_name=sampled_sheet_name, index=False)

    print("Success! Audit sampling file created.")

if __name__ == "__main__":
    process_audit_data()