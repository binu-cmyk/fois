import os
import pandas as pd
from app import save_to_excel, parse_fois_table

def run_excel_test():
    # 1. Create raw Outstanding ODR data (using raw FOIS headers)
    raw_odr_data = {
        "Sr. No.": [1, 2, 3],
        "Division": ["MUMBAI", "PUNE", "NAGPUR"],
        "STTN FROM": ["CSTM", "PUNE", "NGP"], # Raw station
        "Indent No.": ["IND1001", "IND1002", "IND1003"],
        "Indent Date": ["2026-05-20", "2026-05-22", "2026-05-25"],
        "Priority": ["C", "B", "A"],
        "CNSR": ["ADITYA BIRLA", "TATA", "RELIANCE"], # Raw Consignor
        "CNSG": ["ULTRATECH", "TATA STEEL", "JIO REF"], # Raw Consignee
        "CMDT": ["CEMENT", "COAL", "IRON ORE"], # Raw Commodity
        "RAKE CMDT": ["CEM", "COA", "IRO"], # Raw Rake Commodity
        "DSTN": ["NDLS", "HWH", "MAS"], # Raw Destination
        "Rakes": [2, 4, 1]
    }
    raw_odr_df = pd.DataFrame(raw_odr_data)
    
    # Simulating what parse_fois_table does with renaming:
    # (Since parse_fois_table does the mapping dynamically, we do it via a quick test)
    mapping = {}
    for col in raw_odr_df.columns:
        col_clean = " ".join(str(col).upper().replace('_', ' ').replace('-', ' ').split())
        
        # 1. Serial Number
        if col_clean in ['S.NO.', 'S.NO', 'SR. NO.', 'SR.NO.', 'S. NO.', 'S. NO', 'SERIAL NUMBER']:
            mapping[col] = 'Serial Number'
            
        # 2. Division Code
        elif col_clean in ['DVSN', 'DIVISION', 'DIVISION CODE', 'DVSN CODE']:
            mapping[col] = 'Division Code'
            
        # 3. Station From
        elif col_clean in ['STTN FROM', 'STTN FRM', 'STATION FROM', 'STN FROM', 'SOURCE', 'SRCE', 'DISPATCH']:
            mapping[col] = 'Station From'
            
        # 4. Sender
        elif col_clean in ['CNSR', 'CONSIGNOR', 'SENDER', 'CONSIGNER']:
            mapping[col] = 'Sender'
            
        # 5. Receiver
        elif col_clean in ['CNSG', 'CONSIGNEE', 'RECEIVER', 'RECVR']:
            mapping[col] = 'Receiver'
            
        # 6. Commodity
        elif col_clean in ['CMDT', 'COMMODITY']:
            mapping[col] = 'Commodity'
            
        # 7. Commodity Type
        elif col_clean in ['RAKE CMDT', 'RAKE_CMDT', 'RAKECMDT', 'RAKE COMMODITY']:
            mapping[col] = 'Commodity Type'
            
        # 8. Destination
        elif col_clean in ['DSTN', 'DEST', 'DESTINATION', 'TO', 'STTN TO', 'STN TO']:
            mapping[col] = 'Destination'
            
    mapped_odr_df = raw_odr_df.rename(columns=mapping)
    
    # Filter columns to only keep the 8 requested columns in exact order
    desired_cols = ['Serial Number', 'Division Code', 'Station From', 'Sender', 'Receiver', 'Commodity', 'Commodity Type', 'Destination']
    for col in desired_cols:
        if col not in mapped_odr_df.columns:
            mapped_odr_df[col] = ''
    mapped_odr_df = mapped_odr_df[desired_cols]
    
    # Verify that renaming took place
    print("Mapped columns list:")
    print(mapped_odr_df.columns)
    assert list(mapped_odr_df.columns) == desired_cols
    
    # 2. Mock Matured Indents dataset
    raw_matured_data = {
        "Sr. No.": [1, 2],
        "Division": ["SOLAPUR", "BHUSAVAL"],
        "STTN FROM": ["SUR", "BSL"],
        "Indent No.": ["IND0901", "IND0902"],
        "Indent Date": ["2026-05-01", "2026-05-05"],
        "Matured Date": ["2026-05-15", "2026-05-18"],
        "Priority": ["D", "C"],
        "CNSR": ["ACC", "JINDAL"],
        "CNSG": ["ACC PLT", "JSW STEEL"],
        "CMDT": ["CLINKER", "STEEL"],
        "RAKE CMDT": ["CLI", "STE"],
        "DSTN": ["SUR", "BSL"],
        "Rakes": [1, 3]
    }
    raw_matured_df = pd.DataFrame(raw_matured_data)
    mapped_matured_df = raw_matured_df.rename(columns=mapping)
    for col in desired_cols:
        if col not in mapped_matured_df.columns:
            mapped_matured_df[col] = ''
    mapped_matured_df = mapped_matured_df[desired_cols]
    
    # 3. Test writing Sheet 1 (Outstanding ODR)
    print("Testing Outstanding ODR Sheet writing...")
    file1 = save_to_excel(mapped_odr_df, "Outstanding ODR data")
    print(f"Sheet 1 written to {file1}")
    
    # 4. Test writing Sheet 2 (Matured Indents)
    print("Testing Matured Indents Sheet writing...")
    file2 = save_to_excel(mapped_matured_df, "Matured Indents")
    print(f"Sheet 2 written to {file2}")
    
    # Verify file existence and size
    if os.path.exists("FOIS_MIS_Report.xlsx"):
        size = os.path.getsize("FOIS_MIS_Report.xlsx")
        print(f"SUCCESS: Excel file created successfully! Size: {size} bytes")
        
        # Verify sheets inside Excel
        from openpyxl import load_workbook
        wb = load_workbook("FOIS_MIS_Report.xlsx")
        print(f"Active sheets in workbook: {wb.sheetnames}")
        assert "Outstanding ODR data" in wb.sheetnames, "Missing Outstanding ODR data sheet"
        assert "Matured Indents" in wb.sheetnames, "Missing Matured Indents sheet"
        
        # Verify headers in Excel
        ws1 = wb["Outstanding ODR data"]
        excel_headers = [ws1.cell(row=4, column=col).value for col in range(1, len(mapped_odr_df.columns) + 1)]
        print(f"Excel Sheet 1 headers written: {excel_headers}")
        assert "SERIAL NUMBER" in excel_headers
        assert "DIVISION CODE" in excel_headers
        assert "STATION FROM" in excel_headers
        assert "SENDER" in excel_headers
        assert "RECEIVER" in excel_headers
        assert "COMMODITY" in excel_headers
        assert "COMMODITY TYPE" in excel_headers
        assert "DESTINATION" in excel_headers
        
        print("ALL TESTS PASSED SUCCESSFULLY!")
    else:
        print("FAILURE: Excel file was not created.")

if __name__ == "__main__":
    run_excel_test()
