import pandas as pd
import io
from openpyxl import load_workbook
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils import get_column_letter

def process_call_plan_logic(df, job_number):
    """
    Logic: แยกข้อมูลตามรหัสพนักงาน และสร้าง Excel ที่มี Table Style
    """
    output_bundles = {}

    # ลบแถวที่เป็น header ซ้ำ
    col_filter_idx = 9
    col_filter = df.columns[col_filter_idx]
    df = df[~df[col_filter].astype(str).str.contains("รหัสพนักงาน|^\\s*$", na=False)]

    # กำหนดคอลัมน์หลัก
    columns = df.columns.tolist()
    col_box = columns[4]
    col_store = columns[5]
    col_emp_id = columns[6]
    date_columns = columns[12:43]  # คอลัมน์วันที่ 1-31

    # กรองรหัสพนักงานที่ไม่ซ้ำ
    emp_ids = df[col_filter].dropna().astype(str).str.strip()
    emp_ids = emp_ids[emp_ids != ""].unique()

    for emp_id in emp_ids:
        result_rows = []
        emp_data = df[df[col_filter].astype(str).str.strip() == emp_id]

        for i, col_date in enumerate(date_columns):
            for _, row in emp_data.iterrows():
                if row[col_date] == 1:
                    day = i + 1
                    result_rows.append([
                        row[col_box],
                        row[col_store],
                        row[col_emp_id],
                        day
                    ])

        if result_rows:
            temp_df = pd.DataFrame(result_rows, columns=["กล่อง", "รหัสร้านค้า", "รหัสพนักงาน", "วันที่"])
            temp_df.sort_values(by="วันที่", inplace=True)
            temp_df["ลำดับ"] = temp_df.groupby("วันที่").cumcount() + 1
            output_df = temp_df[["ลำดับ", "กล่อง", "รหัสร้านค้า", "รหัสพนักงาน", "วันที่"]]

            # สร้างไฟล์ Excel ใน Memory (BytesIO)
            excel_out = io.BytesIO()
            with pd.ExcelWriter(excel_out, engine='openpyxl') as writer:
                output_df.to_excel(writer, index=False, sheet_name='Data')

            # โหลดกลับมาเพื่อใส่ Table Style
            excel_out.seek(0)
            wb = load_workbook(excel_out)
            ws = wb['Data']
            max_row = ws.max_row
            max_col = ws.max_column
            col_letter = get_column_letter(max_col)
            
            tab = Table(displayName=f"Table_{emp_id.replace('-', '_')}", ref=f"A1:{col_letter}{max_row}")
            style = TableStyleInfo(name="TableStyleMedium9", showFirstColumn=False, showLastColumn=False,
                                   showRowStripes=True, showColumnStripes=False)
            tab.tableStyleInfo = style
            ws.add_table(tab)
            
            # บันทึกผลลัพธ์ลงใน BytesIO อีกครั้ง
            final_out = io.BytesIO()
            wb.save(final_out)
            
            filename = f"Call Plan_{job_number}_{emp_id}.xlsx"
            output_bundles[filename] = final_out.getvalue()

    return True, output_bundles