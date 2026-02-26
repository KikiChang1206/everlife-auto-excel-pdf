import streamlit as st
import openpyxl
from openpyxl.styles import Alignment, Border, Side
import io

# 定義格線樣式
thin_side = Side(style='thin')
thin_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

def process_invoice(file_bytes):
    # 讀取 Excel
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
    ws = wb.active 

    # --- 1. 公司資訊更換 (A2-A4) ---
    company_val = str(ws['A2'].value) if ws['A2'].value else ""
    if 'EVERLIFE-AL' in company_val:
        ws['A2'] = '歐瑞生醫科技有限公司 Allre Biological Technology Co., Ltd.'
        ws['A3'] = 'TEL : (02)29531399'
        ws['A4'] = 'Adress : 新北市板橋區中山路一段69號十樓'
    elif 'EVERLIFE-MK' in company_val:
        ws['A2'] = '蜜凱生技有限公司 MK BIOTECHNOLOGY Co., Ltd'
        ws['A3'] = 'TEL : (02)29531399'
        ws['A4'] = 'Adress : 236新北市土城區永豐路96巷8號'

    # --- 2. 判定格線結束位置 (尋找 I 欄最後一個有值的列) ---
    grid_end_row = 12
    for r in range(ws.max_row, 11, -1):
        if ws.cell(row=r, column=9).value is not None:
            grid_end_row = r
            break

    # --- 3. 畫格線與對齊設定 (從第 12 列到總金額列) ---
    for row in ws.iter_rows(min_row=12, max_row=grid_end_row, min_col=1, max_col=9):
        for cell in row:
            cell.border = thin_border
            # 預設置中
            cell.alignment = Alignment(horizontal='center', vertical='center', wrapText=True)
            
            # 特殊對齊：Description (B, C 欄) 靠左
            if cell.column_letter in ['B', 'C']:
                cell.alignment = Alignment(horizontal='left', vertical='center', wrapText=True)
            # 特殊對齊：金額相關 (H, I 欄) 靠右
            if cell.column_letter in ['H', 'I']:
                cell.alignment = Alignment(horizontal='right', vertical='center')

    # --- 4. 套用指定欄寬 (依照您的數值) ---
    col_widths = {
        'A': 11.91,
        'B': 23.73,
        'C': 23.73,
        'D': 11.36,
        'E': 5.64,
        'F': 5.36,
        'G': 7.36,
        'H': 9.18,
        'I': 11.09
    }
    
    for col_letter, width in col_widths.items():
        ws.column_dimensions[col_letter].width = width

    # 輸出檔案
    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()

# Streamlit 介面
st.title("🚢 報單格式精確優化 (固定欄寬版)")
st.write("已將 A-I 欄寬設定為您指定的精確數值，且格線僅畫至總金額。")

uploaded_file = st.file_uploader("請上傳原始報單 Excel", type=["xlsx"])

if uploaded_file:
    try:
        processed_data = process_invoice(uploaded_file.read())
        st.success("✅ 格式優化完成！")
        st.download_button(
            label="📥 下載最終報單",
            data=processed_data,
            file_name=f"Fixed_Width_{uploaded_file.name}",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        st.error(f"錯誤：{e}")
