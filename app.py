import streamlit as st
import openpyxl
from openpyxl.styles import Alignment, Border, Side
from openpyxl.utils import get_column_letter  # 修正處 1：匯入工具
import io

# 設定格線樣式
thin_border = Border(
    left=Side(style='thin'), 
    right=Side(style='thin'), 
    top=Side(style='thin'), 
    bottom=Side(style='thin')
)

def process_invoice(file_bytes):
    # 使用 data_only=True 可以讀取公式產生的值
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
    ws = wb.active 

    # 1. 判斷公司並填寫資訊 (A2-A4)
    company_val = str(ws['A2'].value) if ws['A2'].value else ""
    
    if 'EVERLIFE-AL' in company_val:
        ws['A2'] = '歐瑞生醫科技有限公司 Allre Biological Technology Co., Ltd.'
        ws['A3'] = 'TEL:(02)29531399'
        ws['A4'] = 'Address:新北市板橋區中山路一段69號十樓'
    elif 'EVERLIFE-MK' in company_val:
        ws['A2'] = '蜜凱生技有限公司 MK BIOTECHNOLOGY Co., Ltd'
        ws['A3'] = 'TEL:(02)29531399'
        ws['A4'] = 'Address:236新北市土城區永豐路96巷8號'

    # 2. 自動調整欄寬
    for col in ws.columns:
        max_length = 0
        # 修正處 2：正確取得欄位字母
        column_letter = get_column_letter(col[0].column) 
        
        for cell in col:
            try:
                if cell.value:
                    val_str = str(cell.value)
                    # 簡單計算長度，中文字元長度約為 2
                    length = sum(2 if ord(char) > 127 else 1 for char in val_str)
                    if length > max_length:
                        max_length = length
            except: pass
        
        # 設定寬度，最小不低於 10，最大不超過 50
        ws.column_dimensions[column_letter].width = min(max(max_length + 2, 10), 50)

    # 3. 畫格線與置中 (針對 13 列以後的資料區)
    last_row = ws.max_row
    for row in ws.iter_rows(min_row=13, max_row=last_row, min_col=1, max_col=9):
        for cell in row:
            cell.border = thin_border
            cell.alignment = Alignment(horizontal='center', vertical='center', wrapText=True)

    # 4. 在最後一筆資料下面第二格加入 Terms
    terms_row = last_row + 2
    ws.cell(row=terms_row, column=1, value="Terms：FOB")
    ws.cell(row=terms_row, column=1).alignment = Alignment(horizontal='left')

    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()

# Streamlit 介面
st.title("🚢 報單格式自動化優化工具")
st.write("修正了欄位判斷錯誤，請重新上傳測試！")

uploaded_file = st.file_uploader("請上傳原始報單 Excel", type=["xlsx"])

if uploaded_file:
    try:
        processed_data = process_invoice(uploaded_file.read())
        st.success("✅ 處理完成！")
        st.download_button(
            label="📥 下載優化後的報單",
            data=processed_data,
            file_name=f"優化後_{uploaded_file.name}",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        st.error(f"發生錯誤：{e}")
