import streamlit as st
import openpyxl
from openpyxl.styles import Alignment, Border, Side
import io

# 設定格線樣式
thin_border = Border(
    left=Side(style='thin'), 
    right=Side(style='thin'), 
    top=Side(style='thin'), 
    bottom=Side(style='thin')
)

def process_invoice(file_bytes):
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes))
    ws = wb.active # 預設處理第一個分頁

    # 1. 判斷公司並填寫資訊 (A2-A4)
    company_val = str(ws['A2'].value)
    
    if 'EVERLIFE-AL' in company_val:
        ws['A2'] = '歐瑞生醫科技有限公司 Allre Biological Technology Co., Ltd.'
        ws['A3'] = 'TEL:(02)29531399'
        ws['A4'] = 'Address:新北市板橋區中山路一段69號十樓'
    elif 'EVERLIFE-MK' in company_val:
        ws['A2'] = '蜜凱生技有限公司 MK BIOTECHNOLOGY Co., Ltd'
        ws['A3'] = 'TEL:(02)29531399'
        ws['A4'] = 'Address:236新北市土城區永豐路96巷8號'

    # 2. 自動調整欄寬 (遍歷 A 到 I 欄)
    for col in ws.columns:
        max_length = 0
        column = col[0].column_letter
        for cell in col:
            try:
                if cell.value and len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except: pass
        ws.column_dimensions[column].width = max_length + 2

    # 3. 畫格線與置中 (針對 13 列以後的資料區)
    # 假設資料到 I 欄，我們找最後一列
    last_row = ws.max_row
    for row in ws.iter_rows(min_row=13, max_row=last_row, min_col=1, max_col=9):
        for cell in row:
            cell.border = thin_border
            cell.alignment = Alignment(horizontal='center', vertical='center', wrapText=True)

    # 4. 在最後一筆資料下面第二格加入 Terms
    terms_row = last_row + 2
    ws.cell(row=terms_row, column=1, value="Terms：FOB")
    ws.cell(row=terms_row, column=1).alignment = Alignment(horizontal='left')

    # 儲存結果
    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()

# Streamlit 介面
st.title("🚢 報單格式自動化優化工具")
st.write("上傳 Excel 後，我會幫你改地址、調欄寬、畫格線並加 Terms！")

uploaded_file = st.file_uploader("請上傳原始報單 Excel", type=["xlsx"])

if uploaded_file:
    processed_data = process_invoice(uploaded_file.read())
    st.success("✅ 處理完成！")
    st.download_button(
        label="📥 下載優化後的報單",
        data=processed_data,
        file_name=f"優化後_{uploaded_file.name}",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
