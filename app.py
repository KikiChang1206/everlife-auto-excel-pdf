import streamlit as st
import openpyxl
from openpyxl.styles import Alignment, Border, Side
from openpyxl.utils import get_column_letter
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

    # --- 2. 判定格線結束位置 ---
    # 從最後一列往回找，直到在 I 欄找到總金額（通常是數字）
    grid_end_row = 12
    for r in range(ws.max_row, 11, -1):
        if ws.cell(row=r, column=9).value is not None: # I 欄
            grid_end_row = r
            break

    # --- 3. 畫格線與格式化 (從第 12 列開始，到總金額列) ---
    for row in ws.iter_rows(min_row=12, max_row=grid_end_row, min_col=1, max_col=9):
        for cell in row:
            cell.border = thin_border
            # 預設置中
            cell.alignment = Alignment(horizontal='center', vertical='center', wrapText=True)
            
            # B, C 欄 (Description) 靠左
            if cell.column_letter in ['B', 'C']:
                cell.alignment = Alignment(horizontal='left', vertical='center', wrapText=True)
            # H, I 欄 (金額) 靠右
            if cell.column_letter in ['H', 'I']:
                cell.alignment = Alignment(horizontal='right', vertical='center')

    # --- 4. 自動調整欄寬 ---
    for col in ws.columns:
        max_length = 0
        column_letter = get_column_letter(col[0].column)
        
        # 掃描資料區來決定寬度
        for cell in col[:grid_end_row + 1]:
            try:
                if cell.value:
                    val_str = str(cell.value)
                    length = sum(2 if ord(char) > 127 else 1 for char in val_str)
                    if length > max_length:
                        max_length = length
            except: pass
        
        if column_letter in ['B', 'C']:
            ws.column_dimensions[column_letter].width = min(max_length + 5, 45)
        elif column_letter == 'A': # 品項編號欄位不需要太寬
            ws.column_dimensions[column_letter].width = max(max_length, 5)
        else:
            ws.column_dimensions[column_letter].width = max_length + 3

    # 輸出檔案
    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()

# Streamlit 介面
st.title("🚢 報單格式精確優化")
st.write("格線現在會精確結束在 I 欄總金額的位置。")

uploaded_file = st.file_uploader("請上傳原始報單 Excel", type=["xlsx"])

if uploaded_file:
    try:
        processed_data = process_invoice(uploaded_file.read())
        st.success("✅ 處理完成！")
        st.download_button(
            label="📥 下載最終報單",
            data=processed_data,
            file_name=f"Final_{uploaded_file.name}",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        st.error(f"錯誤：{e}")
