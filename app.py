import streamlit as st
import openpyxl
from openpyxl.styles import Alignment, Border, Side
from openpyxl.utils import get_column_letter
import io

# 定義格線樣式
thin_side = Side(style='thin')
thin_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

def process_invoice(file_bytes):
    # 讀取 Excel，data_only=True 確保讀到公式結果
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

    # --- 2. 畫格線與格式化品項區域 ---
    # 自動尋找資料結束列（以 A 欄編號或 I 欄金額判斷）
    last_row = 12
    for r in range(12, ws.max_row + 1):
        if ws.cell(row=r, column=1).value or ws.cell(row=r, column=9).value:
            last_row = r

    # 針對 A12 到 I(資料最後一列) 畫格線
    for row in ws.iter_rows(min_row=12, max_row=last_row, min_col=1, max_col=9):
        for cell in row:
            cell.border = thin_border
            # 預設置中
            cell.alignment = Alignment(horizontal='center', vertical='center', wrapText=True)
            
            # 依照你的樣式需求：Description (B, C 欄) 靠左
            if cell.column_letter in ['B', 'C']:
                cell.alignment = Alignment(horizontal='left', vertical='center', wrapText=True)
            # 金額相關 (H, I 欄) 靠右
            if cell.column_letter in ['H', 'I']:
                cell.alignment = Alignment(horizontal='right', vertical='center')

    # --- 3. 自動調整欄寬 (解決字被遮住的問題) ---
    for col in ws.columns:
        max_length = 0
        column_letter = get_column_letter(col[0].column)
        
        # 掃描前 20 列來決定寬度，避免備註影響全域
        for cell in col[:25]:
            try:
                if cell.value:
                    val_str = str(cell.value)
                    # 中文字計算長度為 2，英文為 1
                    length = sum(2 if ord(char) > 127 else 1 for char in val_str)
                    if length > max_length:
                        max_length = length
            except: pass
        
        # 針對 Description 欄位給予更多寬度補償
        if column_letter in ['B', 'C']:
            ws.column_dimensions[column_letter].width = min(max_length + 5, 45)
        else:
            ws.column_dimensions[column_letter].width = max_length + 3

    # 輸出檔案
    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()

# Streamlit 介面
st.title("🚢 報單格式自動優化工具")
st.write("已移除重複的 Terms 寫入邏輯，保留原始文件中的 FOB 欄位。")

uploaded_file = st.file_uploader("請上傳原始報單 Excel", type=["xlsx"])

if uploaded_file:
    try:
        processed_data = process_invoice(uploaded_file.read())
        st.success("✅ 處理完成！")
        st.download_button(
            label="📥 下載優化後的 Excel",
            data=processed_data,
            file_name=f"Processed_{uploaded_file.name}",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        st.error(f"錯誤：{e}")
