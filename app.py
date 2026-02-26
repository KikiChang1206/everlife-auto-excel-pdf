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

    # --- 2. 安全地合併 E7~I7 ---
    # 為了避免 Excel 報錯，我們先清除 F7:I7 的內容，並確保沒有舊的合併衝突
    try:
        # 如果原本有合併，先解除 (避免重複合併衝突)
        for merged_range in list(ws.merged_cells.ranges):
            if 'E7' in merged_range or 'F7' in merged_range:
                ws.unmerge_cells(str(merged_range))
        
        # 清除 F7 到 I7 的隱藏資料，確保只留 E7
        for col_idx in range(6, 10): # F 到 I
            ws.cell(row=7, column=col_idx).value = None
            
        # 執行合併
        ws.merge_cells('E7:I7')
        # 設定格式：靠左、垂直置中、自動換行
        ws['E7'].alignment = Alignment(horizontal='left', vertical='center', wrapText=True)
    except Exception as e:
        st.warning(f"合併 E7:I7 時發生小提示：{e}")

    # --- 3. 判定格線結束位置 ---
    grid_end_row = 12
    for r in range(ws.max_row, 11, -1):
        if ws.cell(row=r, column=9).value is not None:
            grid_end_row = r
            break

    # --- 4. 畫格線與對齊設定 ---
    for row in ws.iter_rows(min_row=12, max_row=grid_end_row, min_col=1, max_col=9):
        for cell in row:
            cell.border = thin_border
            
            # --- 標題列 (第 12 列) 強制置中 ---
            if cell.row == 12:
                cell.alignment = Alignment(horizontal='center', vertical='center', wrapText=True)
            else:
                # 預設內容垂直置中
                cell.alignment = Alignment(horizontal='center', vertical='center', wrapText=True)
                
                # B, C 欄 (Description) 靠左
                if cell.column_letter in ['B', 'C']:
                    cell.alignment = Alignment(horizontal='left', vertical='center', wrapText=True)
                # H, I 欄 (金額) 靠右
                if cell.column_letter in ['H', 'I']:
                    cell.alignment = Alignment(horizontal='right', vertical='center')

    # --- 5. 套用指定欄寬 ---
    col_widths = {
        'A': 11.91, 'B': 23.73, 'C': 23.73, 'D': 11.36,
        'E': 5.64, 'F': 5.36, 'G': 7.36, 'H': 9.18, 'I': 11.09
    }
    for col_letter, width in col_widths.items():
        ws.column_dimensions[col_letter].width = width

    # 輸出檔案
    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()

# Streamlit 介面
st.title("🚢 報單格式精確優化")
st.write("已加入『安全合併』機制，解決開啟檔案時的修正提示。")

uploaded_file = st.file_uploader("請上傳原始報單 Excel", type=["xlsx"])

if uploaded_file:
    try:
        processed_data = process_invoice(uploaded_file.read())
        st.success("✅ 處理完成！")
        st.download_button(
            label="📥 下載最終報單",
            data=processed_data,
            file_name=f"Fixed_Final_{uploaded_file.name}",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        st.error(f"錯誤：{e}")
