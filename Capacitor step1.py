import pandas as pd

try:
    print("🚀 開始讀取 BOM 檔案...")
    bom_report_path = "ABR0003427_Agile PLM BOM Report.xlsx"
    derating_table_path = "C419E HPM Derating table.xlsx"

    # 初步讀取原始資料
    df_raw = pd.read_excel(bom_report_path, sheet_name=0, skiprows=6)

    # 尋找欄位列
    header_row_index = None
    for i, row in df_raw.iterrows():
        if "Child Part" in row.values:
            header_row_index = i
            break

    if header_row_index is None:
        raise ValueError("❌ 找不到包含 'Child Part' 的欄位列")

    print(f"✅ 表頭在第 {header_row_index + 7} 列（含前面 skiprows）")

    # 再次讀取正確表頭的 DataFrame
    df_bom = pd.read_excel(bom_report_path, sheet_name=0, skiprows=header_row_index + 7)
    df_bom.columns = df_raw.iloc[header_row_index]
    df_bom = df_bom.reset_index(drop=True)
    df_bom = df_bom.loc[:, ~df_bom.columns.isna()]  # 移除空欄

    # 處理合併儲存格造成的空值
    df_bom['Child Part'] = df_bom['Child Part'].fillna(method='ffill')

    # 篩選 B78 零件
    filtered = df_bom[df_bom['Child Part'].astype(str).str.startswith("B78")]
    print(f"🔎 找到 {len(filtered)} 筆 B78 零件")

    # 擷取目標欄位（容錯）
    def safe_get(col):
        return filtered[col] if col in filtered.columns else ""

    result = pd.DataFrame({
        "Child Part": safe_get("Child Part"),
        "Part Description": safe_get("Part Description"),
        "Manufacturer Part Number": safe_get("Manufacturer Part Number"),
        "Manufacturer": safe_get("Manufacturer"),
        "Status": safe_get("Status")
    })

    # 寫入 Derating Table Capacitor Sheet
    existing_df = pd.read_excel(derating_table_path, sheet_name="Capacitor")
    startrow = len(existing_df) + 1

    with pd.ExcelWriter(derating_table_path, mode='a', engine='openpyxl', if_sheet_exists='overlay') as writer:
        result.to_excel(writer, sheet_name="Capacitor", index=False, header=False, startrow=startrow)

    print("✅ 寫入 Derating 表成功")

    # 額外備份輸出
    output_path = "step2_capacitors.xlsx"
    result.to_excel(output_path, index=False)
    print(f"✅ 已另存 B78 篩選資料為：{output_path}")

except Exception as e:
    print(f"❌ 發生錯誤：{e}")

input("📌 執行結束，請按 Enter 鍵關閉...")
