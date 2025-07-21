# Derating-Table

# 🧮 Resistor Derating 自動化工具

這是一套用來自動化處理電阻（Resistor）料號 Derating 表的工具，分為兩個步驟：

1. 從 Agile BOM 報表中篩選出 B64 開頭的料號，並寫入 Derating 表格（Resistor sheet）。
2. 對 Excel 輸出的資料進行 A 欄儲存格合併，相同料號會合併成單一儲存格。

---

## 📂 檔案結構

📁 your-folder/
├─ Resistor step1.py # Step 1：篩選 B64 → 寫入 Derating + 輸出
├─ Resistor step2.py # Step 2：將 A欄相同值合併儲存格
├─ ABR0003427_Agile PLM BOM Report.xlsx # Agile PLM 報表來源
├─ C419E HPM Derating table.xlsx # 目標 Derating 表格（有 Resistor sheet）
├─ step2_Resistors.xlsx # Step1 的輸出（原始）
├─ step2_Resistors_merged.xlsx # Step2 的輸出（已合併儲存格）


---

## 🚀 使用方法

### ✅ 安裝必要套件

```bash
pip install pandas openpyxl

✅ 第一步：篩選 B64 電阻並寫入 Derating 表

python "Resistor step1.py"

功能：

自動尋找欄位名稱（支援有 metadata 的 BOM）

篩選出 Child Part 為 B64 開頭的零件

擷取以下欄位寫入 C419E HPM Derating table.xlsx 中的 Resistor 工作表

Child Part

Part Description

Manufacturer Part Number

Manufacturer

Status

輸出篩選結果為：step2_Resistors.xlsx


✅ 第二步：合併 A 欄相同料號儲存格

python "Resistor step2.py"
功能：

對 step2_Resistors.xlsx 執行 Excel 合併格式處理

將 A 欄相同值的連續列合併儲存格

輸出成：step2_Resistors_merged.xlsx

📌 注意事項
所有 Excel 檔案請與 .py 檔放在同一資料夾中

執行前請確保未打開 Excel 檔案（避免 Permission denied）

若欄位名稱或 sheet 名稱不同，請在程式內部手動修改

🛠️ 延伸功能（可選）
打包成 .exe 可執行工具

建立 GUI 圖形介面版

加入自動備份 Derating 表功能

延伸應用於 B78（電容）、B20（connector）等其他料號類別。


