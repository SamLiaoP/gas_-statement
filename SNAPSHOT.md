# 對帳系統 SNAPSHOT

## 2026-02-14 - 初始版本

### 變更內容
- 建立 `main.py`：加油站支付對帳系統
- 支援 LINE PAY 和中油PAY 兩個支付管道的自動對帳
- 自動偵測同目錄下的內帳檔案（NNN.MM.xlsx 格式）
- 產出對帳結果 Excel，含差異標記和紅色高亮

### 架構決策
- 使用 ChannelConfig dataclass + registry 字典模式，方便擴充新支付管道
- 單一 main.py 檔案，保持簡單
- LinePay 21:00 規則：交易時間 >= 21:00 歸入隔天

### 驗證結果
- 中油PAY Day 1 = 9783 兩邊一致 → V
- 中油PAY Day 7: 內帳 9234 vs 明細 9665 → X（差異 -431）
- LinePay 21:00 規則正確套用（12/31 21:13 的 500 元歸入 1/1）

## 2026-02-14 - LinePay 匯款明細整理程式

### 變更內容
- 新增 `LinePay匯款明細整理/main.py`：將 linepay明細.xlsx 按撥款預定日分組整理
- 每個撥款預定日底下按交易日分組，加總付款金額、手續費、實收
- 輸出格式化 Excel（合併儲存格、粗體底色、千分位、小計與總計）

### 驗證結果
- 總計：付款金額 80,562、手續費 1,377.13、實收 79,184.87（與原始資料一致）
- 撥款預定日 20260102：付款金額 2,672、手續費 45.68、實收 2,626.32（正確）
- 共 21 個撥款預定日、141 筆交易

## 2026-02-14 - 改為讀取 報表/ 底下月份資料夾

### 變更內容
- 修改 `電子支付對帳程式/main.py`：改為掃描 `報表/` 底下所有 YYYYMM 資料夾
  - 抽出 `process_folder(folder_path, folder_name)` 函式
  - `main()` 改為迴圈掃描所有月份資料夾
  - 輸出檔已存在時自動跳過（`對帳結果_YYYYMM.xlsx`）
  - 找不到內帳或明細時印警告並跳過
  - `find_internal_file()` 不再呼叫 `sys.exit`，改回傳 None
- 修改 `LinePay匯款明細整理/main.py`：同樣改為掃描 `報表/` 底下所有資料夾
  - 抽出 `process_folder(folder_path, folder_name)` 函式
  - 輸出檔已存在時自動跳過（`LinePay匯款明細整理_YYYYMM.xlsx`）
  - 找不到 linepay明細.xlsx 時跳過

### 架構決策
- 輸入/輸出統一放在 `報表/{YYYYMM}/` 資料夾內
- 資料夾名稱須為6位數字（YYYYMM），其他名稱會被忽略
- 已有輸出檔的月份自動跳過，避免重複處理

## 2026-02-21 - 資料夾與腳本路徑改為英文（解決 Windows 中文路徑問題）

### 變更內容
- 資料夾重新命名：
  - `電子支付對帳程式/` → `reconciliation/`
  - `LinePay匯款明細整理/` → `linepay_summary/`
  - `報表/` → `reports/`
- 執行腳本重新命名：
  - `執行對帳.bat/.command` → `run_reconciliation.bat/.command`
  - `執行匯款明細整理.bat/.command` → `run_linepay_summary.bat/.command`
- Excel 檔名維持中文（輸入：`linepay明細.xlsx`、`中油pay明細.xls`；輸出：`對帳結果_YYYYMM.xlsx`、`LinePay匯款明細整理_YYYYMM.xlsx`）

### 原因
- Windows 上中文路徑可能造成編碼錯誤，資料夾與腳本改為英文確保跨平台相容
- Excel 檔名維持中文，方便使用者辨識
