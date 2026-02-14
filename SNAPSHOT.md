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
