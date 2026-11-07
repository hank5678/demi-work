# 訂單拆分與可售庫存工具

直接以瀏覽器開啟 `index.html`，或將專案作為靜態網站提供。請保留 `app.js`、`workbook-tools.js` 與 `vendor/` 的相對位置；所有檔案都在瀏覽器本機處理，不需網路或後端。

## 使用

- 輸入僅接受 `.xlsx`，不接受 CSV 或舊版 `.xls`。每份檔案可選擇要使用的工作表，第一列必須為欄位標題。
- 訂單拆分需要「訂單號碼」、「選項」。任一商品含「預購」或「预购」，整張訂單歸入預購。輸出 `預購單.xlsx` 和 `現貨單.xlsx`。
- 可售庫存需要訂單表的「商品貨號」、「數量」，以及庫存表的「卖家SKU」、「可用库存」。各自依貨號加總後相減，輸出 `可售庫存.xlsx`。
- 拆單保留欄位、儲存格值、數值格式及資料順序。公式輸出為 Excel 已儲存的結果；不重新計算公式，也不複製其他工作表、圖表或完整工作表版面。公式缺少結果時，請先在 Excel 重新計算並儲存。
- 貨號以文字或儲存格顯示格式比對，保留前導零，去除前後空白。Excel 已經捨去的前導零或長數字精度無法復原。
- 庫存結果一般以數值輸出，超過 15 位數的整數以文字保留精度。

## 驗證

執行 `node --test tests/workbook-tools.test.cjs`。測試包括實際 XLSX 寫入與重新讀取、整單分類、庫存加總、前導零、日期格式、公式結果、多工作表及輸入錯誤。

## 第三方元件

使用 [SheetJS Community Edition 0.20.3](https://docs.sheetjs.com/docs/getting-started/installation/standalone/)，已置於 `vendor/xlsx.full.min.js` 供離線使用。Apache 2.0 授權見 `vendor/SHEETJS-LICENSE.txt`。
