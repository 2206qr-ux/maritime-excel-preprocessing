# 航運行政人員 MS Excel 資料前處理實務教學與實戰案例庫
Maritime Administration MS Excel Data Preprocessing Practical Guide & 300-Case Challenge

本專案為專為**航運行政人員初學者**量身打造的 MS Excel 基礎資料分析前處理教學套件，內容涵蓋軟體操作介面認知、資料前處理基本常識、30 筆示範案例逐步教學、核心功能指令、高頻函數公式，以及 300 筆多情境航運實戰案例題庫與完整參考解題過程。

---

## 專案核心成果清單

| 檔案名稱 | 格式 | 說明 |
| :--- | :---: | :--- |
| **`航運行政人員Excel資料前處理實務教學手冊.docx`** | Word (DOCX) | 37 頁完整實務操作手冊，含封面、目錄、圖表目錄、頁碼、15 張高解析註解截圖與章節分頁。 |
| **`航運行政Excel資料前處理實務教學簡報.pptx`** | PowerPoint (PPTX) | 16:9 寬螢幕專業簡報（共 18 頁），深海藍航運商務風格，內嵌完整圖文卡片、公式呼叫與演講備忘稿。 |
| **`航運行政Excel資料前處理教學簡報.html`** | HTML5 / JS | 18 頁互動式投影片網頁，支援鍵盤導航、投影片縮圖抽屜、即時公式模擬器與儀表板篩選。 |
| **`01_航運事務資料前處理_30筆示範檔.xlsx`** | Excel (XLSX) | 30 筆海運出口併櫃裝箱單原始資料、清洗工作表與欄位對照規範。 |
| **`02_航運案例情境練習_300筆練習檔.xlsx`** | Excel (XLSX) | 涵蓋海運出口、海運進口、空運、內陸集疏運、倉儲五大情境 300 筆練習題組與題意說明。 |
| **`03_航運案例情境練習_300筆參考解答檔.xlsx`** | Excel (XLSX) | 完整清洗結果（288 筆唯一單證、264 筆放行合格、24 筆異常分類）與全動態 KPI 統計儀表板。 |
| **`build_pptx_deck.py`** | Python | 透過 `python-pptx` 自動建構高品質投影片之原始程式碼。 |

---

## 教材章節架構

- **第一章：MS Excel 操作介面與基礎導航**
  - 五大核心功能區塊全覽（快速存取工具列、功能區 Ribbon、名稱方塊與編輯列、工作表網格、狀態列）
  - 「常用」、「資料」、「公式」與「檢視（凍結窗格）」航運關鍵指令群
- **第二章：資料前處理應具備的基本常識**
  - 資料前處理核心概念：Garbage In, Garbage Out
  - 航運業務常見「資料五毒」（重複鍵入、隱形空白、日期顛倒、型態錯置、複合欄位）
  - 前處理標準作業流程五部曲（探索評估 → 結構清洗 → 格式轉換 → 邏輯驗證 → 成果輸出）
- **第三章：30筆航運事務示範資料集初探**
  - 海運出口併櫃（LCL Consolidation）業務情境與 11 欄位結構剖析
  - 原始資料人為缺陷與清洗目標規劃
- **第四章：MS Excel 核心前處理操作功能詳解**
  - 篩選與排序（自動篩選、多層級自訂排序）
  - 特殊目標定位（F5 定位空白儲存格與 Ctrl+Enter 批量補值）
  - 尋找與取代（標準代碼統一與萬用字元 `*`, `?` 實務）
  - 移除重複項（主鍵去重保證唯一性）
  - 資料剖析（依分隔符號拆分貨櫃號碼與封條號）
  - 資料驗證（防呆機制：下拉清單、長度限制、日期合理性防護）
- **第五章：航運高頻實用函數公式精講**
  - 文字清洗與轉換：`TRIM()`, `CLEAN()`, `UPPER()`, `SUBSTITUTE()`, `TEXT()`
  - 邏輯判斷與容錯：`IF()`, `AND()`, `OR()`, `ISBLANK()`, `IFERROR()`
  - 跨表關聯查詢：`VLOOKUP()`, `XLOOKUP()` 動態帶入航商全名與費率
  - 航期計算與品質檢核：`DATEDIF()`, `NETWORKDAYS()`, `COUNTIF()`, `ISNUMBER()`
- **第六章：300筆航運情境練習與解題成效評估**
  - 五大情境題庫架構設計與 24 筆業務異常缺陷佈署
  - 標準解題五階段 SOP 流程
  - 成效評估儀表板（288 筆單證、合格 264 筆、異常 24 筆、總運費 $693,376.68 USD、總重量 2,367,924 KGS）
- **第七章：航運資料治理守則與參考資源**
  - 一源多用原則、前端防呆、破壞性操作前留底備份、巨集與公式規範化

---

## 國際標準與參考規範

1. **UN/CEFACT**：United Nations Code for Trade and Transport Locations (UN/LOCODE)
2. **ISO 6346**：Freight containers — Coding, identification and marking
3. **財政部關務署**：關港貿單一窗口海空運進出口通關訊息規範 (EDI Standards)
4. **FIATA**：International Federation of Freight Forwarders Associations Standard Rules
5. **Microsoft Learn**：Excel Functions and Data Preprocessing Best Practices

---

## 適用版本與環境

- **作業系統**：Windows 10 / 11 / macOS
- **辦公軟體**：Microsoft Excel 2016 / 2019 / 2021 / Microsoft 365
- **簡報軟體**：Microsoft PowerPoint 2016 或更新版本
- **文字軟體**：Microsoft Word 2016 或更新版本
- **瀏覽器**：Google Chrome, Microsoft Edge, Firefox, Safari（支援 HTML5/ES6 規格）

---

*教材製作：航運行政數位轉型培訓專案*  
*GitHub 帳號：2206qr@gmail.com*
