# 1688 待發貨訂單運單抓取流程

## 流程圖

```mermaid
flowchart TD
    UI[網頁/桌面 GUI<br/>app_web.py / app_gui.py] -->|一鍵執行| S1

    subgraph GAS["☁️ Google Apps Script Web App"]
        direction LR
        GASEP{{HTTPS endpoint<br/>?action=...&token=...}}
        GASEP --> EP1[getDateRange<br/>讀 B1 + 計算昨天]
        GASEP --> EP2[getEmptyWaybills<br/>讀 cell 背景色 + 過濾]
        GASEP --> EP3[updateWaybills<br/>套規則 + setBackground]
    end

    A[(Google Sheet<br/>運單對照 beta)]
    GAS <-->|SpreadsheetApp| A

    S1[fetch_empty_waybill.py] -->|HTTP GET<br/>?action=getEmptyWaybills| GASEP
    S1 --> C[empty_waybill.csv<br/>含 order_time C欄訂購時間]

    UI --> S2[filter_orders_by_date.py]
    S2 -->|HTTP GET<br/>?action=getDateRange| GASEP
    S2 -->|cookies / storage_state| F[Playwright<br/>開 1688 訂單頁]
    F --> G[點下單時間下拉]
    G --> H[JS 設 q-date value<br/>穿透 shadow DOM]
    H --> I[搜索 → 導出當前條件]
    I --> K[對話框 確認導出]
    K -->|等 10 秒| L[對話框 outline<br/>展開記錄]
    L --> M[抓最新下載連結]
    M -->|expect_download| N[xxx.xlsx<br/>~225 列 49 欄]

    N --> T[trim_xlsx 函式]
    T -->|保留 A 訂單編號 + AF 運單號<br/>移除重複合併列| P[xxx_trimmed.xlsx<br/>~138 列 2 欄]

    P --> U[update_waybills 函式]
    U -->|HTTP POST<br/>?action=updateWaybills<br/>body={mapping}| GASEP

    style A fill:#e1f5ff
    style C fill:#fff4e1
    style N fill:#fff4e1
    style P fill:#e1ffe1
    style UI fill:#ffe1f5
    style GAS fill:#fff8d4,stroke:#d4a017,stroke-width:2px
```

## 架構說明

**所有與 Google Sheet 的互動,都改走 GAS Web App HTTPS endpoint**:
- 客戶端**零 OAuth**(不需要 `client_secret.json` / `token.json`)
- 只需 `gas_config.json` 內含 URL + SECRET_TOKEN 兩個字串
- 若 `gas_config.json` 不存在,腳本自動 fallback 到原本的 Sheets API + OAuth 路徑(向下相容)

## 兩階段(GUI 一鍵)

### 階段 1: 取空運單清單 (`fetch_empty_waybill.py`)

**API 路徑**:`GET ?action=getEmptyWaybills&token=...`

**GAS 端邏輯**(`gas_backend.gs` 的 `getEmptyWaybills()`):
- `range.getBackgrounds()` 拿整片背景色矩陣
- 過濾條件:
  - A 欄背景 = 白色 `#ffffff`(排除淺灰 1 `#d9d9d9` 已處理)
  - 訂單編號是長數字(`/^\d{11,}$/`)
  - F 欄(運單號)為空
- 用 `range.getDisplayValues()` 拿格式化字串(日期 = `YYYY-MM-DD HH:mm:ss`)
- 回傳 `{empty: [{row, order_id, order_time, pay_time, address, ...}], total}`

**🔑 1688 日期偵測邏輯**(階段 2 透過 `getDateRange` 取):

| 步驟 | 來源 / 動作 |
|---|---|
| 1. 鎖定 cell | Sheet **B1** ("下次新增數據抓取起始日期" 的值,例如 `5/13`)|
| 2. GAS 端解析 | `getDisplayValue()` → regex `(\d{1,2})/(\d{1,2})` |
| 3. 組合日期 | 月日 + 今年 → 若超過今天則回推一年(跨年情境)|
| 4. 結束日 | `today - 1 day`(昨天)|
| 5. 回傳 | `{start: "YYYY-MM-DD", end: "YYYY-MM-DD", b1Raw}` |

**輸出**:`empty_waybill.csv` — 待補運單訂單明細

### 階段 2: Playwright 自動化 + 精簡 + 回填 (`filter_orders_by_date.py`)

#### 2-1. 1688 自動操作
1. 用 `storage_state.json`(來自 `setup_login.py` 互動掃 QR)或 `cookies_header.txt` 登入
2. 進頁面 → **偵測 `#pc-login-modal`**,若有 = cookie 過期,提示跑 setup_login.py
3. `GASClient().get_date_range()` 取得 `start ~ end`
4. 點「下單時間」`.q-select-selector`
5. **shadow DOM 穿透**:`<q-date>` 是 Web Component(內含 `<ui-datetime readonly>`),不能直接打字。用 Playwright `locator.evaluate()` 設 `value` 屬性 + 派發 `input/change/q-change` 事件
6. 搜索 → 導出當前條件 → 確認對話框 → 等 10 秒 → 展開記錄 → 抓最新下載連結
7. `expect_download()` 接 xlsx → 存到本地

#### 2-2. 精簡 xlsx (`trim_xlsx` 函式)
- `openpyxl` 重建只含「訂單編號 / 運單號」兩欄
- **去重**:訂單編號為空 + 運單號重複 → 跳過(合併儲存格副作用)

#### 2-3. 回填 Google Sheet (`update_waybills` 函式)

**API 路徑**:`POST ?action=updateWaybills` body=`{token, mapping: {orderId: waybill}}`

**GAS 端邏輯**(`gas_backend.gs` 的 `updateWaybills()`):

| Case | 條件 | F 運單 | G 更新日 | H 狀態 | 底色 |
|---|---|---|---|---|---|
| **A1 (升級)** | F 已有 + F == BC(抓出貨表) + H ≠ `>> TW` | (不動) | (不動) | ← `>> TW` | ← 淺灰 1 |
| A2 (跳過) | F 已有(其他情況) | — | — | — | — |
| **B (新填)** | F 空 + 在 xlsx mapping + F == BC | ← 寫入 | ← 今天 | ← `>> TW` | ← 淺灰 1 |
| **B' (新填)** | F 空 + 在 xlsx mapping + F ≠ BC | ← 寫入 | ← 今天 | ← `廠商 >> 倉庫` | (不動) |
| **C (待發)** | F 空 + 不在 mapping + H 空 | (不動) | (不動) | ← `廠商未發貨` | (不動) |

實作:`setValue()` 逐 cell 寫 + `setBackground()` 整列塗灰

## 介面層

### 網頁版 `app_web.py` (推薦)
Flask + SSE 即時串流,深色 Consolas 主題

啟動:
```bash
python d:/1688excel/app_web.py
```
自動開瀏覽器到 `http://127.0.0.1:5000`

按鈕:
- ▶ **一鍵執行** — 依序跑階段 1→2,後端 thread 不阻塞
- 📋 **複製全部 LOG** — 含時間戳的完整日誌
- ⚠ **複製錯誤** — 標紅錯誤行 + 階段標籤
- ✓ **複製動作摘要** — 「已X / 完成 / 統計」類關鍵動作 + 編號
- 🗑 **清除** — 重置面板

技術:
- SSE `/stream` 推送,連線中斷自動重連 + 15s keep-alive
- 1.5s 輪詢狀態列(目前階段 / 計數)

### 桌面版 `app_gui.py`
tkinter 內建,功能對等。執行 `python app_gui.py`

## 檔案清單

| 檔案 | 用途 |
|---|---|
| **`gas_backend.gs`** | **GAS 後端**(複製貼到 script.google.com 部署)|
| **`gas_client.py`** | **Python HTTP 客戶端**(urllib + JSON)|
| **`gas_config.json`** | **GAS URL + SECRET_TOKEN**(機密,別上傳 git)|
| `gas_config.example.json` | 設定檔範本 |
| `GAS_SETUP.md` | GAS 部署完整指南 |
| **`fetch_empty_waybill.py`** | **階段 1**(優先 GAS,fallback OAuth)|
| **`filter_orders_by_date.py`** | **階段 2:Playwright + trim + 回填** |
| **`trim_xlsx.py`** | xlsx 精簡(獨立可跑也可 import)|
| **`update_waybills.py`** | Sheet 回填邏輯(優先 GAS,fallback OAuth)|
| **`app_web.py`** | **網頁 GUI(Flask + SSE)** |
| **`app_gui.py`** | **桌面 GUI(tkinter)** |
| `setup_login.py` | 一次性互動掃 QR 登入 1688,儲存 storage_state.json |
| `cookies_header.txt` | (舊)手動複製 1688 cookies(fallback)|
| `storage_state.json` | Playwright 完整 storage state(setup_login 產生)|
| `client_secret_*.json` | (舊)Google OAuth 憑證 — 設定 GAS 後可移除 |
| `token.json` | (舊)OAuth token — 設定 GAS 後可移除 |
| `empty_waybill.csv` | 階段 1 輸出 |
| `*.xlsx` / `*_trimmed.xlsx` | 階段 2 下載 / 精簡 |
| `FLOW.md` | 本文件 |

## 一鍵執行

**網頁版**:
```bash
python d:/1688excel/app_web.py
```

**命令列**:
```bash
python d:/1688excel/fetch_empty_waybill.py     # 階段 1(走 GAS)
python d:/1688excel/filter_orders_by_date.py   # 階段 2(GAS + Playwright + 回填)
```

**單獨工具**:
```bash
python d:/1688excel/gas_client.py ping          # GAS 連通性測試
python d:/1688excel/gas_client.py daterange     # 取日期範圍
python d:/1688excel/gas_client.py empty         # 取空運單清單
python d:/1688excel/setup_login.py              # 互動掃 QR 更新 1688 storage state
python d:/1688excel/trim_xlsx.py                # 精簡最新 xlsx
python d:/1688excel/update_waybills.py          # 用現有 trimmed.xlsx 回填 Sheet
```

## 關鍵踩雷

1. **GAS 部署後沒改 SECRET_TOKEN**:預設 `change-me-please-2026` 太弱,改完要重新部署(管理部署 → 編輯 → 新版本)
2. **GAS 改 Code.gs 沒生效**:Apps Script 是「版本」概念,要重新部署「新版本」,URL 不變
3. **GAS 回傳 Date 物件**:用 `getValues()` 拿日期 cell 會得到 JS Date,要改用 `getDisplayValues()` 拿格式化字串
4. **GAS Web App quota**:每天 ~20,000 次呼叫上限、單次執行 6 分鐘上限 — 個人用足夠
5. **1688 cookie 過期**:DevTools Application Tab 複製 cookie 不一定能登入(partition / SameSite / Secure 限制),改跑 `setup_login.py` 互動掃 QR 拿完整 storage state
6. **Web Component shadow DOM**:`document.querySelector('q-date')` 找不到,Playwright `locator.evaluate()` 才能穿透 shadow
7. **`readonly` 偽輸入框**:`<q-date>` 不能 `.fill()`,要設 `value` 屬性 + 派發事件
8. **隱藏對話框模板**:1688 把多個 dialog 都塞 DOM,選器要加 `:visible`
9. **`accept_downloads=True`**:`browser.new_context()` 預設不接收下載,要顯式開
10. **openpyxl `delete_cols`**:不會清最大欄寬參考,改用「複製到新 workbook」做法
11. **訂單編號合併儲存格**:1688 匯出 xlsx 同訂單多 SKU 共用第一列,去重要鎖「重複運單 + 空編號」雙條件
12. **subprocess 串流**:GUI 跑子腳本要 `python -u` + `PYTHONIOENCODING=utf-8` + `bufsize=1`,輸出才會即時
