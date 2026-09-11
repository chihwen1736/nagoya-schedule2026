# 2026 名古屋亞運｜中華台北代表團每日賽程查詢系統

供內部工作人員以手機快速查詢中華台北代表團每日賽程的純前端網站。無登入、無資料庫、無需伺服器，可直接部署到 GitHub Pages。視覺風格比照姊妹系統「2026名古屋亞運中華台北代表團據點查詢系統」（深藍＋白＋淺灰為主，橘色線條作為系列識別），維持同一系列內部工具的一致觀感。

## 檔案結構與用途

```
index.html                      網站主頁面（結構）
style.css                       樣式（版面、配色、RWD）
script.js                       主要邏輯（日期切換、搜尋、篩選、排序、渲染）
data/
  schedule.js                   網站實際載入的賽程資料 + 更新日期（window.SCHEDULE_DATA / window.SCHEDULE_META）
  schedule.json                 賽程資料的純 JSON 版本（供工具 / 除錯 / 未來其他系統使用）
  meta.json                     單獨的更新日期 / 總筆數（與 schedule.js 內的 SCHEDULE_META 相同，方便查看）
  validation_report.json        每次轉換後自動產生的資料驗證報告
scripts/
  convert_excel_to_json.py      Excel → schedule.json / schedule.js / meta.json 轉換工具（★ 未來更新賽程只需要這一支程式）
README.md                       本說明文件
```

**Excel 是唯一正式資料來源。** 網站本身**不會**在瀏覽器中直接讀取 Excel；Excel 只在「更新資料」這個階段使用，執行轉換工具後會產生 `data/schedule.js`，網站只依賴這個檔案運作。GitHub Pages 上不需要放原始 Excel。

---

## 更新賽程（給未來自行操作的你）

整體流程固定為：

```
原始 Excel（修改賽程）
      ↓ 執行 convert_excel_to_json.py
自動重新產生 schedule.js / schedule.json / meta.json / validation_report.json
      ↓ 上傳到 GitHub
網站自動顯示最新資料
```

你**不需要**手動修改 `schedule.js`、`schedule.json`、或 `index.html` 裡的任何賽程內容或日期文字——這些都由轉換工具自動產生。

### 1. 要修改哪一個 Excel

修改你原本收到的官方賽程 Excel（例如 `2026名古屋亞運每日賽程表.xlsx`）。**請直接在這份 Excel 上修改**：改時間、改項目、改選手、改對手、改場館、改備註、填成績、填名次、新增一列（一筆賽程）、或整列刪除（取消的賽程）都可以，格式只要維持：
- 每個日期一個分頁（例如 `賽程0920`）
- 分頁第 3 列的 `B3` 儲存格是該分頁的日期
- 第 5 列是標題列
- 第 6 列開始才是實際賽程資料（欄位順序：運動種類／時間／項目(量級)及階段／參賽選手／對手國(姓名)／成績(比數)／地點／名次／備註）

### 2. Excel 要放在哪個資料夹

放在你電腦上任何位置都可以（例如放在跟這個網站資料夾同一層，方便輸入路徑）。轉換工具會用你提供的路徑去讀取它，不會自動去別的地方找檔案。

### 3. Mac 使用者如何執行轉換工具（第一次設定，之後不用再做）

1. 打開「終端機」（Terminal，可用 Spotlight 搜尋「終端機」開啟）。
2. 確認電腦已有 Python 3（Mac 通常內建）：
   ```
   python3 --version
   ```
3. 安裝轉換工具需要的套件（只需要做一次）：
   ```
   pip3 install openpyxl
   ```

### 4. 之後每次更新賽程要執行哪一個指令

在終端機中，先切換到這個網站的資料夾（範例，請換成你實際存放的路徑）：

```
cd /Users/你的名字/Desktop/nagoya-schedule
```

接著執行轉換工具，**把下面路徑換成你這次要用的 Excel 檔案**：

```
python3 scripts/convert_excel_to_json.py "/Users/你的名字/Desktop/2026名古屋亞運每日賽程表.xlsx" data
```

若這次的 Excel 是在你更新的當天處理，工具會**自動**把「資料更新日期」設為執行當天的日期，不需要手動輸入。如果你想指定成別的日期（例如提前製作、隔天才上架），可以加上 `--date`：

```
python3 scripts/convert_excel_to_json.py "/Users/你的名字/Desktop/2026名古屋亞運每日賽程表.xlsx" data --date "2026年9月20日"
```

### 5. 執行後會更新哪些檔案

指令執行完成後，`data/` 資料夾內以下 4 個檔案會被**覆蓋、自動重新產生**：

- `data/schedule.js`（網站實際使用，內含賽程資料 + 更新日期）
- `data/schedule.json`
- `data/meta.json`
- `data/validation_report.json`

其他檔案（`index.html`、`style.css`、`script.js`、`scripts/convert_excel_to_json.py`）**不會被更動**。

### 6. 哪些檔案需要重新上傳 GitHub

只需要上傳 / push 更新過的 `data/` 資料夾（裡面的 4 個檔案，尤其是 `schedule.js`）。如果 `index.html`、`style.css`、`script.js` 沒有異動，不需要重新上傳。

常見做法（有安裝 GitHub Desktop 或 git 指令都可以）：
```
git add data/
git commit -m "更新賽程資料 2026年9月20日"
git push
```
push 完成後，GitHub Pages 網站約 1 分鐘內會自動更新，不需要重新設定 Pages。

### 7. 如何確認資料轉換成功

轉換工具執行完成後，終端機會直接印出摘要，例如：

```
=== 轉換完成 ===
Excel 分頁總數：25（扣除範例分頁後日期分頁：25）
成功匯入日期數：25
系統資料筆數：977
無法解析時間筆數：0
多時段資料筆數：33
疑似重複群組數：0
資料更新日期（將顯示於網站頁尾）：2026年9月11日
已輸出：data/schedule.json
已輸出：data/schedule.js
已輸出：data/meta.json
已輸出：data/validation_report.json
```

請確認：
- **系統資料筆數**是否符合你這次修改後的預期（例如新增 1 筆，總數應該多 1）。
- 沒有出現 Python 錯誤訊息（紅字 Traceback）。
- 想看更詳細的檢查結果，可以打開 `data/validation_report.json`，裡面有：
  - `totalRecords`：系統資料總筆數
  - `recordsPerDate` / `recordsPerSport`：每個日期、每個運動種類各幾筆，方便對照你這次改了哪裡
  - `unparseableTimeSamples`：無法辨識的時間格式（需要人工確認，但資料仍會照常顯示，不會消失）
  - `multiTimeSamples`：時間欄位包含多個時段的資料（例如上午、下午都有場次）
  - `suspectedDuplicateGroups`：疑似重複資料的提示（工具只會提示，**不會自動刪除或修改任何資料**，需要你自己確認是否為真的重複）
  - `warnings`：例如分頁日期跟分頁名稱對不上等異常提示
- 最後打開網站確認畫面（本機直接用瀏覽器打開 `index.html`，或打開部署好的 GitHub Pages 網址），確認：
  - 頁尾「資料更新：」的日期已經變成你這次指定 / 執行的日期
  - 你這次修改的賽程（新增 / 刪除 / 改時間等）已經正確顯示

轉換工具的原則（寫在程式最上方註解中，請勿修改破壞）：
- 不自動補齊、推測、修改任何賽程內容（時間、對手、場館、選手、比賽是否晉級等）——Excel 裡寫什麼，系統就顯示什麼。
- 不因看起來重複就自動刪除資料，只會在驗證報告中提示疑似重複，由人工確認。
- 僅進行「顯示 / 搜尋標準化」：去除多餘空白、統一少數已知的運動種類別名寫法（例如「田俓」→「田徑」僅影響篩選比對，原始文字仍保留在 `sport` 欄位）。
- 時間格式不一致時，會盡量解析出「開始時間」以利排序；無法解析則保留原文字，並在畫面上以 ⚠ 提示，但**一律照常顯示**，不會因排序失敗而消失。
- 「資料更新日期」預設自動取執行當下的日期，寫入 `data/schedule.js` 的 `window.SCHEDULE_META.updateDate`，網站頁尾自動讀取這裡，不需要在 HTML / JS 手動修改多處。

---

## 如何部署到 GitHub Pages（第一次上架時）

1. 建立一個新的 GitHub Repository（或使用現有的）。
2. 將本資料夾內所有檔案（`index.html`、`style.css`、`script.js`、`data/`、`scripts/`、`README.md`）上傳 / push 到該 repository 的預設分支（例如 `main`）。
   - **不需要**上傳原始 Excel 檔案。
3. 到 Repository 的 **Settings → Pages**。
4. 在「Build and deployment」的 Source 選擇 **Deploy from a branch**，Branch 選 `main`（或你使用的分支）、資料夾選 `/ (root)`，儲存。
5. 等待約 1 分鐘，GitHub 會提供一個網址，例如：
   `https://<你的帳號>.github.io/<repository名稱>/`
6. 之後若賽程更新，只要依照上一節「更新賽程」重新產生 `data/` 底下的檔案並 push 上去，網站就會自動更新（不需要重新設定 Pages）。

---

## 系統操作說明（給工作人員）

- **日期切換**：畫面上方可左右滑動 / 點選日期，選到的日期會立即顯示當天賽程，不需要重新整理頁面。今天的日期會特別標示。
- **搜尋**：在搜尋框輸入選手姓名、項目、場館或對手，會即時篩選，不需要按搜尋鍵。
- **搜尋範圍**：預設只搜尋「當日」；切換成「全部日期」可以查到某位選手在整個賽會期間的所有賽程。
- **項目篩選**：可篩選特定運動種類，項目清單是直接從資料自動產生的。
- **全部 / 接下來**：「接下來」會依目前日本時間（JST，每次操作皆重新取得，不會停留在剛開啟頁面時的時間），只顯示尚未開始或即將開始的賽程；若同一筆資料包含多個時段（例如同時列出上午、下午場次），只要其中任一時段尚未開始就會保留；無法判斷時間的資料一律照常顯示，不會被隱藏。
- 若某天沒有中華台北代表隊賽程，會顯示提示並提供「查看下一個賽程日」的按鈕。
- 若搜尋 / 篩選沒有結果，會顯示「找不到符合條件的賽程」並提供「清除搜尋條件」按鈕。

## 目前版本刻意不包含的功能（依需求規劃）

登入、成績回填、資料庫、Google Sheets、Apps Script、Firebase、推播通知、留言、後台管理、Google Maps、AI 功能、複雜動畫。

資料結構已保留「成績（比數）」與「名次」欄位（若 Excel 本身已有資料會顯示，空白則不顯示），方便未來擴充。
