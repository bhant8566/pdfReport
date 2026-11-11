# PDF Report Extractor

獨立的 Windows PowerShell 5.1 工具，將具有**可讀文字與繪圖格線**的 PDF 報表儲存至 Excel。函式庫已包含在本資料夾，執行時不需上網、不需安裝 Microsoft Excel、Python、Node.js 或 Java，也不使用 LocalDataCenter 的程式或資料。

## 執行

1. 保留整個資料夾，尤其是 `lib`、`licenses` 與 `Extractor.Core.psm1`。
2. 將正式 PDF 放進 `Input`，在本工具資料夾開啟 Windows PowerShell。
3. 執行：

```powershell
.\Import-PdfReports.ps1
```

正式報告**固定略過檔案的實際第一頁封面**。也可指定 PDF 檔案或資料夾：

```powershell
.\Import-PdfReports.ps1 -InputPath 'D:\Reports'
.\Import-PdfReports.ps1 -InputPath 'D:\Reports\report-a.pdf','D:\Reports\report-b.pdf'
```

資料夾輸入只讀該層的 `.pdf`，不遞迴子資料夾。提供的三份範例已經裁掉封面，請使用：

```powershell
.\Import-PdfReports.ps1 -InputPath 'D:\Samples' -StartPage 1
```

可將整包搬到任何位置。資料預設存放在工具內的 `Data`；若要分開保存：

```powershell
.\Import-PdfReports.ps1 -InputPath 'D:\Reports' -DataRoot 'D:\ReportHistory'
```

需要 Windows PowerShell 5.1 與 .NET Framework 4.6.2 或更新版本。下載的檔案若被 Windows 封鎖，確認來源後可在本工具資料夾執行 `Get-ChildItem -File -Recurse | Unblock-File`；公司執行原則仍以管理設定為準。

## 輸出及備份

預設資料結構：

```text
Data/
  Reports.xlsx                 持續累加的歷史工作簿
  Reports.previous.xlsx        上次成功儲存前的工作簿
  .session.lock                避免同時匯入同一份歷史資料
  Output/<時間-批次識別碼>/
    manifest.json              此批次的進度、結果與失敗原因
```

`-OutputRoot` 可另行指定批次紀錄根目錄。每次建立新的批次紀錄資料夾；不匯出 CSV 或 JSON 資料集，已有的舊匯出檔不刪除、不覆寫。`manifest.json` 僅供追蹤進度與排錯，不包含報告明細資料。

每個頁首小標題對應一個 Excel 工作表。同類報告往下累加；新增欄位按欄名追加，舊資料該欄留空，欄位順序改變也不會按位置錯接。長名稱會產生穩定的短工作表名稱，完整名稱可在 `_Reports` 查詢。

本次提供的 12 種報告、各自欄位與核對筆數，見 [REPORT-FORMATS.md](REPORT-FORMATS.md)。這份清單是檢查紀錄，程式不會把這些欄名或筆數寫死。

- `_Reports`：報告類型、工作表名稱、累積筆數與更新時間。
- `_Imports`：每個來源檔案的批次、SHA-256、起始頁、成功／失敗、筆數與錯誤。
- 報告明細：保留 PDF 原始欄名及格內換行；不將數字、日期或前導零自動改寫，不依單元格內容執行公式。

每筆明細的來源欄位：

| 欄位 | 意義 |
|---|---|
| `__SourceFile` | 原始 PDF 檔名 |
| `__SourcePage` | PDF 檔案中的實際頁碼，從 1 開始，不是頁尾印出的頁碼 |

報告明細僅保留這兩個額外來源欄位，其他欄位均來自 PDF 表格，包括原本的 `Source`。批次與錯誤紀錄仍保留在 `_Imports` 及 `manifest.json`，報告名稱對應仍由 `_Reports` 管理，不再逐筆重複寫入明細。

匯入既有舊版 `Reports.xlsx` 時，會在下一次 PDF 成功儲存時，一併移除所有報告明細的七個舊欄位：`__ReportType`、`__SourcePath`、`__SourceRow`、`__ReportRange`、`__ImportedAt`、`__BatchId`、`__FileHash`；既有 PDF 資料及兩個保留欄位不變。轉換仍先寫暫存檔再替換，並保留儲存前備份。

匯入程序只將資料存入 Excel，不另匯出 CSV 或 JSON 資料集。可另行使用下方指令，將 Excel 資料嵌入單一離線 HTML 報表。

## 產生離線前端報表

在工具資料夾執行：

```powershell
.\New-OfflineReport.ps1
```

預設讀取 `Data/Reports.xlsx`，產生 `Data/Report.html`。直接雙擊 HTML 即可使用，不需要啟動伺服器、安裝 Excel 或連接網際網路；CSS、JavaScript 及資料都包含在同一檔案內，沒有另外匯出 JSON 或 CSV。

也可指定來源與輸出：

```powershell
.\New-OfflineReport.ps1 -WorkbookPath 'D:\ReportHistory\Reports.xlsx' -OutputPath 'D:\Reports\Report.html'
```

- `_Reports` 索引中的每張報告工作表對應一個頁面，以側邊分頁切換；完整報告名稱保留，管理用索引表不顯示。
- 12 種已知報告依下表指定欄位精確分組，以欄名定位，區分 key 大小寫、不自行改寫識別值；空白 key 集中為一組並明確標示。未列出的報告依第一個 PDF 原始欄位分組，其餘欄位按原順序顯示；`__` 開頭的來源／管理欄位不顯示。
- 頁面保留報告標題、搜尋、排序及分組明細，移除頂端導覽列、報告副標題、產生時間、統計列、群組說明及搜尋結果摘要。每組標題保留 key 與總筆數，展開後不再重複顯示明細總數。
- 表格第一欄表頭顯示 key 欄名（例如 User Name），下方以合併儲存格顯示該 key 一次；跨頁時，每頁各顯示一次，合併高度依該頁實際筆數調整。其餘指定欄位逐筆保留，內容相同也不合併，沒有「共同資訊」區或卡片裝飾。
- `Time Generated` 或 `Resource Accessed Time` 作為觸發時間，顯示為 `MM/dd HH:mm:ss`，24 小時制；每組依完整年月日時間由舊到新排序，同時間保留原始順序。原始時間與來源檔名／頁碼可在時間儲存格的滑鼠提示中查看。
- 支援目前報告的中文月份、上午／下午及換行，也支援英文月份。無法可靠轉換的時間放在該組最後，保留原文、以黃色標示並顯示筆數，不自行推定時間。
- 群組可展開／收合，首組預設展開，預設依筆數由多到少排序，也可依 key 名稱或最近觸發時間排序。搜尋會保留符合群組的全部明細；每頁最多 12 個群組，每個群組每頁 100 筆明細，可按上一頁／下一頁查看全部紀錄。
- HTML 是 Excel 最後儲存資料的快照。Excel 資料更新後，請重新執行指令；未儲存的 Excel 編輯不會帶入。報表生成失敗時保留原 HTML，不修改 Excel。

### 前端顯示欄位

下表只控制離線 HTML 的顯示，不刪除 Excel 或 HTML 內嵌資料中的原始欄位。已知報告缺少必要欄位時，產生器會指出報告與缺少的欄名，保留既有 HTML。欄名比對會忽略前後空白並合併換行／連續空白。

| 報告 | key（表格第一欄合併顯示） | 其餘明細欄位（依序） |
|---|---|---|
| Active Alerts (Trouble) | Source | Time Generated、Alert Message |
| All First Time Resource Access | Reference Resource | Resource Accessed Time、Message |
| All Threshold Breach Activity | Name | Time Generated、Mean Count、Threshold Count、Message |
| All Unusual Time Activity | Name | Time Generated、General Start Time、General End Time、Message |
| First Time Remote Access on Host | Unusual Remote Access From Computer | Resource Accessed Time、Message |
| Unusual Logon Activity Time | User Name | Time Generated、General Start Time、General End Time、Message |
| Unusual Logon Activity Time on Host | Computer Name | Time Generated、General Start Time、General End Time、Message |
| Unusual User Activity Time | User Name | Time Generated、General Start Time、General End Time、Message |
| Unusual Volume of Lockout | User Name | Time Generated、Mean Count、Threshold Count、Message |
| Unusual Volume of Logon Failure | User Name | Time Generated、Mean Count、Threshold Count、Message |
| Unusual Volume of Logon Failure on Host | Computer Name | Time Generated、Mean Count、Threshold Count、Message |
| Unusual Volume of User Management Activity | User Name | Time Generated、Mean Count、Threshold Count、Message |

前端更新後，既有 HTML 必須重新執行 `New-OfflineReport.ps1` 產生，才會套用新版欄位與版面。

報表產生工具只使用 `PdfReportExtractor` 內的程式及函式庫，與 LocalDataCenter 的事件資料、前端及服務分開。

這一版**不去重、不清除七天以前資料、不做日期分類或統計**。重複匯入相同 PDF 會完整追加；日期與七天保留規則待下一階段確認。匯入前請關閉 `Reports.xlsx`，避免被 Excel 鎖定。

備份時保留整個 `Data`，原 PDF 則由使用者保存於原位置；工具只讀取原檔，不搬移或刪除。工作簿先寫入暫存檔再替換正式檔；更新失敗不覆寫舊工作簿。

**每份 PDF 完整解析成功後，立即儲存一次 Excel，再處理下一份。** `Reports.previous.xlsx` 保留最近一次儲存前的版本，可能是同批上一份 PDF 完成時的狀態。中途停止或後續儲存失敗，前面已顯示「已儲存」的資料仍會保留；正在處理的 PDF 不會寫入部分資料。

中斷後請依 `manifest.json` 的 `Files` 確認已成功的檔案，重新執行時只指定尚未成功的 PDF；本版沒有自動續跑或去重，重跑已儲存的檔案仍會重複追加。

## 萃取方式與失敗處理

Tabula Sharp 根據格線還原儲存格邊界；格內文字由隨附的 `PdfReportExtractor.TableText.dll` 依 PdfPig 原始字元與基線重組，不使用 Tabula 的文字合併結果，避免遺失重複字母及空格。此 DLL 已預先編譯，執行時不編譯或安裝套件；原始碼在 `src`，開發者可用 `tools/Build-TextReader.ps1` 重建。不同高度的資料列、長 SID、日期與 Message 換行都保存在同一個儲存格。表頭與頁首頁尾不當成明細。

報告標題支援日期範圍放在右側，或像 `Active Alerts (Trouble) (From ... To ...)` 接在標題後方；查詢期間存入 `__ReportRange`，不會因月份改變而建立新的報告分頁。

部分來源頁面沒有重複印出表頭：當欄數與各欄界線符合前頁，且首列高度大於已知表頭時，沿用先前確認的欄名並保留第一筆資料。無法確認的結構會回報錯誤。預設從第 2 頁匯入時，若封面提供 `Number of Records`，也會核對整份報告筆數；不符則整份不寫入，避免漏列後仍回報成功。

支援範圍是可讀取文字、具有規則格線、單層且非空白或重名表頭的表格。欄名使用 `__` 開頭為工具保留名稱。旋轉、合併儲存格、多層表頭、掃描頁、無格線或無法確認結構的頁面會列出失敗原因，不猜測欄位，也不自動 OCR。單一 PDF 任一資料頁失敗，該檔案全部不匯入；其他正常 PDF 可繼續。

執行結束後看 `manifest.json` 的 `Status` 與 `WorkbookCommitted`：

- `Success`：所有 PDF 解析並儲存成功。
- `PartialSuccess`：正常 PDF 已儲存，部分 PDF 失敗；細節見 `Files`。
- `Failed`：解析或儲存失敗。若 `WorkbookCommitted` 為 false，歷史工作簿未更新；若為 true，前面已成功的 PDF 仍保留，請看 `Files` 確認。
- `Interrupted`：執行中斷，請確認最後進度與 `WorkbookCommitted`；強制結束程序時可能仍留下 `Processing`。

### 大檔看起來沒有反應時

啟動後會立即顯示載入訊息、輸入檔數與實際資料位置；解析中顯示目前檔案、頁碼、總頁數、筆數和經過秒數。每 25 頁另印一行進度，方便看純文字輸出。`manifest.json` 的 `Current` 會在每頁開始與完成時更新，包含 `Page`、`TotalPages`、`ParsedRows`、`Phase` 與 `UpdatedAt`。

**第一份 PDF 完整解析成功後，就會產生或更新正式 Excel**；之後每成功一份再儲存一次，畫面會顯示「已儲存」。第一份尚未跑完時還看不到新工作簿是正常的；數百頁月報會比兩頁測試檔花更多時間。請依畫面印出的資料路徑查找。匯入期間請保持 `Reports.xlsx` 關閉，避免後續儲存被 Excel 鎖定。

若執行的是 `tests/Test-Extractor.ps1`，它會使用新的 `%TEMP%\pdf-extractor-tests-*` 資料夾，結束時印出路徑，不寫入正式 `Data`。`Current` 長時間停在同一頁、或出現錯誤時，請保留最後訊息及 `manifest.json` 供排查；舊版已在執行中的程序不會自動套用更新，需停止後重新執行。

結束碼為 `0`（成功）、`1`（部分或全部 PDF 解析失敗）、`2`（環境、輸入或儲存錯誤）。

## 驗證及依賴

使用原始三份範例所在資料夾執行完整回歸；所有輸出寫入新的系統暫存資料夾，不修改來源 PDF 或正式資料：

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tests\Test-Extractor.ps1 -SampleDirectory 'D:\Samples'
```

完整檢查 Input 中整批正式報告（從第 2 頁起），在新的暫存資料夾匯入，並逐份核對封面的 `Number of Records`、Excel 筆數、來源頁碼及索引，確認沒有額外的資料集匯出：

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tests\Test-InputBatch.ps1
```

此驗證用於帶有上述封面筆數欄位的報告，可能需要數分鐘。若已有一份全新資料目錄的成功批次，可加 `-ExistingDataRoot '該資料目錄'` 只做讀取驗證，避免重複萃取。

依賴版本、官方來源與套件雜湊見 `dependencies.json`，授權見 `licenses`。`tools/Get-Dependencies.ps1` 僅供開發者重新封裝使用，正式匯入不會呼叫它。

離線報表的瀏覽器測試使用獨立合成資料，驗證分組、時間轉換、共同／不同欄位、每筆時間、搜尋、分頁、內容安全與三種畫面寬度：

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tests\Test-OfflineReport.ps1 -ActualReport .\Data\Report.html
```

測試使用本機 Edge 的無介面模式，並將頁面網路設為離線；測試輸出與截圖放在新的系統暫存資料夾。`-ActualReport` 可省略，省略時只測合成資料。
