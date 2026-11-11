# 本次 12 種報告的格式核對

以下是這次 Input 月報的檢查紀錄，不是程式內的固定欄位設定。正式匯入仍從 PDF 表頭動態取得欄名及數量；以頁首小標題分工作表，保留各類自己的欄位，不將 `Name`、`User Name`、`Computer Name` 等欄位自行混為同一欄。

頁數包含第一頁封面，資料從實際第 2 頁開始。筆數依這批檔案封面的 `Number of Records` 核對，不代表以後每批都應有相同數字。

| 報告小標題 | 頁數 | 原始欄位數 | 本批筆數 |
|---|---:|---:|---:|
| Active Alerts (Trouble) | 1027 | 6 | 16579 |
| All First Time Resource Access | 2 | 6 | 3 |
| All Threshold Breach Activity | 699 | 8 | 5659 |
| All Unusual Time Activity | 539 | 8 | 5098 |
| First Time Remote Access on Host | 2 | 6 | 3 |
| Unusual Logon Activity Time | 160 | 10 | 2529 |
| Unusual Logon Activity Time on Host | 130 | 9 | 2549 |
| Unusual User Activity Time | 3 | 10 | 20 |
| Unusual Volume of Lockout | 2 | 9 | 7 |
| Unusual Volume of Logon Failure | 164 | 10 | 2596 |
| Unusual Volume of Logon Failure on Host | 172 | 9 | 2852 |
| Unusual Volume of User Management Activity | 14 | 10 | 204 |
| **合計** | **2914** | | **38099** |

## 欄位差異

- **Active Alerts (Trouble)**：Source、Domain、Severity、Time Generated、Alert Message、Alert Profile Name。這份的查詢日期接在報告標題後方，已獨立取出為 `__ReportRange`，標題內原本的 `(Trouble)` 保留。
- **All First Time Resource Access**：Reference Resource、Domain Name、Unusual resource accessed、Resource Accessed Time、Activity Type、Message。
- **First Time Remote Access on Host**：Unusual Remote Access From Computer、Domain Name、Server Name、Resource Accessed Time、Activity Type、Message。
- **All Threshold Breach Activity**：Name、Domain Name、Hour Of Activity、Time Generated、Mean Count、Threshold Count、Activity Type、Message。
- **All Unusual Time Activity**：Name、Domain Name、Unusual Activity Hour、Time Generated、General Start Time、General End Time、Activity Type、Message。
- **Unusual Logon Activity Time／Unusual User Activity Time**：User Name、SID、Domain Name、Unusual Activity Hour、Time Generated、General Start Time、General End Time、Activity Type、Message、Activity Analyzer。
- **Unusual Logon Activity Time on Host**：Computer Name、Domain Name、Unusual Activity Hour、Time Generated、General Start Time、General End Time、Activity Type、Message、Activity Analyzer。
- **Unusual Volume of Logon Failure／Unusual Volume of User Management Activity**：User Name、SID、Domain Name、Hour Of Activity、Time Generated、Mean Count、Threshold Count、Activity Type、Message、Activity Analyzer。
- **Unusual Volume of Logon Failure on Host**：Computer Name、Domain Name、Hour Of Activity、Time Generated、Mean Count、Threshold Count、Activity Type、Message、Activity Analyzer。
- **Unusual Volume of Lockout**：User Name、Domain Name、Hour Of Activity、Time Generated、Mean Count、Threshold Count、Activity Type、Message、Activity Analyzer。

所有報告僅另加 `__SourceFile`、`__SourcePage` 兩個來源欄位，所以 Excel 的實際欄數為原始欄數加 2。PDF 原本的 `Source` 等欄位完整保留。不同高度的資料列依格線判斷，文字換行不會變成新紀錄。不另匯出 CSV 或 JSON 資料集。

## 檢查方式

- 逐份檢查資料起始頁、中間頁、末頁的表頭、格線及欄數。
- 比對 3,151 個抽查儲存格：預先編譯文字重組器與原本保留空格、換行的重組結果一致。
- 完整解析所有資料頁，再核對封面筆數、Excel 分頁及來源頁碼。
- 可使用 `tests/Test-InputBatch.ps1` 在獨立暫存目錄重跑整批驗證，不修改正式 Data。

本批完整實跑通過：12 份、2,914 頁（含封面）、38,099 筆，每份 Excel 報告工作表均與封面筆數相符。回歸檢查涵蓋兩個來源欄位、舊欄位轉換、保留 PDF 的 Source 欄位、累加、新欄位、換行、逐份儲存、後續儲存失敗時保留前份資料、失敗不寫入、檔案鎖定、離線搬移、預設路徑及不額外匯出資料集。

另以這 12 張工作表的舊版 Excel 副本驗證欄位轉換：保留 38,099 筆既有紀錄、364,051 個保留欄位的儲存格（含表頭），逐格比對一致；再追加 3 筆測試資料成功，且轉換前備份與舊檔完全一致。測試只使用獨立暫存資料。

特別處理兩個沒有重複表頭的頁面：`Active_Alerts_(Trouble).pdf` 實際第 538 頁，以及 `Unusual_Volume_of_Logon_Failure.pdf` 實際第 127 頁。沿用已確認的欄位格線與列高判斷，保留該頁第一筆資料；程式沒有寫死這兩個檔名或頁碼。
