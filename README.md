# automatic-ai-order

由 RDLife selector 篩選符合條件的工單，找到規格文件，上傳附件並建立 Tracker AI Coding 工單。

> ⚠️ **這是 Winton 內部工具**。它需要連線到內部系統（`crd.winton.com.tw` 的 RDLife / Tracker）、內部 SVN 規格庫與內部程式碼 repo，**僅供 Winton 內部、連得到內網的同仁使用**。外部使用者無法運作。

## 前置需求（clone 後請先確認）

| 需求 | 說明 |
|------|------|
| **PowerShell 7+ (pwsh)** | 所有腳本以 pwsh 執行 |
| **TortoiseSVN** | 預設路徑 `C:\Program Files\TortoiseSVN\bin\TortoiseProc.exe`，用於更新規格庫（可用 `run -SkipSvnUpdate` 略過） |
| **Microsoft Word** | 透過 Word COM 抽取 `.doc/.docx` 規格內文 |
| **內網連線** | 需連得到 `crd.winton.com.tw`（RDLife / Tracker） |
| **RDLife / Tracker 帳號** | 用 `setcred` 設定，DPAPI 加密存本機 |
| **SVN 規格文件 checkout** | 本機一份規格庫資料夾，路徑用 `config -SpecRoot` 指定 |
| **程式碼 repo** | 用來判斷「相關檔案」是否存在；路徑用 `config -ProjectRoot` 指定（預設 `C:\Project\agent1`） |

clone 下來後，照下方「[新同事第一次設定](#新同事第一次設定四步)」設定好個人參數與帳密即可開始使用。

## 推薦入口：`ai-order.ps1`

`ai-order.ps1` 是統一入口，把帳密與參數外部化（**per-user 設定 + DPAPI 加密憑證**），所以整套腳本與對應 skill **可安全分享給同事**——每個人各自設定自己的帳密與規格書路徑，互不影響、密碼不外洩。

設定儲存於 `%USERPROFILE%\.winton-ai-order\`：

| 檔案 | 內容 |
|------|------|
| `config.json` | 非機密參數：工號、規格書路徑（specRoot）、程式碼 repo 路徑（projectRoot）、MaxFunctionPoints、MaxCalculatedHours、aiOrderRoot |
| `cred.xml` | 帳號密碼，**DPAPI 加密**（只有「該 Windows 使用者 + 該機器」能解，不進版控、不明碼） |
| `last-run.json` | 最近一次建單的完整 JSON 結果（供查閱，不會灌進 AI context） |

### 四個動作

```powershell
# 1) 查看目前設定與是否已設帳密（極簡 JSON，唯讀）
& "C:\Project\automatic-ai-order\ai-order.ps1" status

# 2) 設定 / 重設參數（非互動，只帶要改的）
& "C:\Project\automatic-ai-order\ai-order.ps1" config -EmployeeNo CQ1 -SpecRoot "C:\Users\xxx\Desktop\規格文件" -ProjectRoot "C:\Project\agent1" -MaxFunctionPoints 12 -MaxCalculatedHours 4

# 3) 設定帳號密碼（互動，會跳出 Windows 帳密對話框；只需一次）
& "C:\Project\automatic-ai-order\ai-order.ps1" setcred

# 4) 執行建單（只輸出極簡摘要，省 token）
& "C:\Project\automatic-ai-order\ai-order.ps1" run
```

`run` 也接受 `-TargetWorkNo <工號>`、`-DryRun`、`-SkipSvnUpdate`。

### 各動作的使用時機（重要）

| 動作 | 何時用 | 互動性 / 是否彈窗 |
|------|--------|------------------|
| **`config`** | **隨時**。第一次使用前必須先設工號與規格書路徑；之後任何想改參數時（換工號、規格書搬路徑、調整 MaxFP/MaxCH）都能下。它與建單流程**解耦**，只更新存檔、**不會觸發建單**。 | 非互動，無機密，**不需彈窗**，可直接帶參數呼叫 |
| **`setcred`** | **幾乎只需一次**。第一次使用、或密碼變更 / 換帳號時。DPAPI 憑證跨重開機持續有效，平常不用再碰。 | **互動**，需在獨立視窗執行（PowerShell `-NonInteractive` 內會卡住） |
| **`status`** | 任何時候想確認設定狀態，或建單前預檢「是否已 config、是否已設帳密」。 | 非互動，唯讀 |
| **`run`** | 要實際建單時。會先檢查設定與憑證，缺則回 `STATUS=not-configured` / `no-credential` / `spec-root-missing` 並附 `HINT`。 | 非互動 |

> 設計取捨：帳密用 DPAPI 一次設定即持久，因此**沒有 login/logout 狀態管理**；唯一需要「重設機會」的是非機密參數，由 `config` 隨時提供。

### 調整設定的講法（透過 Claude / skill 時）

**只有一份 config，兩個流程共用**：`batch-generate-task`（只建單）與 `batch-generate-task-and-execute`（建單 + 開 session）讀的是同一份 `config.json`。所以**不需要指定是哪個流程的 config**，調一次兩邊同時生效。

**但別只說「調 config」**：Claude 環境裡還有別的 config 會搶觸發（如 `update-config` skill 改的是 Claude Code `settings.json`、`/config` 內建指令）。只說「調 config」可能被當成改 Claude Code 設定而跑錯。

請帶「建單 / automatic-ai-order」字眼，Claude 才會正確載入本流程的 skill 並跑 `ai-order.ps1 config`：

- ✅「調**建單**的 config，MaxFunctionPoints 改 20」
- ✅「改 **automatic-ai-order** 的規格書路徑為 `D:\specs`」
- ✅「**RDLife 建單**的工號改成 AB2」
- ⚠️「調 config」（太籠統，可能被誤判為 Claude Code 設定）

一句話：**「說『調建單的 config』+ 要改什麼值就好，不用分哪個流程。」**

### 工號（EmployeeNo）的兩個作用

`config` 的 `-EmployeeNo`（如 `CQ1`）同時決定：

- **查詢用 designerId**：自動轉小寫（`CQ1` → `cq1`），作為 RDLife selector 的 designer 過濾。
- **commit 分支後綴**：`workingBranch + ".<EmployeeNo>"`（如 `26.07.SP2` → `26.07.SP2.CQ1`）。

### `run` 的輸出（省 token）

完整 JSON 寫入 `last-run.json`，stdout 只回幾行摘要（不論工單多寡）：

```
STATUS=ok
COUNTS=created=2;skipped=5;failed=0
CREATED_TASKS=TS26060111:26.07.SP2.CQ1,TS26060112:26.05.SP2.FEE.CQ1
FULL_RESULT_FILE=C:\Users\xxx\.winton-ai-order\last-run.json
```

`STATUS` 可能值：`ok` / `not-configured` / `no-credential` / `spec-root-missing` / `error` / `parse-error`。

### 新同事第一次設定（四步）

```powershell
# 1. 設工號、自己的規格書資料夾、程式碼 repo 路徑（MaxFP/MaxCH 不帶則用預設 12 / 4）
& "C:\Project\automatic-ai-order\ai-order.ps1" config -EmployeeNo <你的工號> -SpecRoot "<你的規格文件資料夾>" -ProjectRoot "<你的程式碼 repo>"

# 2. 設帳密（會跳出 Windows 帳密對話框，DPAPI 加密儲存）
& "C:\Project\automatic-ai-order\ai-order.ps1" setcred

# 3. 確認設定（configured / credValid 應為 true）
& "C:\Project\automatic-ai-order\ai-order.ps1" status

# 4. 建單
& "C:\Project\automatic-ai-order\ai-order.ps1" run
```

## 流程（建單主邏輯）

1. 登入 Tracker。
2. 呼叫 `RdLifeWorkItemSelector` 取得 RDLife 工單。
3. 篩選狀態為 `W04`，且（功能點 ≤ `MaxFunctionPoints` 或計算工時 ≤ `MaxCalculatedHours`）的項目。
4. 查目前自己的未結 Tracker 工單，避免同一個 RDLife 工單重複建立。
5. 對 `SpecRoot` 做 SVN Update。
6. 使用 `Export-RDLifeSpecMht.ps1` 查 RDLife 工單內容、解析規格檔（預設取原始 `.doc/.docx`，不產 `.mht`）。
7. 比對 RDLife「內容」與規格書內文，若找到對應日期，需求描述會追加「只針對該日期的部分做改動」；若找不到則保留原始需求描述。
8. 若 `ProjectRoot` 下存在與**程式代號同名的資料夾**（忽略大小寫），需求描述再追加一段「## 相關檔案 / 請優先參考 `<程式代號>` 的資料夾」；找不到資料夾則不加。
9. 呼叫 `UploadFile` 上傳規格文件。
10. 以 upload response 的 `systemFileName` / `fileName` 呼叫 `ProcessTempSpecFile` 做規格預處理。
11. 呼叫 `SaveTask` 建立 Tracker 工單（帶入 `workingBranch` / `commitBranch` / `workflowPrompt`）。

## 檔案

- `ai-order.ps1`: **推薦入口**（status / config / setcred / run；per-user 設定 + DPAPI 憑證 + 省 token 摘要）。
- `Invoke-AutomaticAiOrder.ps1`: 舊版包裝器（參數寫死、需自行帶 `-Credential`）。仍可用，但建議改用 `ai-order.ps1`。
- `Invoke-TrackerRdLifeTaskImport.ps1`: Tracker 建單主流程（已參數化 `-CommitBranchSuffix`）。
- `Export-RDLifeSpecMht.ps1`: 查 RDLife 與解析規格檔（`.mht` 產出已由 `-ExportMht` 開關鎖住，預設不產）。
- `query-rdlife.ps1`: RDLife DataSnap 查詢工具。
- `work\rdlife-output`: 查詢與解析暫存輸出。
- `work\upload-temp`: 上傳前暫存檔。

## 注意

- 不會把密碼寫入專案。`ai-order.ps1` 用 DPAPI 加密存於 `%USERPROFILE%`；舊版 `Invoke-AutomaticAiOrder.ps1` 執行時會跳認證視窗或由呼叫端傳入 `-Credential`。
- DPAPI 憑證綁定「使用者 + 機器」：換電腦或重灌後需重新 `setcred`。
