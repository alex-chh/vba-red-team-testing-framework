# VBA Red Team Testing Framework

以新手友善為目標，整合並完善 `README_TESTING.md` 與 `TestingGuide.txt`，提供一步一步的操作與背景知識，幫助你在授權的內部測試環境中安全、可控地驗證 VBA 端點防護與偵測能力。

---

## ⚠️ 法律與道德聲明（必讀）
- 僅供授權的內部安全測試與紅隊演練使用。
- 部署與執行前必須取得明確書面授權並遵守法律與組織政策。
- 請在隔離的測試環境中使用，避免影響生產系統與非目標資產。

---

## 你會得到什麼
- 可導入 Microsoft Word 的 5 個 VBA 模組，包含 C2 通訊、系統偵察、反分析混淆、自動執行與啟動流程。
- 一套可逐步操作的測試方法：手動指令、文件自動執行、綜合測試腳本。
- 安全與合規提醒、監控與清理建議、故障排除與常見問答。

---

## 系統需求與前置作業
- Windows（建議 Windows 10/11）。
- Microsoft Word（支援 VBA，建議 Office 2016 以上）。
- 測試環境可連到 C2 伺服器或本機 HTTP（預設 `C2MainModule.bas` 中 `C2_SERVER_URL` 為 `http://172.31.44.17:8080`）。
- 啟用宏（僅在測試環境）：
  - Word → 檔案 → 選項 → 信任中心 → 信任中心設定 → 巨集設定 → 選擇允許巨集（或以受信任位置進行）。

---

## 專案結構與模組說明
- `C2MainModule.bas`
  - `InitializeC2Connection`：初始化 C2 連線並收集系統資訊。
  - `EstablishEdgeC2`：以 `msedge.exe` 建立瀏覽器型 C2 通道（測試合法程序濫用）。
  - `ExecuteCommand`：透過 `cmd.exe` 執行系統命令。
  - `SendBeacon`：HTTP GET 信標（含 `URLEncode`）。
- `UtilityFunctions.bas`
  - 系統偵察：`GetFullSystemInfo`、使用者/電腦名、處理程序、網路連線。
  - 檔案操作：`FileExists`、`ReadFileContents`、`WriteToFile`、`ListFilesInDirectory`。
  - PowerShell 執行、Base64 編解碼、登錄檔讀寫、URL 編碼。
- `ObfuscationTechniques.bas`
  - 字串混淆/還原（XOR）、動態呼叫與程式流程混淆。
  - 環境檢測：`IsDebuggerPresent`、`IsSandbox`、`AntiAnalysisChecks`。
  - 多型執行：`ExecutePolymorphicCode`。
  - 清理：`CleanArtifacts`、（測試用）`DisableMacroWarnings`。
- `AutoExecModule.bas`
  - 自動執行巨集：`AutoExec`、`AutoOpen`、`Document_Open`。
  - 延遲與隨機化、環境安全檢查、隱蔽 C2 初始化、簡易持久化示意。
- `StartupModule.bas`
  - 文件啟動事件與延遲執行、簡易調試器/沙箱檢測、隱蔽 C2 測試流程。

---

## 快速開始（五步走）
1. 下載或克隆本倉庫到測試機。
2. 開啟 Word，按 `ALT + F11` 進入 VBA 編輯器。
3. 在左側「Project」視窗對 `Normal` 右鍵 → `Import File`，一次導入五個 `.bas`：
   - `C2MainModule.bas`
   - `UtilityFunctions.bas`
   - `ObfuscationTechniques.bas`
   - `AutoExecModule.bas`
   - `StartupModule.bas`
4. 開啟「即時運算視窗」（`Ctrl + G`），執行基本測試：
   - 輸入並按 Enter：`SimpleTest` 或 `TestAllFunctions`
5. 若要測試自動執行：
   - 將文件另存為 `*.docm` → 關閉再開啟 → 觀察自動事件 `AutoOpen`/`Document_Open` 行為。

---

## 測試方法與範例
- 手動測試（即時運算視窗）：
  - `InitializeC2Connection`（初始化 C2）
  - `EstablishEdgeC2`（建立 Edge C2 通道）
  - `ExecuteCommand("whoami")`（執行系統命令）
  - `GetFullSystemInfo()`（完整系統資訊）
  - `GetRunningProcesses()`、`GetNetworkConnections()`
  - `MainObfuscated`（混淆主程式）
- 自動執行測試（`.docm` 重新開啟）：
  - `AutoExecModule` 與 `StartupModule` 會觸發 `AutoOpen`/`Document_Open` 與延遲執行，隨機選擇 C2 行為（發送信標、初始化、Edge C2、基本偵察）。
- 綜合測試腳本：
  - `TestModule.bas` 中：`SimpleTest`、`TestAllFunctions`、`TestSystemInfo`、`TestCommandExecution`、`TestFileOperations`、`TestObfuscation`、`TestNetworkFunctions`、`TestEdgeC2`。

示例輸出（`SimpleTest`）：
```
=== Simple Test Started ===
User: YourUsername
Computer: YourComputerName
Current User: domain\username
Hostname: COMPUTER01
=== Simple Test Completed ===
```

---

## 基礎指令速查
- 取得系統資訊：
  - `GetFullSystemInfo()`
- 執行命令並回傳：
  - `ExecuteCommand("echo Hello World")`
- 進程與連線：
  - `GetRunningProcesses()`、`GetNetworkConnections()`
- 字串混淆：
  - `ObfuscateString("secret")`、`DeobfuscateString(obfuscated)`
- C2 測試：
  - `InitializeC2Connection`、`EstablishEdgeC2`、`SendBeacon("data")`

---

## 進階操作與情境
- 偵測逃避測試：
  - `IsDebuggerPresent()`、`IsSandbox()`、`AntiAnalysisChecks`
  - `ExecutePolymorphicCode`（時間與路徑隨機化）
- 通訊測試：
  - Beacon 成功行為：HTTP `200-299` 與 `404`（路徑不存在但連通）都視為連線測試成功。
- 持久化/Office 設定測試（僅限測試）：
  - `DisableMacroWarnings`、登錄檔寫入（請審慎，避免污染環境）。

---

## 監控與記錄建議（防守視角）
- Windows 事件：4688（Process Creation）、PowerShell Operational Logs、Office 巨集事件。
- 網路監控：HTTP/DNS 模式、憑證異常、非常規 User-Agent。
- 端點偵測：子進程樹、臨時目錄檔案、登錄檔修改。

---

## 清理與還原
- 執行暫存清理：`CleanArtifacts`。
- 關閉由測試啟動的 `msedge.exe` 等程序。
- 還原宏安全設定與登錄檔（如有更動）。

---

## 故障排除
- 編譯錯誤：確認五個 `.bas` 全部成功導入，且在標準模組（非 Class）。
- 無輸出：請確定在「即時運算視窗」執行並按 Enter。
- 宏被阻擋：檢查信任中心設定或使用受信任位置。
- 無法連線 C2：請確認 `C2_SERVER_URL` 指向可用的測試伺服器或改為本機服務。

---

## 常見問答（FAQ）
- Q：可以在生產環境測試嗎？
  - A：不可以。僅限授權、隔離的測試環境。
- Q：如何修改 C2 伺服器位址？
  - A：在 `C2MainModule.bas` 調整 `C2_SERVER_URL` 為你的測試 URL。
- Q：Edge C2 一定會連線成功嗎？
  - A：不保證，請以信標與網路層監控佐證；測試目的在於流程與偵測。

---

## 維護與建議
- 在每次測試前確認環境安全與授權邊界，避免產生誤報與誤傷。
- 若要擴充：新增測試案例至 `TestModule.bas`，並維持簡單清楚的介面與輸出。

---

> 責任提醒：請你在合法、合規且已明確授權的前提下使用本專案。任何不當使用造成的後果需自行承擔。
