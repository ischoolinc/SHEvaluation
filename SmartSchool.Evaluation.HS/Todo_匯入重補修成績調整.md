# OpenSpec + Cursor Todo.md

## Change
- Change ID: `adjust-import-makeup-retake-validation`
- Feature: 匯入重補修成績
- Target: `ImportMakeupRetakeScore.cs`（實際專案路徑由 Cursor 搜尋確認）
- Reference: `ImportMakeupRetakeScore(4).cs`, `ImportStudentV2.cs`（共用精靈反編譯參考，**不得修改**）
- Mode: 先 `explore` / `propose`，規格確認後再 `apply`、`verify`、`archive`。
- Implement **one Todo at a time**. After each Todo, report changes and verification results before proceeding.
- Change log: `匯入重補修成績調整1013.md`（不存在則新增，保留每項結果與測試證據）。

## Scope and invariants
1. 僅修改本功能必要的驗證／原始成績更新控制。不要修改 `SmartSchool.Core.General.dll`、`ImportStudentV2.cs` 或 `ImportWizard` 介面。
2. 第一頁仍依 `RequiredFields` 顯示必要欄位；第二頁保留現有邏輯，只顯示 Excel 中存在的非必要欄位。**不實作自動跳頁**；若無選填欄位，第二頁可空白，仍可按「下一步」。
3. 不更動 `LearningHistoryDataAccess.SaveScores43`（補修 4.3）及 `SaveScores52`（重修 5.2）的呼叫、資料分組及既有名冊欄位對應。
4. 不調整重修擇優／及格標準截斷規則、自動／手動取得學分公式；但必須驗證本次資料更新與既有公式協作正確。
5. 不新增學期科目成績；維持現有「只能更新既有科目成績」驗證。
6. 不能因無關欄位而清除 `hasChanged`，避免應更新資料被略過。處理同一筆資料補修、重修同時符合的情況時，不可將原本兩個獨立分支改成互斥。

## Confirmed requirements
### R1. 修課及格標準為必填
- 直接使用現有欄位名 `修課及格標準`，**不要新增 `及格標準` 欄位**。
- 加入 `requiredImportFields` / `wizard.RequiredFields`；原本已在 `wizard.ImportableFields`，不可重複新增。
- Excel 必須有此欄位，且每筆不可空白；數字範圍 `0 <= value <= 100`，小數允許；非數字、超出範圍、不存在欄位皆拒絕。
- 移除該欄位原本「空白可接受」的共用檢查，避免重複 switch case。
- 沿用既有寫入 `修課及格標準` 的機制。

### R2. 匯入畫面維持原流程（非此次修改）
- 第一頁顯示必要欄位（含 `修課及格標準`）。
- 第二頁只顯示 Excel 中存在且不在 `RequiredFields` 的非必要欄位，預設勾選、允許取消。
- 不跳頁、不修改共用 DLL。不需為空白第二頁添加假的選填欄位。

### R3. 補修原始成績條件式必填
- `是否補修成績 = 是` 時，`補修學年度`、`補修學期`、`原始成績` 均需有值。
- `原始成績` 必須為可解析的數字；空白／非數字視為驗證錯誤。
- 其他既有補修驗證沿用。

### R4. 補修及重修學期格式
- 只要 `補修學期` 或 `重修學期` **有值**，就必須是整數 `1` 或 `2`。
- 原本條件式必填繼續維持：補修是時須有補修學年度/學期；重修成績有值時須有重修學年度/學期。
- 空白在不觸發條件式必填時允許。

### R5. 非補修完全忽略 Excel 原始成績
- `是否補修成績 = 否`：Excel `原始成績` 無論空白、數字、文字，都**不驗證、不寫回原有學期科目成績**。
- 現有學期科目的原始成績要保留，仍可依原本公式參與自動取得學分判斷。
- `是否補修成績 = 是`：Excel 原始成績驗證成功後，照原有補修流程更新原始成績與產生 4.3 名冊。
- 要檢查 `case "原始成績"` 與 `case "是否補修成績"` 兩個分支，避免先寫回或再被其他分支覆蓋；不能重置先前已成立的 `hasChanged`。

## Key source locations (reference file line numbers; verify in repository)
- Required / Importable fields: `ImportMakeupRetakeScore(4).cs` lines 36–91.
- Row format and conditional checks: lines 119–315, 484–523.
- 4.3 / 5.2 record preparation: lines 595–1080.
- Existing semester detail field updates: lines 1129–1385, especially 1195–1203 and 1254–1287.
- Auto credit logic: lines 1341–1383.
- SaveScores43 / SaveScores52: lines 1633–1668.
- Shared wizard filters optional fields: `ImportStudentV2.cs` `CheckFile()` lines 462–501; builds validation selection: lines 559–592.

## OpenSpec process
### Explore
- [ ] Confirm real source path, branch, solution/framework, build commands and test strategy.
- [ ] Confirm OpenSpec CLI syntax by checking installed version/help; do not assume command names or flags.
- [ ] Summarize current validation, writing, and name-list flows, noting any ambiguity before editing.

### Propose
- [ ] Create OpenSpec change `adjust-import-makeup-retake-validation` using the repository's OpenSpec conventions.
- [ ] Write proposal, affected behavior/spec deltas, implementation tasks, and acceptance tests.
- [ ] Mark shared wizard UI modification explicitly **out of scope**.
- [ ] Obtain confirmation of the generated spec before applying.

### Apply — one Todo at a time
- [ ] **Todo 01 — 修課及格標準**：列為必要欄位，檢查 Excel 欄位存在、每筆必填、數字 0–100（含小數）；沿用現有寫入。
- [ ] **Todo 02 — 保留匯入畫面**：只確認第一頁顯示新增必要欄位、第二頁選填運作正常；**不修改共用元件**。
- [ ] **Todo 03 — 補修原始成績**：補修是時原始成績必填且可解析；錯誤欄位訊息清楚。
- [ ] **Todo 04 — 學期範圍**：補修學期、重修學期有值就限制 `1` 或 `2`，仍保留條件式必填。
- [ ] **Todo 05 — 非補修原始成績**：分離原始／重修格式驗證；非補修不驗證、不更新原始成績；補修仍更新。檢查 `hasChanged` 不被覆寫。
- [ ] **Todo 06 — 回歸保護**：驗證自動／手動取得學分與 4.3／5.2 名冊，除需求必要變動外保持原樣。
- [ ] **Todo 07 — 變更紀錄**：逐項寫入 `匯入重補修成績調整1013.md`，記錄檔案、方法、測試與結果。

### Verify
- [ ] 編譯專案；不能編譯時提出實際限制，不宣稱通過。
- [ ] 測試必要欄位缺少（第一頁阻擋）及欄位內容空白／`-1`／`0`／`40.5`／`100`／`101`／`ABC`。
- [ ] 測試補修是 + 原始成績空白、非數字、正常數字；補修否 + 原始成績空白、數字、`ABC`，確認原始資料不被覆蓋。
- [ ] 測試補修／重修學期空白、`0`、`1`、`2`、`3`、`1.5`、文字；在原有條件式必填情況下確認空白是否應報錯。
- [ ] 測試補修單獨、重修單獨、同筆補修與重修都符合、無名冊寫入等情境。
- [ ] 驗證 4.3 `SaveScores43`、5.2 `SaveScores52` 寫入仍使用原有年度學期分組，確認筆數、欄位及既有資料是否受影響。
- [ ] 驗證自動／手動取得學分；非補修時原始成績用系統既有值計算。
- [ ] 查明並記錄「Excel 新修課及格標準 vs 名冊在學期成績回寫前讀到的舊修課及格標準」的不一致風險。**不得在未確認規格前自行修改名冊計算公式**。
- [ ] 確認第一頁 `修課及格標準` 是必要欄位；第二頁只出現可用選填欄位，若空白仍可按下一步。

### Archive
- [ ] 驗證與人工測試通過後，依目前 OpenSpec CLI 版本執行 archive。
- [ ] 將規格、實作檔案、測試結果及變更紀錄納入版本控制。

## Cursor execution rules
1. Do not bulk-edit all Todo items. Perform one Todo, show exact diff and affected files, wait for approval.
2. Do not edit unrelated modules, generated vendor/decompiled sources or DLLs.
3. Do not silently alter business rules, UI steps, default pass-score fallback, score clipping or name-list semantics.
4. If a requirement conflicts with existing code or a dependency is missing, stop and explain with file/method/line evidence.
5. Make minimal, reviewable changes, avoid redundant validations or code paths, and provide test evidence.
