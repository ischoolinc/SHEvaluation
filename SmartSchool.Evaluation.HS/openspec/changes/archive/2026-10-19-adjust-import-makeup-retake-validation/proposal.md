# Proposal

## Why

匯入重補修成績目前把「修課及格標準」當選填，空白也可通過；補修列的「原始成績」沒有條件式必填；補修學期、重修學期只要是整數即可。非補修列仍會把 Excel 原始成績寫回既有學期科目，而且「是否補修成績」分支會把已成立的 `hasChanged` 清掉，導致同一列其他變更被略過。

## What Changes

- 將既有欄位「修課及格標準」列為必要欄位：Excel 必須有此欄，每筆必填，且為 `0` 到 `100`（含端點、允許小數）。沿用現有寫入，不新增「及格標準」欄位。
- 「是否補修成績」為「是」時，「原始成績」必填且必須可解析為數字。為「否」時，Excel 原始成績不驗證、不寫回，既有學期科目原始成績保留，並繼續參與自動取得學分。
- 「補修學期」「重修學期」只要有值就必須是整數 `1` 或 `2`。既有條件式必填維持：補修為是時補修學年度／學期必填；重修成績有值時重修學年度／學期必填。未觸發時空白仍允許。
- 修正「是否補修成績」分支對 `hasChanged` 的清除。同一列補修與重修仍是兩個獨立分支，不可改成互斥。
- 不修改共用匯入精靈、不跳頁、不改名冊寫入方法與取得學分／擇優截斷公式。名冊在回寫前讀到舊修課及格標準的不一致只記錄，不在本變更改公式。

## Capabilities

### New Capabilities

- `import-makeup-retake-score`: 匯入重補修成績的欄位必要條件、格式驗證，以及原始成績是否寫回既有學期科目成績。

### Modified Capabilities

- （無。專案目前沒有既有 spec。）

## Impact

- 僅 `ImportExport/ImportMakeupRetakeScore.cs` 的必要欄位清單與列驗證、以及原始成績寫回與 `hasChanged` 控制。
- 第一頁必要欄位會多出「修課及格標準」，因為它進入 `wizard.RequiredFields`。第二頁選填規則不變，由共用精靈決定；本變更不改 `SmartSchool.Core.General.dll`、`ImportWizard`、`ImportStudentV2.cs`。
- `LearningHistoryDataAccess.SaveScores43`、`SaveScores52` 的呼叫、分組與欄位對應不變。不新增學期科目成績。
- 驗證方式：編譯 `SmartSchool.Evaluation.HS.csproj`（.NET Framework 4.8 類別庫），並以人工情境記錄於 `匯入重補修成績調整1013.md`。專案內沒有既有單元測試專案。
