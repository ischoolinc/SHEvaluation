# Proposal

## Why

2026-11-02 更新：依使用者最新要求，再次修習的原始成績改為本學期有效成績、既有學期科目成績、封存成績三者擇優。此規則取代先前「指定總成績直接覆蓋原始成績」及「既有學期優先、封存僅備用」的規則。

## What Changes

- 本學期有效成績優先使用可解析的修課直接指定總成績（包含 0），否則使用本次課程總成績，沿用既有小數位與進位方式。
- 封存只取早於本次計算學年度、學期的資料。每位學生各學期先選最後更新的封存，再以科目名稱＋級別跨封存學期取最高分。
- 同科目名稱＋級別的本學期有效成績、既有學期成績、封存成績三者取最高值，寫入本學期原始成績。
- 新增與更新套用同一共用方法；既有學期成績排除本次計算學期。
- 5.3 再次修習成績仍使用本次有效成績；重讀成績使用本次有效成績與歷史最高分擇優。

## Capabilities

### New Capabilities

- `semester-subject-repeat-original-score`: 定義再次修習原始成績與 5.3 成績來源。

### Modified Capabilities

- 無。

## Impact

修改 `WearyDogComputer.cs`、`RepeatSubjectOriginalScore.cs` 與 `tests/RepeatSubjectOriginalScoreCheck/Program.cs`。不變更資料表，不主動重寫既有資料；需重新執行學期科目成績計算才會套用。一般修課、重修、補修與產生學習歷程流程未修改。