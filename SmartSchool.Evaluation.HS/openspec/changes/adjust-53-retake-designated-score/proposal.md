# Proposal

## Why

產生學習歷程時，5.3 重讀成績的「再次修習成績」沿用名冊上既有的分數。學期科目 `score_info` 已可寫入「修課直接指定總成績」，但 5.3 回寫沒有讀這個屬性；有直接指定總成績時，再次修習成績仍可能被補考是否存在左右。

## What Changes

- 產生學習歷程、回寫既有 5.3 時，從 `sems_subj_score.score_info` 的 `SemesterSubjectScoreInfo/Subject` 讀取屬性「修課直接指定總成績」。
- 該屬性可解析為數字時，把它寫入既有 5.3 的「再次修習成績」（5.3.7）。有沒有補考成績都不改變這個優先順序。
- 同步重算「再次修習成績及格」（5.3.8）：用同一科目節點的「修課及格標準」比較；標準不存在、空白或無法解析時用 60。
- 直接指定總成績不存在、空白或無法解析時，5.3 的再次修習成績、再次修習成績及格，以及後續補考與重讀成績處理，維持目前行為。
- 重讀成績（5.3.11）仍取再次修習成績、補考成績、原始成績三者中可解析值的最大值。再次修習成績若已改為直接指定總成績，最大值計算使用更新後的值。
- 只改既有 5.3 列的上述分數。產生學習歷程不因指定總成績新增 5.3，也不改 4.2、4.3、4.4、6.2、6.4。
- 計算學期科目成績、`FillSemesterSubjectCalcScore` 建立 5.3 時，5.3.7 與 5.3.8 改跟上述判斷一致。來源是本次修課 `sc_attend.designate_final_score`，對應相同學生、科目、科目級別。`decimal.TryParse` 成功就優先採用，包含 0，與補考成績是否存在無關。
- 5.3.7 先用既有學期科目成績的位數與進位方式處理，再寫入，使數值與同一筆 `score_info` 的「修課直接指定總成績」一致。5.3.8 用寫入後的分數比較有效修課及格標準：大於等於為 `1`，否則為 `0`；標準無效時用 60。
- 指定總成績不存在、空白或無法解析時，5.3.7、5.3.8 仍用原本的 `sacRecord.FinalScore` 與 `HasFinalScore`。本學期科目成績已存在、尚未存在這兩個再次修習分支都要套用。
- 只改這兩個欄位的成績來源與及格判斷。不改 5.3.9 到 5.3.12、擇優、補考、4.2、4.3、4.4、5.2、6.2、6.4，也不放寬原本建立 5.3 的條件。

## Capabilities

### New Capabilities

- `learning-history-retake-score`: 產生學習歷程回寫既有 5.3，以及計算學期科目成績建立 5.3 時，如何決定再次修習成績與再次修習成績及格。

### Modified Capabilities

- （無。既有 `import-makeup-retake-score` 只規範匯入重補修成績的驗證與寫回，不涵蓋學習歷程 5.3。）

## Impact

- `StudentLearningHistoryProcessor.ProcessLearningHistory` 中，讀取既有 5.3 之後、`SaveScores53` 之前的比對。目前 `GetStudentScoreDataAsDictionary` 只回傳補考成績非空白的科目，且未選取「修課直接指定總成績」；直接指定總成績必須能在沒有補考成績時被讀到。
- 回寫仍走既有 `LearningHistoryDataAccess.SaveScores53`。產生學習歷程的對象限於已讀到的 5.3；計算成績仍只把原本會建立的 5.3 放進 `repeatScoreList`。
- 另改 `WearyDogComputer.FillSemesterSubjectCalcScore` 裡兩個再次修習分支的 5.3 `Score`、`ScoreP`。已存在學期科目成績的分支約在 `repeatRec` 第 1475 行；尚未存在的分支約在第 1917 行。指定總成績沿用該分支已用來寫入「修課直接指定總成績」的 `studentFinalScoreDict` 列。
- 不改 4.2、4.3、4.4、5.2、6.2、6.4 的組裝與儲存，也不改 `studentFinalScoreDict` 的建立方式。同一學生、科目、級別若有多筆修課，目前後一筆會蓋掉前一筆，而且查詢沒有排序；這次只記錄風險。
- 驗證方式：編譯 `SmartSchool.Evaluation.HS.csproj`（.NET Framework 4.8 類別庫）。專案內沒有既有單元測試專案。
