# Verify deferred

OpenSpec Verify for `adjust-import-makeup-retake-validation` found no Critical issues. On 2026-10-19 the 4 warnings and 1 suggestion below were left unchanged. Program code was not modified for archive.

## Warning

1. 手動模式不強制「修課及格標準」。自動模式會要求此欄；手動模式可缺欄或空白，空白不覆蓋既有值。`ImportMakeupRetakeScore.cs` 約第 89、136、253、1338 行。
2. 補修 4.3、重修 5.2 與重修截斷經 `ResolveCoursePassStandard`（約第 1758 行）優先使用 Excel 的修課及格標準，含 0。規格仍要求名冊使用回寫前的既有標準。
3. 原始成績只在 `case "原始成績"`（約第 1215 行）寫入。補修分支（約第 1282 行）不再寫第二次。補修為「是」時目前仍會更新。
4. 匯入精靈第一頁與空白第二頁的下一步沒有實際執行。

## Suggestion

1. `匯入重補修成績調整1013.md` 仍描述較早行為：此欄一律必填，名冊用舊標準，小於等於 0 用 60。現行程式與此不同。
