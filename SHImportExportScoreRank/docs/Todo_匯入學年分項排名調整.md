# Todo：調整匯入學年分項排名寫入方式

## 目標

修改專案：

```text
C:\ischoolProject\SHEvaluation\SHImportExportScoreRank
```

主要修改檔案：

```text
DAO\RankOverrideDataAccess.cs
```

主要修改方法：

```csharp
UpsertSchoolYearEntryRanks(
    List<SchoolYearEntryRankImportRecord> records)
```

調整「匯入學年分項排名」寫入 `rank_override` 的新增、更新與資料比對方式。

完成後，將實際修改內容、SQL、測試結果及注意事項記錄於：

```text
匯入學年分項排名調整.md
```

---

# 一、參考資料

請依下列檔案內容實作：

```text
ed01.txt
DAO\RankOverrideDataAccess.cs
```

不得修改其他匯入流程的主要架構。

不得改動 `rank_override` 資料表結構。

---

# 二、rank_override 資料表欄位

```text
id              bigint
ref_student_id  bigint
school_year     integer
semester        integer
grade_year      integer
item_type       character varying
ref_exam_id     bigint
item_name       character varying
rank_type       character varying
rank_name       character varying
rank            integer
matrix_count    integer
extension       jsonb
```

---

# 三、匯入資料寫入規格

`UpsertSchoolYearEntryRanks` 寫入 `rank_override` 時，必須依下列規格處理：

| rank_override 欄位 | 寫入規則 |
|---|---|
| `ref_student_id` | `SchoolYearEntryRankImportRecord.RefStudentID` |
| `school_year` | 學年度，來源為 `SchoolYearEntryRankImportRecord.SchoolYear` |
| `semester` | 固定 `-1` |
| `grade_year` | 年級，來源為 `SchoolYearEntryRankImportRecord.GradeYear` |
| `item_type` | 固定 `學年/分項成績` |
| `ref_exam_id` | 固定 `-1` |
| `rank_type` | 固定 `班排名` |
| `rank_name` | 班級名稱，來源為 `SchoolYearEntryRankImportRecord.RankName` |
| `rank` | 學年學業排名，來源為 `SchoolYearEntryRankImportRecord.Rank` |
| `matrix_count` | 不寫入，由資料庫維持 `NULL` 或既有值 |
| `extension` | 新增時寫入 SQL `NULL`，不可再寫入 `'{}'::jsonb` |

固定值優先使用既有常數：

```csharp
RankOverrideConstants.SchoolYearSemester
RankOverrideConstants.ItemType
RankOverrideConstants.RankType
```

請確認：

```csharp
RankOverrideConstants.SchoolYearSemester == -1
RankOverrideConstants.ItemType == "學年/分項成績"
RankOverrideConstants.RankType == "班排名"
```

若常數值不符合，才修改常數定義；不要在 SQL 多處重複硬編碼。

---

# 四、item_name 處理原則

`ed01.txt` 沒有指定 `item_name` 的新寫入值，也沒有將 `item_name` 列入唯一值比對條件。

請依以下原則處理：

1. 不得將 `item_name` 放入本次 Upsert 唯一值比對條件。
2. 不得因本次修改任意改變既有 `item_name` 業務規則。
3. 若目前新增 SQL 必須寫入 `item_name`，沿用現有 `RankOverrideConstants.ItemName`。
4. 更新既有資料時，不更新 `item_name`。
5. 在 `匯入學年分項排名調整.md` 中記錄實際採用方式。

---

# 五、新增／更新唯一值比對條件

以下欄位全部相同時，視為同一筆資料：

```text
school_year
+ semester
+ ref_student_id
+ grade_year
+ item_type
+ rank_type
```

對應 SQL Join 條件必須為：

```sql
ro.school_year = ir.school_year
AND ro.semester = -1
AND ro.ref_student_id = ir.ref_student_id
AND ro.grade_year = ir.grade_year
AND ro.item_type = '學年/分項成績'
AND ro.rank_type = '班排名'
```

## 不可加入唯一值比對的欄位

以下欄位不得加入本次唯一值比對：

```text
rank
rank_name
matrix_count
extension
item_name
ref_exam_id
```

特別注意：

- `rank` 是要被更新的資料，不能作為比對鍵。
- `rank_name` 是班級名稱，不能作為比對鍵。
- `ref_exam_id` 固定寫入 `-1`，但依 `ed01.txt` 規格，不列入唯一值條件。
- `item_name` 不列入本次唯一值條件。

---

# 六、matched CTE 調整

目前 `matched` 使用的比對條件需重新調整。

請將：

```sql
LEFT JOIN rank_override ro
    ON ...
```

修改為依下列自然鍵比對：

```sql
LEFT JOIN rank_override ro
    ON ro.ref_student_id = ir.ref_student_id
    AND ro.school_year = ir.school_year
    AND ro.semester = -1
    AND ro.grade_year = ir.grade_year
    AND ro.item_type = '學年/分項成績'
    AND ro.rank_type = '班排名'
```

不得再使用：

```sql
ro.ref_exam_id IS NULL
```

因為新規格新增資料的：

```text
ref_exam_id = -1
```

但請注意，本次唯一值條件本身不包含 `ref_exam_id`。

---

# 七、更新既有資料規格

比對到相同資料時，只更新：

```text
rank
```

SQL 應接近：

```sql
UPDATE rank_override
SET rank = matched.rank
FROM matched
WHERE rank_override.id = matched.current_id
  AND matched.current_id IS NOT NULL
```

## 不可更新的欄位

本次 Update 不可修改：

```text
ref_student_id
school_year
semester
grade_year
item_type
ref_exam_id
item_name
rank_type
rank_name
matrix_count
extension
```

即使匯入資料中的班級名稱與資料庫不同，也不要在本次 Update 修改 `rank_name`。

本次調整目的為：

```text
唯一值相同 → 只更新學年學業排名 rank
```

---

# 八、新增資料規格

比對不到相同資料時，新增一筆 `rank_override`。

Insert 至少包含：

```text
ref_student_id
school_year
semester
grade_year
item_type
ref_exam_id
item_name
rank_type
rank_name
rank
extension
```

建議 SQL：

```sql
INSERT INTO rank_override (
    ref_student_id,
    school_year,
    semester,
    grade_year,
    item_type,
    ref_exam_id,
    item_name,
    rank_type,
    rank_name,
    rank,
    extension
)
SELECT
    matched.ref_student_id,
    matched.school_year,
    -1,
    matched.grade_year,
    '學年/分項成績',
    -1,
    '<沿用現有 item_name 固定值>',
    '班排名',
    matched.rank_name,
    matched.rank,
    NULL
FROM matched
WHERE matched.current_id IS NULL
```

## matrix_count

`matrix_count` 不列入 Insert 欄位。

不可寫入：

```sql
matrix_count = 0
```

讓資料庫維持：

```text
NULL
```

## extension

新增時必須改成：

```sql
NULL
```

不可再使用：

```sql
'{}'::jsonb
```

---

# 九、輸入 JSON 與 input_rows

目前 JSON 輸入資料可保留下列欄位：

```text
ref_student_id
school_year
grade_year
rank_name
rank
```

請確認序列化內容：

```csharp
List<object> jsonRows = records.Select(x => new
{
    ref_student_id = x.RefStudentID,
    school_year = x.SchoolYear,
    grade_year = x.GradeYear,
    rank_name = x.RankName,
    rank = x.Rank
}).Cast<object>().ToList();
```

`input_rows` 需正確轉型：

```sql
(obj->>'ref_student_id')::BIGINT
(obj->>'school_year')::INT
(obj->>'grade_year')::INT
obj->>'rank_name'
(obj->>'rank')::INT
```

保留 JSONB CTE 批次處理方式。

不得改成 `foreach` 逐筆執行 SQL。

---

# 十、重複資料風險處理

新唯一值條件為：

```text
ref_student_id
+ school_year
+ semester
+ grade_year
+ item_type
+ rank_type
```

請檢查資料庫是否可能已有多筆符合相同條件的資料。

若 `matched` 對到多筆 `rank_override`，可能造成：

- 同一匯入資料更新多筆資料。
- 更新筆數大於匯入筆數。
- CTE 結果不符合預期。

請調整或新增資料庫重複檢查，使用與本次 Upsert 相同的唯一值條件。

檢查 SQL 應類似：

```sql
SELECT
    ref_student_id,
    school_year,
    semester,
    grade_year,
    item_type,
    rank_type,
    COUNT(*) AS cnt
FROM rank_override
WHERE semester = -1
  AND item_type = '學年/分項成績'
  AND rank_type = '班排名'
GROUP BY
    ref_student_id,
    school_year,
    semester,
    grade_year,
    item_type,
    rank_type
HAVING COUNT(*) > 1;
```

若匯入涉及的自然鍵在資料庫已有多筆：

1. 不可自行刪除資料。
2. 不可任意挑一筆更新。
3. 應停止匯入或回報清楚錯誤訊息。
4. 錯誤訊息至少包含：

```text
學生系統編號
學號
學年度
年級
重複筆數
```

---

# 十一、FindDuplicateNaturalKeys 同步調整

目前：

```csharp
FindDuplicateNaturalKeys(
    List<SchoolYearEntryRankImportRecord> records)
```

其查詢條件與 Group By 必須同步改為新的唯一值規格。

## 必須加入

```text
grade_year
```

## 必須移除或停止使用

```text
ref_exam_id IS NULL
item_name 作為唯一值判斷
```

## importKeys

目前匯入比對鍵若為：

```csharp
RefStudentID + "_" + SchoolYear
```

必須改成至少包含：

```csharp
RefStudentID + "_" + SchoolYear + "_" + GradeYear
```

建議建立共用方法產生 Key，避免不同方法的自然鍵規則不一致，例如：

```csharp
private static string BuildNaturalKey(
    long studentID,
    int schoolYear,
    int gradeYear)
{
    return studentID + "_" + schoolYear + "_" + gradeYear;
}
```

`FindDuplicateNaturalKeys` 與 `UpsertSchoolYearEntryRanks` 必須使用相同的自然鍵定義。

---

# 十二、UpsertResult

保留：

```csharp
public class UpsertResult
{
    public int InsertedCount { get; set; }
    public int UpdatedCount { get; set; }
}
```

SQL 最後仍需回傳：

```text
inserted_count
updated_count
```

驗證：

```text
InsertedCount + UpdatedCount
```

原則上應等於本次有效匯入資料筆數。

若不同，需檢查是否有重複自然鍵或 SQL 比對到多筆資料。

---

# 十三、不可變更項目

本次修改限制：

1. 不可修改 `rank_override` 資料表結構。
2. 不可修改 Excel 欄位格式。
3. 不可改寫學生學號查詢流程。
4. 不可改成逐筆 SQL。
5. 不可新增、刪除其他排名類型資料。
6. 不可將 `matrix_count` 寫成 `0`。
7. 不可將 `extension` 寫成空 JSON 物件。
8. 不可在 Update 修改 `rank_name`。
9. 不可將 `ref_exam_id IS NULL` 繼續當作本次資料判斷條件。
10. 不可自行清理資料庫重複資料。
11. 不可更動匯出功能及學生資料項目主要架構。
12. 不可將 `grade_year` 從唯一值條件移除。

---

# 十四、建議測試資料

## 測試 1：新增資料

資料庫不存在相同自然鍵：

```text
StudentID = 1001
SchoolYear = 114
GradeYear = 1
Semester = -1
ItemType = 學年/分項成績
RankType = 班排名
```

匯入：

```text
班級名稱 = 高一甲
學年學業排名 = 5
```

預期：

- 新增一筆資料。
- `ref_exam_id = -1`。
- `extension IS NULL`。
- `matrix_count IS NULL`。
- `rank = 5`。
- `rank_name = 高一甲`。

## 測試 2：更新排名

資料庫已有相同自然鍵：

```text
rank = 5
rank_name = 高一甲
```

匯入：

```text
rank = 3
rank_name = 高一乙
```

預期：

- 不新增資料。
- 只將 `rank` 更新為 `3`。
- `rank_name` 仍維持資料庫原值「高一甲」。
- 其他欄位不變。

## 測試 3：不同年級

資料庫已有：

```text
StudentID = 1001
SchoolYear = 114
GradeYear = 1
```

匯入：

```text
StudentID = 1001
SchoolYear = 114
GradeYear = 2
```

預期：

- 因 `grade_year` 不同，新增一筆資料。
- 不更新原本一年級資料。

## 測試 4：不同學年度

資料庫已有：

```text
StudentID = 1001
SchoolYear = 113
GradeYear = 1
```

匯入：

```text
StudentID = 1001
SchoolYear = 114
GradeYear = 1
```

預期：

- 新增一筆 114 學年度資料。
- 113 學年度資料不變。

## 測試 5：資料庫重複自然鍵

資料庫已有兩筆：

```text
StudentID = 1001
SchoolYear = 114
Semester = -1
GradeYear = 1
ItemType = 學年/分項成績
RankType = 班排名
```

預期：

- 匯入前檢查出重複資料。
- 不執行 Upsert。
- 顯示清楚錯誤訊息。
- 不自動刪除或合併資料。

## 測試 6：extension 與 matrix_count

新增後查詢：

```sql
SELECT
    matrix_count,
    extension,
    ref_exam_id
FROM rank_override
WHERE ...;
```

預期：

```text
matrix_count = NULL
extension = NULL
ref_exam_id = -1
```

---

# 十五、完成後檢查

修改完成後確認：

- [ ] `UpsertSchoolYearEntryRanks` 可正常編譯。
- [ ] `semester` 固定為 `-1`。
- [ ] `ref_exam_id` 新增時寫入 `-1`。
- [ ] `extension` 新增時為 SQL `NULL`。
- [ ] `matrix_count` 沒有出現在 Insert 欄位。
- [ ] `grade_year` 已加入唯一值條件。
- [ ] `item_name` 沒有加入唯一值條件。
- [ ] `ref_exam_id` 沒有加入唯一值條件。
- [ ] Update 只修改 `rank`。
- [ ] Update 不修改 `rank_name`。
- [ ] 新增與更新筆數正確。
- [ ] 資料庫重複自然鍵可被偵測。
- [ ] 沒有使用 `foreach` 逐筆 SQL。
- [ ] 原本匯入驗證流程沒有被破壞。
- [ ] 原本匯出及資料項目功能沒有被修改。

---

# 十六、完成紀錄

建立：

```text
匯入學年分項排名調整.md
```

至少記錄：

1. 修改日期。
2. 修改檔案。
3. 修改方法。
4. 原本 Upsert 比對條件。
5. 新的唯一值條件。
6. Update 實際更新欄位。
7. Insert 實際寫入欄位。
8. `ref_exam_id` 從 `NULL` 改為 `-1` 的處理。
9. `extension` 從 `{}` 改為 `NULL` 的處理。
10. `matrix_count` 不寫入的處理。
11. `item_name` 實際保留方式。
12. 資料庫重複自然鍵處理方式。
13. 新增測試結果。
14. 更新測試結果。
15. 不同年級及不同學年度測試結果。
16. 編譯結果。
17. 已知限制或待確認事項。
