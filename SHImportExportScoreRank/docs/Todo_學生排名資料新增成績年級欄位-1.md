## Goal

Add a new **Grade Year (`成績年級`)** column to `frmStudentRankDetail`.

The new column must be added to `dgvRankDetail` **before the existing `成績類別` column**.

After completing the modification, record all changes in:

```text
匯入學年分項排名調整0814.md
```

---

# 1. Current Status

File:

```text
frmStudentRankDetail.cs
```

The current DataGridView column order is:

```text
成績類別
排名方式
排名分數
排名範圍
母群
排名
PR值
百分比
```

The current code starts with:

```csharp
AddTextColumn("colScoreCategory", "成績類別", "ScoreCategory", 90);
```

The form header already supports Grade Year through:

```csharp
_detailData.GradeYear
```

and:

```csharp
lblGradeYearValue
```

Do not duplicate or redesign the existing Grade Year header.

This task specifically adds Grade Year to:

```text
dgvRankDetail
```

---

# 2. Expected DataGridView Column Order

Change the column order to:

```text
成績年級
成績類別
排名方式
排名分數
排名範圍
母群
排名
PR值
百分比
```

`成績年級` must be the first column.

---

# 3. Modify `SetupGridColumns()`

Update:

```csharp
private void SetupGridColumns()
```

Add the Grade Year column before `colScoreCategory`.

Expected structure:

```csharp
private void SetupGridColumns()
{
    dgvRankDetail.AutoGenerateColumns = false;
    dgvRankDetail.Columns.Clear();

    AddTextColumn("colGradeYear", "成績年級", "GradeYear", 70);
    AddTextColumn("colScoreCategory", "成績類別", "ScoreCategory", 90);
    AddTextColumn("colRankMethod", "排名方式", "RankMethod", 70);
    AddTextColumn("colScore", "排名分數", "Score", 70);
    AddTextColumn("colRankType", "排名範圍", "RankType", 80);
    AddTextColumn("colRankName", "母群", "RankName", 90);
    AddTextColumn("colRank", "排名", "RankDisplay", 70);
    AddTextColumn("colPR", "PR值", "PR", 60);
    AddTextColumn("colPercentage", "百分比", "Percentage", 70);
}
```

---

# 4. Check `StudentRankDetailItem`

Locate:

```csharp
StudentRankDetailItem
```

Check whether it already contains:

```csharp
public int? GradeYear { get; set; }
```

or an equivalent Grade Year property.

If it does not exist, add:

```csharp
public int? GradeYear { get; set; }
```

Prefer:

```csharp
int?
```

instead of `string`, because the source field:

```text
rank_override.grade_year
```

is numeric and may be NULL.

Do not display a missing Grade Year as `0`.

---

# 5. Populate Grade Year in Detail Data

Inspect:

```csharp
RankOverrideDataAccess.GetStudentRankDetail(...)
```

The SQL/detail mapping must populate each:

```csharp
StudentRankDetailItem
```

with the correct Grade Year.

The data source should be:

```text
rank_override.grade_year
```

If `ro.grade_year` is already included in the SQL query, reuse it.

If not, add it to the SELECT list.

Example:

```sql
SELECT
    ro.id,
    ro.school_year,
    ro.grade_year,
    ro.item_name,
    ro.rank_type,
    ro.rank_name,
    ro.rank,
    ro.matrix_count,
    ...
```

Do not create a separate database query only for Grade Year.

---

# 6. Map `grade_year` to `StudentRankDetailItem.GradeYear`

When constructing each detail item, map:

```text
ro.grade_year
```

to:

```csharp
item.GradeYear
```

Reuse the existing nullable integer conversion helper if available, such as:

```csharp
ToNullableInt(...)
```

Expected concept:

```csharp
GradeYear = ToNullableInt(row["grade_year"])
```

Use the actual object initialization style already used by the project.

---

# 7. Keep Grade Year Consistent with the Selected Rank Group

The selected record already has Grade Year in the detail context:

```csharp
_context.GradeYear
```

The detail query should already use Grade Year to identify the correct rank data group.

Verify that:

```text
StudentRankDetailContext.GradeYear
```

and:

```text
StudentRankDetailItem.GradeYear
```

represent the same source:

```text
rank_override.grade_year
```

Do not calculate Grade Year from School Year.

Do not calculate Grade Year from the student's current class.

Use the Grade Year stored in the rank record.

---

# 8. Expected Data Flow

The expected data flow is:

```text
rank_override.grade_year
        ↓
GetStudentRankDetail()
        ↓
StudentRankDetailItem.GradeYear
        ↓
dgvRankDetail DataSource
        ↓
colGradeYear
        ↓
成績年級
```

The `DataPropertyName` must therefore match:

```text
GradeYear
```

---

# 9. Do Not Change Existing Columns

Do not change the existing bindings for:

```text
成績類別 -> ScoreCategory
排名方式 -> RankMethod
排名分數 -> Score
排名範圍 -> RankType
母群 -> RankName
排名 -> RankDisplay
PR值 -> PR
百分比 -> Percentage
```

Only insert the new Grade Year column before them.

---

# 10. Preserve Existing `RankDisplay` Logic

Do not modify the existing `RankDisplay` rule.

Existing expected behavior must remain:

```text
Only Rank exists:
RankDisplay = Rank
```

```text
Rank and MatrixCount both exist:
RankDisplay = Rank / MatrixCount
```

The Grade Year modification must not affect this logic.

---

# 11. Preserve Existing Header Grade Year

The current `frmStudentRankDetail.cs` already includes:

```csharp
data.GradeYear = _context.GradeYear;
```

and:

```csharp
lblGradeYearValue.Text = _detailData.GradeYear.HasValue
    ? _detailData.GradeYear.Value.ToString()
    : string.Empty;
```

Keep this behavior.

Do not remove the existing Grade Year header just because Grade Year is also being added to the grid.

The final form may show Grade Year in both:

```text
Header:
成績年級
```

and:

```text
dgvRankDetail:
成績年級
```

This is intentional for this task.

---

# 12. Null Handling

If:

```text
grade_year = NULL
```

display:

```text
(empty)
```

Do not display:

```text
0
```

Do not throw an exception.

---

# 13. Do Not Modify Unrelated Logic

Do not unnecessarily modify:

- `SchoolYear`
- `ScoreType`
- `ScoreItem`
- `CreateType`
- `CreateTime`
- `BatchName`
- Rank import logic
- Rank insert/update logic
- `rank_override.extension`
- `matrix_count`
- PR
- Percentage
- `RankDisplay`
- Import refresh mechanism
- `UCStudentRank`
- Existing BackgroundWorker logic
- Existing form save restrictions

Keep this change focused on:

```text
StudentRankDetailItem.GradeYear
+
dgvRankDetail colGradeYear
```

---

# 14. Verification

## Test 1 - Normal Grade Year

Prepare detail data:

```text
school_year = 114
grade_year = 2
```

Expected grid:

```text
成績年級 | 成績類別 | 排名方式 | 排名分數 | 排名範圍 | 母群 | 排名 | PR值 | 百分比
2        | ...
```

---

## Test 2 - Multiple Detail Rows

If the selected ranking group contains multiple rows, verify that each row displays the correct:

```text
GradeYear
```

Example:

```text
2 | 學業 | ...
2 | 國文 | ...
2 | 英文 | ...
```

Do not leave the Grade Year cells blank when `grade_year` exists.

---

## Test 3 - NULL Grade Year

For:

```text
grade_year = NULL
```

Expected:

```text
成績年級 = blank
```

No exception should occur.

---

## Test 4 - Column Alignment

Verify the final order is exactly:

```text
成績年級
成績類別
排名方式
排名分數
排名範圍
母群
排名
PR值
百分比
```

Ensure that adding the new first column does not cause any existing data to appear under the wrong header.

---

## Test 5 - Existing Rank Display

Verify:

```text
rank = 5
matrix_count = NULL
```

Expected:

```text
排名 = 5
```

Verify:

```text
rank = 5
matrix_count = 30
```

Expected:

```text
排名 = 5 / 30
```

The new Grade Year column must not affect this behavior.

---

# 15. Completion Record

After implementation and testing, update:

```text
匯入學年分項排名調整0814.md
```

Record:

- Modified files
- Added `成績年級` as the first `dgvRankDetail` column
- Added/confirmed `StudentRankDetailItem.GradeYear`
- Added/confirmed `ro.grade_year` in the detail query
- Added/confirmed Grade Year data mapping
- Confirmed nullable Grade Year handling
- Confirmed existing header Grade Year remains unchanged
- Confirmed existing rank display logic remains unchanged
- Test results
- Any issues discovered during implementation

Do not modify unrelated code.