## Goal

Add the **Grade Year (`成績年級`)** field to the Student Rank Detail form.

The new field must be displayed **before `成績類型` (Score Type)**.

The Grade Year already exists in the rank data source as:

```text
rank_override.grade_year
```

The modification must pass `GradeYear` correctly through the complete data flow instead of only adding a UI label.

After completing the modification, record all changes in:

```text
匯入學年分項排名調整0814.md
```

---

# 1. Expected UI

Current header information in `frmStudentRankDetail` includes:

```text
學年度
成績類型
成績項目
建立方式
建立時間
批次名稱
```

Change it so that `成績年級` is inserted before `成績類型`:

```text
學年度
成績年級
成績類型
成績項目
建立方式
建立時間
批次名稱
```

Do not change the meaning of the existing fields.

---

# 2. Modify `UCStudentRank.cs`

When opening `frmStudentRankDetail`, the selected record currently creates a `StudentRankDetailContext`.

The current data flow already has:

```csharp
record.GradeYear
```

Add `GradeYear` to the context.

Expected logic:

```csharp
StudentRankDetailContext context = new StudentRankDetailContext
{
    StudentId = this.PrimaryKey,
    SchoolYear = record.SchoolYear,
    GradeYear = record.GradeYear,
    ScoreType = record.ScoreType,
    ScoreItem = record.ScoreItem,
    CreateTime = record.CreateTime,
    CreateMethod = record.CreateMethod
};
```

Do not query Grade Year again from `UCStudentRank`.

Use the existing:

```csharp
SchoolYearEntryRankRecord.GradeYear
```

---

# 3. Modify `StudentRankDetailContext`

Locate the definition of:

```csharp
StudentRankDetailContext
```

Add a nullable Grade Year property:

```csharp
public int? GradeYear { get; set; }
```

The property must remain nullable because historical or incomplete data may not contain `grade_year`.

---

# 4. Modify `StudentRankDetailData`

Locate the definition of:

```csharp
StudentRankDetailData
```

Add:

```csharp
public int? GradeYear { get; set; }
```

This field will be used by `frmStudentRankDetail` when binding the header.

Do not store Grade Year only as display text.

Keep it as:

```csharp
int?
```

to match the database/model semantics.

---

# 5. Modify `RankOverrideDataAccess.GetStudentRankDetail()`

The current detail query does not fully use `grade_year`.

Add:

```sql
ro.grade_year
```

to the required SELECT data where appropriate.

Example:

```sql
SELECT
    ro.id,
    ro.school_year,
    ro.grade_year,
    ro.item_name,
    ro.rank_name,
    ro.rank,
    ro.rank_type,
    ro.matrix_count,
    ...
```

Preserve the existing query structure.

---

# 6. Add Grade Year to the detail query condition

This is important.

The selected row in `UCStudentRank` represents a specific:

```text
Student
+ School Year
+ Grade Year
+ Score Type
+ Score Item
```

Therefore, `GetStudentRankDetail()` must use `GradeYear` as part of the query condition when it is available.

Expected concept:

```sql
AND ro.grade_year = ...
```

Use the existing parameterized/query-building style of the project.

Do not concatenate unsafe user input directly into SQL.

If `context.GradeYear` is nullable, preserve the existing project's handling style for nullable query conditions.

The purpose is to prevent rank data from different Grade Years from being mixed together when they share the same School Year and rank category.

---

# 7. Map Grade Year into `StudentRankDetailData`

When building the result from the query, assign:

```csharp
data.GradeYear = ...
```

using the existing integer conversion/helper method.

For example, if the project already uses:

```csharp
ToNullableInt(...)
```

reuse that helper rather than creating duplicate conversion logic.

Expected concept:

```csharp
GradeYear = ToNullableInt(row["grade_year"])
```

Use the actual mapping structure already used by `RankOverrideDataAccess`.

---

# 8. Modify `frmStudentRankDetail.ApplyDetailHeaderFromContext()`

The current fallback logic copies header information from `_context` when the database query fails or returns fallback data.

Add:

```csharp
data.GradeYear = _context.GradeYear;
```

Expected section:

```csharp
data.SchoolYear = _context.SchoolYear;
data.GradeYear = _context.GradeYear;
data.ScoreType = _context.ScoreType ?? string.Empty;
data.ScoreItem = _context.ScoreItem ?? string.Empty;
data.CreateType = _context.CreateMethod ?? string.Empty;
data.CreateTime = _context.CreateTime ?? string.Empty;
```

This ensures the Grade Year header can still be displayed consistently when using context fallback data.

---

# 9. Modify `frmStudentRankDetail.BindHeader()`

Add Grade Year binding before Score Type.

Expected logic:

```csharp
lblSchoolYearValue.Text = _detailData.SchoolYear.HasValue
    ? _detailData.SchoolYear.Value.ToString()
    : string.Empty;

lblGradeYearValue.Text = _detailData.GradeYear.HasValue
    ? _detailData.GradeYear.Value.ToString()
    : string.Empty;

lblScoreTypeValue.Text = _detailData.ScoreType ?? string.Empty;
```

Do not convert a missing Grade Year to `0`.

If no value exists, display:

```text
(empty)
```

---

# 10. Modify `frmStudentRankDetail.Designer.cs`

Add the required controls for Grade Year.

Use naming consistent with the existing controls:

```text
lblGradeYear
lblGradeYearValue
```

Display text:

```text
成績年級
```

Place this field **before `成績類型`** in the header area.

Adjust the positions of existing controls only as necessary to maintain a clean layout.

Do not unnecessarily redesign the form.

---

# 11. Do Not Add Grade Year to `dgvRankDetail`

The requested `成績年級` is header-level information for the selected rank data group.

Do not add another Grade Year column to:

```csharp
dgvRankDetail
```

unless the existing architecture specifically requires it.

Keep the existing detail grid columns unchanged:

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

---

# 12. Expected Data Flow

After modification, the complete data flow should be:

```text
rank_override.grade_year
        ↓
GetSchoolYearEntryRanksByStudentID()
        ↓
SchoolYearEntryRankRecord.GradeYear
        ↓
UCStudentRank
        ↓
StudentRankDetailContext.GradeYear
        ↓
GetStudentRankDetail()
        ↓
grade_year query condition
        ↓
StudentRankDetailData.GradeYear
        ↓
frmStudentRankDetail.BindHeader()
        ↓
lblGradeYearValue
```

Do not break this chain by re-querying or manually calculating Grade Year in the UI.

---

# 13. Preserve Existing Behavior

Do not modify unrelated behavior, including:

- Rank import rules
- Rank insert/update keys
- `rank_override.extension`
- `matrix_count`
- `RankDisplay`
- PR calculation/display
- Percentage calculation/display
- Rank Type
- Rank Name
- Score Category
- Score Item
- Score Type
- Create Time
- Create Method
- Batch Name
- Existing `BackgroundWorker` refresh logic
- Existing rank import refresh mechanism
- Existing save restrictions in `frmStudentRankDetail`

Keep the change focused on adding and correctly propagating:

```text
GradeYear / 成績年級
```

---

# 14. Verification

## Test 1 - Grade Year display

Prepare rank data:

```text
school_year = 114
grade_year = 2
```

Open the student rank detail.

Expected header:

```text
學年度：114
成績年級：2
成績類型：學年
成績項目：分項
```

---

## Test 2 - Correct detail query

Prepare two records for the same student and School Year but different Grade Years:

```text
Student A
School Year = 114
Grade Year = 1
```

and:

```text
Student A
School Year = 114
Grade Year = 2
```

Open the Grade Year 1 record.

Confirm that only Grade Year 1 detail data is displayed.

Open the Grade Year 2 record.

Confirm that only Grade Year 2 detail data is displayed.

Data from the two Grade Years must not be mixed.

---

## Test 3 - Null Grade Year

Test an older/incomplete record where:

```text
grade_year = NULL
```

Expected behavior:

- No exception.
- `成績年級` displays blank.
- Do not display `0`.
- Other detail information continues to load normally.

---

## Test 4 - Existing detail fields

Confirm that these fields still display correctly:

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

No existing DataGrid column may shift to the wrong data source.

---

# 15. Completion Record

After implementation and verification, update:

```text
匯入學年分項排名調整0814.md
```

Record:

- Modified files
- Added `GradeYear` to `StudentRankDetailContext`
- Added `GradeYear` to `StudentRankDetailData`
- Passed `record.GradeYear` from `UCStudentRank`
- Added `grade_year` to the detail query
- Added Grade Year to the detail query condition
- Added Grade Year result mapping
- Added `成績年級` to `frmStudentRankDetail`
- Added `lblGradeYear` / `lblGradeYearValue`
- Updated `BindHeader()`
- Updated context fallback handling
- Verification results
- Any compatibility or data issues discovered

Do not modify unrelated code.