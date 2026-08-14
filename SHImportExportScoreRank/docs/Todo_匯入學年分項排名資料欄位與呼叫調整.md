## Goal

Update the student rank detail content so that:

1. The `lvData` list displays the **Grade Year** column.
2. The Grade Year column is placed between **School Year** and **Score Type**.
3. After importing School Year Entry Rank data, the `lvData` content is automatically refreshed.
4. The refresh mechanism must remain asynchronous and must **not freeze or block the UI**.
5. Do not change unrelated rank import, database, or detail-form behavior.

After completing the modification, record all changes in:

`匯入學年分項排名調整0814.md`

---

# 1. Modify `UCStudentRank.cs`

## 1.1 Add the Grade Year column

Current `lvData` column order:

```text
學年度
成績類型
成績項目
排名範圍
建立時間
建立方式
```

Change it to:

```text
學年度
成績年級
成績類型
成績項目
排名範圍
建立時間
建立方式
```

Update `SetupColumns()`.

Expected structure:

```csharp
private void SetupColumns()
{
    lvData.Columns.Clear();

    AddColumn("學年度", 80);
    AddColumn("成績年級", 80);
    AddColumn("成績類型", 90);
    AddColumn("成績項目", 90);
    AddColumn("排名範圍", 110);
    AddColumn("建立時間", 140);
    AddColumn("建立方式", 90);
}
```

---

# 2. Display `GradeYear` in `lvData`

The rank data source already contains:

```csharp
SchoolYearEntryRankRecord.GradeYear
```

The DAO already reads:

```text
rank_override.grade_year
```

Do not add another SQL query or duplicate the Grade Year data source.

Update `BindListView()` so `GradeYear` is inserted immediately after `SchoolYear`.

Expected logic:

```csharp
ListViewItem item = new ListViewItem(
    record.SchoolYear.HasValue
        ? record.SchoolYear.Value.ToString()
        : string.Empty);

item.SubItems.Add(
    record.GradeYear.HasValue
        ? record.GradeYear.Value.ToString()
        : string.Empty);

item.SubItems.Add(record.ScoreType ?? string.Empty);
item.SubItems.Add(record.ScoreItem ?? string.Empty);
item.SubItems.Add(record.RankType ?? string.Empty);
item.SubItems.Add(record.CreateTime ?? string.Empty);
item.SubItems.Add(record.CreateMethod ?? string.Empty);
```

Do not change the existing meaning of the other columns.

---

# 3. Fix the rank-data refresh invocation

`UCStudentRank` currently registers the refresh feature using:

```csharp
FISCA.Features.TryRegister(
    "RankDetailContent",
    ...
);
```

The import side currently uses a different Feature name:

```csharp
FISCA.Features.Invoke("SchoolYearEntryRankDetailContent");
```

These names are inconsistent.

Use the same Feature name on both sides:

```text
RankDetailContent
```

After a successful import, invoke:

```csharp
FISCA.Features.Invoke("RankDetailContent");
```

Do not create another refresh Feature unless required by the framework.

---

# 4. Make the refresh UI-safe and non-blocking

The refresh callback in `UCStudentRank.cs` must safely return execution to the UI thread without blocking the import process.

Update the registered Feature callback similar to:

```csharp
FISCA.Features.TryRegister(
    "RankDetailContent",
    x =>
    {
        if (this.IsDisposed)
            return;

        if (this.InvokeRequired)
        {
            this.BeginInvoke(new Action(() =>
            {
                if (!this.IsDisposed)
                    ReloadData();
            }));
        }
        else
        {
            ReloadData();
        }
    });
```

Important:

- Prefer `BeginInvoke()` instead of synchronous `Invoke()`.
- Do not perform database queries directly on the UI thread.
- Do not call `Thread.Sleep()`.
- Do not wait synchronously for the refresh to finish.
- Do not create refresh loops.

---

# 5. Preserve the existing asynchronous data-loading mechanism

Keep the existing flow:

```text
ReloadData()
    ↓
LoadDataAsync()
    ↓
BackgroundWorker.RunWorkerAsync()
    ↓
GetSchoolYearEntryRanksByStudentID()
    ↓
RunWorkerCompleted
    ↓
BindListView()
```

Database access must continue to run through the existing `BackgroundWorker`.

Do not replace it with a synchronous database query.

---

# 6. Preserve the existing busy-state protection

Keep the existing protection:

```csharp
if (_bgWorker.IsBusy)
{
    _isBusy = true;
    return;
}
```

and the corresponding reload behavior in `RunWorkerCompleted`.

This is required so that multiple refresh requests do not create multiple concurrent database queries.

Expected behavior:

```text
Refresh requested
    ↓
BackgroundWorker is idle
    ↓
Start one database query

Another refresh requested while query is running
    ↓
Set _isBusy = true
    ↓
Do not start another concurrent query

Current query completes
    ↓
If _isBusy == true
    ↓
Run one additional query to obtain the latest data
```

Do not remove this mechanism.

---

# 7. Refresh only after the import is completed

Do not refresh `lvData` for every imported row.

Incorrect:

```text
Import row 1 → Refresh
Import row 2 → Refresh
Import row 3 → Refresh
...
```

Required:

```text
Import all rows
    ↓
Database insert/update completed
    ↓
Invoke RankDetailContent once
    ↓
UCStudentRank reloads the latest data
```

The refresh invocation should occur only after the import/write operation has successfully completed.

---

# 8. Do not modify unrelated behavior

Do not unnecessarily change:

- Rank import validation
- `rank_override` insert/update rules
- Rank key matching rules
- `extension` JSONB handling
- `frmStudentRankDetail`
- Existing School Year / Score Type / Score Item / Rank Type values
- Existing `PrimaryKey` switching behavior
- Existing edit/detail behavior
- Existing XML or other data structures
- Unrelated UI controls

Keep the modification focused on:

```text
1. Grade Year display
2. RankDetailContent refresh invocation
3. Non-blocking UI-safe refresh
```

---

# 9. Verification

After modification, verify the following.

## Test A - Grade Year display

Prepare data such as:

```text
school_year = 114
grade_year = 2
```

Expected `lvData`:

```text
學年度 | 成績年級 | 成績類型 | 成績項目 | 排名範圍 | 建立時間 | 建立方式
114    | 2        | ...
```

Confirm that all subsequent columns remain correctly aligned.

## Test B - Import refresh

1. Open a student's `排名資料` DetailContent.
2. Keep `lvData` visible.
3. Import a new School Year Entry Rank record for the same student.
4. Complete the import.
5. Confirm that `lvData` automatically reloads.
6. Confirm that switching to another student and back is not required.

## Test C - Update existing data

1. Keep the student's rank detail page open.
2. Import data that updates an existing `rank_override` record.
3. Confirm the updated value appears after import.

## Test D - UI responsiveness

During and immediately after import:

- The main application window must remain responsive.
- No long UI freeze should occur.
- No synchronous database query should execute on the UI thread.
- No cross-thread UI exception should occur.

## Test E - Repeated refresh requests

Trigger multiple refresh requests while data is already loading.

Confirm:

- No `BackgroundWorker is currently busy` exception occurs.
- No duplicate concurrent queries are created.
- The final `lvData` shows the latest database state.

---

# 10. Completion Record

After all modifications and verification are completed, create/update:

`匯入學年分項排名調整0814.md`

Include:

- Modified files
- Added `成績年級` column
- `GradeYear` binding changes
- Original refresh Feature-name mismatch
- Final unified Feature name: `RankDetailContent`
- Import-completion refresh changes
- `BeginInvoke()` UI-thread handling
- Existing `BackgroundWorker` behavior retained
- `_isBusy` protection retained
- Verification results
- Any issues discovered during implementation

Do not modify unrelated code.