
## Target File

* `Program.cs`
* Feature: `期末成績通知單(固定排名)`

## Objective

Fix the semester text evaluation display logic so that the existing Word merge field always shows the text evaluation for the semester selected by the user.

Required behavior:

* When the user selects Semester 1, display the Semester 1 text evaluation.
* When the user selects Semester 2, display the Semester 2 text evaluation.
* Prevent Semester 1 text evaluation from overwriting Semester 2 text evaluation when generating a Semester 2 report.
* The result must remain correct when printing one student, several students, or approximately 200 students.

The existing merge field must continue to work:

```text
綜合表現：評語(含社團、競賽、服務學習、禮節等綜合表現及建議)
```

Do not require changes to the existing Word report template unless absolutely necessary.

---

## Current Problem

The report loads text evaluation data from:

```csharp
stuRec.SemesterMoralScoreList
```

The actual text evaluation values are read from:

```csharp
info.Detail.SelectNodes("TextScore/Morality")
```

The current semester and previous semester processing blocks both write values into the same legacy merge field:

```csharp
row["綜合表現：" + face] = each.InnerText;
```

When printing Semester 2, the program processes both:

* Semester 1 text evaluation.
* Semester 2 text evaluation.

Because both semesters write to the same `綜合表現：Face` field, the final displayed value may depend on the order of records in `SemesterMoralScoreList`.

Possible incorrect result:

```text
Semester 2 data is written first.
Semester 1 data is processed later.
Semester 1 overwrites Semester 2.
```

This may appear correct when testing a few students but fail for some students when printing a large batch.

---

## Required Modification

### 1. Preserve the Existing Main Structure

Do not redesign the report generation process.

Do not make unnecessary changes to:

* `BackgroundWorker`
* Student loading
* Score calculation
* Attendance calculation
* Reward calculation
* Ranking calculation
* Word Mail Merge
* Report template loading
* Existing `DataTable` structure
* Existing text evaluation mapping table
* Existing teacher comment fields
* Existing Semester 1–10 text evaluation fields

Only adjust the semester text evaluation assignment logic.

---

### 2. Current-Semester Rule

The legacy merge fields beginning with:

```text
綜合表現：
```

must only be written by the record matching the selected report semester:

```csharp
info.SchoolYear.ToString() == conf.SchoolYear
info.Semester.ToString() == conf.Semester
```

Expected behavior:

```text
conf.Semester == "1"
    → 綜合表現：Face receives Semester 1 content

conf.Semester == "2"
    → 綜合表現：Face receives Semester 2 content
```

Keep the current-semester assignment:

```csharp
row["綜合表現：" + face] = each.InnerText;
```

only inside the current selected semester block.

Before assigning, retain or add a safe column check:

```csharp
string columnName = "綜合表現：" + face;

if (row.Table.Columns.Contains(columnName))
{
    row[columnName] = each.InnerText;
}
```

---

### 3. Previous-Semester Rule

When generating a Semester 2 report, the program may continue loading Semester 1 text evaluation into the existing dedicated previous-semester fields:

```text
上學期文字評量名稱1～上學期文字評量名稱10
上學期文字評量1～上學期文字評量10
上學期導師評語
```

However, the Semester 1 processing block must not write to:

```csharp
row["綜合表現：" + face]
```

Remove or disable this assignment from the previous-semester block:

```csharp
row["綜合表現：" + face] = each.InnerText;
```

The previous-semester block should only populate:

```csharp
row["上學期導師評語"]
row["上學期文字評量名稱" + i]
row["上學期文字評量" + i]
```

This prevents Semester 1 values from overwriting the Semester 2 legacy merge fields.

---

### 4. Keep Current-Semester Text Evaluation Fields

Continue populating the current-semester dedicated fields:

```text
文字評量名稱1～文字評量名稱10
文字評量1～文字評量10
導師評語
```

The current selected semester block should continue populating:

```csharp
row["導師評語"] = info.SupervisedByComment;
row["文字評量名稱" + i] = faceList[i - 1];
row["文字評量" + i] = commentList[i - 1];
```

Do not change the meaning or output format of these fields.

---

### 5. Preserve Text Evaluation Mapping Validation

Continue validating each `Face` against:

```csharp
SmartSchool.Customization.Data.SystemInformation
    .Fields["文字評量對照表"]
```

Retain the existing validation logic using:

```csharp
Content/Morality[@Face='...']
```

Do not change the `Face` name, text content, or mapping structure.

---

### 6. Limit Text Evaluation Output to 10 Items

Continue supporting a maximum of 10 text evaluation items.

Use a safe loop condition such as:

```csharp
for (int i = 1; i <= faceList.Count && i <= 10; i++)
{
    ...
}
```

Do not allow an index outside the existing fields:

```text
文字評量1～10
上學期文字評量1～10
```

---

## Expected Logic

### Selected Semester 1

```text
Semester 1 record
    ├─ 導師評語
    ├─ 文字評量名稱1～10
    ├─ 文字評量1～10
    └─ 綜合表現：Face
```

The output must display Semester 1 text evaluation.

### Selected Semester 2

```text
Semester 1 record
    ├─ 上學期導師評語
    ├─ 上學期文字評量名稱1～10
    └─ 上學期文字評量1～10

Semester 2 record
    ├─ 導師評語
    ├─ 文字評量名稱1～10
    ├─ 文字評量1～10
    └─ 綜合表現：Face
```

The Semester 1 record must not write to `綜合表現：Face`.

The output must display Semester 2 text evaluation in:

```text
綜合表現：評語(含社團、競賽、服務學習、禮節等綜合表現及建議)
```

---

## Important Restrictions

* Do not change the existing report merge field name.
* Do not change the Word template unless the current template is confirmed to use an incorrect field.
* Do not remove support for previous-semester text evaluation fields.
* Do not alter `SemesterMoralScoreList` loading.
* Do not alter `FillSemesterMoralScore`.
* Do not change teacher comment behavior unrelated to this issue.
* Do not change score, ranking, attendance, reward, credit, or student data logic.
* Do not depend on the order of records in `SemesterMoralScoreList`.
* Do not solve the issue by sorting alone.
* Do not introduce shared state between students.
* Each student's `DataRow` must continue to contain only that student's data.

---

## Verification

Test the following cases.

### Test 1: Single Student, Semester 1

Student data contains different Semester 1 and Semester 2 text evaluations.

Generate a Semester 1 report.

Expected:

```text
綜合表現：評語(...) = Semester 1 content
```

### Test 2: Single Student, Semester 2

Generate a Semester 2 report.

Expected:

```text
綜合表現：評語(...) = Semester 2 content
```

Also verify:

```text
上學期文字評量1～10 = Semester 1 content
文字評量1～10 = Semester 2 content
```

### Test 3: Reversed Record Order

Test a student whose `SemesterMoralScoreList` order is:

```text
Semester 2
Semester 1
```

Expected:

* Semester 2 report still displays Semester 2 content.
* Semester 1 must not overwrite the `綜合表現：Face` field.

### Test 4: Normal Record Order

Test a student whose record order is:

```text
Semester 1
Semester 2
```

Expected:

* Semester 2 report displays Semester 2 content.

### Test 5: Semester 2 Missing the Face

Semester 1 contains:

```text
評語(含社團、競賽、服務學習、禮節等綜合表現及建議)
```

Semester 2 does not contain that Face.

Expected:

* Semester 2 report must not display the Semester 1 value in the current-semester `綜合表現：Face` field.
* The current-semester field should remain blank unless Semester 2 contains a value.

### Test 6: Large Batch

Generate a report for approximately 200 students.

Verify:

* Each student displays their own text evaluation.
* No student receives another student's content.
* Semester 1 reports display Semester 1 values.
* Semester 2 reports display Semester 2 values.
* Record order does not affect the result.
* Page breaks and Word Mail Merge do not shift student data.

### Test 7: Existing Fields

Confirm that these existing fields still work:

```text
導師評語
上學期導師評語
文字評量名稱1～10
文字評量1～10
上學期文字評量名稱1～10
上學期文字評量1～10
綜合表現：Face
```

---

## Optional Debug Verification

Before Word Mail Merge, temporarily export the `DataTable` for verification:

```csharp
table.TableName = "test";
table.WriteXml(
    Path.Combine(
        Application.StartupPath,
        "期末成績通知單_文字評量_debug.xml"
    ),
    XmlWriteMode.WriteSchema
);
```

Check the affected students' fields:

```xml
<姓名>...</姓名>
<學期>...</學期>
<導師評語>...</導師評語>
<上學期導師評語>...</上學期導師評語>
<文字評量1>...</文字評量1>
<上學期文字評量1>...</上學期文字評量1>
<綜合表現_x003A_評語...>...</綜合表現_x003A_評語...>
```

Remove or comment out temporary debug output after verification.

---

## Completion Documentation

After completing and testing the modification, create or update:

```text
期末成績單(固定排名)調整0803.md
```

Document the following:

1. Original issue and reproduction conditions.
2. Root cause:

   * Previous and current semesters wrote to the same `綜合表現：Face` field.
   * Output depended on the order of `SemesterMoralScoreList`.
3. Modified file and code location.
4. Exact logic changed.
5. Fields intentionally preserved.
6. Single-student test results.
7. Semester 1 and Semester 2 test results.
8. Reversed-record-order test result.
9. Approximately 200-student batch test result.
10. Confirmation that the main report structure and unrelated logic were not changed.
