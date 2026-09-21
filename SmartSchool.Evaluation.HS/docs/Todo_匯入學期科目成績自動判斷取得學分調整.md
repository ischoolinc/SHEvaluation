# Goal

Update `ImportSemesterSubjectScore.cs`.

When the user selects **「自動判斷取得學分」**, change the passing-score priority as follows:

1. If the subject has a valid **「修課及格標準」**, use it first.
2. If **「修課及格標準」** is empty or does not contain a valid numeric value, use the existing student passing score:

   ```csharp
   _StudentPassScore[studentRec][gy]
   ```
3. Do not change any other existing behavior.

After completing the modification, record the changes in:

```text
匯入學期科目成績調整0921.md
```

---

# Target File

```text
ImportSemesterSubjectScore.cs
```

---

# Current Behavior

When **「自動判斷取得學分」** is selected, the current code determines the passing score only from the student's existing score calculation rule:

```csharp
_StudentPassScore[studentRec][gy]
```

The subject field:

```text
修課及格標準
```

is currently:

* validated during import
* stored into the subject score XML
* updated when importing existing subject scores

However, it is **not currently used when automatically determining whether the student receives credit**.

---

# Required Behavior

Change only the automatic credit determination logic.

Use this priority:

```text
修課及格標準有有效數值
    ↓
優先使用修課及格標準

修課及格標準沒有值
或不是有效數值
    ↓
使用原本學生及格標準
_StudentPassScore[studentRec][gy]
```

The final logic should conceptually be:

```csharp
decimal passScore = _StudentPassScore[studentRec][gy];

decimal coursePassScore;
if (decimal.TryParse(
    subjectElement.GetAttribute("修課及格標準"),
    out coursePassScore))
{
    passScore = coursePassScore;
}
```

Then change the existing comparison from:

```csharp
maxScore >= _StudentPassScore[studentRec][gy]
```

to:

```csharp
maxScore >= passScore
```

---

# Important: Keep Existing Credit Logic

Do not change the existing highest-score calculation.

Continue using the highest valid score from:

```text
原始成績
學年調整成績
擇優採計成績
補考成績
重修成績
```

Existing behavior must remain:

```text
不需評分 = 是
    → 是否取得學分 = 是

otherwise

maxScore >= passScore
    → 是否取得學分 = 是

maxScore < passScore
    → 是否取得學分 = 否
```

Only the source of `passScore` is being adjusted.

---

# Modification Locations

There are three automatic credit calculation paths that must use the same rule.

## 1. Update Existing Subject Score

Locate the existing:

```csharp
if (autoCheckPass.Checked)
```

logic that operates on:

```csharp
score.Detail
```

Current comparison uses:

```csharp
_StudentPassScore[studentRec][gy]
```

Change it so that:

```csharp
decimal passScore = _StudentPassScore[studentRec][gy];

decimal coursePassScore;
if (decimal.TryParse(
    score.Detail.GetAttribute("修課及格標準"),
    out coursePassScore))
{
    passScore = coursePassScore;
}
```

Then use:

```csharp
maxScore >= passScore
```

Important:

If this import includes a new value for `修課及格標準`, the code already updates:

```csharp
score.Detail.SetAttribute("修課及格標準", value);
```

before automatic credit determination.

Therefore, automatic credit calculation should use the newly imported value.

Do not change that existing update order.

---

## 2. Add New Subject to Existing Semester

Locate the `autoCheckPass.Checked` logic when creating:

```csharp
XmlElement newScore
```

for a new subject being added to an existing semester.

After calculating `maxScore`, determine:

```csharp
decimal passScore = _StudentPassScore[studentRec][gy];

decimal coursePassScore;
if (decimal.TryParse(
    newScore.GetAttribute("修課及格標準"),
    out coursePassScore))
{
    passScore = coursePassScore;
}
```

Then use:

```csharp
maxScore >= passScore
```

Do not modify other new-subject creation logic.

---

## 3. Add Subject Score to a Completely New Semester

Locate the other `autoCheckPass.Checked` logic used while creating `newScore` inside:

```text
處理新增成績學期
```

Apply the same passing-score priority:

```csharp
decimal passScore = _StudentPassScore[studentRec][gy];

decimal coursePassScore;
if (decimal.TryParse(
    newScore.GetAttribute("修課及格標準"),
    out coursePassScore))
{
    passScore = coursePassScore;
}
```

Then use:

```csharp
maxScore >= passScore
```

Do not modify other insert behavior.

---

# Examples

Student original passing score:

```text
60
```

### Case 1

```text
最高成績 = 65
修課及格標準 = 空白
```

Use:

```text
60
```

Result:

```text
取得學分 = 是
```

---

### Case 2

```text
最高成績 = 65
修課及格標準 = 70
```

Use:

```text
70
```

Result:

```text
取得學分 = 否
```

---

### Case 3

```text
最高成績 = 55
修課及格標準 = 50
```

Use:

```text
50
```

Result:

```text
取得學分 = 是
```

---

### Case 4

```text
最高成績 = 55
修課及格標準 = 空白
學生原本及格標準 = 60
```

Result:

```text
取得學分 = 否
```

---

### Case 5

```text
不需評分 = 是
```

Regardless of passing score or subject passing score:

```text
取得學分 = 是
```

Keep this existing behavior unchanged.

---

# Do Not Change

Do not change any unrelated logic.

Especially do not change:

* `ValidateRow` validation behavior
* `修課及格標準` numeric validation
* `修課補考標準`
* `修課直接指定總成績`
* student score calculation rule loading
* student category matching
* grade-specific passing score lookup
* fallback passing score of `60`
* `_StudentPassScore` structure
* highest-score calculation
* score priority / max-score behavior
* `不需評分` handling
* manual credit determination
* `取得學分` manual import behavior
* semester grade-year handling
* existing subject matching
* insert/update determination
* XML structure
* logging
* threading
* package size
* database/API update behavior

Do not refactor unrelated code.

Do not extract or redesign the existing score calculation logic unless absolutely required.

Prefer the smallest possible code change.

---

# Verification

Verify all three automatic-credit paths.

Test at least:

```text
1. Existing subject + 修課及格標準 has value
2. Existing subject + 修課及格標準 empty
3. New subject in existing semester + 修課及格標準 has value
4. New subject in existing semester + 修課及格標準 empty
5. Completely new semester + 修課及格標準 has value
6. Completely new semester + 修課及格標準 empty
7. 不需評分 = 是
8. 手動判斷取得學分
```

Confirm that manual credit determination is completely unaffected.

---

# Completion Record

After implementation, create/update:

```text
匯入學期科目成績調整0921.md
```

Record:

* modified file
* modified locations
* original behavior
* new passing-score priority
* confirmation that all three automatic-credit paths were updated
* test cases
* confirmation that unrelated logic was not changed

Final behavior:

```text
自動判斷取得學分：

修課及格標準有有效數值
    → 優先使用修課及格標準

修課及格標準沒有有效數值
    → 使用原本學生及格標準

其他取得學分判斷及匯入邏輯維持不變
```
