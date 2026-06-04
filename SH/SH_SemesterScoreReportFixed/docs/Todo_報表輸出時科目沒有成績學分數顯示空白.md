## 目標

調整 `Program.cs` 報表輸出邏輯。

當報表輸出時，若某一科目沒有成績，該科目的「學分數N」也要顯示空白。

修改完成後，請完整記錄在：

```text
期末成績通知單(固定排名)0602.md
````

---

# 一、重要限制

請嚴格遵守：

```text
1. 只調整報表輸出 DataRow 合併欄位的顯示值。
2. 不可影響原本學分計算邏輯。
3. 不可修改 SemesterSubjectScore.CreditDec()。
4. 不可修改 ExamScoreInfo.CreditDec()。
5. 不可修改本學期取得學分數計算。
6. 不可修改累計取得學分數計算。
7. 不可修改必修 / 選修學分計算。
8. 不可修改科目排序。
9. 不可修改成績計算邏輯。
10. 不可修改獎勵、懲戒、缺曠計算。
11. 不要重構 Program.cs。
12. 只做最小修改。
```

---

# 二、修改檔案

主要修改：

```text
Program.cs
```

---

# 三、目前問題

目前程式在產生報表合併欄位時，會填入：

```csharp
row["學分數" + subjectIndex] = semesterSubjectScore.CreditDec();
```

或：

```csharp
row["學分數" + subjectIndex] = sceTakeRecord.CreditDec();
```

但如果該科目最後沒有成績，例如：

```text
科目成績N = 空白
```

或：

```text
科目成績N = 未輸入
```

目前「學分數N」仍可能顯示數字。

本次要調整為：

```text
科目沒有成績 → 學分數N 顯示空白
```

---

# 四、修改原則

## 4.1 只改報表顯示欄位

只允許修改：

```csharp
row["學分數" + subjectIndex]
```

不要修改任何原始成績物件或學分計算來源。

也就是說，只能在報表輸出前，針對 `DataRow row` 的合併欄位做顯示調整。

---

## 4.2 不影響這些欄位

不得影響以下欄位計算：

```text
本學期已修必修學分
本學期已修選修學分
本學期取得學分數
本學期實得必修學分
本學期實得選修學分
累計取得學分數
累計取得必修學分
累計取得選修學分
在校期間實得必修學分
在校期間實得選修學分
```

---

# 五、建議新增 Helper Method

請在 `Program.cs` 的 `Program` 類別中，新增以下方法。

建議位置：

```text
放在 GetNumber()、InitAbsenceDefault() 或其他 private static helper method 附近。
```

新增方法：

```csharp
/// <summary>
/// 報表輸出用：如果科目沒有成績，則清空該科目的學分數顯示值。
/// 注意：只調整 DataRow 合併欄位，不可影響原本學分計算邏輯。
/// </summary>
private static void ClearCreditIfNoSubjectScore(DataRow row, int subjectIndex)
{
    if (row == null || row.Table == null)
        return;

    string creditField = "學分數" + subjectIndex;

    if (!row.Table.Columns.Contains(creditField))
        return;

    string[] scoreFields = new string[]
    {
        "科目成績" + subjectIndex,
        "學期科目成績" + subjectIndex,
        "學期科目原始成績" + subjectIndex,
        "學期科目補考成績" + subjectIndex,
        "學期科目重修成績" + subjectIndex,
        "學期科目手動調整成績" + subjectIndex,
        "學期科目學年調整成績" + subjectIndex,
        "上學期科目成績" + subjectIndex,
        "上學期科目原始成績" + subjectIndex,
        "上學期科目補考成績" + subjectIndex,
        "上學期科目重修成績" + subjectIndex,
        "上學期科目手動調整成績" + subjectIndex,
        "上學期科目學年調整成績" + subjectIndex,
        "學年科目成績" + subjectIndex
    };

    bool hasScore = false;

    foreach (string field in scoreFields)
    {
        if (!row.Table.Columns.Contains(field))
            continue;

        string value = "";

        if (row[field] != null && row[field] != DBNull.Value)
            value = Convert.ToString(row[field]).Trim();

        // 0 是有效成績，不可視為空白。
        // 「免」或其他文字成績，只要不是空白與未輸入，都視為有成績。
        if (!string.IsNullOrEmpty(value) && value != "未輸入")
        {
            hasScore = true;
            break;
        }
    }

    if (!hasScore)
        row[creditField] = "";
}
```

---

# 六、呼叫位置

請在處理每一個科目的迴圈中，找到：

```csharp
subjectIndex++;
```

在它前面加入：

```csharp
ClearCreditIfNoSubjectScore(row, subjectIndex);
```

修改後：

```csharp
ClearCreditIfNoSubjectScore(row, subjectIndex);

subjectIndex++;
```

---

# 七、放置位置注意

這行必須放在：

```text
該科目的所有合併欄位都填完之後
subjectIndex++ 之前
```

也就是必須等以下欄位都處理完後再判斷：

```text
學分數N
科目成績N
學期科目成績N
學期科目原始成績N
學期科目補考成績N
學期科目重修成績N
學期科目手動調整成績N
學期科目學年調整成績N
上學期科目成績N
學年科目成績N
```

---

# 八、不得放置的位置

不要放在以下位置：

```text
1. 學分統計加總迴圈內。
2. foreach (var semesterSubjectScore in stuRec.SemesterSubjectScoreList) 的學分累計邏輯內。
3. 本學期取得學分數計算前。
4. 累計取得學分數計算前。
5. row["本學期取得學分數"] 賦值前。
6. row["累計取得學分數"] 賦值前。
```

原因：

```text
本次只要改報表顯示，不可以影響原本計算結果。
```

---

# 九、判斷規則

## 9.1 要保留學分數的情況

以下都視為「有成績」，學分數要保留：

```text
科目成績N = 0
科目成績N = 60
科目成績N = 免
科目成績N = 缺
科目成績N = 任一非空白文字
學期科目成績N 有值
學期科目原始成績N 有值
上學期科目成績N 有值
學年科目成績N 有值
```

---

## 9.2 要清空學分數的情況

以下視為「沒有成績」，學分數要空白：

```text
所有成績相關欄位都是空白
所有成績相關欄位都是 DBNull
所有成績相關欄位都是 null
成績欄位只有「未輸入」
```

處理結果：

```csharp
row["學分數" + subjectIndex] = "";
```

---

# 十、測試案例

## 測試 1：科目有正常成績

資料：

```text
科目成績1 = 80
學分數1 = 2
```

預期：

```text
學分數1 = 2
```

---

## 測試 2：科目成績是 0

資料：

```text
科目成績1 = 0
學分數1 = 2
```

預期：

```text
學分數1 = 2
```

注意：

```text
0 是有效成績，不可清空學分數。
```

---

## 測試 3：科目沒有成績

資料：

```text
科目成績1 = 空白
學期科目成績1 = 空白
學期科目原始成績1 = 空白
學分數1 = 2
```

預期：

```text
學分數1 = 空白
```

---

## 測試 4：科目成績是未輸入

資料：

```text
科目成績1 = 未輸入
學分數1 = 2
```

預期：

```text
學分數1 = 空白
```

---

## 測試 5：科目成績是免

資料：

```text
科目成績1 = 免
學分數1 = 2
```

預期：

```text
學分數1 = 2
```

---

## 測試 6：確認統計學分不受影響

確認以下欄位與修改前一致：

```text
本學期取得學分數
累計取得學分數
本學期已修必修學分
本學期已修選修學分
本學期實得必修學分
本學期實得選修學分
累計取得必修學分
累計取得選修學分
```

---

# 十一、修改後檢查

## 11.1 程式搜尋

確認新增方法：

```text
ClearCreditIfNoSubjectScore
```

確認有在 `subjectIndex++` 前呼叫：

```csharp
ClearCreditIfNoSubjectScore(row, subjectIndex);
subjectIndex++;
```

---

## 11.2 確認沒有修改計算邏輯

請確認沒有改到：

```text
本學期取得學分數
累計取得學分數
CreditDec()
Pass
Require
不計學分
RewardList
AttendanceList
```

---

## 11.3 編譯檢查

確認 `Program.cs` 可正常編譯。

如果缺少命名空間，請確認已有：

```csharp
using System;
using System.Data;
```

通常 `Program.cs` 原本已經會有 `System.Data`。

---

# 十二、完成後紀錄

請建立或更新：

```text
期末成績通知單(固定排名)0602.md
```

紀錄內容請包含：

````md
# 期末成績通知單(固定排名)0602

## 修改目標

調整報表輸出時，當科目沒有成績，該科目的學分數顯示空白。

## 修改檔案

- Program.cs

## 修改內容

### 1. 新增報表顯示用判斷方法

新增：

```csharp
ClearCreditIfNoSubjectScore(DataRow row, int subjectIndex)
````

用途：

```text
只針對 DataRow 合併欄位判斷。
當科目沒有成績時，清空 學分數N。
不影響原本學分計算邏輯。
```

### 2. 在科目欄位輸出完成後呼叫

在每一科處理完成後、`subjectIndex++` 前呼叫：

```csharp
ClearCreditIfNoSubjectScore(row, subjectIndex);
```

### 3. 判斷規則

```text
若科目成績N、學期科目成績N、學期科目原始成績N、上學期科目成績N、學年科目成績N 等成績欄位都沒有值，則 學分數N 顯示空白。
```

### 4. 不影響原本計算

本次沒有修改：

```text
本學期取得學分數
累計取得學分數
必修 / 選修學分
CreditDec()
Pass
Require
不計學分
獎勵
懲戒
缺曠
```

## 測試結果

請記錄：

1. 科目有成績時，學分數是否保留。
2. 科目成績為 0 時，學分數是否保留。
3. 科目沒有成績時，學分數是否空白。
4. 科目成績為未輸入時，學分數是否空白。
5. 本學期取得學分數是否與修改前一致。
6. 累計取得學分數是否與修改前一致。
7. 編譯是否成功。
