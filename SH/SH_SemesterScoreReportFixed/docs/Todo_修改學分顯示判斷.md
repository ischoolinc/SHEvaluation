## 目標

調整 `Program.cs`。

依照目前結論，修正科目學分數顯示判斷：

```text
學分數N：只用本學期成績判斷。
上學期學分數N：只用上學期成績判斷。
學年成績不影響學分數N判斷。
````

修改完成後，請完整記錄在：

```text
期末成績通知單(固定排名)調整0604.md
```

---

# 一、重要限制

請嚴格遵守：

```text
1. 只修改 Program.cs。
2. 不要變動本學期科目成績原本填值邏輯。
3. 不要變動上學期科目成績原本填值邏輯。
4. 不要變動學年科目成績原本填值邏輯。
5. 不要變動本學期取得學分數計算。
6. 不要變動累計取得學分數計算。
7. 不要變動排名邏輯。
8. 不要變動 Word 合併流程。
9. 不要重構整個 Program.cs。
10. 只調整學分數N與上學期學分數N的顯示判斷。
```

---

# 二、目前問題

目前 `學分數N` 會受到多種成績欄位影響，例如：

```text
學期科目成績N
上學期科目成績N
學年科目成績N
```

因此會發生：

```text
本學期沒有分數
但上學期有分數
或學年成績有分數
=> 學分數N 仍然顯示
```

例如：

```text
金融與證券投資實務Ⅰ

第二學期沒有分數
但上學期科目成績有值
學年科目成績有值
所以學分數N 被保留下來
```

這不是正確顯示規則。

---

# 三、修正後規則

## 3.1 學分數N

```text
學分數N 只用本學期成績判斷。
```

判斷來源只允許使用本學期相關欄位，例如：

```text
科目成績N
學期科目成績N
學期科目原始成績N
學期科目補考成績N
學期科目重修成績N
學期科目手動調整成績N
學期科目學年調整成績N
```

不可使用：

```text
上學期科目成績N
上學期科目原始成績N
上學期科目補考成績N
上學期科目重修成績N
上學期科目手動調整成績N
上學期科目學年調整成績N
學年科目成績N
```

---

## 3.2 上學期學分數N

```text
上學期學分數N 只用上學期成績判斷。
```

判斷來源只允許使用上學期相關欄位，例如：

```text
上學期科目成績N
上學期科目原始成績N
上學期科目補考成績N
上學期科目重修成績N
上學期科目手動調整成績N
上學期科目學年調整成績N
```

不可使用：

```text
科目成績N
學期科目成績N
學期科目原始成績N
學年科目成績N
```

---

## 3.3 學年科目成績

```text
學年科目成績N 不可以影響 學分數N 是否顯示。
```

也就是：

```text
即使學年科目成績N有值，
只要本學期成績N沒有值，
學分數N 就應該空白。
```

---

# 四、有效成績判斷

請確認或新增 Helper Method：

```csharp
/// <summary>
/// 判斷成績欄位是否有有效成績值。
/// 0 是有效成績；空白、null、DBNull、未輸入視為無成績。
/// </summary>
private static bool HasScoreValue(object score)
{
    if (score == null || score == DBNull.Value)
        return false;

    string value = Convert.ToString(score).Trim();

    if (string.IsNullOrEmpty(value))
        return false;

    if (value == "未輸入")
        return false;

    return true;
}
```

如果 `Program.cs` 已經有相同功能的 Helper，請沿用，不要重複新增。

---

# 五、調整 ClearCreditIfNoSubjectScore

請找到目前類似這個方法：

```csharp
ClearCreditIfNoSubjectScore(row, subjectIndex);
```

或：

```csharp
private static void ClearCreditIfNoSubjectScore(DataRow row, int subjectIndex)
```

目前此方法可能會檢查：

```text
科目成績N
學期科目成績N
上學期科目成績N
學年科目成績N
```

請調整為：

```text
ClearCreditIfNoSubjectScore 只檢查本學期相關成績欄位。
```

---

## 5.1 修正後建議寫法

```csharp
private static void ClearCreditIfNoSubjectScore(DataRow row, int subjectIndex)
{
    if (row == null)
        return;

    bool hasCurrentSemesterScore =
        HasScoreValue(GetRowValue(row, "科目成績" + subjectIndex)) ||
        HasScoreValue(GetRowValue(row, "學期科目成績" + subjectIndex)) ||
        HasScoreValue(GetRowValue(row, "學期科目原始成績" + subjectIndex)) ||
        HasScoreValue(GetRowValue(row, "學期科目補考成績" + subjectIndex)) ||
        HasScoreValue(GetRowValue(row, "學期科目重修成績" + subjectIndex)) ||
        HasScoreValue(GetRowValue(row, "學期科目手動調整成績" + subjectIndex)) ||
        HasScoreValue(GetRowValue(row, "學期科目學年調整成績" + subjectIndex));

    if (!hasCurrentSemesterScore && row.Table.Columns.Contains("學分數" + subjectIndex))
    {
        row["學分數" + subjectIndex] = "";
    }
}
```

如果目前程式沒有 `GetRowValue`，可新增：

```csharp
private static object GetRowValue(DataRow row, string columnName)
{
    if (row == null || row.Table == null)
        return null;

    if (!row.Table.Columns.Contains(columnName))
        return null;

    return row[columnName];
}
```

如果已有相同功能方法，請沿用既有方法。

---

# 六、不可再讓學年成績影響學分數N

請確認 `ClearCreditIfNoSubjectScore()` 內不要再出現：

```csharp
HasScoreValue(GetRowValue(row, "學年科目成績" + subjectIndex))
```

也不要再用：

```text
學年科目成績N
```

判斷是否保留：

```text
學分數N
```

---

# 七、上學期學分數N判斷

請確認已經有欄位：

```csharp
table.Columns.Add("上學期學分數" + subjectIndex);
```

若尚未新增，請在上學期科目欄位附近新增：

```csharp
table.Columns.Add("上學期學分數" + subjectIndex);
```

建議位置：

```csharp
table.Columns.Add("上學期科目成績" + subjectIndex);
table.Columns.Add("上學期學分數" + subjectIndex);
```

---

## 7.1 上學期學分數N填值規則

請在上學期科目成績填值後加入或確認：

```csharp
if (HasScoreValue(semesterSubjectScore.Score))
{
    row["上學期學分數" + subjectIndex] = semesterSubjectScore.CreditDec();
}
else
{
    row["上學期學分數" + subjectIndex] = "";
}
```

如果上學期成績不一定只看 `semesterSubjectScore.Score`，也可以改成用 row 裡的上學期欄位判斷：

```csharp
bool hasPrevSemesterScore =
    HasScoreValue(GetRowValue(row, "上學期科目成績" + subjectIndex)) ||
    HasScoreValue(GetRowValue(row, "上學期科目原始成績" + subjectIndex)) ||
    HasScoreValue(GetRowValue(row, "上學期科目補考成績" + subjectIndex)) ||
    HasScoreValue(GetRowValue(row, "上學期科目重修成績" + subjectIndex)) ||
    HasScoreValue(GetRowValue(row, "上學期科目手動調整成績" + subjectIndex)) ||
    HasScoreValue(GetRowValue(row, "上學期科目學年調整成績" + subjectIndex));

if (hasPrevSemesterScore)
{
    row["上學期學分數" + subjectIndex] = semesterSubjectScore.CreditDec();
}
else
{
    row["上學期學分數" + subjectIndex] = "";
}
```

原則：

```text
上學期學分數N 只能由上學期成績欄位決定。
```

---

# 八、不要修改原本學分統計

請不要修改以下欄位與計算：

```text
本學期取得學分數
累計取得學分數
本學期已修必修學分
本學期已修選修學分
本學期實得必修學分
本學期實得選修學分
累計取得必修學分
累計取得選修學分
在校期間實得必修學分
在校期間實得選修學分
上學期實得學分數
上學期累計取得學分數
學年實得學分數
學年累計取得學分數
```

本次只處理：

```text
學分數N
上學期學分數N
```

的顯示判斷。

---

# 九、Word 樣板注意

Word 樣板使用方式應該是：

```text
本學期成績旁邊學分：使用 學分數N
上學期成績旁邊學分：使用 上學期學分數N
```

請確認上學期區塊不要再使用：

```text
學分數N
```

應改用：

```text
上學期學分數N
```

---

# 十、測試案例

## 測試 1：本學期有成績

資料：

```text
學期科目成績8 = 85
學分數8 原始 credit = 2
```

預期：

```text
學分數8 = 2
```

---

## 測試 2：本學期沒有成績，但上學期有成績

資料：

```text
學期科目成績15 = 空白
上學期科目成績15 = 99
學年科目成績15 = 99.0
credit = 2
```

預期：

```text
學分數15 = 空白
上學期學分數15 = 2
```

---

## 測試 3：本學期沒有成績，但學年成績有值

資料：

```text
學期科目成績15 = 空白
上學期科目成績15 = 空白
學年科目成績15 = 99.0
credit = 2
```

預期：

```text
學分數15 = 空白
```

說明：

```text
學年科目成績15 不可以影響學分數15。
```

---

## 測試 4：上學期沒有成績，但本學期有成績

資料：

```text
學期科目成績8 = 85
上學期科目成績8 = 空白
credit = 2
```

預期：

```text
學分數8 = 2
上學期學分數8 = 空白
```

---

## 測試 5：成績為 0

資料：

```text
學期科目成績10 = 0
credit = 2
```

預期：

```text
學分數10 = 2
```

資料：

```text
上學期科目成績10 = 0
credit = 2
```

預期：

```text
上學期學分數10 = 2
```

說明：

```text
0 是有效成績，不可視為空白。
```

---

## 測試 6：成績為未輸入

資料：

```text
學期科目成績10 = 未輸入
credit = 2
```

預期：

```text
學分數10 = 空白
```

資料：

```text
上學期科目成績10 = 未輸入
credit = 2
```

預期：

```text
上學期學分數10 = 空白
```

---

# 十一、完成後檢查

請確認：

```text
1. 學分數N 只受本學期成績欄位影響。
2. 上學期學分數N 只受上學期成績欄位影響。
3. 學年科目成績N 不影響學分數N。
4. 本學期沒有分數、上學期有分數時，學分數N 為空白。
5. 本學期沒有分數、學年有分數時，學分數N 為空白。
6. 上學期有分數時，上學期學分數N 有值。
7. 上學期沒有分數時，上學期學分數N 空白。
8. 原本學分統計欄位沒有被修改。
9. 原本排名邏輯沒有被修改。
10. 編譯成功。
```

---

# 十二、完成後紀錄

請建立或更新：

```text
期末成績通知單(固定排名)調整0604.md
```

紀錄內容請包含：

````md
# 期末成績通知單(固定排名)調整0604

## 修改目標

調整科目學分數顯示判斷：

```text
學分數N 只用本學期成績判斷。
上學期學分數N 只用上學期成績判斷。
學年科目成績N 不影響學分數N 判斷。
````

## 修改檔案

* Program.cs

## 修改內容

### 1. 調整學分數N判斷

```text
學分數N 只檢查本學期相關成績欄位。
不再檢查上學期科目成績N。
不再檢查學年科目成績N。
```

### 2. 調整上學期學分數N判斷

```text
上學期學分數N 只檢查上學期相關成績欄位。
不檢查本學期科目成績N。
不檢查學年科目成績N。
```

### 3. 有效成績判斷

```text
0 為有效成績。
null、DBNull、空字串、未輸入 視為無成績。
```

## 不變動內容

本次未修改：

```text
1. 本學期科目成績填值邏輯
2. 上學期科目成績填值邏輯
3. 學年科目成績填值邏輯
4. 本學期取得學分數計算
5. 累計取得學分數計算
6. 本學期實得必修/選修學分
7. 排名邏輯
8. Word 合併流程
```

## 測試結果

請記錄：

```text
1. 本學期有成績時，學分數N 是否顯示。
2. 本學期沒有成績但上學期有成績時，學分數N 是否空白。
3. 本學期沒有成績但學年有成績時，學分數N 是否空白。
4. 上學期有成績時，上學期學分數N 是否顯示。
5. 上學期沒有成績時，上學期學分數N 是否空白。
6. 成績為 0 時是否仍顯示學分。
7. 成績為 未輸入 時是否清空學分。
8. 原本學分統計是否不變。
9. 編譯是否成功。
```

```

重點：**`學分數N` 不可以再被 `上學期科目成績N` 或 `學年科目成績N` 影響。**
```
