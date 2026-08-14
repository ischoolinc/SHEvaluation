## 目標

修正 `frmStudentRankDetail` 讀取學生排名資料時，`matrix_count` 沒有資料或無法轉成整數時可能發生：

```text
讀取排名資料失敗：輸入字串格式不正確。
```

的問題。

同時維持目前排名顯示規則：

```text
只有排名，沒有排名母數
→ 顯示排名

有排名，也有排名母數
→ 顯示 排名 / 排名母數

沒有排名
→ 顯示空白
```

修改完成後，將本次調整追加記錄至：

```text
學生排名資料調整0812.md
```

---

# 一、修改原則

- [ ] 採最小修改原則。
- [ ] 不修改 `rank_override` Schema。
- [ ] 不新增 `matrix_count` DB 欄位。
- [ ] 不修改學年分項排名匯入 Excel 格式。
- [ ] 不要求學年分項排名一定要有 `matrix_count`。
- [ ] 不修改既有 `rank` 資料。
- [ ] 不修改 extension JSON 結構。
- [ ] 不修改與本問題無關的排名功能。
- [ ] 不把 DataGrid 的 `RankDisplay` 改成 `Rank`。

---

# 二、先檢查目前相關程式

請檢查：

```text
RankOverrideDataAccess.cs
frmStudentRankDetail.cs
StudentRankDetailItem
```

搜尋：

```text
matrix_count
MatrixCount
Rank
RankDisplay
ToNullableInt
MapDetailItem
GetStudentRankDetail
```

確認目前實際程式碼後再修改。

---

# 三、目前 GetStudentRankDetail SQL

目前 SQL 已讀取：

```sql
ro.rank,
ro.matrix_count
```

這部分原則上不需要修改。

目前 Mapping：

```csharp
Rank = ToNullableInt(row["rank"]),
MatrixCount = ToNullableInt(row["matrix_count"]),
```

請保留 `Rank` 與 `MatrixCount` 分開 Mapping 的設計。

---

# 四、修正 ToNullableInt

目前 `ToNullableInt()` 類似：

```csharp
private static int? ToNullableInt(object value)
{
    if (value == null || value == DBNull.Value)
        return null;

    return Convert.ToInt32(value);
}
```

問題：

只處理：

```text
null
DBNull.Value
```

若取得的值是：

```text
""
" "
其他無法解析成 int 的內容
```

則：

```csharp
Convert.ToInt32(value)
```

可能產生：

```text
FormatException
輸入字串格式不正確
```

---

# 五、調整 ToNullableInt 為安全解析

請修改為安全判斷。

建議：

```csharp
private static int? ToNullableInt(object value)
{
    if (value == null || value == DBNull.Value)
        return null;

    string text = Convert.ToString(value);

    if (string.IsNullOrWhiteSpace(text))
        return null;

    int result;
    if (int.TryParse(text.Trim(), out result))
        return result;

    return null;
}
```

## 預期轉換

```text
NULL
→ null

DBNull.Value
→ null

""
→ null

" "
→ null

"15"
→ 15

15
→ 15

"25"
→ 25
```

不可因為 `matrix_count` 沒有資料而中斷整個 `frmStudentRankDetail` 的讀取。

---

# 六、matrix_count 規則

`matrix_count` 為排名母數。

本次學年分項排名允許：

```text
rank 有資料
matrix_count 沒有資料
```

這是合法情況。

例如：

```text
rank = 3
matrix_count = NULL
```

不應視為錯誤。

DataAccess 應得到：

```text
Rank = 3
MatrixCount = null
```

---

# 七、RankDisplay 顯示規則

確認 `StudentRankDetailItem.RankDisplay`。

必須符合以下規則。

## 7.1 有排名，沒有排名母數

資料：

```text
Rank = 3
MatrixCount = null
```

顯示：

```text
3
```

---

## 7.2 有排名，也有排名母數

資料：

```text
Rank = 3
MatrixCount = 25
```

顯示：

```text
3 / 25
```

---

## 7.3 沒有排名

資料：

```text
Rank = null
MatrixCount = null
```

顯示：

```text
空白
```

---

## 7.4 沒有排名，但有排名母數

資料：

```text
Rank = null
MatrixCount = 25
```

仍顯示：

```text
空白
```

不可顯示：

```text
/ 25
```

---

# 八、RankDisplay 建議寫法

若目前尚未符合規則，調整為：

```csharp
public string RankDisplay
{
    get
    {
        if (!Rank.HasValue)
            return string.Empty;

        if (!MatrixCount.HasValue)
            return Rank.Value.ToString();

        return Rank.Value + " / " + MatrixCount.Value;
    }
}
```

如果目前 `RankDisplay` 已符合以上規則，則不要重複修改。

---

# 九、DataGrid 綁定不要修改

目前 `frmStudentRankDetail` 的排名欄位應繼續使用：

```csharp
AddTextColumn(
    "colRank",
    "排名",
    "RankDisplay",
    70);
```

不要改成：

```csharp
AddTextColumn(
    "colRank",
    "排名",
    "Rank",
    70);
```

原因：

`RankDisplay` 必須同時支援：

```text
3
```

以及：

```text
3 / 25
```

兩種資料。

---

# 十、學年分項排名

目前學年分項排名匯入不一定提供：

```text
matrix_count
```

因此：

```text
rank = 1
matrix_count = NULL
```

必須是正常資料。

畫面應顯示：

```text
1
```

不可：

- 顯示空白。
- 發生 FormatException。
- 顯示「輸入字串格式不正確」。
- 因一筆 `matrix_count` 為空造成整個 DataGrid 無法顯示。

---

# 十一、既有排名資料

如果既有排名資料具有：

```text
rank = 15
matrix_count = 25
```

仍必須正常顯示：

```text
15 / 25
```

本次修改不可破壞有排名母數的既有資料。

---

# 十二、錯誤處理

重點確認：

```csharp
MatrixCount = ToNullableInt(row["matrix_count"])
```

在以下資料狀況都不會拋出 Exception：

```text
NULL
DBNull.Value
空字串
空白字串
正常整數
可解析的整數字串
```

若內容無法解析：

```text
abc
-
其他非整數內容
```

本次以：

```text
null
```

處理，避免因單筆排名母數異常導致整個學生排名 Detail 無法開啟。

---

# 十三、不要修改 Upsert

本次不要為了解決 Detail 顯示問題修改：

```text
UpsertSchoolYearEntryRanks()
```

尤其不要自行新增：

```text
matrix_count
```

寫入邏輯。

本次問題是：

```text
Detail 讀取時必須允許 matrix_count 沒有資料
```

不是：

```text
匯入時強制建立 matrix_count
```

---

# 十四、測試

## Case 1：只有排名

DB：

```text
rank = 2
matrix_count = NULL
```

預期：

```text
RankDisplay = 2
```

且不發生：

```text
輸入字串格式不正確
```

---

## Case 2：排名 + 排名母數

DB：

```text
rank = 2
matrix_count = 30
```

預期：

```text
RankDisplay = 2 / 30
```

---

## Case 3：排名與母數都 NULL

```text
rank = NULL
matrix_count = NULL
```

預期：

```text
RankDisplay = 空白
```

且 Form 可以正常載入。

---

## Case 4：matrix_count 為空內容

若資料來源取得：

```text
matrix_count = ""
```

或：

```text
matrix_count = " "
```

預期：

```text
MatrixCount = null
```

不可發生 Exception。

---

## Case 5：非法 matrix_count

若取得：

```text
matrix_count = "abc"
```

預期：

```text
MatrixCount = null
```

且其他排名資料仍可顯示。

---

## Case 6：學年分項排名

開啟：

```text
UCStudentRank
→ 學年
→ 分項
→ frmStudentRankDetail
```

確認：

- DataGrid 可以正常顯示。
- 有 `rank` 即顯示排名。
- 沒有 `matrix_count` 不影響顯示。
- 不再出現「輸入字串格式不正確」。

---

# 十五、確認其他欄位沒有被影響

修改後確認：

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

原本資料讀取流程沒有因本次修改受到額外影響。

此次只處理：

```text
matrix_count
ToNullableInt
RankDisplay
```

必要範圍。

---

# 十六、Build

完成後：

- [ ] 專案 Build 成功。
- [ ] 無 Compiler Error。
- [ ] `matrix_count = null` 可以正常讀取。
- [ ] `rank` 可以正常顯示。
- [ ] 有 `matrix_count` 時可以顯示 `rank / matrix_count`。
- [ ] 不再因 `matrix_count` 發生 FormatException。
- [ ] 不影響既有匯入功能。

---

# 十七、完成紀錄

修改完成後，將本次內容追加至：

```text
學生排名資料調整0812.md
```

至少記錄：

## 修改原因

`frmStudentRankDetail` 讀取排名資料時，需要允許：

```text
matrix_count
```

沒有資料。

避免數值轉換時發生：

```text
FormatException
輸入字串格式不正確
```

## 修改內容

記錄實際修改：

```text
ToNullableInt
MatrixCount
RankDisplay
```

哪些程式。

## 最終顯示規則

```text
Rank 有值
MatrixCount 無值
→ Rank

Rank 有值
MatrixCount 有值
→ Rank / MatrixCount

Rank 無值
→ 空白
```

## 測試結果

記錄：

```text
只有排名
排名 + 排名母數
matrix_count NULL
matrix_count 空白
無排名
學年分項排名
```

測試結果。

---

# 最終要求

本次核心目標：

```text
允許學年分項排名只有 rank、沒有 matrix_count，
不要因 matrix_count 空值或格式問題造成學生排名資料讀取失敗。
```

並維持：

```text
只有排名：
RankDisplay = 排名

有排名與排名母數：
RankDisplay = 排名 / 排名母數

沒有排名：
RankDisplay = 空白
```

完成後更新：

```text
學生排名資料調整0812.md
```