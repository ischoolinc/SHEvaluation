## 目標

調整 `Program.cs`。

依照目前本學期「學期學業成績排名」相關欄位的處理方式，新增「上學期學業成績排名」相關欄位。

正式欄位名稱請使用：

```text
上學期學業成績班排名
````

不要再使用：

```text
上學期學業成績名次
```

修改完成後，請完整記錄在：

```text
期末成績通知單(固定排名)調整0604.md
```

---

# 一、重要限制

請嚴格遵守：

```text
1. 只修改 Program.cs。
2. 不要變動本學期學業成績排名原本邏輯。
3. 不要變動本學期學業成績排名原本欄位名稱。
4. 不要變動本學期學業成績排名原本填值方式。
5. 不要變動科目成績計算。
6. 不要變動學分數計算。
7. 不要變動本學期取得學分數、累計取得學分數計算。
8. 不要變動 Word 合併流程。
9. 不要重構整個 Program.cs。
10. 只新增上學期學業成績排名相關欄位與填值。
11. 不要再新增或使用「上學期學業成績名次」。
```

---

# 二、欄位命名修正

本次正式欄位名稱是：

```text
上學期學業成績班排名
```

不要使用：

```text
上學期學業成績名次
```

如果目前 `Program.cs` 已經有：

```csharp
table.Columns.Add("上學期學業成績名次");
```

請移除，或不要再填值使用。

如果目前已經有：

```csharp
row["上學期學業成績名次"] = ...
```

請移除或改成：

```csharp
row["上學期學業成績班排名"] = ...
```

---

# 三、上學期排名資料來源

目前本學期排名資料來源通常是：

```csharp
Utility.GetSemsScoreRankMatrixData(conf.SchoolYear, conf.Semester, selectedStudents);
```

上學期排名資料來源請使用：

```csharp
Utility.GetSemsScoreRankMatrixData(conf.SchoolYear, "1", selectedStudents);
```

但只有目前報表學期是第 2 學期時才讀取：

```csharp
Dictionary<string, Dictionary<string, DataRow>> PrevSemsScoreRankMatrixDataDict =
    new Dictionary<string, Dictionary<string, DataRow>>();

if (conf.Semester == "2")
{
    PrevSemsScoreRankMatrixDataDict =
        Utility.GetSemsScoreRankMatrixData(conf.SchoolYear, "1", selectedStudents);
}
```

規則：

```text
conf.Semester == "2" 時：
讀取 conf.SchoolYear 第 1 學期排名資料。

conf.Semester == "1" 時：
不讀取上學期排名資料。
上學期相關欄位保持空白。
```

---

# 四、新增 DataTable 欄位

請在建立本學期「學期學業成績排名」欄位附近，新增上學期欄位。

## 4.1 新增一般學業成績排名欄位

新增：

```csharp
table.Columns.Add("上學期學業成績班排名");
table.Columns.Add("上學期學業成績科排名");
table.Columns.Add("上學期學業成績全校排名");
table.Columns.Add("上學期學業成績類別1排名");
table.Columns.Add("上學期學業成績類別2排名");

table.Columns.Add("上學期學業成績班排名母數");
table.Columns.Add("上學期學業成績科排名母數");
table.Columns.Add("上學期學業成績全校排名母數");
table.Columns.Add("上學期學業成績類別1排名母數");
table.Columns.Add("上學期學業成績類別2排名母數");
```

---

## 4.2 新增原始學業成績排名欄位

新增：

```csharp
table.Columns.Add("上學期學業(原始)成績班排名");
table.Columns.Add("上學期學業(原始)成績科排名");
table.Columns.Add("上學期學業(原始)成績全校排名");
table.Columns.Add("上學期學業(原始)成績類別1排名");
table.Columns.Add("上學期學業(原始)成績類別2排名");

table.Columns.Add("上學期學業(原始)成績班排名母數");
table.Columns.Add("上學期學業(原始)成績科排名母數");
table.Columns.Add("上學期學業(原始)成績全校排名母數");
table.Columns.Add("上學期學業(原始)成績類別1排名母數");
table.Columns.Add("上學期學業(原始)成績類別2排名母數");
```

---

## 4.3 新增類別名稱欄位

如果本學期目前有：

```text
學期類別排名1
學期類別排名2
```

請新增：

```csharp
table.Columns.Add("上學期類別排名1");
table.Columns.Add("上學期類別排名2");
```

---

## 4.4 新增 r2List 相關欄位

請檢查本學期是否有透過 `r2List` 建立排名統計欄位，例如：

```text
學期學業成績班排名_percentile
學期學業成績班排名_pr
學期學業成績班排名_avg_top_25
學期學業成績班排名_avg_top_50
學期學業成績班排名_avg_bottom_50
學期學業成績班排名_avg
學期學業成績班排名_std_dev
學期學業成績班排名_level_gte100
```

如果本學期有，請上學期也依照相同 `r2List` 新增欄位，只是前綴改成 `上學期`。

範例：

```csharp
foreach (string item2 in r2List)
{
    table.Columns.Add("上學期學業成績班排名_" + item2);
    table.Columns.Add("上學期學業成績科排名_" + item2);
    table.Columns.Add("上學期學業成績全校排名_" + item2);
    table.Columns.Add("上學期學業成績類別1排名_" + item2);
    table.Columns.Add("上學期學業成績類別2排名_" + item2);

    table.Columns.Add("上學期學業(原始)成績班排名_" + item2);
    table.Columns.Add("上學期學業(原始)成績科排名_" + item2);
    table.Columns.Add("上學期學業(原始)成績全校排名_" + item2);
    table.Columns.Add("上學期學業(原始)成績類別1排名_" + item2);
    table.Columns.Add("上學期學業(原始)成績類別2排名_" + item2);
}
```

注意：

```text
不要額外發明本學期沒有的 r2List 欄位。
上學期欄位要跟本學期欄位結構一致，只是前綴改成「上學期」。
```

---

# 五、Ranking Key 對應

上學期使用的 ranking key 與本學期相同，只是資料來源改用：

```text
PrevSemsScoreRankMatrixDataDict
```

---

## 5.1 一般學業成績排名 key

| 上學期輸出欄位      | Ranking Key      |
| ------------ | ---------------- |
| 上學期學業成績班排名   | 學期/分項成績_學業_班排名   |
| 上學期學業成績科排名   | 學期/分項成績_學業_科排名   |
| 上學期學業成績全校排名  | 學期/分項成績_學業_年排名   |
| 上學期學業成績類別1排名 | 學期/分項成績_學業_類別1排名 |
| 上學期學業成績類別2排名 | 學期/分項成績_學業_類別2排名 |

注意：

```text
Ranking Key 裡的「年排名」對應輸出欄位名稱「全校排名」。
請保持與本學期邏輯一致。
```

---

## 5.2 原始學業成績排名 key

| 上學期輸出欄位          | Ranking Key          |
| ---------------- | -------------------- |
| 上學期學業(原始)成績班排名   | 學期/分項成績(原始)_學業_班排名   |
| 上學期學業(原始)成績科排名   | 學期/分項成績(原始)_學業_科排名   |
| 上學期學業(原始)成績全校排名  | 學期/分項成績(原始)_學業_年排名   |
| 上學期學業(原始)成績類別1排名 | 學期/分項成績(原始)_學業_類別1排名 |
| 上學期學業(原始)成績類別2排名 | 學期/分項成績(原始)_學業_類別2排名 |

---

# 六、填值方式

## 6.1 建議新增 Helper Method

為了避免動到本學期原本排名邏輯，建議新增 Helper，只給上學期使用。

如果 `Program.cs` 已經有類似方法，請沿用現有方法，不要重複新增。

```csharp
private static void FillRankField(
    DataRow row,
    Dictionary<string, DataRow> rankData,
    string rankKey,
    string outputFieldName,
    List<string> r2List,
    List<string> r2ParseList)
{
    if (row == null || rankData == null)
        return;

    if (!rankData.ContainsKey(rankKey))
        return;

    DataRow rankRow = rankData[rankKey];

    if (rankRow == null)
        return;

    if (row.Table.Columns.Contains(outputFieldName) &&
        rankRow.Table.Columns.Contains("rank"))
    {
        row[outputFieldName] = "" + rankRow["rank"];
    }

    if (row.Table.Columns.Contains(outputFieldName + "母數") &&
        rankRow.Table.Columns.Contains("matrix_count"))
    {
        row[outputFieldName + "母數"] = "" + rankRow["matrix_count"];
    }

    if (rankRow.Table.Columns.Contains("rank_name"))
    {
        if (outputFieldName.Contains("類別1") &&
            row.Table.Columns.Contains("上學期類別排名1"))
        {
            row["上學期類別排名1"] = "" + rankRow["rank_name"];
        }

        if (outputFieldName.Contains("類別2") &&
            row.Table.Columns.Contains("上學期類別排名2"))
        {
            row["上學期類別排名2"] = "" + rankRow["rank_name"];
        }
    }

    if (r2List != null)
    {
        foreach (string item2 in r2List)
        {
            string output2 = outputFieldName + "_" + item2;

            if (row.Table.Columns.Contains(output2) &&
                rankRow.Table.Columns.Contains(item2))
            {
                row[output2] = "" + rankRow[item2];
            }
        }
    }

    if (r2ParseList != null)
    {
        foreach (string item2 in r2ParseList)
        {
            string output2 = outputFieldName + "_" + item2;

            if (row.Table.Columns.Contains(output2) &&
                rankRow.Table.Columns.Contains(item2))
            {
                row[output2] = "" + rankRow[item2];
            }
        }
    }
}
```

---

## 6.2 新增上學期學業排名填值 Helper

新增：

```csharp
private static void FillPrevSemesterAcademicRankFields(
    DataRow row,
    Dictionary<string, DataRow> rankData,
    List<string> r2List,
    List<string> r2ParseList)
{
    if (row == null || rankData == null)
        return;

    // 一般學業成績排名
    FillRankField(row, rankData,
        "學期/分項成績_學業_班排名",
        "上學期學業成績班排名",
        r2List,
        r2ParseList);

    FillRankField(row, rankData,
        "學期/分項成績_學業_科排名",
        "上學期學業成績科排名",
        r2List,
        r2ParseList);

    FillRankField(row, rankData,
        "學期/分項成績_學業_年排名",
        "上學期學業成績全校排名",
        r2List,
        r2ParseList);

    FillRankField(row, rankData,
        "學期/分項成績_學業_類別1排名",
        "上學期學業成績類別1排名",
        r2List,
        r2ParseList);

    FillRankField(row, rankData,
        "學期/分項成績_學業_類別2排名",
        "上學期學業成績類別2排名",
        r2List,
        r2ParseList);

    // 原始學業成績排名
    FillRankField(row, rankData,
        "學期/分項成績(原始)_學業_班排名",
        "上學期學業(原始)成績班排名",
        r2List,
        r2ParseList);

    FillRankField(row, rankData,
        "學期/分項成績(原始)_學業_科排名",
        "上學期學業(原始)成績科排名",
        r2List,
        r2ParseList);

    FillRankField(row, rankData,
        "學期/分項成績(原始)_學業_年排名",
        "上學期學業(原始)成績全校排名",
        r2List,
        r2ParseList);

    FillRankField(row, rankData,
        "學期/分項成績(原始)_學業_類別1排名",
        "上學期學業(原始)成績類別1排名",
        r2List,
        r2ParseList);

    FillRankField(row, rankData,
        "學期/分項成績(原始)_學業_類別2排名",
        "上學期學業(原始)成績類別2排名",
        r2List,
        r2ParseList);
}
```

注意：

```text
不要在這裡寫：
row["上學期學業成績名次"] = row["上學期學業成績班排名"];

本次不再使用「上學期學業成績名次」。
```

---

# 七、呼叫位置

請在每位學生 `DataRow row` 填值時，找到目前處理 `PrevSemsScoreRankMatrixDataDict` 或 `上學期學業成績名次` 的區塊。

如果目前有類似：

```csharp
row["上學期學業成績名次"] = "";
```

請移除，或改成初始化正式欄位：

```csharp
row["上學期學業成績班排名"] = "";
```

請新增或調整成：

```csharp
if (conf.Semester == "2" &&
    PrevSemsScoreRankMatrixDataDict.ContainsKey(stuRec.StudentID))
{
    FillPrevSemesterAcademicRankFields(
        row,
        PrevSemsScoreRankMatrixDataDict[stuRec.StudentID],
        r2List,
        r2ParseList);
}
```

規則：

```text
conf.Semester == "1" 時，不呼叫 FillPrevSemesterAcademicRankFields。
上學期排名欄位保持空白。
```

---

# 八、不要修改本學期排名邏輯

請不要修改本學期原本這些欄位的處理：

```text
學期學業成績班排名
學期學業成績科排名
學期學業成績全校排名
學期學業成績類別1排名
學期學業成績類別2排名
學期學業(原始)成績班排名
學期學業(原始)成績科排名
學期學業(原始)成績全校排名
學期學業(原始)成績類別1排名
學期學業(原始)成績類別2排名
```

也不要改本學期原本使用的：

```text
SemsScoreRankMatrixDataDict
Utility.GetSemsScoreRankMatrixData(conf.SchoolYear, conf.Semester, selectedStudents)
```

本次只新增：

```text
PrevSemsScoreRankMatrixDataDict
上學期學業...
```

---

# 九、測試項目

## 測試 1：第 1 學期

設定：

```text
conf.Semester = 1
```

預期：

```text
上學期學業成績班排名 = 空白
上學期學業成績科排名 = 空白
上學期學業成績全校排名 = 空白
上學期學業成績類別1排名 = 空白
上學期學業成績類別2排名 = 空白
```

---

## 測試 2：第 2 學期有上學期班排名資料

設定：

```text
conf.Semester = 2
```

上學期排名資料：

```text
學期/分項成績_學業_班排名 rank = 3
matrix_count = 40
```

預期：

```text
上學期學業成績班排名 = 3
上學期學業成績班排名母數 = 40
```

---

## 測試 3：第 2 學期有上學期科排名資料

上學期排名資料：

```text
學期/分項成績_學業_科排名 rank = 10
matrix_count = 120
```

預期：

```text
上學期學業成績科排名 = 10
上學期學業成績科排名母數 = 120
```

---

## 測試 4：第 2 學期有上學期全校排名資料

上學期排名資料 key：

```text
學期/分項成績_學業_年排名
```

預期填入：

```text
上學期學業成績全校排名
上學期學業成績全校排名母數
```

注意：

```text
key 是「年排名」，欄位是「全校排名」。
```

---

## 測試 5：第 2 學期有上學期類別排名資料

上學期排名資料：

```text
學期/分項成績_學業_類別1排名
rank = 5
matrix_count = 120
rank_name = 類別A
```

預期：

```text
上學期學業成績類別1排名 = 5
上學期學業成績類別1排名母數 = 120
上學期類別排名1 = 類別A
```

---

## 測試 6：不要再出現上學期學業成績名次

請檢查：

```text
Program.cs 不要再新增 table.Columns.Add("上學期學業成績名次");
Program.cs 不要再填 row["上學期學業成績名次"];
debug.xml 不需要再輸出 上學期學業成績名次。
```

正式欄位請使用：

```text
上學期學業成績班排名
```

---

## 測試 7：本學期排名不變

請比對修改前後：

```text
學期學業成績班排名
學期學業成績科排名
學期學業成績全校排名
學期學業成績類別1排名
學期學業成績類別2排名
```

預期：

```text
本學期排名欄位與修改前完全一致。
```

---

# 十、完成後紀錄

請建立或更新：

```text
期末成績通知單(固定排名)調整0604.md
```

紀錄內容請包含：

````md
# 期末成績通知單(固定排名)調整0604

## 修改目標

新增上學期學業成績排名相關欄位與填值處理。

正式欄位名稱使用：

```text
上學期學業成績班排名
````

不再使用：

```text
上學期學業成績名次
```

## 修改檔案

* Program.cs

## 修改內容

### 1. 新增上學期排名資料來源

只有目前報表學期是第 2 學期時，讀取：

```csharp
Utility.GetSemsScoreRankMatrixData(conf.SchoolYear, "1", selectedStudents)
```

### 2. 新增上學期學業成績排名欄位

新增：

```text
上學期學業成績班排名
上學期學業成績科排名
上學期學業成績全校排名
上學期學業成績類別1排名
上學期學業成績類別2排名
```

以及對應母數欄位：

```text
上學期學業成績班排名母數
上學期學業成績科排名母數
上學期學業成績全校排名母數
上學期學業成績類別1排名母數
上學期學業成績類別2排名母數
```

### 3. 新增上學期學業原始成績排名欄位

新增：

```text
上學期學業(原始)成績班排名
上學期學業(原始)成績科排名
上學期學業(原始)成績全校排名
上學期學業(原始)成績類別1排名
上學期學業(原始)成績類別2排名
```

以及對應母數欄位。

### 4. 欄位命名修正

移除或不再使用：

```text
上學期學業成績名次
```

改用：

```text
上學期學業成績班排名
```

## 不變動內容

本次未修改：

```text
1. 本學期學業成績排名原本邏輯
2. 本學期排名資料來源
3. 本學期排名欄位名稱
4. 科目成績計算
5. 學分數計算
6. 本學期取得學分數
7. 累計取得學分數
8. Word 合併流程
```

## 測試結果

請記錄：

```text
1. 第 1 學期時，上學期排名欄位是否空白。
2. 第 2 學期時，上學期班排名是否正確。
3. 第 2 學期時，上學期科排名是否正確。
4. 第 2 學期時，上學期全校排名是否正確。
5. 第 2 學期時，上學期類別排名是否正確。
6. 是否不再輸出 上學期學業成績名次。
7. 是否改用 上學期學業成績班排名。
8. 本學期排名欄位是否與修改前一致。
9. 編譯是否成功。
```

```

重點：正式合併欄位統一用 **`上學期學業成績班排名`**，不要再用 **`上學期學業成績名次`**。
```
