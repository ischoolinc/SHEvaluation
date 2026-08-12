## 目標

調整 `frmStudentRankDetail` 學生排名資料的「排名」顯示規則。

修改完成後，將本次調整內容追加記錄至：

```text
學生排名資料調整0812.md
```

---

## 一、先檢查相關程式

請先確認目前專案中的：

```text
frmStudentRankDetail.cs
RankOverrideDataAccess.cs
StudentRankDetailItem
```

搜尋：

```text
RankDisplay
Rank
MatrixCount
StudentRankDetailItem
```

確認 `StudentRankDetailItem` 實際定義位置。

---

## 二、目前資料流

目前 `frmStudentRankDetail.cs` 的 DataGrid「排名」欄位使用：

```csharp
AddTextColumn(
    "colRank",
    "排名",
    "RankDisplay",
    70);
```

因此：

```text
DataGrid 排名
    ↓
StudentRankDetailItem.RankDisplay
```

請保留此設計。

不要改成直接綁：

```text
Rank
```

---

## 三、DAO 已有資料

目前 `RankOverrideDataAccess.GetStudentRankDetail()` SQL 已讀取：

```sql
ro.rank,
ro.matrix_count
```

並在 `MapDetailItem()` 中 Mapping：

```csharp
Rank = ToNullableInt(row["rank"]),
MatrixCount = ToNullableInt(row["matrix_count"]),
```

因此此次不需要修改 SQL 查詢，也不需要新增 DB 欄位。

---

## 四、調整 RankDisplay

找到 `StudentRankDetailItem` 的：

```csharp
RankDisplay
```

將顯示規則調整如下。

### 規則 1：有排名，沒有排名母數

例如：

```text
Rank = 15
MatrixCount = null
```

畫面顯示：

```text
15
```

---

### 規則 2：有排名，也有排名母數

例如：

```text
Rank = 15
MatrixCount = 25
```

畫面顯示：

```text
15 / 25
```

---

### 規則 3：沒有排名

例如：

```text
Rank = null
```

不論 `MatrixCount` 是否有值，畫面顯示空白。

---

## 五、建議實作

`RankDisplay` 可調整為：

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

請依目前 `StudentRankDetailItem` 實際 Property 型別調整，不要為了套用範例而改變既有 Model 型別。

---

## 六、預期結果

| Rank | MatrixCount | RankDisplay |
| ---: | ----------: | ----------- |
|   15 |        null | `15`        |
|   15 |          25 | `15 / 25`   |
|    1 |         120 | `1 / 120`   |
| null |          25 | 空白          |
| null |        null | 空白          |

---

## 七、學年分項排名相容性

目前學年分項排名匯入資料主要寫入：

```text
rank
```

目前不一定有：

```text
matrix_count
```

因此這類資料應正常顯示：

```text
15
```

不可因為：

```text
matrix_count = null
```

造成整個 `RankDisplay` 為空白。

---

## 八、既有排名資料相容性

如果既有 `rank_override` 資料同時具有：

```text
rank = 15
matrix_count = 25
```

則應繼續顯示：

```text
15 / 25
```

不可因本次修改只顯示：

```text
15
```

---

## 九、不要修改

本次只處理排名顯示邏輯。

不要因本次調整修改：

```text
rank_override Schema
學年分項排名 Excel 匯入欄位
UpsertSchoolYearEntryRanks()
extension JSON 結構
rank
matrix_count
DataGrid 欄位綁定名稱
其他排名計算邏輯
```

尤其不要把：

```csharp
DataPropertyName = "RankDisplay"
```

改為：

```csharp
DataPropertyName = "Rank"
```

因為 `RankDisplay` 本來就是用來處理：

```text
排名
或
排名 / 排名母數
```

兩種顯示情況。

---

## 十、測試案例

### Case 1：學年分項排名只有 Rank

測試：

```text
Rank = 5
MatrixCount = null
```

預期：

```text
5
```

---

### Case 2：排名與排名母數都有資料

測試：

```text
Rank = 5
MatrixCount = 30
```

預期：

```text
5 / 30
```

---

### Case 3：兩者皆空

```text
Rank = null
MatrixCount = null
```

預期：

```text
空白
```

---

### Case 4：Rank 空白但 MatrixCount 有值

```text
Rank = null
MatrixCount = 30
```

預期：

```text
空白
```

不可顯示：

```text
/ 30
```

---

### Case 5：重新開啟 frmStudentRankDetail

確認：

```text
UCStudentRank
    ↓
frmStudentRankDetail
    ↓
GetStudentRankDetail()
    ↓
Rank / MatrixCount
    ↓
RankDisplay
    ↓
DataGrid 排名
```

資料正確顯示。

---

## 十一、確認其他欄位不受影響

修改完成後確認 DataGrid 其他欄位仍可正常載入：

```text
成績類別
排名方式
排名分數
排名範圍
母群
PR值
百分比
```

此次不要順便修改上述欄位邏輯。

若發現其他欄位仍有問題，請記錄問題，但不要擴大此次修改範圍。

---

## 十二、Build 檢查

完成修改後：

* [ ] Solution / Project 可以正常 Build。
* [ ] 沒有 Compiler Error。
* [ ] `RankDisplay` 可以正常編譯。
* [ ] 沒有 NullReferenceException。
* [ ] 沒有改變既有 DB Schema。
* [ ] 沒有影響學年分項排名匯入。
* [ ] 沒有影響有 `matrix_count` 的既有排名資料。

---

## 十三、修改完成紀錄

修改完成後，追加記錄至：

```text
學生排名資料調整0812.md
```

記錄至少包含：

### 修改目標

調整學生排名資料的排名顯示方式。

### 修改檔案

列出實際修改的檔案。

### 修改前問題

原本 `RankDisplay` 在只有：

```text
Rank
```

沒有：

```text
MatrixCount
```

時，無法正確顯示排名。

### 修改後規則

```text
Rank 有值 + MatrixCount 無值
→ 顯示 Rank

Rank 有值 + MatrixCount 有值
→ 顯示 Rank / MatrixCount

Rank 無值
→ 顯示空白
```

### 測試結果

記錄至少：

```text
只有排名
排名 + 排名母數
排名空白
學年分項排名
既有排名資料
```

的測試結果。

---

# 最終要求

本次採「最小修改」原則。

核心只調整：

```text
StudentRankDetailItem.RankDisplay
```

達成：

```text
只有排名：
RankDisplay = 排名

有排名與排名母數：
RankDisplay = 排名 / 排名母數

沒有排名：
RankDisplay = 空白
```

完成後記錄：

```text
學生排名資料調整0812.md
```
