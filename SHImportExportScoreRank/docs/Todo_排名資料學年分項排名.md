
# Todo - 排名資料：學年分項排名讀取

## 目標

調整學生資料項目「排名資料」中的學年分項排名讀取與顯示。

本次只處理：

- rank_override 學年分項排名資料讀取
- extension JSONB 排名資料資訊讀取
- UCStudentRank.lvData 顯示欄位調整

完成修改後，請將修改內容記錄到：

`排名資料-學年分項排名調整0811.md`

---

# 1. 修改範圍

主要檢查與修改：

- `DetailContent/UCStudentRank.cs`
- `DAO/RankOverrideDataAccess.cs`
- `SchoolYearEntryRankRecord` 的 class 定義檔案

如需調整其他檔案，只能修改本功能直接需要的資料 Model，
不要改動其他匯入、匯出或排名計算功能。

---

# 2. 學年分項排名資料判斷規則

本次讀取的是 `rank_override` 中的「學年分項排名」。

學年資料規則：

```text
semester = -1
ref_exam_id = -1
````

注意：

目前 `GetSchoolYearEntryRanksByStudentID()` 使用：

```sql
AND ro.ref_exam_id IS NULL
```

這個條件不符合目前學年排名資料格式。

請修改為：

```sql
AND ro.ref_exam_id = -1
```

不要將學年資料的 `ref_exam_id` 判斷成 NULL。

---

# 3. Extension 排名資料讀取

`rank_override.extension` 為 JSONB Array。

排名資料 extension 格式目前為：

```json
[
    {
        "extension_name": "排名資料",
        "create_time": "...",
        "建立方式": "匯入",
        "成績類型": "學年",
        "成績項目": "分項"
    }
]
```

讀取時只取：

```text
extension_name = 排名資料
```

可參考以下 PostgreSQL 寫法：

```sql
SELECT
    ro.id,
    ext->>'extension_name' AS extension_name,
    ext->>'create_time' AS create_time,
    ext->>'建立方式' AS "建立方式",
    ext->>'成績類型' AS "成績類型",
    ext->>'成績項目' AS "成績項目"
FROM rank_override ro
CROSS JOIN LATERAL
    jsonb_array_elements(
        COALESCE(ro.extension, '[]'::jsonb)
    ) AS ext
WHERE ext->>'extension_name' = '排名資料'
ORDER BY ro.id DESC;
```

請將此概念整合到：

```csharp
GetSchoolYearEntryRanksByStudentID()
```

不要另外建立不必要的第二次資料庫查詢。

---

# 4. 學年分項排名查詢條件

`GetSchoolYearEntryRanksByStudentID(studentID)` 必須：

1. 只讀取目前學生：

   ```sql
   ro.ref_student_id = studentID
   ```

2. 學年資料：

   ```sql
   ro.semester = -1
   ```

3. 學年沒有評量 ID：

   ```sql
   ro.ref_exam_id = -1
   ```

4. extension 必須是：

   ```sql
   ext->>'extension_name' = '排名資料'
   ```

5. 保留目前學年分項排名相關的：

   * item_type
   * item_name
   * rank_type
     等既有條件，避免讀到其他種類的 rank_override 資料。

不要因本次修改而擴大到其他排名資料類型。

---

# 5. lvData 欄位調整

目前 `UCStudentRank.SetupColumns()` 為：

```text
學年度
年級
班級
學年學業排名
```

請改成以下固定 6 欄：

```text
學年度
成績類型
成績項目
排名範圍
建立時間
建立方式
```

建議寬度可依目前畫面適度設定，例如：

```csharp
AddColumn("學年度", 80);
AddColumn("成績類型", 90);
AddColumn("成績項目", 90);
AddColumn("排名範圍", 110);
AddColumn("建立時間", 140);
AddColumn("建立方式", 90);
```

如畫面大小需要，可微調 Width，但不要增加其他欄位。

---

# 6. 六個欄位的資料來源

請嚴格依以下來源顯示：

| lvData 欄位 | rank_override / extension 來源 |
| --------- | ---------------------------- |
| 學年度       | `ro.school_year`             |
| 成績類型      | `ext->>'成績類型'`               |
| 成績項目      | `ext->>'成績項目'`               |
| 排名範圍      | `ro.rank_type`               |
| 建立時間      | `ext->>'create_time'`        |
| 建立方式      | `ext->>'建立方式'`               |

注意：

### 成績類型

不要使用：

```text
ro.item_type
```

畫面「成績類型」必須讀取：

```sql
ext->>'成績類型'
```

例如：

```text
學年
```

### 成績項目

不要直接顯示：

```text
ro.item_name
```

畫面「成績項目」必須讀取：

```sql
ext->>'成績項目'
```

例如：

```text
分項
```

### 排名範圍

使用：

```text
ro.rank_type
```

例如：

```text
班排名
```

### 建立方式

使用：

```sql
ext->>'建立方式'
```

例如：

```text
匯入
```

---

# 7. 建立時間

「建立時間」來源固定為：

```sql
ext->>'create_time'
```

目前 extension 寫入方式為：

```sql
'create_time', now()
```

因此可能取得：

```text
2026-08-11T00:37:34.601675+08:00
```

C# 顯示時請安全解析時間。

建議使用：

```csharp
DateTimeOffset.TryParse(...)
```

解析成功後顯示成容易閱讀的格式，例如：

```text
2026/8/11 00:37
```

格式：

```text
yyyy/M/d HH:mm
```

如果舊資料或異常資料無法解析：

* 不要拋出 Exception
* 保留原始 create_time 文字，或安全顯示空白

不要因單筆 create_time 格式異常造成整個排名資料頁無法顯示。

---

# 8. SchoolYearEntryRankRecord

檢查 `SchoolYearEntryRankRecord`。

為支援新的 lvData 六欄，可新增必要的顯示屬性，例如：

```csharp
public string ScoreType { get; set; }
public string ScoreItem { get; set; }
public string RankType { get; set; }
public string CreateTime { get; set; }
public string CreateMethod { get; set; }
```

命名可以依目前專案 Coding Style 微調，但用途必須清楚。

不要任意刪除現有：

```text
ID
StudentID
SchoolYear
GradeYear
RankName
Rank
```

等既有欄位。

這些資料後續可能仍需用於排名詳細內容。

---

# 9. MapRankRecord

修改：

```csharp
MapRankRecord(...)
```

使 SQL 查詢出的：

```text
score_type
score_item
rank_type
create_time
create_method
```

可以正確 mapping 到 `SchoolYearEntryRankRecord`。

NULL 值必須安全處理。

可以沿用目前：

```csharp
NullToEmpty(...)
```

的作法。

---

# 10. BindListView

修改：

```csharp
UCStudentRank.BindListView()
```

目前顯示：

```text
SchoolYear
GradeYear
RankName
Rank
```

請改成依序顯示：

```text
SchoolYear
ScoreType
ScoreItem
RankType
CreateTime
CreateMethod
```

仍保留：

```csharp
item.Tag = record;
```

方便後續雙擊或「修改」功能取得完整 `SchoolYearEntryRankRecord`。

本次不要實作詳細資料視窗。

---

# 11. 不要修改的功能

本次目標只有：

「排名資料 - 學年分項排名讀取與主清單顯示」

不要修改：

* 匯入 Excel 欄位
* 匯出 Excel 欄位
* UpsertSchoolYearEntryRanks 的主要寫入流程
* rank_override 資料表結構
* extension JSON 結構
* FeatureCode
* 權限架構
* Ribbon
* Program.cs 功能註冊
* 排名計算邏輯
* 新增功能
* 修改功能
* 刪除功能
* 排名詳細內容視窗

不要進行不相關的 Refactor。

---

# 12. 相容性

維持目前專案技術環境。

注意：

* .NET Framework 相容性
* PostgreSQL JSONB 語法相容性
* 不導入不必要的新 NuGet package
* 優先沿用目前 FISCA.Data.QueryHelper
* 優先沿用目前 DAO / Record 架構
* 保留 BackgroundWorker 非同步讀取方式

---

# 13. 驗證案例

至少驗證一筆由「匯入學年分項排名」建立的資料。

假設資料：

```text
school_year = 113
semester = -1
ref_exam_id = -1
rank_type = 班排名
```

extension：

```json
[
    {
        "extension_name": "排名資料",
        "create_time": "2026-08-11T00:37:34.601675+08:00",
        "建立方式": "匯入",
        "成績類型": "學年",
        "成績項目": "分項"
    }
]
```

`lvData` 應顯示類似：

```text
113 | 學年 | 分項 | 班排名 | 2026/8/11 00:37 | 匯入
```

並確認：

1. `ref_exam_id = -1` 可以正常讀取。
2. 不再使用 `ref_exam_id IS NULL`。
3. extension_name 不是「排名資料」的 extension 不顯示。
4. 沒有 extension 排名資料的 rank_override 不應誤顯示。
5. 切換學生後 lvData 正常重新載入。
6. 沒有排名資料的學生顯示空清單，不發生錯誤。
7. create_time 無法解析時不造成整個畫面錯誤。

---

# 14. 完成後紀錄

完成後建立或更新：

`排名資料-學年分項排名調整0811.md`

內容至少記錄：

* 修改檔案
* 修改的方法名稱
* `ref_exam_id IS NULL` 改為 `ref_exam_id = -1`
* extension JSONB 查詢方式
* 六個 lvData 欄位與資料來源
* create_time 解析與顯示方式
* SchoolYearEntryRankRecord 新增的欄位
* 測試方式與測試結果
* 是否有未完成事項

---

# 完成條件

本 Todo 完成時，學生「排名資料」資料項目必須可以正確讀取學年分項排名，並在 lvData 顯示固定六欄：

`學年度、成績類型、成績項目、排名範圍、建立時間、建立方式`

且學年資料必須使用：

```text
semester = -1
ref_exam_id = -1
```

extension 資料必須使用：

```text
extension_name = 排名資料
```

本次只完成主清單讀取與顯示，不提前實作排名詳細內容及 CRUD 功能。

```

這份 Todo 特別把範圍鎖在「**讀取 + 六欄主清單**」，不讓 Cursor 順手去改匯入、詳細畫面或 CRUD。這樣比較適合你現在分階段調整。
```
