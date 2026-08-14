

## 目標

調整「匯入學年分項排名」功能。

當匯入資料新增或更新 `rank_override` 時，同步寫入 `rank_override.extension` JSONB 欄位。

完成修改後，將本次修改內容記錄到：

`匯入學年分項排名調整.md`

---

## 一、修改範圍

主要檢查並修改：

* `RankOverrideDataAccess.cs`
* `UpsertSchoolYearEntryRanks(...)`

參考現有匯入流程：

`ImportSchoolYearEntryRank.cs`

目前匯入最後會呼叫：

```csharp
dataAccess.UpsertSchoolYearEntryRanks(records);
```

請維持既有匯入架構，不要重新設計整個匯入流程。

---

## 二、extension 欄位規格

資料表：

```text
rank_override
```

欄位：

```text
extension
```

資料型態：

```text
jsonb
```

只有 `extension` 是 JSONB。

其他 `rank_override` 欄位維持原本資料型態，不要改成 JSONB。

---

## 三、extension 寫入內容

新增或更新資料時，`extension` 必須寫入：

```sql
jsonb_build_array(
    jsonb_build_object(
        'extension_name', '排名資料',
        'create_time', now(),
        '建立方式', '匯入',
        '成績類型', '學年',
        '成績項目', '分項'
    )
)
```

實際 JSON 結構應類似：

```json
[
  {
    "extension_name": "排名資料",
    "create_time": "資料庫 now() 產生的時間",
    "建立方式": "匯入",
    "成績類型": "學年",
    "成績項目": "分項"
  }
]
```

---

## 四、新增 INSERT 處理

找到 `UpsertSchoolYearEntryRanks(...)` 中新增 `rank_override` 的 SQL。

原本 INSERT 的其他欄位及邏輯全部保留。

只調整 `extension` 寫入內容。

INSERT 應包含：

```sql
extension
```

並寫入：

```sql
jsonb_build_array(
    jsonb_build_object(
        'extension_name', '排名資料',
        'create_time', now(),
        '建立方式', '匯入',
        '成績類型', '學年',
        '成績項目', '分項'
    )
)
```

若目前 INSERT 是：

```sql
'{}'::jsonb
```

或使用預設空 JSON，請改成上述內容。

---

## 五、更新 UPDATE 處理

找到 `UpsertSchoolYearEntryRanks(...)` 中更新既有 `rank_override` 的 SQL。

維持目前既有資料判斷及 UPDATE 條件，不要因本次需求任意修改自然鍵或 Upsert 判斷規則。

原本例如更新：

```text
grade_year
rank_name
rank
```

仍維持原本邏輯。

另外加入：

```sql
extension = jsonb_build_array(
    jsonb_build_object(
        'extension_name', '排名資料',
        'create_time', now(),
        '建立方式', '匯入',
        '成績類型', '學年',
        '成績項目', '分項'
    )
)
```

也就是：

* 新增資料：寫入 extension
* 更新資料：重新寫入 extension
* 更新時 `create_time` 使用本次匯入的 `now()`

---

## 六、extension 更新方式

本功能請採用「覆寫」方式：

```sql
extension = jsonb_build_array(...)
```

不要使用：

```sql
extension = extension || ...
```

原因：

同一筆學年分項排名重新匯入時，只需要保留本次最新的「排名資料」資訊。

不要因每次重新匯入而產生：

```json
[
  { "extension_name": "排名資料", ... },
  { "extension_name": "排名資料", ... },
  { "extension_name": "排名資料", ... }
]
```

同一筆資料應維持一組最新的「排名資料」。

---

## 七、不要修改的地方

除非編譯上確實需要，否則不要修改：

* `ImportSchoolYearEntryRank.cs` 的 Excel 欄位規則
* `SchoolYearEntryRankImportRecord.cs`
* 學號判斷
* 姓名比對
* 學年度判斷
* 年級判斷
* 排名驗證
* 原本 Insert / Update 筆數統計
* 原本重複自然鍵檢查
* 原本 Log 紀錄
* 原本匯入流程

`SchoolYearEntryRankImportRecord` 不需要新增 `Extension` 屬性。

extension 內容是系統固定產生，不是 Excel 匯入欄位。

---

## 八、查詢驗證 SQL

修改完成後，可以使用以下 SQL 驗證：

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

應可以看到：

```text
extension_name = 排名資料
建立方式       = 匯入
成績類型       = 學年
成績項目       = 分項
create_time    = 本次匯入時間
```

---

## 九、測試新增

準備一筆資料庫中尚不存在的學年分項排名資料並匯入。

確認：

1. `rank_override` 成功新增。
2. 原本排名欄位資料正確。
3. `extension` 不是 NULL。
4. `extension` 不是空 JSON。
5. `extension_name = 排名資料`。
6. `建立方式 = 匯入`。
7. `成績類型 = 學年`。
8. `成績項目 = 分項`。
9. `create_time` 有值。
10. 匯入結果新增筆數正確。

---

## 十、測試更新

將同一筆資料重新匯入，但修改「學年學業排名」。

確認：

1. 不新增第二筆 `rank_override`。
2. 原本資料執行 UPDATE。
3. `rank` 更新成新排名。
4. `extension` 仍有資料。
5. `extension_name = 排名資料`。
6. `create_time` 更新成第二次匯入時間。
7. extension 不會累積重複的「排名資料」object。
8. 匯入結果更新筆數正確。

---

## 十一、回歸測試

確認此次修改沒有影響：

* 學號查詢
* 學生姓名比對
* 學年度
* 年級
* 班級
* 學年學業排名
* 新增判斷
* 更新判斷
* 重複資料檢查
* Import Summary
* ApplicationLog
* `SchoolYearEntryRankDetailContent` 更新

---

## 十二、程式修改原則

1. 優先修改既有 `UpsertSchoolYearEntryRanks(...)` SQL。
2. 不大幅重構現有程式。
3. 不改變原本 Insert / Update 判斷邏輯。
4. 不增加不必要的 Model 欄位。
5. 不把一般欄位改成 JSONB。
6. 只有 `rank_override.extension` 使用 JSONB。
7. SQL 保持 PostgreSQL 相容。
8. 維持目前專案既有 C# coding style。
9. 修改後確認專案可以正常編譯。

---

## 十三、完成紀錄

修改完成後建立或更新：

`匯入學年分項排名調整.md`

內容至少記錄：

### 修改目標

匯入學年分項排名時，新增及更新 `rank_override.extension` 排名資料。

### 修改檔案

列出實際修改的檔案。

### INSERT 修改

記錄新增資料時如何寫入 extension。

### UPDATE 修改

記錄更新資料時如何重新寫入 extension。

### extension JSON 格式

```json
[
  {
    "extension_name": "排名資料",
    "create_time": "now()",
    "建立方式": "匯入",
    "成績類型": "學年",
    "成績項目": "分項"
  }
]
```

### 驗證結果

記錄：

* 新增測試
* 更新測試
* extension 查詢結果
* Insert / Update 筆數
* 編譯結果

### 注意事項

說明 extension 採覆寫方式，不累積重複「排名資料」JSON object。
