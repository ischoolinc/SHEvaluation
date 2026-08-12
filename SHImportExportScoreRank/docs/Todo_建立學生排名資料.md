## 目標

建立 `frmStudentRankDetail` 學生排名資料明細功能。

依目前學生排名資料畫面設計，完成：

1. 從 `UCStudentRank` 開啟學生排名明細。
2. 顯示選定學生、學年度的排名基本資料。
3. 顯示 `rank_override` 的排名明細。
4. 支援「排名方式」顯示與必要的修改功能。
5. 完成儲存及離開功能。
6. 不破壞目前既有排名匯入、匯出及資料項目功能。
7. 修改完成後，將實作內容記錄在：

```text
學生排名資料調整0812.md
```

---

# 一、修改原則

* [ ] 先閱讀目前專案相關程式碼及資料流，再進行修改。
* [ ] 不大幅重構既有架構。
* [ ] 不修改與本功能無關的程式。
* [ ] 優先沿用現有 `RankOverrideDataAccess`。
* [ ] 不要將大量 SQL 直接寫在 `frmStudentRankDetail`。
* [ ] 資料存取統一放在 DataAccess 層。
* [ ] 保持目前 .NET Framework / WinForms 專案既有程式風格。
* [ ] 不任意改變目前 `rank_override` 資料結構。
* [ ] 不任意改變既有 extension JSON 結構。
* [ ] 不因本功能影響目前學年分項排名匯入功能。

---

# 二、確認目前程式

請先檢查：

```text
frmStudentRankDetail.cs
frmStudentRankDetail.Designer.cs
UCStudentRank.cs
UCStudentRank.Designer.cs
RankOverrideDataAccess.cs
```

並搜尋與排名資料相關的：

```text
rank_override
StudentRank
RankOverride
extension
ref_exam_id
matrix_count
rank_type
rank_name
item_type
item_name
```

確認目前既有 Model / DTO 是否已經能重複使用。

若已有相同用途的類別，優先沿用，不要重複建立。

---

# 三、frmStudentRankDetail 畫面

依設計畫面建立學生排名明細。

## 3.1 基本資料區

畫面上方顯示：

```text
學生排名資料

學年度
成績類型
成績項目
建立方式
建立時間
排名批次
```

建議控制項名稱：

```text
lblSchoolYear
lblScoreType
lblScoreItem
lblCreateType
lblCreateTime
lblBatchName
```

以上資料原則上皆為唯讀顯示。

---

# 四、基本資料來源

## 4.1 學年度

來源：

```text
rank_override.school_year
```

---

## 4.2 成績類型

來源：

```text
rank_override.extension
```

讀取：

```text
extension_name = 排名資料
成績類型
```

目前學年分項資料預期：

```text
成績類型 = 學年
```

---

## 4.3 成績項目

來源：

```text
rank_override.extension
```

目前學年分項資料預期：

```text
成績項目 = 分項
```

---

## 4.4 建立方式

讀取：

```text
extension -> 建立方式
```

例如：

```text
匯入
```

---

## 4.5 建立時間

必須讀取：

```text
extension -> create_time
```

例如：

```text
2026-08-11T00:37:34.601675+08:00
```

請安全解析日期時間。

若解析成功，畫面顯示較容易閱讀的格式，例如：

```text
2026/08/11 00:37
```

若資料為 null、空白或無法解析：

* 不可發生 Exception。
* 顯示空白或保留原字串。
* 不影響其他排名資料讀取。

---

# 五、排名明細 Grid

建立排名明細表格。

建議使用：

```text
dgvRankDetail
```

欄位順序：

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

建議欄位名稱：

```text
colScoreCategory
colRankMethod
colScore
colRankType
colRankName
colRank
colPR
colPercentage
```

---

# 六、排名資料欄位對應

開始實作前，請先依目前 `rank_override` 實際欄位確認資料來源。

目前已知重要欄位包括：

```text
id
ref_student_id
school_year
semester
grade_year
item_type
ref_exam_id
item_name
rank_type
rank_name
rank
matrix_count
extension
```

不要自行猜測不存在的資料欄位。

---

# 七、排名顯示

## 7.1 排名

資料庫原始資料：

```text
rank
matrix_count
```

畫面「排名」欄位組合顯示：

```text
rank / matrix_count
```

例如：

```text
15/25
21/63
```

不要為了畫面顯示而改變資料庫原始欄位。

---

# 八、排名方式

畫面需要支援：

```text
原始
擇優
```

請先搜尋目前排名資料中「原始」、「擇優」實際儲存位置。

可能來源需依現有程式確認，例如：

```text
item_name
rank_type
rank_name
extension
其他既有欄位
```

## 重要

* [ ] 不可尚未確認資料結構就自行新增 DB 欄位。
* [ ] 不可自行改變 `rank_override` Schema。
* [ ] 若專案目前已有原始／擇優定義，必須沿用既有定義。
* [ ] 將確認結果記錄在最後的調整文件。

---

# 九、排名分數、PR值、百分比

請搜尋現有排名資料 Model、SQL 或相關排名功能，確認：

```text
排名分數
PR值
百分比
```

實際資料來源。

若 `rank_override` 本身沒有這些欄位：

* 不要自行假造資料。
* 請確認是否由現有欄位計算。
* 請確認是否有其他排名資料表。
* 請確認既有排名功能是否已有計算方法。

必須優先重複使用現有排名計算邏輯。

不要另外建立不同算法。

---

# 十、排名範圍與母群

依現有資料結構確認：

```text
rank_type
rank_name
```

與畫面的：

```text
排名範圍
母群
```

之間的對應。

例如可能顯示：

```text
班排 / 101
科排 / 普通科
年排 / 二年級
```

實際內容以目前資料庫既有定義為準。

---

# 十一、資料查詢條件

本次主要處理：

```text
學年分項排名
```

查詢時至少需考慮：

```text
ref_student_id
school_year
semester = -1
ref_exam_id = -1
```

extension 需要符合：

```text
extension_name = 排名資料
成績類型 = 學年
成績項目 = 分項
```

SQL 邏輯可參考目前專案既有寫法：

```sql
CROSS JOIN LATERAL
    jsonb_array_elements(
        COALESCE(ro.extension, '[]'::jsonb)
    ) AS ext
```

並依：

```sql
ext->>'extension_name'
ext->>'create_time'
ext->>'建立方式'
ext->>'成績類型'
ext->>'成績項目'
```

取得排名資訊。

但請優先整合到既有 `RankOverrideDataAccess`，不要讓 Form 自行負責 SQL。

---

# 十二、建立 Detail Data Model

若目前沒有合適的 Model，可以建立學生排名 Detail Model。

例如概念：

```text
StudentRankDetailData
    SchoolYear
    ScoreType
    ScoreItem
    CreateTime
    CreateType
    BatchName
    Items
```

明細概念：

```text
StudentRankDetailItem
    RankOverrideId
    ScoreCategory
    RankMethod
    Score
    RankType
    RankName
    Rank
    MatrixCount
    PR
    Percentage
```

實際型別及欄位必須依目前專案既有資料結構調整。

如果已有相同 Model，直接沿用。

---

# 十三、frmStudentRankDetail 參數

不要讓 `frmStudentRankDetail` 自己猜目前學生。

應由 `UCStudentRank` 傳入必要識別資料。

至少包含：

```text
StudentId
SchoolYear
```

若 UCStudentRank 已經取得：

```text
成績類型
成績項目
create_time
建立方式
```

可以考慮包成一個資料物件傳入。

避免 constructor 出現大量獨立參數。

---

# 十四、資料讀取流程

建議流程：

```text
UCStudentRank
    ↓
使用者選擇排名資料
    ↓
開啟 frmStudentRankDetail
    ↓
傳入 StudentId / SchoolYear / 排名類型資訊
    ↓
RankOverrideDataAccess
    ↓
查詢 rank_override
    ↓
解析 extension
    ↓
回傳 Detail Model
    ↓
frmStudentRankDetail 顯示 Header
    ↓
frmStudentRankDetail 顯示排名明細
```

---

# 十五、Load 方法

不要把所有程式寫在 Form Load Event。

建議拆分：

```text
LoadData()
BindHeader()
BindDetail()
```

例如：

```csharp
private void frmStudentRankDetail_Load(object sender, EventArgs e)
{
    LoadData();
}
```

由：

```text
LoadData()
```

統一取得資料，再交給 UI Binding。

---

# 十六、UCStudentRank 整合

檢查目前 `UCStudentRank` 的排名資料清單。

增加開啟 Detail Form 的方式。

可依目前 UI 設計使用：

```text
雙擊
按鈕
右鍵
```

優先沿用目前專案其他資料項目的操作習慣。

流程：

```text
選取一筆學生排名資料
    ↓
frmStudentRankDetail
    ↓
顯示該筆學年度 / 成績類型 / 成績項目
```

不要開啟後顯示其他學年度或其他排名類型資料。

---

# 十七、儲存功能

畫面包含：

```text
儲存
離開
```

建議控制項：

```text
btnSave
btnExit
```

第一版只允許修改真正有業務需求且已有資料結構支援的欄位。

## 原則

以下排名計算結果原則上先設為唯讀：

```text
排名分數
排名範圍
母群
排名
PR值
百分比
```

避免使用者直接修改計算結果。

如果「排名方式」確定可修改，則只有：

```text
原始
擇優
```

允許修改。

---

# 十八、排名方式控制項

如果排名方式可以修改，建議使用：

```text
DataGridViewComboBoxColumn
```

選項：

```text
原始
擇優
```

但只有在確認目前 DB 已有適當儲存位置之後才能實作更新。

---

# 十九、儲存資料

按下：

```text
btnSave
```

流程：

```text
檢查是否有修改
    ↓
驗證資料
    ↓
呼叫 RankOverrideDataAccess
    ↓
更新 rank_override
    ↓
成功後 DialogResult.OK
```

更新 SQL 不要直接寫在 Form。

---

# 二十、離開功能

按下：

```text
btnExit
```

如果沒有異動：

```text
直接 Close()
```

如果資料已經修改但尚未儲存：

提示：

```text
資料尚未儲存，是否確定離開？
```

避免誤關閉造成修改遺失。

---

# 二十一、Detail 儲存後更新 UCStudentRank

`UCStudentRank` 開啟 Detail Form 後：

若：

```text
DialogResult == OK
```

重新載入目前學生排名資料。

目的：

Detail 畫面修改完成後，外層清單立即反映最新資料。

---

# 二十二、錯誤處理

必須處理：

* [ ] 查不到排名資料。
* [ ] extension 為 null。
* [ ] extension 為空陣列。
* [ ] extension 中不存在 `排名資料`。
* [ ] create_time 為空白。
* [ ] create_time 格式無法解析。
* [ ] rank 為 null。
* [ ] matrix_count 為 null。
* [ ] rank_name 為空白。
* [ ] rank_type 為空白。
* [ ] DB Query 發生 Exception。
* [ ] 單筆異常資料不可造成整個 Form NullReferenceException。

---

# 二十三、UI 顯示

畫面需接近提供的設計圖：

```text
學生排名資料

學年度    113
成績類型  學年
成績項目  分項
建立方式  匯入

建立時間  2024/7/1 10:35
排名批次  Batch-20240701-001


成績類別 | 排名方式 | 排名分數 | 排名範圍 | 母群 | 排名 | PR值 | 百分比
---------------------------------------------------------------------------
學業分項 | 原始     | 70       | 班排     | 101  |15/25 | 15   | 15
學業分項 | 擇優     | 75       | 科排     |普通科|21/63 | 21   | 21
```

下方：

```text
儲存
離開
```

---

# 二十四、畫面尺寸

請確認：

* [ ] 視窗大小適合目前欄位。
* [ ] Grid 欄位不被截斷。
* [ ] Windows DPI 125% / 150% 下不嚴重跑版。
* [ ] 長文字可以正常顯示。
* [ ] 表頭與資料對齊。
* [ ] Form 開啟位置符合目前專案慣例。

---

# 二十五、測試案例

至少完成以下測試。

## Case 1：正常學年分項排名

資料：

```text
school_year = 113
semester = -1
ref_exam_id = -1
成績類型 = 學年
成績項目 = 分項
```

確認 Header 正確。

---

## Case 2：多筆排名

同一學生存在：

```text
班排名
科排名
年排名
```

確認全部正確顯示，不可只顯示第一筆。

---

## Case 3：原始／擇優

若 DB 同時存在：

```text
原始
擇優
```

確認 Grid 各自顯示正確。

---

## Case 4：extension 空值

確認：

```text
extension = []
```

不會發生 Exception。

---

## Case 5：create_time 含時區

測試：

```text
2026-08-11T00:37:34.601675+08:00
```

確認可以正常顯示。

---

## Case 6：排名母群

測試：

```text
rank = 15
matrix_count = 25
```

畫面：

```text
15/25
```

---

## Case 7：沒有排名資料

Form 必須可以正常開啟或顯示提示。

不可發生：

```text
NullReferenceException
```

---

## Case 8：UCStudentRank → Detail

從 `UCStudentRank` 開啟某筆：

```text
113 / 學年 / 分項
```

Detail 必須只顯示：

```text
113 / 學年 / 分項
```

不可混入其他學年度或其他排名類型。

---

# 二十六、不要修改

除非本功能確實需要，不要修改：

```text
學年分項排名匯入主要流程
Excel 匯入欄位
Excel 匯出欄位
rank_override 既有 Schema
現有 extension JSON 結構
其他學生基本資料功能
其他排名計算功能
其他 Ribbon 功能
其他專案
```

---

# 二十七、完成後程式碼檢查

修改完成後請自行檢查：

* [ ] 專案可以 Build。
* [ ] 沒有新增 Compiler Error。
* [ ] 沒有明顯 Warning。
* [ ] 沒有未使用變數。
* [ ] 沒有重複 Model。
* [ ] 沒有重複 SQL。
* [ ] 沒有把 DB 邏輯散落到 Form。
* [ ] 所有 null 都有適當處理。
* [ ] 不影響現有 `UCStudentRank` 載入。
* [ ] 不影響現有學年分項排名匯入。

---

# 二十八、完成紀錄

全部修改及測試完成後，建立：

```text
學生排名資料調整0812.md
```

內容至少記錄：

## 1. 修改目標

說明本次建立 `frmStudentRankDetail` 學生排名資料功能的目的。

## 2. 修改檔案

列出實際修改的：

```text
frmStudentRankDetail.cs
frmStudentRankDetail.Designer.cs
UCStudentRank.cs
RankOverrideDataAccess.cs
相關 Model
```

只列實際有修改的檔案。

## 3. 資料來源

清楚記錄：

```text
rank_override
extension
```

及各畫面欄位實際 DB 對應。

尤其需要記錄：

```text
排名方式
排名分數
排名範圍
母群
排名
PR值
百分比
```

最後實際來源。

## 4. 查詢條件

記錄：

```text
ref_student_id
school_year
semester
ref_exam_id
extension_name
成績類型
成績項目
```

實際使用條件。

## 5. extension 解析

記錄：

```text
create_time
建立方式
成績類型
成績項目
```

如何解析。

## 6. UI 修改

說明：

```text
Header
Grid
Save
Exit
```

實作方式。

## 7. 儲存功能

記錄：

* 哪些欄位允許修改。
* 更新到哪個 DB 欄位。
* 哪些欄位保持唯讀。

## 8. 測試結果

記錄實際測試：

```text
正常資料
多筆資料
空 extension
create_time 時區
無排名資料
原始／擇優
UCStudentRank 開啟 Detail
儲存後 Reload
```

## 9. 未完成或待確認事項

如果發現現有資料庫沒有提供：

```text
排名分數
PR值
百分比
排名方式
排名批次
```

等資料來源，不要自行設計 DB Schema。

請將調查結果及待確認事項記錄於此區。

---

# 最終要求

本次工作的核心目標是：

```text
建立 frmStudentRankDetail 學生排名資料明細功能，
正確讀取 rank_override 與 extension，
依指定學生、學年度、成績類型與成績項目顯示排名資料，
並依目前既有資料結構完成必要的排名方式修改與儲存功能。
```

請以「最小必要修改」原則完成。

不要為了完成 UI 而自行改變既有排名資料結構。

修改完成後務必建立：

```text
學生排名資料調整0812.md
```

完整記錄本次修改與測試結果。
