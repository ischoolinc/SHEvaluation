# Todo：建立匯入學年分項排名功能

## 目標

在 `SHImportExportScoreRank` 專案中建立「匯入學年分項排名」功能。

實作完成後，請將修改內容、資料規則、SQL Upsert 規則、權限代碼及測試結果記錄在：

```text
匯入學年分項排名功能.md
```

---

## 一、參考程式

請參考 `LHDB_SH_Core.zip` 中以下檔案的程式架構：

```text
LHDB_SH_Core\ImportExport\ImportStudentLearningHistory4_2.cs
LHDB_SH_Core\ImportExport\ImportStudentLearningHistory4_2.xml
LHDB_SH_Core\Program.cs
LHDB_SH_Core\ValidationRule\StudentLearningHisRowValidatorFactory.cs
LHDB_SH_Core\ValidationRule\RowValidator\StudCheckStudentNumberStatusVal.cs
```

參考重點：

1. 使用 `Campus.Import2014.ImportWizard` 建立匯入精靈。
2. 使用 XML 定義匯入欄位、重複資料判斷及驗證規則。
3. 使用 `IRowValidatorFactory` 註冊自訂資料列驗證器。
4. 使用 `FISCA.MainMethod` 註冊模組啟動入口。
5. 使用 `RoleAclSource`、`RibbonFeature`、`UserAcl` 註冊及控制功能權限。
6. 使用 JSONB CTE 一次完成批次新增及更新。
7. 回傳新增筆數及更新筆數。

只參考架構，不可沿用 `student_learning_history` 的資料模型、欄位或 SQL。

本功能必須直接操作：

```text
rank_override
```

---

## 二、本次實作範圍

本次只實作：

```text
匯入學年分項排名
```

本次不要實作：

```text
匯出學年分項排名
學年分項排名資料項目
刪除學年分項排名
```

不要修改 `LHDB_SH_Core` 參考專案。

不要更動其他既有功能或主要架構。

---

## 三、資料表規格

資料表：

```text
rank_override
```

欄位：

```text
id               bigint
ref_student_id   bigint
school_year      integer
semester         integer
grade_year       integer
item_type        varchar
ref_exam_id      bigint
item_name        varchar
rank_type        varchar
rank_name        varchar
rank             integer
matrix_count     integer
extension        jsonb
```

---

## 四、固定資料規則

所有匯入資料固定寫入：

```text
semester   = -1
item_type  = 學年/分項成績
item_name  = 學業
rank_type  = 班排名
ref_exam_id = NULL
```

請建立共用常數類別，不可將固定值分散硬編碼在 SQL、匯入類別及驗證器中。

建議檔案：

```text
Constants\RankOverrideConstants.cs
```

建議內容：

```csharp
namespace SHImportExportScoreRank.Constants
{
    public static class RankOverrideConstants
    {
        public const int SchoolYearSemester = -1;
        public const string ItemType = "學年/分項成績";
        public const string ItemName = "學業";
        public const string RankType = "班排名";
    }
}
```

---

## 五、Excel 匯入檔規格

檔案格式：

```text
.xlsx
```

標準欄位順序：

```text
學號
班級
座號
科別
姓名
學年度
年級
學年學業排名
```

### 欄位與資料表對照

| Excel 欄位 | 用途 | `rank_override` 對應 |
|---|---|---|
| 學號 | 對應學生系統編號 | 查詢後寫入 `ref_student_id` |
| 班級 | 學年排名所屬班級 | `rank_name` |
| 座號 | 人工辨識及核對 | 不寫入 `rank_override` |
| 科別 | 人工辨識及核對 | 不寫入 `rank_override` |
| 姓名 | 學生身分核對 | 不寫入 `rank_override` |
| 學年度 | 排名所屬學年度 | `school_year` |
| 年級 | 排名所屬年級 | `grade_year` |
| 學年學業排名 | 學年學業班排名 | `rank` |

### 必要欄位

以下欄位必須存在且不可空白：

```text
學號
班級
姓名
學年度
年級
學年學業排名
```

以下欄位必須列在標準匯入格式中，但允許個別資料為空白：

```text
座號
科別
```

不要將 Excel 的「座號」、「科別」、「姓名」寫入 `rank_override.extension`。

---

## 六、學生對應規則

以 Excel「學號」查詢 `student.student_number`，取得：

```text
student.id
student.name
```

### 對應原則

1. 學號查無學生：資料列驗證錯誤，不可匯入。
2. 同一學號只找到一位學生：使用該學生 `id`。
3. 同一學號找到多個不同學生：資料列驗證錯誤，不可猜測或自動選擇。
4. Excel「姓名」與系統學生姓名不同：資料列驗證錯誤。
5. 不可使用目前班級、目前座號或目前科別強制判斷學生，因為本資料可能是歷史學年度資料。
6. 班級欄位是排名資料本身，直接寫入 `rank_name`，不可改成學生目前班級。

批次匯入前，先收集所有學號，一次查詢學生資料並建立 Dictionary。

不可在每一筆資料的 `foreach` 中個別查詢學生。

---

## 七、匯入驗證規則

建立：

```text
ImportExport\ImportSchoolYearEntryRank.xml
```

### 重複資料判斷

同一個 Excel 檔案內，以下組合不可重複：

```text
學號 + 學年度
```

不要將年級放入檔案內重複判斷鍵。

理由：年級屬於可修正資料；若原年級錯誤，重新匯入應更新原資料，不應新增另一筆。

### 欄位驗證

#### 學號

- 不可空白。
- 必須在系統中存在。
- 同一學號若對應多個學生，視為錯誤。

#### 班級

- 不可空白。
- 寫入前必須 `Trim()`。
- 寫入 `rank_name`。

#### 姓名

- 不可空白。
- 必須與系統學生姓名一致。
- 比對前雙方都執行 `Trim()`。

#### 學年度

- 不可空白。
- 必須為整數。
- 建議合法範圍：`1～3000`。

#### 年級

- 不可空白。
- 必須為整數。
- 必須大於 `0`。
- 建議合法範圍：`1～9`，不要在程式中硬限制只能 `1～3`。

#### 學年學業排名

- 不可空白。
- 必須為整數。
- 必須大於 `0`。

#### 座號

- 可空白。
- 有值時必須為整數。
- 只作辨識用，不寫入資料表。

#### 科別

- 可空白。
- 只作辨識用，不寫入資料表。

---

## 八、自訂驗證器

建立：

```text
ValidationRule\SchoolYearEntryRankRowValidatorFactory.cs
ValidationRule\RowValidator\CheckSchoolYearEntryRankStudentValidator.cs
```

建議自訂驗證器 Type 名稱：

```text
SHIMPORTEXPORTSCORERANKCHECKSTUDENT
```

驗證器負責：

1. 學號是否存在。
2. 學號是否只對應一位學生。
3. Excel 姓名是否與系統姓名一致。

驗證器建立時一次載入需要的學生對照資料，不可在 `Validate()` 中逐筆執行 SQL。

在 `Program.Main()` 中註冊：

```csharp
FactoryProvider.RowFactory.Add(
    new ValidationRule.SchoolYearEntryRankRowValidatorFactory());
```

避免重複註冊；`Program` 必須有初始化旗標。

---

## 九、匯入類別

建立：

```text
ImportExport\ImportSchoolYearEntryRank.cs
```

類別架構：

```csharp
namespace SHImportExportScoreRank.ImportExport
{
    public class ImportSchoolYearEntryRank
        : Campus.Import2014.ImportWizard
    {
    }
}
```

### 建構子

```csharp
public ImportSchoolYearEntryRank()
{
    this.IsSplit = false;
}
```

### 支援動作

只支援新增或更新：

```csharp
public override ImportAction GetSupportActions()
{
    return ImportAction.InsertOrUpdate;
}
```

本次不要增加「刪除」欄位及刪除流程。

### GetValidateRule

從嵌入資源讀取：

```text
SHImportExportScoreRank.ImportExport.ImportSchoolYearEntryRank.xml
```

不要依賴外部實體 XML 路徑，避免部署後找不到驗證檔。

請在 `.csproj` 將 XML 設定為：

```xml
<EmbeddedResource Include="ImportExport\ImportSchoolYearEntryRank.xml" />
```

### Prepare

保存 `ImportOption`：

```csharp
private ImportOption _option;

public override void Prepare(ImportOption option)
{
    _option = option;
}
```

### Import

建議處理流程：

1. 設定 `ImportProgress = 5`。
2. 收集全部 Excel 學號。
3. 一次查詢學生資料。
4. 將每一列轉成 DTO。
5. 寫入前再次防禦性驗證必要值。
6. 將 DTO 序列化為 JSON。
7. 使用單一 JSONB CTE 執行批次更新及新增。
8. 取得新增筆數及更新筆數。
9. 設定 `ImportProgress = 100`。
10. 寫入 ApplicationLog。
11. 回傳執行摘要。

不要在 `foreach` 中逐筆執行 `INSERT` 或 `UPDATE`。

---

## 十、DTO 與 DAO

建立：

```text
DAO\SchoolYearEntryRankImportRecord.cs
DAO\StudentLookupRecord.cs
DAO\RankOverrideDataAccess.cs
```

### SchoolYearEntryRankImportRecord

至少包含：

```csharp
public long RefStudentID { get; set; }
public int SchoolYear { get; set; }
public int GradeYear { get; set; }
public string RankName { get; set; }
public int Rank { get; set; }
```

座號、科別、姓名不需要放入實際寫入 DTO；若為記錄匯入錯誤，可另外建立顯示欄位，但不可寫入 `rank_override`。

### StudentLookupRecord

至少包含：

```csharp
public long StudentID { get; set; }
public string StudentNumber { get; set; }
public string StudentName { get; set; }
```

### RankOverrideDataAccess

負責：

1. 批次取得學生對照資料。
2. 檢查資料庫中是否存在重複自然鍵資料。
3. 執行批次 Upsert。
4. 回傳新增、更新筆數。

不要將大型 SQL 直接全部放在 `Program.cs`。

---

## 十一、Upsert 自然鍵

判斷資料是否已存在時，使用：

```text
ref_student_id
+ school_year
+ semester = -1
+ item_type = 學年/分項成績
+ item_name = 學業
+ rank_type = 班排名
+ ref_exam_id IS NULL
```

不要將以下欄位放入自然鍵：

```text
grade_year
rank_name
rank
```

原因：以上欄位可能需要透過重新匯入修正。

### 既有資料更新欄位

找到既有資料時，只更新：

```text
grade_year
rank_name
rank
```

更新時必須保留：

```text
matrix_count
extension
```

不要清空或覆蓋既有 `matrix_count`、`extension`。

### 新增資料欄位

新增時寫入：

```text
ref_student_id = 由學號取得
school_year = Excel 學年度
semester = -1
grade_year = Excel 年級
item_type = 學年/分項成績
ref_exam_id = NULL
item_name = 學業
rank_type = 班排名
rank_name = Excel 班級
rank = Excel 學年學業排名
extension = '{}'::jsonb
```

`matrix_count` 本次 Excel 沒有提供：

- INSERT 時不要自行填入 `0`。
- 若欄位允許 NULL，保持 NULL。
- 若資料表有預設值，使用資料表預設值。
- 若實際 Schema 為 NOT NULL 且無預設值，停止自行猜值，將此問題記錄在完成文件中。

---

## 十二、資料庫既有重複資料處理

執行 Upsert 前，檢查同一自然鍵是否已存在兩筆以上資料。

若同一自然鍵在 `rank_override` 已有多筆：

1. 不可一次更新全部重複資料。
2. 不可自動刪除多餘資料。
3. 該筆匯入應回報錯誤或中止整批匯入。
4. 回傳訊息需包含學生 ID、學號及學年度，方便人工處理。

---

## 十三、批次 SQL 規則

SQL 必須相容：

```text
PostgreSQL 9.6
PostgreSQL 16
PostgreSQL 17
```

不可使用：

```text
MERGE
```

建議使用：

```text
jsonb_to_recordset
WITH CTE
UPDATE ... FROM
INSERT ... WHERE NOT EXISTS
RETURNING
```

單一 SQL 應回傳：

```text
inserted_count
updated_count
```

### SQL 安全

不可直接將未處理的 Excel 字串串接進 SQL。

JSON 序列化後至少必須正確處理單引號；優先使用參數化查詢。

若現有 `FISCA.Data.QueryHelper` 不支援參數：

1. 集中建立 SQL Literal Escape Helper。
2. JSON 文字中的單引號必須轉成兩個單引號。
3. 不可在多個地方自行 `Replace`。

---

## 十四、功能權限註冊

建立：

```text
Permissions\FeatureCode.cs
```

固定使用以下 FeatureCode：

```text
6134D009-48C3-484A-912E-5460804AB33B
```

建議常數：

```csharp
namespace SHImportExportScoreRank.Permissions
{
    public static class FeatureCode
    {
        public const string ImportSchoolYearEntryRank =
            "6134D009-48C3-484A-912E-5460804AB33B";
    }
}
```

### 權限設定畫面路徑

```text
學生
└─ 學年分項排名
   └─ 匯入學年分項排名
```

註冊類型：

```csharp
RibbonFeature
```

權限判斷：

```csharp
UserAcl.Current[FeatureCode.ImportSchoolYearEntryRank].Executable
```

---

## 十五、Ribbon 功能路徑

功能位置：

```text
學生
└─ 資料統計
   └─ 匯入
      └─ 成績相關匯入
         └─ 匯入學年分項排名
```

建議註冊：

```csharp
RibbonBarButton importRoot =
    MotherForm.RibbonBarItems["學生", "資料統計"]["匯入"];

RibbonBarButton importButton =
    importRoot["成績相關匯入"]["匯入學年分項排名"];
```

按鈕啟用：

```csharp
importButton.Enable =
    UserAcl.Current[
        FeatureCode.ImportSchoolYearEntryRank
    ].Executable;
```

Click 事件內必須再次檢查權限：

```csharp
if (!UserAcl.Current[
        FeatureCode.ImportSchoolYearEntryRank
    ].Executable)
{
    MsgBox.Show("您沒有「匯入學年分項排名」權限。");
    return;
}
```

有權限才啟動：

```csharp
ImportExport.ImportSchoolYearEntryRank importer =
    new ImportExport.ImportSchoolYearEntryRank();

importer.Execute();
```

不要只依賴按鈕的 `Enable` 狀態。

---

## 十六、Program.cs

目前 `Program.cs` 是空白類別，請補上正式啟動入口。

建議架構：

```csharp
using Campus.DocumentValidator;
using FISCA.Permission;
using FISCA.Presentation;
using FISCA.Presentation.Controls;
using SHImportExportScoreRank.Permissions;
using System;

namespace SHImportExportScoreRank
{
    public class Program
    {
        private static bool _initialized;

        [FISCA.MainMethod()]
        public static void Main()
        {
            if (_initialized)
                return;

            _initialized = true;

            RegisterPermissions();
            RegisterImportFeature();
            RegisterValidators();
        }

        private static void RegisterPermissions()
        {
            Catalog catalog =
                RoleAclSource.Instance["學生"]["學年分項排名"];

            catalog.Add(
                new RibbonFeature(
                    FeatureCode.ImportSchoolYearEntryRank,
                    "匯入學年分項排名"));
        }

        private static void RegisterImportFeature()
        {
            RibbonBarButton importRoot =
                MotherForm.RibbonBarItems["學生", "資料統計"]["匯入"];

            RibbonBarButton importButton =
                importRoot["成績相關匯入"]["匯入學年分項排名"];

            importButton.Enable =
                UserAcl.Current[
                    FeatureCode.ImportSchoolYearEntryRank
                ].Executable;

            importButton.Click += delegate
            {
                if (!UserAcl.Current[
                        FeatureCode.ImportSchoolYearEntryRank
                    ].Executable)
                {
                    MsgBox.Show("您沒有「匯入學年分項排名」權限。");
                    return;
                }

                try
                {
                    ImportExport.ImportSchoolYearEntryRank importer =
                        new ImportExport.ImportSchoolYearEntryRank();

                    importer.Execute();
                }
                catch (Exception ex)
                {
                    FISCA.LogAgent.ApplicationLog.Log(
                        "匯入學年分項排名",
                        "啟動失敗",
                        ex.ToString());

                    MsgBox.Show(
                        "啟動匯入學年分項排名功能失敗：" +
                        ex.Message);
                }
            };
        }

        private static void RegisterValidators()
        {
            FactoryProvider.RowFactory.Add(
                new ValidationRule
                    .SchoolYearEntryRankRowValidatorFactory());
        }
    }
}
```

本功能不需要建立 UDT，因此不要為了參考專案架構加入不必要的 `BackgroundWorker`。

---

## 十七、ApplicationLog

匯入完成後寫入：

```text
功能名稱：匯入學年分項排名
動作：匯入
```

內容至少包含：

```text
匯入檔資料筆數
新增筆數
更新筆數
執行帳號
執行時間
```

失敗時記錄：

```text
Exception.Message
Exception.StackTrace
```

Log 中不要寫入完整學生敏感資料；需要定位時只記錄學號及學年度。

---

## 十八、回傳訊息

匯入成功後回傳範例：

```text
匯入學年分項排名完成。
新增筆數：10
更新筆數：25
```

若沒有任何新增或更新：

```text
匯入學年分項排名完成。
新增筆數：0
更新筆數：0
```

不可顯示成匯入失敗。

---

## 十九、專案檔調整

目前 `SHImportExportScoreRank.csproj` 只包含基本 .NET 參考及 `Program.cs`。

請依實際使用情況，從參考專案複製相同版本及 HintPath 的必要 Reference：

```text
Campus.DocumentValidator
Campus.Import2014
FISCA
FISCA.Data
FISCA.LogAgent
FISCA.Permission
FISCA.Presentation
Newtonsoft.Json
System.Windows.Forms
```

若自訂驗證器實際需要，再加入：

```text
Campus.Validator2014
```

不要加入本功能未使用的報表、Aspose、學習歷程或 UDT 相關 DLL。

本專案為舊式 `.csproj`，新增檔案後必須明確加入：

```xml
<Compile Include="..." />
<EmbeddedResource Include="..." />
```

不可只建立實體檔案而未加入專案，否則不會編譯或部署。

---

## 二十、建議檔案結構

```text
SHImportExportScoreRank
├─ Constants
│  └─ RankOverrideConstants.cs
├─ DAO
│  ├─ RankOverrideDataAccess.cs
│  ├─ SchoolYearEntryRankImportRecord.cs
│  └─ StudentLookupRecord.cs
├─ ImportExport
│  ├─ ImportSchoolYearEntryRank.cs
│  └─ ImportSchoolYearEntryRank.xml
├─ Permissions
│  └─ FeatureCode.cs
├─ ValidationRule
│  ├─ SchoolYearEntryRankRowValidatorFactory.cs
│  └─ RowValidator
│     └─ CheckSchoolYearEntryRankStudentValidator.cs
├─ Program.cs
├─ Properties
│  └─ AssemblyInfo.cs
└─ SHImportExportScoreRank.csproj
```

---

## 二十一、不可變更事項

1. 不可修改 `rank_override` 資料表結構。
2. 不可更新其他 `item_type`、`item_name` 或 `rank_type` 的資料。
3. 不可將 `semester` 寫成 `0`、`1` 或 `2`；本功能固定為 `-1`。
4. 不可將「年級」省略或改由學生目前年級取代。
5. 不可將「班級」改由學生目前班級取代。
6. 不可將 `grade_year` 放入 Upsert 自然鍵。
7. 不可將 `rank_name` 放入 Upsert 自然鍵。
8. 不可將 `rank` 放入 Upsert 自然鍵。
9. 不可清空既有 `matrix_count`。
10. 不可清空既有 `extension`。
11. 不可先 DELETE 再 INSERT 來模擬更新。
12. 不可在迴圈中逐筆查詢學生或逐筆寫入排名。
13. 不可使用 PostgreSQL `MERGE`。
14. 不可自動刪除資料庫內既有重複資料。
15. 不可實作本次範圍以外的匯出及學生資料項目功能。

---

## 二十二、測試項目

完成後至少測試以下情境。

### 權限測試

- [ ] 有權限時可看到並啟用功能。
- [ ] 無權限時按鈕不可使用。
- [ ] Click 事件內仍有第二次權限檢查。
- [ ] 權限設定畫面可看到「學生／學年分項排名／匯入學年分項排名」。

### Excel 測試

- [ ] 可讀取 `.xlsx`。
- [ ] 八個標準欄位可正確辨識。
- [ ] 欄位順序不同時仍可依欄名匯入。
- [ ] 缺少必要欄位時不可執行匯入。

### 驗證測試

- [ ] 學號空白。
- [ ] 學號不存在。
- [ ] 同一學號對應多位學生。
- [ ] 姓名與系統姓名不同。
- [ ] 班級空白。
- [ ] 學年度空白或非整數。
- [ ] 年級空白、非整數或小於 1。
- [ ] 學年學業排名空白、非整數、0 或負數。
- [ ] 同一 Excel 有重複「學號＋學年度」。

### 新增測試

- [ ] 新資料正確寫入 `rank_override`。
- [ ] `semester = -1`。
- [ ] `item_type = 學年/分項成績`。
- [ ] `item_name = 學業`。
- [ ] `rank_type = 班排名`。
- [ ] `ref_exam_id IS NULL`。
- [ ] Excel 年級寫入 `grade_year`。
- [ ] Excel 班級寫入 `rank_name`。
- [ ] Excel 學年學業排名寫入 `rank`。

### 更新測試

- [ ] 相同學生及學年度再次匯入時更新原資料。
- [ ] 修正年級時不新增重複資料。
- [ ] 修正班級時不新增重複資料。
- [ ] 修正排名時不新增重複資料。
- [ ] 更新時保留原 `matrix_count`。
- [ ] 更新時保留原 `extension`。

### 安全及相容性測試

- [ ] 班級名稱包含單引號時 SQL 不會失敗。
- [ ] 中文班級名稱可正常寫入。
- [ ] 大量資料為單次批次 SQL，不是逐筆寫入。
- [ ] PostgreSQL 9.6 可執行。
- [ ] PostgreSQL 16 可執行。
- [ ] PostgreSQL 17 可執行。
- [ ] 專案可在 .NET Framework 4.8 編譯。

---

## 二十三、完成紀錄

完成後建立：

```text
匯入學年分項排名功能.md
```

文件至少記錄：

1. 修改及新增的檔案。
2. FeatureCode。
3. Ribbon 功能路徑。
4. 權限設定路徑。
5. Excel 欄位規格。
6. 必要欄位及驗證規則。
7. `rank_override` 欄位對照。
8. 固定資料規則。
9. Upsert 自然鍵。
10. 更新時保留的欄位。
11. JSONB CTE SQL 實作方式。
12. 新增、更新筆數的取得方式。
13. ApplicationLog 寫入內容。
14. 測試資料及測試結果。
15. 尚未處理或需人工確認的問題。

---

## 驗收標準

完成後必須符合：

```text
使用者可在：
學生 → 資料統計 → 匯入 → 成績相關匯入 → 匯入學年分項排名

選擇 .xlsx 檔案，依：
學號、班級、座號、科別、姓名、學年度、年級、學年學業排名

將學年學業排名新增或更新至 rank_override。
```

資料必須固定符合：

```text
semester = -1
item_type = 學年/分項成績
item_name = 學業
rank_type = 班排名
ref_exam_id IS NULL
```

並且相同學生、相同學年度重新匯入時，必須更新原資料，不可新增重複資料。
