# Todo：建立匯出學年分項排名與學年分項排名資料項目

## 目標

在 `SHImportExportScoreRank` 專案中建立以下兩項功能：

```text
1. 匯出學年分項排名
2. 學年分項排名資料項目
```

本次不重新實作「匯入學年分項排名」。既有匯入功能、固定值、權限代碼及主要架構不可任意改寫。

完成後，請將修改檔案、功能路徑、權限代碼、SQL 條件、畫面欄位、測試結果及已知限制記錄在：

```text
匯出與資料項目學年分項排名功能.md
```

---

## 一、參考程式

請參考 `LHDB_SH_Core.zip` 中以下檔案的程式架構：

```text
LHDB_SH_Core\ImportExport\ExprotStudentLearningHistory4_2.cs
LHDB_SH_Core\ImportExport\ExportStudentV2.cs
LHDB_SH_Core\ImportExport\ExportStudentV2.Designer.cs
LHDB_SH_Core\ImportExport\ExportStudentV2.resx
LHDB_SH_Core\DetailContent\UCStudentLearningHistoryItem.cs
LHDB_SH_Core\DetailContent\UCStudentLearningHistoryItem.Designer.cs
LHDB_SH_Core\DetailContent\UCStudentLearningHistoryItem.resx
LHDB_SH_Core\Program.cs
```

參考重點：

1. 匯出類別繼承 `SmartSchool.API.PlugIn.Export.Exporter`。
2. 使用 `ExportStudentV2` 選取學生及輸出 `.xlsx`。
3. 使用 `ExportableFields` 註冊匯出欄位。
4. 使用 `ExportPackage` 依 `e.List` 批次查詢選取學生資料。
5. 學生資料項目繼承 `FISCA.Presentation.DetailContent`。
6. 使用 `BackgroundWorker` 非同步讀取資料，避免切換學生時卡住畫面。
7. 使用 `OnPrimaryKeyChanged` 重新載入目前學生資料。
8. 使用 `FISCA.Features.TryRegister` 建立重新整理入口，避免重複註冊。
9. 使用 `RoleAclSource`、`RibbonFeature`、`DetailItemFeature` 及 `UserAcl` 處理功能權限。
10. 使用 `K12.Presentation.NLDPanels.Student.AddDetailBulider` 註冊學生資料項目。

只參考程式架構，不可沿用 `student_learning_history` 的資料模型、欄位、serial_no 或 SQL。

本功能必須直接讀取：

```text
rank_override
```

---

## 二、本次實作範圍

### 必須實作

```text
匯出學年分項排名
學年分項排名資料項目
兩項功能的權限註冊
兩項功能的 Program.cs 啟動處理
共用查詢與資料模型
.csproj 檔案及參考組件設定
完成紀錄文件
```

### 本次不要實作

```text
匯入學年分項排名
在資料項目直接新增資料
在資料項目直接修改資料
在資料項目直接刪除資料
刪除學年分項排名功能
重新計算排名
自動產生 matrix_count
修改 rank_override 資料表結構
```

學生資料項目必須是唯讀畫面。資料新增及修正統一由「匯入學年分項排名」處理。

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

## 四、固定資料條件

匯出及資料項目只允許讀取符合以下條件的資料：

```text
semester    = -1
item_type   = 學年/分項成績
item_name   = 學業
rank_type   = 班排名
ref_exam_id IS NULL
```

必須沿用匯入功能建立的共用常數：

```text
Constants\RankOverrideConstants.cs
```

預期內容：

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

如果此檔案已由匯入功能建立，直接引用，不可再建立另一組名稱不同、內容相同的常數。

不可在匯出類別、資料項目或 DAO 中各自硬編碼不同條件。

---

# Part A：匯出學年分項排名

## 五、功能路徑

Ribbon 功能位置：

```text
學生
└─ 資料統計
   └─ 匯出
      └─ 成績相關匯出
         └─ 匯出學年分項排名
```

不得放在「成績相關匯入」。

不得放在 `K12.Presentation.NLDPanels.Student.RibbonBarItems` 的學生個人功能列。

本功能使用：

```csharp
FISCA.Presentation.MotherForm.RibbonBarItems["學生", "資料統計"]["匯出"]
```

---

## 六、匯出權限

請在既有或新建的檔案中集中管理權限代碼：

```text
Permissions\FeatureCode.cs
```

若匯入功能已建立此檔案，不可覆寫原本的匯入權限，只能加入缺少的權限常數。

建議權限代碼：

```csharp
public const string ExportSchoolYearEntryRank =
    "0e230785-a1d3-498a-8a9d-11ed7220d613";
```

權限註冊：

```text
權限目錄：學生 → 學年分項排名
權限名稱：匯出學年分項排名
權限類型：RibbonFeature
權限判斷：Executable
```

建議：

```csharp
Catalog catalog = RoleAclSource.Instance["學生"]["學年分項排名"];

catalog.Add(
    new RibbonFeature(
        FeatureCode.ExportSchoolYearEntryRank,
        "匯出學年分項排名"));
```

如果專案中已存在同功能且使用另一組固定 GUID，保留既有 GUID，不可新增第二個重複權限。

---

## 七、匯出 Excel 規格

檔案格式：

```text
.xlsx
```

欄位順序固定為：

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

### 欄位來源

| Excel 欄位 | 資料來源 | 說明 |
|---|---|---|
| 學號 | 學生基本資料 | `student_number` |
| 班級 | `rank_override.rank_name` | 必須使用排名資料中的歷史班級，不可改用學生目前班級 |
| 座號 | 學生目前基本資料 | 只供辨識；不屬於 `rank_override` |
| 科別 | 學生目前科別資料 | 只供辨識；不屬於 `rank_override` |
| 姓名 | 學生基本資料 | 學生姓名 |
| 學年度 | `rank_override.school_year` | 排名所屬學年度 |
| 年級 | `rank_override.grade_year` | 排名所屬年級，不可改用學生目前年級 |
| 學年學業排名 | `rank_override.rank` | 學年學業班排名 |

### 班級欄位的重要規則

「班級」必須輸出：

```text
rank_override.rank_name
```

不可輸出學生目前班級，否則學生升級、轉班、轉科或畢業後，歷史排名班級會錯誤。

### 座號及科別的限制

`rank_override` 沒有儲存歷史座號及歷史科別。

因此本次匯出的「座號」及「科別」使用學生目前資料，只作為人工辨識資訊。請在完成紀錄中明確記載此限制。

不可自行假設 `student_sems_history` 一定存在可靠的歷史座號或科別欄位，除非 Cursor 在目前系統既有 API 或資料表中確認有一致且可用的來源。

本次不要將座號或科別寫入 `rank_override.extension`。

---

## 八、匯出類別

新增：

```text
ImportExport\ExportSchoolYearEntryRank.cs
```

建議類別：

```csharp
namespace SHImportExportScoreRank.ImportExport
{
    public class ExportSchoolYearEntryRank
        : SmartSchool.API.PlugIn.Export.Exporter
    {
    }
}
```

建構式設定：

```csharp
public ExportSchoolYearEntryRank()
{
    this.Image = null;
    this.Text = "匯出學年分項排名";
}
```

建立固定欄位清單：

```csharp
private readonly List<string> _exportFields = new List<string>
{
    "學號",
    "班級",
    "座號",
    "科別",
    "姓名",
    "學年度",
    "年級",
    "學年學業排名"
};
```

在 `InitializeExport` 中：

```csharp
wizard.ExportableFields.AddRange(_exportFields);
```

不得加入以下內部欄位：

```text
rank_override.id
ref_student_id
semester
item_type
item_name
rank_type
ref_exam_id
matrix_count
extension
```

---

## 九、匯出查詢規則

匯出範圍必須以 `ExportPackageEventArgs.e.List` 的學生系統編號為準。

若沒有選取任何學生，直接回傳，不執行 SQL。

查詢必須一次批次取得所有選取學生的資料，不可在每位學生的 `foreach` 中執行 SQL。

SQL 邏輯至少包含：

```sql
SELECT
    ro.id,
    ro.ref_student_id,
    ro.school_year,
    ro.grade_year,
    ro.rank_name,
    ro.rank
FROM rank_override ro
WHERE ro.ref_student_id IN (...選取學生ID...)
  AND ro.semester = -1
  AND ro.item_type = '學年/分項成績'
  AND ro.item_name = '學業'
  AND ro.rank_type = '班排名'
  AND ro.ref_exam_id IS NULL
ORDER BY
    ro.ref_student_id,
    ro.school_year,
    ro.grade_year;
```

實際 SQL 必須使用 `RankOverrideConstants` 組合條件。

學生基本資料可使用既有 K12／SHSchool API 批次取得，或使用 `QueryHelper` 一次 JOIN 查詢。

不得逐筆呼叫學生 API 或逐筆查詢資料庫。

### 重複資料處理

若資料庫意外存在多筆相同自然鍵資料：

```text
ref_student_id
+ school_year
+ semester = -1
+ item_type = 學年/分項成績
+ item_name = 學業
+ rank_type = 班排名
+ ref_exam_id IS NULL
```

匯出時不要自行刪除或更新資料。

應將查到的每一筆資料匯出，並在完成紀錄中註明發現重複資料的情況。不可使用 `MAX(rank)`、`MIN(rank)` 或任意挑一筆掩蓋問題。

---

## 十、匯出 RowData 寫入

每一筆 `rank_override` 資料建立一個 `RowData`：

```csharp
RowData row = new RowData();
row.ID = refStudentID;
```

依 `e.ExportFields` 寫入使用者勾選的欄位。

應處理 `DBNull`：

- `grade_year` 為空時輸出空白。
- `rank_name` 為空時輸出空白。
- `rank` 為空時輸出空白。
- 座號或科別沒有資料時輸出空白，不可輸出 `0`、`null` 或例外訊息。

完成後加入：

```csharp
e.Items.Add(row);
```

單筆資料轉換錯誤不可讓整批匯出直接中止；但不要只使用 `Console.WriteLine` 靜默忽略。

至少要：

1. 收集錯誤資訊。
2. 完成後顯示錯誤筆數。
3. 將錯誤寫入 `ApplicationLog`。

---

## 十一、ExportStudentV2

參考專案使用：

```text
ImportExport\ExportStudentV2.cs
ImportExport\ExportStudentV2.Designer.cs
ImportExport\ExportStudentV2.resx
```

如果 `SHImportExportScoreRank` 尚未有相同功能：

1. 從 `LHDB_SH_Core` 複製這三個檔案。
2. namespace 改成 `SHImportExportScoreRank.ImportExport`。
3. 保留學生選取及 Excel 匯出流程。
4. 移除與學習歷程專案名稱綁定的程式碼。
5. 不可修改共用匯出流程造成其他欄位順序改變。

如果專案已經有可用的 `ExportStudentV2`，直接共用，不可再建立重複表單。

---

## 十二、匯出啟動程式

在 `Program.cs` 註冊：

```csharp
private static void RegisterExportFeature()
{
    RibbonBarButton exportRoot =
        MotherForm.RibbonBarItems["學生", "資料統計"]["匯出"];

    RibbonBarButton exportButton =
        exportRoot["成績相關匯出"]["匯出學年分項排名"];

    exportButton.Enable =
        UserAcl.Current[
            FeatureCode.ExportSchoolYearEntryRank
        ].Executable;

    exportButton.Click += delegate
    {
        if (!UserAcl.Current[
                FeatureCode.ExportSchoolYearEntryRank
            ].Executable)
        {
            FISCA.Presentation.Controls.MsgBox.Show(
                "您沒有「匯出學年分項排名」權限。");
            return;
        }

        SmartSchool.API.PlugIn.Export.Exporter exporter =
            new ImportExport.ExportSchoolYearEntryRank();

        ImportExport.ExportStudentV2 wizard =
            new ImportExport.ExportStudentV2(
                exporter.Text,
                exporter.Image);

        exporter.InitializeExport(wizard);
        wizard.ShowDialog();
    };
}
```

必須同時做兩層權限控制：

```text
1. 使用 Enable 控制按鈕可用狀態
2. Click 事件內再次檢查 Executable
```

不得只依靠按鈕的 `Enable`。

啟動失敗時應顯示可理解的錯誤訊息並寫入 `ApplicationLog`，不可讓例外中止 ischool 主程式。

---

# Part B：學年分項排名資料項目

## 十三、資料項目位置

此功能不是 Ribbon 按鈕。

功能位置：

```text
學生
└─ 選取單一學生
   └─ 學年分項排名
```

資料項目顯示名稱：

```text
學年分項排名
```

權限名稱可使用：

```text
學年分項排名資料項目
```

不要將畫面 Group 顯示成過長的「學年分項排名資料項目」。

---

## 十四、資料項目權限

在 `Permissions\FeatureCode.cs` 加入：

```csharp
public const string StudentSchoolYearEntryRankDetail =
    "e03b9f03-1bf0-473f-90dd-a737e5b42d1a";
```

權限註冊：

```text
權限目錄：學生 → 學年分項排名
權限名稱：學年分項排名資料項目
權限類型：DetailItemFeature
權限判斷：Viewable
```

建議：

```csharp
Catalog catalog = RoleAclSource.Instance["學生"]["學年分項排名"];

catalog.Add(
    new DetailItemFeature(
        FeatureCode.StudentSchoolYearEntryRankDetail,
        "學年分項排名資料項目"));
```

如果專案中已存在相同功能且使用另一組固定 GUID，保留既有 GUID，不可重複註冊。

---

## 十五、資料項目檔案

新增：

```text
DetailContent\UCStudentSchoolYearEntryRank.cs
DetailContent\UCStudentSchoolYearEntryRank.Designer.cs
DetailContent\UCStudentSchoolYearEntryRank.resx
```

建議類別：

```csharp
using SHImportExportScoreRank.Permissions;

namespace SHImportExportScoreRank.DetailContent
{
    [FISCA.Permission.FeatureCode(
        FeatureCode.StudentSchoolYearEntryRankDetail,
        "學年分項排名資料項目")]
    public partial class UCStudentSchoolYearEntryRank
        : FISCA.Presentation.DetailContent
    {
    }
}
```

建構式設定：

```csharp
this.Group = "學年分項排名";
```

---

## 十六、資料項目畫面

畫面使用唯讀 `ListView` 或唯讀 `DataGridView`。

建議欄位：

```text
學年度
年級
班級
學年學業排名
```

建議欄寬：

| 欄位 | 建議寬度 |
|---|---:|
| 學年度 | 80 |
| 年級 | 70 |
| 班級 | 140 |
| 學年學業排名 | 120 |

畫面禁止：

```text
新增按鈕
修改按鈕
刪除按鈕
直接編輯儲存格
雙擊修改
右鍵刪除
```

資料項目只提供查閱。

如果使用 `DataGridView`：

```csharp
ReadOnly = true;
AllowUserToAddRows = false;
AllowUserToDeleteRows = false;
AllowUserToResizeRows = false;
MultiSelect = false;
SelectionMode = FullRowSelect;
```

如果使用 `ListView`：

```csharp
View = Details;
FullRowSelect = true;
MultiSelect = false;
```

---

## 十七、資料項目資料模型

新增：

```text
DAO\SchoolYearEntryRankRecord.cs
```

建議欄位：

```csharp
public long ID { get; set; }
public string StudentID { get; set; }
public int? SchoolYear { get; set; }
public int? GradeYear { get; set; }
public string RankName { get; set; }
public int? Rank { get; set; }
```

顯示用資料轉換：

```text
SchoolYear → 學年度
GradeYear → 年級
RankName → 班級
Rank → 學年學業排名
```

不得將 `matrix_count` 或 `extension` 顯示成主要畫面欄位，因為本次使用者需求沒有這些欄位。

---

## 十八、共用 DAO

建立或擴充：

```text
DAO\RankOverrideDataAccess.cs
```

至少提供：

```csharp
public static List<SchoolYearEntryRankRecord>
    GetSchoolYearEntryRanksByStudentID(string studentID)
```

查詢必須使用目前資料項目的：

```csharp
this.PrimaryKey
```

SQL 邏輯：

```sql
SELECT
    ro.id,
    ro.ref_student_id,
    ro.school_year,
    ro.grade_year,
    ro.rank_name,
    ro.rank
FROM rank_override ro
WHERE ro.ref_student_id = @StudentID
  AND ro.semester = -1
  AND ro.item_type = '學年/分項成績'
  AND ro.item_name = '學業'
  AND ro.rank_type = '班排名'
  AND ro.ref_exam_id IS NULL
ORDER BY
    ro.school_year DESC,
    ro.grade_year DESC,
    ro.id DESC;
```

實際 SQL 必須引用 `RankOverrideConstants`。

StudentID 必須以參數或確認為有效數字後使用，不可直接串接未驗證字串。

查詢失敗時不得清除或修改資料庫資料。

---

## 十九、資料項目載入流程

參考 `UCStudentLearningHistoryItem`，建立：

```text
BackgroundWorker _bgWorker
bool _isBusy
List<SchoolYearEntryRankRecord> _records
```

建議流程：

```text
切換學生
    ↓
OnPrimaryKeyChanged
    ↓
呼叫 LoadDataAsync / BGRun
    ↓
BackgroundWorker 查詢 rank_override
    ↓
RunWorkerCompleted
    ↓
更新 ListView 或 DataGridView
```

必要規則：

1. `PrimaryKey` 為空時清空畫面，不執行 SQL。
2. BackgroundWorker 執行中再次切換學生時，設定 `_isBusy = true`。
3. 完成後若 `_isBusy = true`，重新載入目前最新 `PrimaryKey`。
4. 不可在背景執行緒直接操作 WinForms 控制項。
5. 查詢期間設定 `this.Loading = true`。
6. 完成或錯誤後必須設定 `this.Loading = false`。
7. 查無資料時顯示空清單，不顯示錯誤。
8. 查詢例外必須顯示訊息並寫入 `ApplicationLog`。

---

## 二十、資料項目重新整理機制

在資料項目建構式使用唯一名稱註冊重新整理入口：

```csharp
FISCA.Features.TryRegister(
    "SchoolYearEntryRankDetailContent",
    x =>
    {
        this.ReloadData();
    });
```

若匯入功能完成後已有呼叫重新整理，必須使用相同名稱：

```text
SchoolYearEntryRankDetailContent
```

不要在不同檔案使用多個近似名稱，例如：

```text
StudentSchoolYearRankRefresh
SchoolYearRankDetailRefresh
ReloadRankOverride
```

匯入成功後可呼叫同一功能入口，讓目前開啟的學生資料項目更新。

若目前專案的 `FISCA.Features` API 寫法與參考專案不同，依現有版本可編譯的方式處理，但必須保留「避免重複註冊」與「可由匯入完成後重新整理」兩項目的。

---

## 二十一、資料項目啟動程式

在 `Program.cs` 註冊：

```csharp
private static void RegisterDetailContent()
{
    Catalog catalog =
        RoleAclSource.Instance["學生"]["學年分項排名"];

    catalog.Add(
        new DetailItemFeature(
            FeatureCode.StudentSchoolYearEntryRankDetail,
            "學年分項排名資料項目"));

    if (!UserAcl.Current[
            FeatureCode.StudentSchoolYearEntryRankDetail
        ].Viewable)
    {
        return;
    }

    K12.Presentation.NLDPanels.Student.AddDetailBulider(
        new FISCA.Presentation.DetailBulider<
            DetailContent.UCStudentSchoolYearEntryRank>());
}
```

必須在呼叫 `AddDetailBulider` 前檢查：

```csharp
UserAcl.Current[FeatureCode.StudentSchoolYearEntryRankDetail].Viewable
```

沒有權限時：

```text
不可註冊資料項目
不可建立控制項
不可查詢 rank_override
```

---

# Part C：Program.cs 與專案整合

## 二十二、Program.cs 整合

目前 `Program.cs` 必須維持單一：

```csharp
[FISCA.MainMethod()]
public static void Main()
```

不可為匯入、匯出、資料項目各建立一個 Main。

建議：

```csharp
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
        RegisterImportFeature();       // 既有匯入功能，保留
        RegisterExportFeature();       // 本次新增
        RegisterDetailContent();       // 本次新增
        RegisterImportValidators();    // 既有匯入驗證，保留
    }
}
```

如果匯入功能尚未完成，不要為了讓本次程式編譯而刪除預留結構；但不可呼叫不存在的方法。依目前專案實際完成狀態調整。

### 初始化防重複

使用：

```csharp
private static bool _initialized;
```

避免：

```text
權限重複註冊
Click 事件重複綁定
資料項目重複註冊
同一次點擊開啟兩個匯出視窗
```

---

## 二十三、建議專案結構

```text
SHImportExportScoreRank
├─ Constants
│  └─ RankOverrideConstants.cs
│
├─ DAO
│  ├─ RankOverrideDataAccess.cs
│  └─ SchoolYearEntryRankRecord.cs
│
├─ DetailContent
│  ├─ UCStudentSchoolYearEntryRank.cs
│  ├─ UCStudentSchoolYearEntryRank.Designer.cs
│  └─ UCStudentSchoolYearEntryRank.resx
│
├─ ImportExport
│  ├─ ImportSchoolYearEntryRank.cs                  # 既有匯入功能
│  ├─ ImportSchoolYearEntryRank.xml                 # 既有匯入功能
│  ├─ ExportSchoolYearEntryRank.cs                  # 本次新增
│  ├─ ExportStudentV2.cs                            # 尚未存在時加入
│  ├─ ExportStudentV2.Designer.cs                   # 尚未存在時加入
│  └─ ExportStudentV2.resx                          # 尚未存在時加入
│
├─ Permissions
│  └─ FeatureCode.cs
│
├─ Program.cs
└─ SHImportExportScoreRank.csproj
```

---

## 二十四、.csproj 處理

目前 `SHImportExportScoreRank.csproj` 初始狀態只有基本 .NET 組件及 `Program.cs`。

Cursor 必須檢查並加入實際使用的參考組件。

參考 `LHDB_SH_Core.csproj`，至少確認下列 API 所需組件已引用：

```text
FISCA
FISCA.Data
FISCA.Presentation
FISCA.Permission
FISCA.LogAgent
K12.Data
K12.Presentation
SHSchool.Data
SmartSchool.API.PlugIn
Campus
Campus.Windows
DevComponents.DotNetBar（若 Designer 或控制項需要）
System.Windows.Forms
System.Drawing
System.Data
System.Xml
System.Xml.Linq
```

只加入實際需要的 DLL，不可複製參考專案所有無關組件。

所有新增 `.cs`、`.Designer.cs`、`.resx` 必須加入 `.csproj`。

WinForms 檔案關聯應正確：

```xml
<Compile Include="DetailContent\UCStudentSchoolYearEntryRank.cs">
  <SubType>UserControl</SubType>
</Compile>
<Compile Include="DetailContent\UCStudentSchoolYearEntryRank.Designer.cs">
  <DependentUpon>UCStudentSchoolYearEntryRank.cs</DependentUpon>
</Compile>
<EmbeddedResource Include="DetailContent\UCStudentSchoolYearEntryRank.resx">
  <DependentUpon>UCStudentSchoolYearEntryRank.cs</DependentUpon>
</EmbeddedResource>
```

`ExportStudentV2` 若複製進專案，也要建立相同的 Designer／resx 關聯。

目標框架維持：

```text
.NET Framework 4.8
```

不可改成 .NET 6、.NET 8 或 SDK-style project。

---

## 二十五、資料庫相容性

SQL 必須相容：

```text
PostgreSQL 9.6
PostgreSQL 16
PostgreSQL 17
```

本次只有查詢，不可使用：

```text
MERGE
PostgreSQL 9.6 不支援的新語法
會修改資料的 CTE
暫存刪除或更新 SQL
```

匯出與資料項目都不得變更 `rank_override`。

---

## 二十六、ApplicationLog

建議模組名稱：

```text
學年分項排名
```

匯出成功可記錄：

```text
動作：匯出
內容：匯出學年分項排名，共 N 筆，選取學生 M 人。
```

查詢或轉換失敗應記錄：

```text
功能名稱
目前使用者
學生系統編號（如適用）
例外訊息
StackTrace
```

不可將完整 SQL、身分證字號或其他不必要的個資寫入一般成功紀錄。

資料項目正常切換學生及查無資料不需要每次寫成功 Log，避免產生大量無意義紀錄。

---

## 二十七、不可變更事項

```text
不可修改 rank_override 資料表結構
不可修改既有匯入功能的 Excel 欄位
不可將 semester 改回 0
不可將 item_type 改成其他值
不可將 item_name 改成其他值
不可將 rank_type 改成其他值
不可將學生目前班級取代 rank_override.rank_name
不可在資料項目直接寫入或刪除資料
不可因資料重複而自動刪除或合併
不可清除 extension
不可重新計算 rank
不可自行填入 matrix_count
不可修改 LHDB_SH_Core 參考專案
```

---

# Part D：測試與驗收

## 二十八、匯出功能測試

至少測試：

### 權限

1. 有匯出權限時按鈕可使用。
2. 無匯出權限時按鈕不可使用。
3. Click 事件仍會再次檢查權限。
4. 權限功能樹顯示於「學生 → 學年分項排名」。

### Ribbon 路徑

確認功能位於：

```text
學生 → 資料統計 → 匯出 → 成績相關匯出 → 匯出學年分項排名
```

### 學生選取

1. 選取一位學生匯出。
2. 選取多位學生匯出。
3. 選取學生沒有排名資料時不產生錯誤。
4. 混合有資料與無資料學生時，只匯出有資料者。

### 固定條件

準備其他 `rank_override` 資料，確認以下資料不會被匯出：

```text
semester != -1
item_type != 學年/分項成績
item_name != 學業
rank_type != 班排名
ref_exam_id IS NOT NULL
```

### Excel

確認 `.xlsx` 欄位順序：

```text
學號、班級、座號、科別、姓名、學年度、年級、學年學業排名
```

確認：

1. 班級來自 `rank_override.rank_name`。
2. 年級來自 `rank_override.grade_year`。
3. 排名來自 `rank_override.rank`。
4. 空座號或空科別輸出空白。
5. 學號前導 0 不可被程式主動轉成整數後遺失。
6. 匯出的檔案可由「匯入學年分項排名」讀取相同欄位格式。

### 重複資料

資料庫有重複自然鍵時：

1. 不可自動刪除。
2. 不可只挑一筆。
3. 匯出結果保留實際資料筆數。
4. 完成紀錄註明此情況。

---

## 二十九、資料項目測試

至少測試：

### 權限

1. 有 Viewable 權限時顯示資料項目。
2. 無 Viewable 權限時不顯示資料項目。
3. 無權限時不執行 SQL。
4. 權限功能樹顯示「學年分項排名資料項目」。

### 資料載入

1. 選取有一筆排名的學生。
2. 選取有多學年度排名的學生。
3. 選取沒有排名的學生。
4. 快速切換多位學生，最後畫面必須顯示最後選取學生。
5. 切換學生時畫面不可凍結。
6. 查詢例外後 Loading 狀態必須恢復。

### 固定條件

確認資料項目只顯示：

```text
semester = -1
item_type = 學年/分項成績
item_name = 學業
rank_type = 班排名
ref_exam_id IS NULL
```

### 顯示

確認欄位：

```text
學年度
年級
班級
學年學業排名
```

確認排序：

```text
學年度由新到舊
同學年度依年級由大到小
```

### 唯讀

確認畫面沒有：

```text
新增
修改
刪除
儲存
可編輯儲存格
```

---

## 三十、編譯與整合測試

1. `SHImportExportScoreRank.csproj` 可在 Visual Studio 以 .NET Framework 4.8 編譯。
2. 無缺少 DLL 或 namespace。
3. Designer 可開啟，不發生控制項載入錯誤。
4. 模組載入後 `Program.Main()` 只執行一次註冊。
5. 匯入、匯出及資料項目三個權限代碼互不衝突。
6. 匯出按鈕只出現一次。
7. 學生資料項目只出現一次。
8. 不影響其他既有學生匯入、匯出或資料項目功能。
9. PostgreSQL 9.6、16、17 查詢語法相容。

---

## 三十一、完成紀錄

完成後新增：

```text
匯出與資料項目學年分項排名功能.md
```

至少記錄：

```text
1. 修改及新增的檔案清單
2. 匯出 Ribbon 完整路徑
3. 匯出 FeatureCode
4. 資料項目 FeatureCode
5. Excel 欄位及欄位來源
6. rank_override 固定查詢條件
7. 班級使用 rank_name 的原因
8. 座號及科別使用目前資料的限制
9. 資料項目顯示欄位
10. 背景載入及快速切換學生處理方式
11. ApplicationLog 寫法
12. .csproj 新增的參考及檔案
13. 編譯結果
14. 測試案例與實際結果
15. 尚未處理或需要後續確認的事項
```

---

## 最終驗收標準

完成後必須同時符合：

```text
[ ] 可從「成績相關匯出」開啟匯出學年分項排名
[ ] 匯出結果為 .xlsx
[ ] Excel 欄位順序完全正確
[ ] 匯出只讀取符合固定條件的 rank_override
[ ] 班級使用 rank_override.rank_name
[ ] 年級使用 rank_override.grade_year
[ ] 學生頁籤顯示「學年分項排名」資料項目
[ ] 資料項目為唯讀
[ ] 資料項目只顯示目前學生資料
[ ] 快速切換學生不會顯示前一位學生資料
[ ] 匯出使用 Executable 權限
[ ] 資料項目使用 Viewable 權限
[ ] Program.cs 不重複註冊
[ ] 不修改 rank_override 任何資料
[ ] 不影響既有匯入功能
[ ] 完成紀錄已寫入匯出與資料項目學年分項排名功能.md
```
