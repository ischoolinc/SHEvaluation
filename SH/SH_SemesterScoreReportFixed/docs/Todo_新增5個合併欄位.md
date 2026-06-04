## 目標

調整 `Program.cs`，新增 5 個 Word 合併欄位。

新增欄位：

```text
上學期學業成績名次
上學期實得學分數
上學期累計取得學分數
學年實得學分數
學年累計取得學分數
````

修改完成後，請完整記錄在：

```text
期末成績通知單(固定排名)調整0603.md
```

---

# 一、重要限制

請嚴格遵守：

```text
1. 只修改 Program.cs。
2. 不要修改原本本學期學業成績排名邏輯。
3. 不要修改原本本學期取得學分數計算。
4. 不要修改原本累計取得學分數計算。
5. 不要修改原本本學期實得必修學分計算。
6. 不要修改原本本學期實得選修學分計算。
7. 不要修改科目成績計算。
8. 不要修改排名原本欄位填值。
9. 不要修改學分數原本欄位填值。
10. 不要重構 Program.cs。
11. 只新增 5 個合併欄位與對應填值。
```

---

# 二、比對條件限制

本次新增欄位，比對條件只允許使用：

```text
SchoolYear
Semester
Pass
不計學分
```

不要加入年級判斷。

不可加入：

```text
GradeYear
currentGradeYear
semesterSubjectScore.GradeYear
```

尤其不要加入這類判斷：

```csharp
semesterSubjectScore.GradeYear == currentGradeYear
```

---

# 三、新增 DataTable 合併欄位

請在 `Program.cs` 建立 `DataTable` 欄位的位置，找到目前學分或成績相關欄位附近，例如：

```csharp
table.Columns.Add("本學期取得學分數");
table.Columns.Add("累計取得學分數");
```

在附近新增：

```csharp
table.Columns.Add("上學期學業成績名次");
table.Columns.Add("上學期實得學分數");
table.Columns.Add("上學期累計取得學分數");
table.Columns.Add("學年實得學分數");
table.Columns.Add("學年累計取得學分數");
```

---

# 四、取得上學期學業成績排名資料

目前本學期排名資料可能是這樣取得：

```csharp
Dictionary<string, Dictionary<string, DataRow>> SemsScoreRankMatrixDataDict =
    Utility.GetSemsScoreRankMatrixData(conf.SchoolYear, conf.Semester, selectedStudents);
```

請在這段附近新增上學期排名資料。

只有目前學期是第 2 學期才取得上學期排名。

```csharp
Dictionary<string, Dictionary<string, DataRow>> PrevSemsScoreRankMatrixDataDict =
    new Dictionary<string, Dictionary<string, DataRow>>();

if (conf.Semester == "2")
{
    PrevSemsScoreRankMatrixDataDict =
        Utility.GetSemsScoreRankMatrixData(conf.SchoolYear, "1", selectedStudents);
}
```

注意：

```text
conf.Semester == "1" 時，不取得上學期排名資料。
```

---

# 五、新增加總變數

請在每位學生產生 `DataRow row` 後，或在處理該學生學分統計前，新增以下變數：

```csharp
decimal 上學期實得學分數 = 0;
decimal 上學期累計取得學分數 = 0;
decimal 學年實得學分數 = 0;
decimal 學年累計取得學分數 = 0;
```

如果專案不適合使用中文變數，也可以用英文變數：

```csharp
decimal prevSemesterPassCredits = 0;
decimal prevSemesterTotalPassCredits = 0;
decimal schoolYearPassCredits = 0;
decimal schoolYearTotalPassCredits = 0;
```

---

# 六、在 SemesterSubjectScoreList 迴圈中新增統計

請找到原本處理學分統計的迴圈，類似：

```csharp
foreach (var semesterSubjectScore in stuRec.SemesterSubjectScoreList)
{
    if (semesterSubjectScore.Detail.GetAttribute("不計學分") != "是")
    {
        ...
    }
}
```

請在不改動原本統計邏輯的前提下，於 `不計學分 != 是` 的區塊內新增以下計算。

---

## 6.1 準備學年度與學期變數

可在迴圈內新增：

```csharp
int confSchoolYear = int.Parse(conf.SchoolYear);
int scoreSchoolYear = semesterSubjectScore.SchoolYear;
int scoreSemester = semesterSubjectScore.Semester;
```

如果外面已經有相同變數，請沿用，不要重複宣告造成衝突。

---

## 6.2 上學期實得學分數

規則：

```text
只有 conf.Semester == "2" 才計算。
上學期 = conf.SchoolYear 第 1 學期。
Pass == true
不計學分 != 是
```

請新增：

```csharp
if (conf.Semester == "2" &&
    semesterSubjectScore.Pass &&
    scoreSchoolYear == confSchoolYear &&
    scoreSemester == 1)
{
    prevSemesterPassCredits += semesterSubjectScore.CreditDec();
}
```

如果使用中文變數：

```csharp
if (conf.Semester == "2" &&
    semesterSubjectScore.Pass &&
    scoreSchoolYear == confSchoolYear &&
    scoreSemester == 1)
{
    上學期實得學分數 += semesterSubjectScore.CreditDec();
}
```

---

## 6.3 上學期累計取得學分數

定義：

```text
截至上學期為止的累計取得學分數。
```

若目前報表是：

```text
114 學年度第 2 學期
```

則統計：

```text
114-1 以前，包含 114-1 的所有已取得學分。
```

規則：

```text
只有 conf.Semester == "2" 才計算。
Pass == true
不計學分 != 是
SchoolYear < conf.SchoolYear
或
SchoolYear == conf.SchoolYear 且 Semester <= 1
```

請新增：

```csharp
if (conf.Semester == "2" &&
    semesterSubjectScore.Pass &&
    (
        scoreSchoolYear < confSchoolYear ||
        (scoreSchoolYear == confSchoolYear && scoreSemester <= 1)
    ))
{
    prevSemesterTotalPassCredits += semesterSubjectScore.CreditDec();
}
```

如果使用中文變數：

```csharp
if (conf.Semester == "2" &&
    semesterSubjectScore.Pass &&
    (
        scoreSchoolYear < confSchoolYear ||
        (scoreSchoolYear == confSchoolYear && scoreSemester <= 1)
    ))
{
    上學期累計取得學分數 += semesterSubjectScore.CreditDec();
}
```

---

## 6.4 學年實得學分數

定義：

```text
同一學年度第 1 學期與第 2 學期的實得學分數加總。
```

規則：

```text
SchoolYear == conf.SchoolYear
Semester == 1 或 2
Pass == true
不計學分 != 是
不判斷年級
```

請新增：

```csharp
if (semesterSubjectScore.Pass &&
    scoreSchoolYear == confSchoolYear &&
    (scoreSemester == 1 || scoreSemester == 2))
{
    schoolYearPassCredits += semesterSubjectScore.CreditDec();
}
```

如果使用中文變數：

```csharp
if (semesterSubjectScore.Pass &&
    scoreSchoolYear == confSchoolYear &&
    (scoreSemester == 1 || scoreSemester == 2))
{
    學年實得學分數 += semesterSubjectScore.CreditDec();
}
```

---

## 6.5 學年累計取得學分數

定義：

```text
截至目前學年度上下學期為止的累計取得學分數。
```

若目前報表是：

```text
114 學年度
```

則統計：

```text
所有小於 114 學年度的已取得學分
+
114-1、114-2 的已取得學分
```

規則：

```text
Pass == true
不計學分 != 是
SchoolYear < conf.SchoolYear
或
SchoolYear == conf.SchoolYear 且 Semester == 1 或 2
不判斷年級
```

請新增：

```csharp
if (semesterSubjectScore.Pass &&
    (
        scoreSchoolYear < confSchoolYear ||
        (
            scoreSchoolYear == confSchoolYear &&
            (scoreSemester == 1 || scoreSemester == 2)
        )
    ))
{
    schoolYearTotalPassCredits += semesterSubjectScore.CreditDec();
}
```

如果使用中文變數：

```csharp
if (semesterSubjectScore.Pass &&
    (
        scoreSchoolYear < confSchoolYear ||
        (
            scoreSchoolYear == confSchoolYear &&
            (scoreSemester == 1 || scoreSemester == 2)
        )
    ))
{
    學年累計取得學分數 += semesterSubjectScore.CreditDec();
}
```

---

# 七、填入 DataRow

請在原本填入學分欄位附近，例如：

```csharp
row["本學期取得學分數"] = ...
row["累計取得學分數"] = ...
```

附近新增以下填值。

---

## 7.1 上學期學業成績名次

目前只新增 1 個名次欄位，因此先使用：

```text
上學期學業成績班排名
```

對應 ranking key：

```text
學期/分項成績_學業_班排名
```

請加入：

```csharp
row["上學期學業成績名次"] = "";

if (conf.Semester == "2")
{
    string prevRankKey = "學期/分項成績_學業_班排名";

    if (PrevSemsScoreRankMatrixDataDict.ContainsKey(stuRec.StudentID) &&
        PrevSemsScoreRankMatrixDataDict[stuRec.StudentID].ContainsKey(prevRankKey) &&
        PrevSemsScoreRankMatrixDataDict[stuRec.StudentID][prevRankKey]["rank"] != null)
    {
        row["上學期學業成績名次"] =
            PrevSemsScoreRankMatrixDataDict[stuRec.StudentID][prevRankKey]["rank"].ToString();
    }
}
```

注意：

```text
conf.Semester == "1" 時，上學期學業成績名次保持空白。
```

---

## 7.2 上學期實得學分數與上學期累計取得學分數

```csharp
if (conf.Semester == "2")
{
    row["上學期實得學分數"] = prevSemesterPassCredits;
    row["上學期累計取得學分數"] = prevSemesterTotalPassCredits;
}
else
{
    row["上學期實得學分數"] = "";
    row["上學期累計取得學分數"] = "";
}
```

如果使用中文變數：

```csharp
if (conf.Semester == "2")
{
    row["上學期實得學分數"] = 上學期實得學分數;
    row["上學期累計取得學分數"] = 上學期累計取得學分數;
}
else
{
    row["上學期實得學分數"] = "";
    row["上學期累計取得學分數"] = "";
}
```

---

## 7.3 學年實得學分數與學年累計取得學分數

```csharp
row["學年實得學分數"] = schoolYearPassCredits;
row["學年累計取得學分數"] = schoolYearTotalPassCredits;
```

如果使用中文變數：

```csharp
row["學年實得學分數"] = 學年實得學分數;
row["學年累計取得學分數"] = 學年累計取得學分數;
```

---

# 八、欄位規則整理

## 8.1 上學期學業成績名次

```text
只有 conf.Semester == "2" 才有資料。
使用 Utility.GetSemsScoreRankMatrixData(conf.SchoolYear, "1", selectedStudents)。
使用 key：學期/分項成績_學業_班排名。
填入 rank。
conf.Semester == "1" 時空白。
```

---

## 8.2 上學期實得學分數

```text
只有 conf.Semester == "2" 才有資料。
SchoolYear == conf.SchoolYear
Semester == 1
Pass == true
不計學分 != 是
CreditDec() 加總。
conf.Semester == "1" 時空白。
```

---

## 8.3 上學期累計取得學分數

```text
只有 conf.Semester == "2" 才有資料。
統計截至 conf.SchoolYear 第 1 學期為止。
SchoolYear < conf.SchoolYear
或
SchoolYear == conf.SchoolYear 且 Semester <= 1
Pass == true
不計學分 != 是
CreditDec() 加總。
conf.Semester == "1" 時空白。
```

---

## 8.4 學年實得學分數

```text
SchoolYear == conf.SchoolYear
Semester == 1 或 2
Pass == true
不計學分 != 是
CreditDec() 加總。
不判斷年級。
```

---

## 8.5 學年累計取得學分數

```text
統計截至 conf.SchoolYear 第 2 學期為止。
SchoolYear < conf.SchoolYear
或
SchoolYear == conf.SchoolYear 且 Semester == 1 或 2
Pass == true
不計學分 != 是
CreditDec() 加總。
不判斷年級。
```

---

# 九、不要修改的原本邏輯

請確認沒有修改以下邏輯：

```text
1. 本學期學業成績排名
2. 本學期取得學分數
3. 累計取得學分數
4. 本學期已修必修學分
5. 本學期已修選修學分
6. 本學期實得必修學分
7. 本學期實得選修學分
8. 累計取得必修學分
9. 累計取得選修學分
10. 在校期間實得必修學分
11. 在校期間實得選修學分
12. 科目成績輸出
13. 排名原本欄位輸出
```

---

# 十、測試案例

## 測試 1：目前學期是第 1 學期

設定：

```text
conf.SchoolYear = 114
conf.Semester = 1
```

預期：

```text
上學期學業成績名次 = 空白
上學期實得學分數 = 空白
上學期累計取得學分數 = 空白
學年實得學分數 = 114-1 + 114-2 中 Pass=true 的學分加總
學年累計取得學分數 = 114 學年度以前到 114-2 的已取得學分加總
```

---

## 測試 2：目前學期是第 2 學期

設定：

```text
conf.SchoolYear = 114
conf.Semester = 2
```

資料：

```text
113-2 已取得 20 學分
114-1 已取得 25 學分
114-2 已取得 30 學分
```

預期：

```text
上學期實得學分數 = 25
上學期累計取得學分數 = 45
學年實得學分數 = 55
學年累計取得學分數 = 75
```

---

## 測試 3：不計學分

資料：

```text
114-1 科目 A，Pass=true，Credit=2，不計學分=是
114-1 科目 B，Pass=true，Credit=3，不計學分=否
```

預期：

```text
上學期實得學分數只加 3
學年實得學分數只加 3
```

---

## 測試 4：不判斷年級

資料：

```text
114-1 GradeYear=2，Pass=true，Credit=2
114-2 GradeYear=3，Pass=true，Credit=3
```

預期：

```text
學年實得學分數 = 5
```

說明：

```text
本次新增欄位不使用 GradeYear 判斷。
```

---

## 測試 5：排名資料

當：

```text
conf.Semester = 2
```

且上學期排名資料有：

```text
學期/分項成績_學業_班排名 rank = 8
```

預期：

```text
上學期學業成績名次 = 8
```

當：

```text
conf.Semester = 1
```

預期：

```text
上學期學業成績名次 = 空白
```

---

# 十一、完成後紀錄

請建立或更新：

```text
期末成績通知單(固定排名)調整0603.md
```

紀錄內容請包含：

````md
# 期末成績通知單(固定排名)調整0603

## 修改目標

新增 5 個上學期與學年相關合併欄位。

## 修改檔案

- Program.cs

## 新增合併欄位

```text
上學期學業成績名次
上學期實得學分數
上學期累計取得學分數
學年實得學分數
學年累計取得學分數
````

## 計算方式

### 上學期學業成績名次

```text
只有目前學期是第 2 學期才填入。
資料來源使用 Utility.GetSemsScoreRankMatrixData(conf.SchoolYear, "1", selectedStudents)。
排名 key 使用：學期/分項成績_學業_班排名。
填入 rank。
目前學期是第 1 學期時空白。
```

### 上學期實得學分數

```text
只有目前學期是第 2 學期才填入。
SchoolYear == conf.SchoolYear
Semester == 1
Pass == true
不計學分 != 是
CreditDec() 加總。
目前學期是第 1 學期時空白。
```

### 上學期累計取得學分數

```text
只有目前學期是第 2 學期才填入。
統計截至 conf.SchoolYear 第 1 學期為止。
SchoolYear < conf.SchoolYear
或
SchoolYear == conf.SchoolYear 且 Semester <= 1
Pass == true
不計學分 != 是
CreditDec() 加總。
目前學期是第 1 學期時空白。
```

### 學年實得學分數

```text
SchoolYear == conf.SchoolYear
Semester == 1 或 2
Pass == true
不計學分 != 是
CreditDec() 加總。
不判斷年級。
```

### 學年累計取得學分數

```text
統計截至 conf.SchoolYear 第 2 學期為止。
SchoolYear < conf.SchoolYear
或
SchoolYear == conf.SchoolYear 且 Semester == 1 或 2
Pass == true
不計學分 != 是
CreditDec() 加總。
不判斷年級。
```

## 不變動內容

本次未修改：

```text
1. 本學期學業成績排名
2. 本學期取得學分數
3. 累計取得學分數
4. 本學期實得必修學分
5. 本學期實得選修學分
6. 科目成績計算
7. 原本排名欄位填值
8. Word 合併流程
```

## 測試結果

請記錄：

```text
1. 第 1 學期時，上學期三個欄位是否空白。
2. 第 2 學期時，上學期學業成績名次是否可帶出。
3. 第 2 學期時，上學期實得學分數是否正確。
4. 第 2 學期時，上學期累計取得學分數是否正確。
5. 學年實得學分數是否為同學年度第 1、2 學期加總。
6. 學年累計取得學分數是否為截至同學年度第 2 學期為止的累計。
7. 是否未使用 GradeYear 判斷。
8. 原本本學期與累計欄位是否與修改前一致。
9. 編譯是否成功。
```
