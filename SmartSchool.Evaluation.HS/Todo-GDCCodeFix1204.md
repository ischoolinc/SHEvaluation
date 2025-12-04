# Cursor Todo：修正 SQL 語法問題（自動寫入學期歷程 GDCCode 取得方式）
檔名建議：Todo-學期歷程GDCCode修正1204.md  
完成後紀錄：高中自動寫入學習歷程修正1204.md

---

# 目標
1. 修正 CalcSemesterSubjectScoreWizard.cs 內處理學期歷程所使用的 GDCCode SQL。
2. 將舊版「COALESCE(student.gdc_code, class.gdc_code)」改成使用 graduation_plan.moe_group_code。
3. 程式其他邏輯保持不變。
4. 修改完成後記錄在《高中自動寫入學習歷程修正1204.md》。

---

# 需修改的檔案
CalcSemesterSubjectScoreWizard.cs

---

# Step 1：搜尋並確認舊 SQL 程式碼

搜尋以下內容：

```csharp
string qry = "SELECT " +
    "student.id AS student_id" +
    ",COALESCE(student.gdc_code,class.gdc_code) AS gdc_code " +
    "FROM student " +
    "LEFT OUTER JOIN " +
    "class " +
    " ON student.ref_class_id = class.id " +
    " WHERE student.id IN(" + string.Join(",", studIDList.ToArray()) + ");";
```

---

# Step 2：替換成下列新版 SQL（使用畢業規劃 moe_group_code）

```csharp
string qry = @"
WITH stud_gp_id AS(
    SELECT
        student.id,
        COALESCE(
            student.ref_graduation_plan_id,
            class.ref_graduation_plan_id
        ) AS graduation_plan_id
    FROM
        student
        LEFT JOIN class ON student.ref_class_id = class.id
)
SELECT
    stud_gp_id.id AS student_id,
    graduation_plan.moe_group_code AS gdc_code
FROM
    stud_gp_id
    INNER JOIN graduation_plan 
        ON stud_gp_id.graduation_plan_id = graduation_plan.id
WHERE
    stud_gp_id.id IN(" + string.Join(",", studIDList.ToArray()) + @");";
```

---

# Step 3：保持後續程式碼不變

```csharp
foreach (DataRow dr in dt.Rows)
{
    string sid = dr["student_id"] + "";
    string groupCode = dr["gdc_code"] + "";
    if (!studGDCCodeDict.ContainsKey(sid))
        studGDCCodeDict.Add(sid, groupCode);
}
```

---

# Step 4：測試項目

1. 有畢業規劃 / 無畢業規劃學生是否正確產生 GDCCode。
2. 學期歷程 XML `<History GDCCode="">` 是否正確填入。
3. 舊版 student/class.gdc_code 是否不再被使用。
4. 已有相同學期歷程時，更新流程是否正確（只更新變動欄位）。

---

# Step 5：完成後請記錄

將本次修改內容與測試結果紀錄於：

```
高中自動寫入學習歷程修正1204.md
```

