# Todo-ImportMakeupRetakeScore_成績年級必填1121.md

## 目標
依 2025-11-21 與 ChatGPT 討論結果，調整 `ImportMakeupRetakeScore.cs`：
- 不論匯入精靈中選擇「自動判斷取得學分」或「手動判斷取得學分」，**「成績年級」欄位都必須勾選並填入正確年級**。
- 確保補修(4.3)、重修(5.2) 名冊的 GradeYear 以及學期科目成績的成績年級邏輯維持正確。
- 完成後，將本次調整記錄整理寫入：`高中成績調整1121.md`。

---

## 一、程式檔案範圍

- 主要檔案：`ImportMakeupRetakeScore.cs`
- 主要影響區塊：
  - 匯入精靈初始化 (`InitializeImport` 或相關初始化流程)
  - RequiredFields 設定
  - 「自動判斷取得學分 / 手動判斷取得學分」 RadioButton CheckedChanged 事件

---

## 二、需求重點整理（成績年級行為）

1. **成績年級永遠為必填欄位**
   - 不論使用者選擇「自動判斷取得學分」或「手動判斷取得學分」。
   - 在匯入精靈畫面中，「成績年級」必須出現在「必要欄位(RequiredFields)」中，使用者一定要對應欄位才能下一步。

2. **自動判斷取得學分（autoCheckPass）**
   - 繼續維持原有行為：
     - 「成績年級」 + 「取得學分」都是必填欄位。
   - 成績年級在 RequiredFields 中不可以被移除。

3. **手動判斷取得學分（manulCheckPass）**
   - 調整行為：
     - 「成績年級」維持必填（留在 RequiredFields 中）。
     - 只有「取得學分」在手動模式下可以從 RequiredFields 移除。
   - 也就是說：在任何情況下，「成績年級」都不能從 RequiredFields 被 Remove。

4. **下游影響（確認不破壞既有邏輯）**
   - ValidateRow 中對「成績年級」的格式檢查（必須為整數）照舊。
   - 同一學生、同一學年度學期的「成績年級一致性」檢查照舊。
   - 補修/重修名冊 (`SubjectScoreRec108.GradeYear`) 會一定有匯入資料。
   - 學期科目成績的 `semesterGradeYear` 邏輯維持原本行為：
     - 有匯入值 → 可更新學期成績年級。
     - 無匯入欄位的情境被禁止（因為現在成績年級必填）。

---

## 三、實作步驟（Cursor 編輯重點）

### Todo 1：在匯入精靈初始化時，加入「成績年級」為固定 RequiredFields

**位置建議：**
- 找到 `ImportMakeupRetakeScore.cs` 中，匯入精靈 `wizard` 初始化、設定 RequiredFields 的區塊，原本大致類似：

```csharp
wizard.RequiredFields.AddRange("科目", "科目級別", "學年度", "學期", "是否補修成績", "補修學年度", "補修學期", "重修學年度", "重修學期");
```

**修改方式：**
- 將「成績年級」加入 RequiredFields，比如：

```csharp
wizard.RequiredFields.AddRange(
    "科目",
    "科目級別",
    "學年度",
    "學期",
    "成績年級",          // ★ 新增：成績年級永遠必填
    "是否補修成績",
    "補修學年度",
    "補修學期",
    "重修學年度",
    "重修學期"
);
```

**預期效果：**
- 匯入精靈欄位對應畫面：「成績年級」會顯示為必填欄位，不對應不能下一步。

---

### Todo 2：「自動判斷取得學分」CheckedChanged 邏輯調整（保險再加入一次成績年級 RequiredFields）

**位置建議：**
- 找到 `autoCheckPass.CheckedChanged += delegate { ... };` 區塊。

**原始邏輯（概念）：**

```csharp
autoCheckPass.CheckedChanged += delegate
{
    if (autoCheckPass.Checked)
    {
        if (!wizard.RequiredFields.Contains("成績年級"))
            wizard.RequiredFields.Add("成績年級");
        if (!wizard.RequiredFields.Contains("取得學分"))
            wizard.RequiredFields.Add("取得學分");
    }
};
```

**調整方式（可以保留邏輯，但註解說明成績年級現在本來就固定必填）：**

```csharp
autoCheckPass.CheckedChanged += delegate
{
    if (autoCheckPass.Checked)
    {
        // 成績年級本來就在 RequiredFields.AddRange 中設定為必填
        // 為了保險，仍然檢查一次，若被移除則補回
        if (!wizard.RequiredFields.Contains("成績年級"))
            wizard.RequiredFields.Add("成績年級");

        if (!wizard.RequiredFields.Contains("取得學分"))
            wizard.RequiredFields.Add("取得學分");
    }
};
```

**重點：**
- 「成績年級」不會被此段程式移除，只會在需要時再補回去。
- 「取得學分」在自動判斷模式下保持必填。

---

### Todo 3：「手動判斷取得學分」CheckedChanged 調整 —— 不得移除成績年級

**位置建議：**
- 找到 `manulCheckPass.CheckedChanged += delegate { ... };` 區塊。

**原始邏輯（概念）：**

```csharp
manulCheckPass.CheckedChanged += delegate
{
    if (manulCheckPass.Checked)
    {
        if (wizard.RequiredFields.Contains("成績年級"))
            wizard.RequiredFields.Remove("成績年級");
        if (wizard.RequiredFields.Contains("取得學分"))
            wizard.RequiredFields.Remove("取得學分");
    }
};
```

**修改後邏輯：**

```csharp
manulCheckPass.CheckedChanged += delegate
{
    if (manulCheckPass.Checked)
    {
        // 「成績年級」現在設計為無論自動或手動都必填，因此不可移除
        // if (wizard.RequiredFields.Contains("成績年級"))
        //     wizard.RequiredFields.Remove("成績年級");

        // 手動判斷時，只移除「取得學分」的必填限制
        if (wizard.RequiredFields.Contains("取得學分"))
            wizard.RequiredFields.Remove("取得學分");
    }
};
```

**重點：**
- 將原本移除「成績年級」的程式碼註解或刪除。
- 保留／新增註解，說明新設計：即使在手動判斷模式，「成績年級」仍然是必填。

---

## 四、驗證項目（完成修改後測試）

### 測試 1：自動判斷取得學分模式

1. 開啟匯入精靈，選取這個匯入功能。
2. 在選項頁面勾選「自動判斷取得學分」。
3. 進入「欄位對應」畫面：
   - 確認「成績年級」顯示為必填。
   - 不對應「成績年級」時，應無法下一步。
4. 使用一個完整的測試 Excel（有成績年級、取得學分），確認匯入流程可成功完成。

### 測試 2：手動判斷取得學分模式

1. 開啟匯入精靈，切換為「手動判斷取得學分」。  
2. 進入「欄位對應」畫面：
   - 確認「成績年級」仍為必填欄位。
   - 「取得學分」不再是必填欄位，可不對應或不填值。
3. 用測試 Excel：
   - 有填「成績年級」，不填「取得學分」。
   - 驗證匯入可通過。
   - ValidateRow 對「成績年級」格式不正確時仍會報錯。

### 測試 3：補修/重修名冊 GradeYear

1. 準備測試資料，包含：
   - 有補修成績的資料列（是否補修成績 = 是）。
   - 成績年級有填值（例如 1、2、3）。
2. 匯入後檢查：
   - 4.3 補修名冊與 5.2 重修名冊中，`GradeYear` 是否正確寫入匯入年級。
   - 不再出現 GradeYear 空白情形。

### 測試 4：學期科目成績的成績年級變更

1. 找一位已經有學期成績的學生，原本某學年度學期的成績年級已存在（例如原本為 1）。
2. 匯入一筆資料，將同一學年度學期的「成績年級」改成 2。
3. 匯入後檢查：
   - 該學年度學期的學期成績年級是否改成 2。
   - Log 記錄中有顯示成績年級由 1 變更為 2 的紀錄。

---

## 五、文件與紀錄

- 完成以上程式調整與測試後：
  1. 在專案文件或維護紀錄中新增一筆說明：  
     - 說明「成績年級」欄位改為不論自動/手動模式都必填。
     - 說明理由：確保補修/重修名冊與學期成績年級資料一致、避免 GradeYear 空白。
  2. 在 `高中成績調整1121.md` 中紀錄：
     - 調整檔案：`ImportMakeupRetakeScore.cs`
     - 調整內容摘要：
       - RequiredFields 新增「成績年級」為固定必填。
       - 自動/手動判斷取得學分的 RadioButton 事件中，移除成績年級可被取消必填的行為。
       - 驗證步驟與測試狀況（可貼上本 Todo 的測試條目簡化版）。

---

## 六、備註

- 若後續需要再調整「及格標準」判斷（_StudentPassScore）或其他以年級為條件的邏輯，請再重新檢查：
  - `ValidateRow` 中與成績年級相關的判斷。
  - `ImportPackage` 中 `semesterGradeYear` 的使用情況。
- 若發現其他匯入流程也有「成績年級」欄位且存在「自動/手動判斷取得學分」選項，可考慮比照本次作法統一行為。
