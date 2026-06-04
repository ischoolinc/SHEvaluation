
## 目標

調整 `Program.cs`，新增「上學期學分數N」合併欄位，用於上學期科目成績旁邊的學分數顯示。

修改完成後，請完整記錄在：

```text
期末成績通知單(固定排名)調整0604.md
````

---

# 一、重要限制

請嚴格遵守：

```text
1. 只修改 Program.cs。
2. 不要變動原本本學期學分數N的填值邏輯。
3. 不要變動原本科目成績計算邏輯。
4. 不要變動原本上學期科目成績填值邏輯。
5. 不要變動原本學期成績、學年成績、排名計算。
6. 不要變動本學期取得學分數、累計取得學分數計算。
7. 不要重構 Program.cs。
8. 只新增上學期學分數N欄位與對應填值。
```

---

# 二、目前問題

目前報表科目欄位共用：

```text
學分數N
```

但同一個科目列可能同時有：

```text
學期科目成績N
上學期科目成績N
學年科目成績N
```

當本學期有成績、上學期沒有成績時，因為共用 `學分數N`，會造成：

```text
上學期科目成績N 空白
但學分數N 仍顯示
```

本次新增：

```text
上學期學分數N
```

讓上學期成績旁邊使用獨立學分欄位。

---

# 三、新增 DataTable 合併欄位

請在建立每一科目欄位的位置，找到類似：

```csharp
table.Columns.Add("上學期科目原始成績" + subjectIndex);
table.Columns.Add("上學期科目補考成績" + subjectIndex);
table.Columns.Add("上學期科目重修成績" + subjectIndex);
table.Columns.Add("上學期科目手動調整成績" + subjectIndex);
table.Columns.Add("上學期科目學年調整成績" + subjectIndex);
table.Columns.Add("上學期科目成績" + subjectIndex);
```

在後面新增：

```csharp
table.Columns.Add("上學期學分數" + subjectIndex);
```

建議位置：

```csharp
table.Columns.Add("上學期科目成績" + subjectIndex);
table.Columns.Add("上學期學分數" + subjectIndex);
```

---

# 四、上學期學分數填值規則

## 規則

```text
只有上學期科目成績N有值時，才顯示上學期學分數N。
如果上學期科目成績N空白，則上學期學分數N也空白。
```

有效成績包含：

```text
0
60
85
免
缺
其他非空白文字
```

無效成績包含：

```text
null
DBNull
空字串
未輸入
```

---

# 五、建議新增 Helper Method

如果 `Program.cs` 內還沒有類似判斷方法，請新增：

```csharp
/// <summary>
/// 判斷成績欄位是否有有效成績值。
/// 0 是有效成績；空白、null、DBNull、未輸入視為無成績。
/// </summary>
private static bool HasScoreValue(object score)
{
    if (score == null || score == DBNull.Value)
        return false;

    string value = Convert.ToString(score).Trim();

    if (string.IsNullOrEmpty(value))
        return false;

    if (value == "未輸入")
        return false;

    return true;
}
```

建議放在 `Program` 類別內，靠近其他 private static helper method。

如果程式內已經有可用的相同功能方法，請沿用，不要重複新增。

---

# 六、在上學期科目成績填值處新增學分數

請找到上學期科目成績填值區塊，類似：

```csharp
row["上學期科目原始成績" + subjectIndex] = semesterSubjectScore.Detail.GetAttribute("原始成績");
row["上學期科目補考成績" + subjectIndex] = semesterSubjectScore.Detail.GetAttribute("補考成績");
row["上學期科目重修成績" + subjectIndex] = semesterSubjectScore.Detail.GetAttribute("重修成績");
row["上學期科目手動調整成績" + subjectIndex] = semesterSubjectScore.Detail.GetAttribute("擇優採計成績");
row["上學期科目學年調整成績" + subjectIndex] = semesterSubjectScore.Detail.GetAttribute("學年調整成績");
row["上學期科目成績" + subjectIndex] = semesterSubjectScore.Score;
```

在後面新增：

```csharp
if (HasScoreValue(semesterSubjectScore.Score))
{
    row["上學期學分數" + subjectIndex] = semesterSubjectScore.CreditDec();
}
else
{
    row["上學期學分數" + subjectIndex] = "";
}
```

---

# 七、不要修改原本學分數N

請不要修改原本：

```csharp
row["學分數" + subjectIndex] = ...
```

原因：

```text
學分數N目前仍供本學期或原本報表欄位使用。
本次只新增上學期專用欄位，避免影響原本邏輯。
```

---

# 八、Word 樣板使用方式

修改 `Program.cs` 後，Word 樣板若要讓上學期成績旁邊的學分數正確顯示，請把原本上學期區塊使用的：

```text
學分數N
```

改成：

```text
上學期學分數N
```

例如：

```text
上學期學分數1
上學期學分數2
...
上學期學分數N
```

如果樣板仍使用 `學分數N`，畫面仍會顯示原本共用學分數。

---

# 九、測試案例

## 測試 1：上學期有成績

資料：

```text
上學期科目成績8 = 85
CreditDec() = 2
```

預期：

```text
上學期學分數8 = 2
```

---

## 測試 2：上學期成績為 0

資料：

```text
上學期科目成績8 = 0
CreditDec() = 2
```

預期：

```text
上學期學分數8 = 2
```

說明：

```text
0 是有效成績，不可清空學分數。
```

---

## 測試 3：上學期沒有成績

資料：

```text
上學期科目成績8 = 空白
CreditDec() = 2
```

預期：

```text
上學期學分數8 = 空白
```

---

## 測試 4：上學期成績為未輸入

資料：

```text
上學期科目成績8 = 未輸入
CreditDec() = 2
```

預期：

```text
上學期學分數8 = 空白
```

---

## 測試 5：本學期有成績，但上學期沒成績

資料：

```text
學期科目成績8 = 85
上學期科目成績8 = 空白
學分數8 = 2
```

預期：

```text
學分數8 = 2
上學期學分數8 = 空白
```

說明：

```text
本學期學分與上學期學分分開顯示。
```

---

# 十、修改後檢查

請確認：

```text
1. Program.cs 已新增 DataTable 欄位：上學期學分數N。
2. 上學期科目成績填值後，有填入上學期學分數N。
3. 上學期科目成績有值時，上學期學分數N 有值。
4. 上學期科目成績空白時，上學期學分數N 空白。
5. 原本學分數N沒有被修改。
6. 原本科目成績計算沒有被修改。
7. 原本學分統計沒有被修改。
8. 原本排名邏輯沒有被修改。
9. 編譯成功。
```

---

# 十一、完成後紀錄

請建立或更新：

```text
期末成績通知單(固定排名)調整0604.md
```

紀錄內容請包含：

````md
# 期末成績通知單(固定排名)調整0604

## 修改目標

新增上學期學分數合併欄位，用於上學期科目成績旁邊的學分數顯示。

## 修改檔案

- Program.cs

## 新增合併欄位

```text
上學期學分數N
````

N 依原本科目欄位數量產生。

## 修改內容

### 1. 新增 DataTable 欄位

在科目欄位建立區塊新增：

```csharp
table.Columns.Add("上學期學分數" + subjectIndex);
```

### 2. 新增上學期學分數填值

在上學期科目成績填值後，新增：

```csharp
if (HasScoreValue(semesterSubjectScore.Score))
{
    row["上學期學分數" + subjectIndex] = semesterSubjectScore.CreditDec();
}
else
{
    row["上學期學分數" + subjectIndex] = "";
}
```

### 3. 判斷規則

```text
上學期科目成績N有值，才顯示上學期學分數N。
上學期科目成績N空白、null、DBNull、未輸入時，上學期學分數N空白。
0 為有效成績。
```

## 不變動內容

本次未修改：

```text
1. 原本學分數N填值邏輯
2. 本學期科目成績計算
3. 上學期科目成績計算
4. 學期成績排名
5. 本學期取得學分數
6. 累計取得學分數
7. Word 合併流程
```

## Word 樣板注意

上學期成績旁邊的學分欄位，需從：

```text
學分數N
```

改為：

```text
上學期學分數N
```

否則報表仍會使用原本共用學分數欄位。

## 測試結果

請記錄：

```text
1. 上學期有成績時，上學期學分數是否顯示。
2. 上學期成績為 0 時，上學期學分數是否顯示。
3. 上學期沒有成績時，上學期學分數是否空白。
4. 本學期有成績但上學期沒有成績時，上學期學分數是否空白。
5. 原本學分數N是否未受影響。
6. 編譯是否成功。
```

