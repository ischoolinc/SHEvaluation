## 目標

調整 `Program.cs`，新增 14 個 Word 合併欄位變數。

新增欄位分成 2 類：

### 1. 學年類：同一學年度上下學期加總

```text
學年大功統計
學年小功統計
學年嘉獎統計
學年大過統計
學年小過統計
學年警告統計
學年留校察看
````

### 2. 累計類：所有學年度學期加總

```text
累計大功統計
累計小功統計
累計嘉獎統計
累計大過統計
累計小過統計
累計警告統計
累計留校察看
```

修改完成後，請完整記錄在：

```text
期末成績通知單(固定排名)調整0603.md
```

---

# 一、重要限制

請嚴格遵守：

```text
1. 只修改 Program.cs。
2. 不要變動原本本學期獎懲計算邏輯。
3. 不要變動原本上學期獎懲計算邏輯。
4. 不要變動原本缺曠統計。
5. 不要變動原本科目成績計算。
6. 不要變動原本學分數計算。
7. 不要變動原本排名計算。
8. 不要重構 Program.cs。
9. 只在原本獎懲統計區塊附近新增欄位、變數、加總與 DataRow 填值。
10. 新增欄位必須可被 Word MailMerge 使用。
```

---

# 二、目前原本計算邏輯不可變動

目前本學期欄位：

```text
大功統計
小功統計
嘉獎統計
大過統計
小過統計
警告統計
留校察看
```

本學期判斷條件保持不變：

```csharp
("" + info.Semester) == conf.Semester
&&
("" + info.SchoolYear) == conf.SchoolYear
```

目前上學期欄位：

```text
上學期大功統計
上學期小功統計
上學期嘉獎統計
上學期大過統計
上學期小過統計
上學期警告統計
上學期留校察看
```

上學期判斷條件保持不變：

```csharp
conf.Semester == "2"
&&
("" + info.Semester) == "1"
&&
("" + info.SchoolYear) == conf.SchoolYear
```

---

# 三、新增 DataTable 合併欄位

請在 `Program.cs` 建立 `DataTable` 欄位的位置，找到既有獎懲欄位附近，例如：

```csharp
table.Columns.Add("大功統計");
table.Columns.Add("小功統計");
table.Columns.Add("嘉獎統計");
table.Columns.Add("大過統計");
table.Columns.Add("小過統計");
table.Columns.Add("警告統計");
table.Columns.Add("留校察看");
```

在獎懲欄位附近新增以下 14 個欄位：

```csharp
// 學年獎懲統計：同一學年度上下學期加總
table.Columns.Add("學年大功統計");
table.Columns.Add("學年小功統計");
table.Columns.Add("學年嘉獎統計");
table.Columns.Add("學年大過統計");
table.Columns.Add("學年小過統計");
table.Columns.Add("學年警告統計");
table.Columns.Add("學年留校察看");

// 累計獎懲統計：所有學年度學期加總
table.Columns.Add("累計大功統計");
table.Columns.Add("累計小功統計");
table.Columns.Add("累計嘉獎統計");
table.Columns.Add("累計大過統計");
table.Columns.Add("累計小過統計");
table.Columns.Add("累計警告統計");
table.Columns.Add("累計留校察看");
```

注意：

```text
只新增欄位。
不要移除或修改原本欄位。
```

---

# 四、新增統計變數

請在原本獎懲統計區塊中，找到類似：

```csharp
int 大功 = 0;
int 小功 = 0;
int 嘉獎 = 0;
int 大過 = 0;
int 小過 = 0;
int 警告 = 0;
bool 留校察看 = false;

int previous大功 = 0;
int previous小功 = 0;
int previous嘉獎 = 0;
int previous大過 = 0;
int previous小過 = 0;
int previous警告 = 0;
bool previous留校察看 = false;
```

在後面新增：

```csharp
// 學年獎懲統計：同一學年度上下學期加總
int 學年大功 = 0;
int 學年小功 = 0;
int 學年嘉獎 = 0;
int 學年大過 = 0;
int 學年小過 = 0;
int 學年警告 = 0;
bool 學年留校察看 = false;

// 累計獎懲統計：所有學年度學期加總
int 累計大功 = 0;
int 累計小功 = 0;
int 累計嘉獎 = 0;
int 累計大過 = 0;
int 累計小過 = 0;
int 累計警告 = 0;
bool 累計留校察看 = false;
```

注意：

```text
新增變數即可。
不要修改原本 大功、小功、嘉獎、大過、小過、警告、留校察看 變數。
不要修改原本 previous 開頭的上學期變數。
```

---

# 五、新增學年統計計算

請在原本：

```csharp
foreach (RewardInfo info in stuRec.RewardList)
{
    ...
}
```

裡面，保留原本本學期與上學期統計邏輯不變。

在同一個 `foreach` 內新增「學年統計」。

## 學年統計規則

學年類是同一學年度上下學期加總。

條件：

```text
info.SchoolYear == conf.SchoolYear
info.Semester == 1 或 2
```

請新增：

```csharp
// 學年獎懲統計：同一學年度上下學期加總
if (("" + info.SchoolYear) == conf.SchoolYear &&
    (("" + info.Semester) == "1" || ("" + info.Semester) == "2"))
{
    學年大功 += info.AwardA;
    學年小功 += info.AwardB;
    學年嘉獎 += info.AwardC;

    if (!info.Cleared)
    {
        學年大過 += info.FaultA;
        學年小過 += info.FaultB;
        學年警告 += info.FaultC;
    }

    if (info.UltimateAdmonition)
        學年留校察看 = true;
}
```

注意：

```text
獎勵：大功、小功、嘉獎直接加總。
懲戒：大過、小過、警告只統計未銷過資料。
留校察看：同一學年度任一筆 UltimateAdmonition = true，即為 是。
```

---

# 六、新增累計統計計算

請同樣在：

```csharp
foreach (RewardInfo info in stuRec.RewardList)
{
    ...
}
```

裡面新增「累計統計」。

## 累計統計規則

累計類是所有學年度學期加總。

條件：

```text
不限制學年度
不限制學期
stuRec.RewardList 內所有資料都納入
```

請新增：

```csharp
// 累計獎懲統計：所有學年度學期加總
累計大功 += info.AwardA;
累計小功 += info.AwardB;
累計嘉獎 += info.AwardC;

if (!info.Cleared)
{
    累計大過 += info.FaultA;
    累計小過 += info.FaultB;
    累計警告 += info.FaultC;
}

if (info.UltimateAdmonition)
    累計留校察看 = true;
```

注意：

```text
累計獎勵直接加總所有 RewardList。
累計懲戒只加總未銷過資料。
累計留校察看只要任一筆 UltimateAdmonition = true，即為 是。
```

---

# 七、填入 DataRow 合併欄位

請在原本本學期與上學期獎懲欄位填值後面，新增 14 個欄位填值。

原本類似：

```csharp
row["大功統計"] = 大功 == 0 ? "0" : ("" + 大功);
row["小功統計"] = 小功 == 0 ? "0" : ("" + 小功);
row["嘉獎統計"] = 嘉獎 == 0 ? "0" : ("" + 嘉獎);
row["大過統計"] = 大過 == 0 ? "0" : ("" + 大過);
row["小過統計"] = 小過 == 0 ? "0" : ("" + 小過);
row["警告統計"] = 警告 == 0 ? "0" : ("" + 警告);
row["留校察看"] = 留校察看 ? "是" : "否";

row["上學期大功統計"] = previous大功 == 0 ? "0" : ("" + previous大功);
row["上學期小功統計"] = previous小功 == 0 ? "0" : ("" + previous小功);
row["上學期嘉獎統計"] = previous嘉獎 == 0 ? "0" : ("" + previous嘉獎);
row["上學期大過統計"] = previous大過 == 0 ? "0" : ("" + previous大過);
row["上學期小過統計"] = previous小過 == 0 ? "0" : ("" + previous小過);
row["上學期警告統計"] = previous警告 == 0 ? "0" : ("" + previous警告);
row["上學期留校察看"] = previous留校察看 ? "是" : "否";
```

在後面新增：

```csharp
// 學年獎懲統計
row["學年大功統計"] = 學年大功 == 0 ? "0" : ("" + 學年大功);
row["學年小功統計"] = 學年小功 == 0 ? "0" : ("" + 學年小功);
row["學年嘉獎統計"] = 學年嘉獎 == 0 ? "0" : ("" + 學年嘉獎);
row["學年大過統計"] = 學年大過 == 0 ? "0" : ("" + 學年大過);
row["學年小過統計"] = 學年小過 == 0 ? "0" : ("" + 學年小過);
row["學年警告統計"] = 學年警告 == 0 ? "0" : ("" + 學年警告);
row["學年留校察看"] = 學年留校察看 ? "是" : "否";

// 累計獎懲統計
row["累計大功統計"] = 累計大功 == 0 ? "0" : ("" + 累計大功);
row["累計小功統計"] = 累計小功 == 0 ? "0" : ("" + 累計小功);
row["累計嘉獎統計"] = 累計嘉獎 == 0 ? "0" : ("" + 累計嘉獎);
row["累計大過統計"] = 累計大過 == 0 ? "0" : ("" + 累計大過);
row["累計小過統計"] = 累計小過 == 0 ? "0" : ("" + 累計小過);
row["累計警告統計"] = 累計警告 == 0 ? "0" : ("" + 累計警告);
row["累計留校察看"] = 累計留校察看 ? "是" : "否";
```

注意：

```text
輸出格式要跟原本一致：
數字為 0 時輸出 "0"
留校察看輸出 "是" 或 "否"
```

---

# 八、計算規則整理

## 8.1 本學期統計，原本邏輯不變

```text
只統計 conf.SchoolYear + conf.Semester。
```

## 8.2 上學期統計，原本邏輯不變

```text
只在 conf.Semester == "2" 時，統計同學年度第 1 學期。
```

## 8.3 學年統計，新增

```text
統計 conf.SchoolYear 的第 1 學期與第 2 學期。
```

## 8.4 累計統計，新增

```text
統計 stuRec.RewardList 內所有學年度、所有學期。
```

## 8.5 獎勵規則

```text
大功、小功、嘉獎直接加總。
不判斷 Cleared。
```

## 8.6 懲戒規則

```text
大過、小過、警告只統計 info.Cleared == false 的資料。
```

## 8.7 留校察看規則

```text
範圍內任一筆 info.UltimateAdmonition == true，輸出「是」。
否則輸出「否」。
```

---

# 九、測試案例

## 測試 1：本學期資料

資料：

```text
114-1 大功 1、小功 2、嘉獎 3
```

報表設定：

```text
學年度 = 114
學期 = 1
```

預期：

```text
大功統計 = 1
小功統計 = 2
嘉獎統計 = 3
學年大功統計 = 1
學年小功統計 = 2
學年嘉獎統計 = 3
累計大功統計 = 1
累計小功統計 = 2
累計嘉獎統計 = 3
```

---

## 測試 2：同學年度上下學期加總

資料：

```text
114-1 嘉獎 2
114-2 嘉獎 3
```

報表設定：

```text
學年度 = 114
學期 = 2
```

預期：

```text
嘉獎統計 = 3
上學期嘉獎統計 = 2
學年嘉獎統計 = 5
累計嘉獎統計 = 5
```

---

## 測試 3：跨學年度累計

資料：

```text
113-2 小功 1
114-1 小功 2
114-2 小功 3
```

報表設定：

```text
學年度 = 114
學期 = 2
```

預期：

```text
小功統計 = 3
上學期小功統計 = 2
學年小功統計 = 5
累計小功統計 = 6
```

---

## 測試 4：懲戒銷過

資料：

```text
114-1 小過 1，Cleared = true
114-1 小過 2，Cleared = false
```

報表設定：

```text
學年度 = 114
學期 = 1
```

預期：

```text
小過統計 = 2
學年小過統計 = 2
累計小過統計 = 2
```

說明：

```text
Cleared = true 的懲戒不列入統計。
```

---

## 測試 5：留校察看

資料：

```text
113-2 UltimateAdmonition = true
114-1 UltimateAdmonition = false
```

報表設定：

```text
學年度 = 114
學期 = 1
```

預期：

```text
留校察看 = 否
學年留校察看 = 否
累計留校察看 = 是
```

---

# 十、修改後檢查

請確認：

```text
1. DataTable 已新增 14 個欄位。
2. 每位學生的 DataRow 都有填入 14 個欄位。
3. 原本本學期獎懲統計沒有被改動。
4. 原本上學期獎懲統計沒有被改動。
5. 獎勵仍直接加總。
6. 懲戒仍只統計未銷過資料。
7. 留校察看仍輸出「是 / 否」。
8. 缺曠、成績、學分、排名計算沒有被改動。
9. 編譯成功。
10. Word 樣板可使用新增合併欄位。
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

新增 14 個獎懲統計合併欄位，不變動原本計算邏輯。

## 修改檔案

- Program.cs

## 新增合併欄位

### 學年類：同一學年度上下學期加總

- 學年大功統計
- 學年小功統計
- 學年嘉獎統計
- 學年大過統計
- 學年小過統計
- 學年警告統計
- 學年留校察看

### 累計類：所有學年度學期加總

- 累計大功統計
- 累計小功統計
- 累計嘉獎統計
- 累計大過統計
- 累計小過統計
- 累計警告統計
- 累計留校察看

## 計算方式

### 學年統計

同一學年度上下學期加總：

```text
info.SchoolYear == conf.SchoolYear
info.Semester == 1 或 2
````

### 累計統計

所有 `stuRec.RewardList` 資料加總，不限制學年度與學期。

### 獎勵

大功、小功、嘉獎直接加總。

### 懲戒

大過、小過、警告只統計未銷過資料：

```csharp
if (!info.Cleared)
```

### 留校察看

只要符合範圍內任一筆：

```csharp
info.UltimateAdmonition == true
```

輸出：

```text
是
```

否則輸出：

```text
否
```

## 不變動內容

本次未修改：

```text
1. 本學期獎懲統計邏輯
2. 上學期獎懲統計邏輯
3. 缺曠統計
4. 科目成績計算
5. 學分數計算
6. 排名計算
7. Word 合併流程
```

## 測試結果

請記錄：

1. 本學期獎懲統計是否與修改前一致。
2. 上學期獎懲統計是否與修改前一致。
3. 學年獎懲統計是否為同一學年度上下學期加總。
4. 累計獎懲統計是否為所有學年度學期加總。
5. 銷過懲戒是否未列入大過、小過、警告。
6. 留校察看是否正確顯示是 / 否。
7. 編譯是否成功。
