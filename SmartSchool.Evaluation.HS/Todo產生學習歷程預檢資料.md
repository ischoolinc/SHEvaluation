# Todo.md — 新增「產生學習歷程預檢資料」功能（0127）

> 目標：新增產生預檢資料功能（來源：學生修課紀錄 sc_attend），寫入 student_learning_history  
> - 日校：serial_no = **41.2**、name = **學期成績(預檢)**、JSON schema **同 4.2**
> - 進校：serial_no = **42.2**、name = **學期成績(預檢)**、JSON schema **同 6.2**
> - 原本 ProcessLearningHistory(...) 流程 **不修改、不影響**
> - 修改完成後，請更新紀錄檔：**產生學習歷程預檢資料0127.md**

---

## 0. 前置確認（避免做錯）
- [ ] 確認資料表/欄位：
  - `student_learning_history` 既有寫入規則（serial_no, name, detail JSON）
  - `sc_attend.subject_code` 來源是否就是課程代碼（CourseCode）
- [ ] 確認預檢資料是否需要「包含不可提交課程代碼（8/9 非 D）」：
  - 建議預檢仍寫入，但在既有欄位（例如 Text / DataKey）寫入提示；**不要因 CodePass=false 而略過寫入**
- [ ] 確認科目級別 `subj_level` 規格：**正整數或空白**（不合格需記錄錯誤並略過寫入或標記 checkPass=false）

---

## 1. 新增資料來源：修課學生查詢（SQL 參數化）
> 來源參考（你提供的 SQL）：course + sc_attend + student + class，並取 `sc_attend.subject_code` 作為課程代碼。

- [ ] 在 DataAccess（或 StudentLearningHistoryProcessor 內部）新增方法：  
  `DataTable GetSCAttendCourseRows(int schoolYear, int semester, List<string> studentIds)`  
  - [ ] 將 SQL 改為參數化：`@SchoolYear`, `@Semester`, `@StudentIDs`
  - [ ] WHERE 建議包含：`student.status IN (1,2)` 且 `course.school_year=@SchoolYear` `course.semester=@Semester`
  - [ ] SELECT 至少包含（供組 DTO）：
    - student：id, name, student_number, seat_no, id_number, birthdate
    - class：class_name, grade_year
    - course：id, course_name, subject, subj_level, credit, period, score_type, school_year, semester
    - 必修/選修、部定/校訂：用 COALESCE 做出文字
    - `sc_attend.subject_code`（課程代碼）
    - `sc_attend.extensions`（若你要沿用 duplicated rule/補修判斷，可加）
  - [ ] ORDER BY 可留在 UI 匯出端，不一定要寫死在 SQL

---

## 2. 新增 Processor：ProcessLearningHistory_FromSCAttend(...)
> 完全獨立於既有 ProcessLearningHistory(...)；只負責「sc_attend -> DTO -> 寫入 41.2 / 42.2」

- [ ] 在 `StudentLearningHistoryProcessor.cs` 新增方法：
  ```csharp
  public void ProcessLearningHistory_FromSCAttend(
      AccessHelper helper,
      List<StudentRecord> students,
      int schoolYear,
      int semester,
      bool isNightSchool,
      BackgroundWorker bkw)
  ```
- [ ] 主要流程：
  1) 建 studentIdList  
  2) 呼叫 `GetSCAttendCourseRows(...)` 取得 DataTable  
  3) 逐列轉成 `SubjectScoreRec108`（或你 writer 現用 DTO）
  4) 驗證欄位（與 UI 同步）：
     - [ ] 身分證號 `id_number`：必填（或至少 checkPass=false）
     - [ ] 科目名稱 `subject`：必填
     - [ ] 科目級別 `subj_level`：空白或正整數
     - [ ] 課程代碼 `subject_code`：必填
  5) 設定 DTO 的必要欄位（務必讓 JSON keys 能完整產出）
     - [ ] SchoolYear/Semester/GradeYear/Credit/Period/ScoreType/Required/RequiredBy/…（依現有 schema）
     - [ ] 成績欄位若 sc_attend 取不到：填空值/預設值（但 key 必須存在）
  6) 寫入 DB：
     - [ ] if (!isNightSchool) → 呼叫 `SaveScores41_2(...)`
     - [ ] else → 呼叫 `SaveScores42_2(...)`

- [ ] 錯誤回報（給 Wizard / ErrorViewer）：
  - [ ] 將「不合法的 subj_level / 缺 id_number / 缺 subject_code」整理成 `Dictionary<StudentRecord, List<string>>`
  - [ ] 使用 `bkw.ReportProgress(percent, errorDict)` 回報（與你既有 wizard 接法一致）

---

## 3. 新增 DataAccess 寫入：SaveScores41_2 / SaveScores42_2（或參數化共用）
> 你的需求是：JSON schema 不變，差別只在 serial_no/name。  
> 最小改動可先新增兩個方法（copy 42/62），之後再抽共用。

### 3.1 最小改動（先做得出來）
- [ ] 在 `LearningHistoryDataAccess.cs` 新增：
  - [ ] `SaveScores41_2(List<SubjectScoreRec108> scores, int schoolYear, int semester)`
  - [ ] `SaveScores42_2(List<SubjectScoreRec108> scores, int schoolYear, int semester)`
- [ ] 兩者邏輯比照既有 `SaveScores42` / `SaveScores62`：
  - [ ] `CreateJSubjectInfo(...)` 的 serialNo/name 改為：
    - 41.2：serialNo="41.2", name="學期成績(預檢)"
    - 42.2：serialNo="42.2", name="學期成績(預檢)"
  - [ ] detail serial 前綴同步改為：
    - 4.2.x（1..20）
    - 6.2.x（1..20）
  - [ ] SQL template 新增兩個（內容同既有 template，但 serial prefix 變更）：
    - `GetSemesterScoreSqlTemplate41_2(...)`
    - `GetSemesterScoreSqlTemplate42_2(...)`
- [ ] 預檢寫入建議：
  - [ ] **不要因 CodePass=false 而略過寫入**（預檢應該能看見不可提交資料）
  - [ ] 若仍要提示：把原因寫入既有欄位（如 Text / DataKey / Memo），但不新增 JSON 欄位（維持 schema 不變）

### 3.2 進階重構（可選，驗證 OK 後再做）
- [ ] 抽共用 `SaveScoresBySerial(...)`，讓 serial_no/name/detailPrefix 參數化
- [ ] 讓 4.2/6.2/41.2/42.2 共用同一套 JSON mapping（避免技術債）

---

## 4. UI / Wizard 串接（新增「產生預檢資料」入口）
> 讓使用者能選擇產生「原本學習歷程」或「預檢資料」

- [ ] 在 `CalcLearningHistoryPrvScoreWizard.cs`（或相關 UI）新增選項：
  - [ ] 勾選「產生預檢資料」
  - [ ] 互斥選擇「日校(41.2) / 進校(42.2)」或沿用你既有日/進 checkbox 邏輯
- [ ] 在背景處理（DoWork）分流：
  - [ ] 若「預檢」→ 呼叫 `ProcessLearningHistory_FromSCAttend(...)`
  - [ ] 否則 → 維持原本 `ProcessLearningHistory(...)`

---

## 5. 驗證與回歸測試
- [ ] 用小樣本（1 班 / 10 人）測：
  - [ ] 41.2 是否寫入 student_learning_history（serial_no/name 正確）
  - [ ] 42.2 是否寫入 student_learning_history（serial_no/name 正確）
  - [ ] JSON keys 是否完整（與 4.2/6.2 一致）
- [ ] 測 subj_level 非法資料：
  - [ ] 空白 OK
  - [ ] "1" OK
  - [ ] "0" / "-1" / "A" 應被擋下並記錄
- [ ] 測 subject_code 缺值：
  - [ ] 應列入錯誤清單或 checkPass=false，不可寫入（或寫入但標記，依你的規格）
- [ ] 回歸：原本 4.2 / 6.2 / 4.3 / 5.3 / 6.3 流程不受影響（跑一次完整 wizard）

---

## 6. 文件紀錄（必做）
- [ ] 新增/更新：`產生學習歷程預檢資料0127.md`
  - [ ] 記錄新增功能目的與範圍
  - [ ] 記錄寫入辨識方式（41.2/42.2）
  - [ ] 記錄 SQL 來源（修課學生）與欄位 mapping
  - [ ] 記錄已知限制（成績欄位空值策略、CodePass 策略、身分證缺漏策略）
  - [ ] 記錄測試案例與結果（至少 41.2/42.2 各一例）

---

## 7. 交付檢查清單
- [ ] 程式可編譯
- [ ] 產生預檢資料按鈕/選項可用
- [ ] DB 有成功寫入 41.2 / 42.2 且 name 正確
- [ ] JSON keys 完整一致（schema 不變）
- [ ] `產生學習歷程預檢資料0127.md` 已更新
