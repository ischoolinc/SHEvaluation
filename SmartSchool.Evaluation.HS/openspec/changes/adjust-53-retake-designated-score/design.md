# Design

## Context

見 `proposal.md` 的 Why。這裡只記已核對的程式關係，以及實作必須守住的界線。

`sems_subj_score.score_info` 的科目節點是 `SemesterSubjectScoreInfo/Subject`。直接指定總成績的 XML 屬性名稱是「修課直接指定總成績」，及格標準是同一節點的「修課及格標準」。計算學期科目成績時，`designate_final_score` 可解析就寫入該屬性（已進位的數字字串），否則寫空字串。匯入學期科目成績與匯入重補修成績也寫同一個屬性名稱。

產生學習歷程的 5.3 更新在 `StudentLearningHistoryProcessor.ProcessLearningHistory`：

1. `GetLearningHistoryRetakeDataAsDictionary53` 只讀 `student_learning_history` 中 `serial_no = 5.3`、`name = 重讀成績`、且為本次學年度學期的既有列。再次修習成績對到 `Score`，再次修習成績及格對到 `ScoreP`。
2. `GetStudentScoreDataAsDictionary` 只服務這段 5.3。它解析同一份 `score_info`，但 SQL 要求「補考成績」非空白，而且目前只取出原始成績、補考成績、修課及格標準。
3. 比對鍵是學生、科目、科目級別，以及本次學年度、學期。有比對到才回寫補考成績、補考及格，並計算重讀成績。
4. 重讀成績取三個可解析值的最大值：該列目前的再次修習成績、補考成績、原始成績。沒有任何可解析值時，重讀成績為空、成績及格為 `-1`。
5. 最後把這份已讀入的清單交給 `SaveScores53`。

`GetStudentScoreDataAsDictionary43` 只給 4.3。4.2、4.4、6.2、6.4 走各自的組裝與 `SaveScores42`、`SaveScores44`、`SaveScores62`、`SaveScores64`。4.2 排除 5.3 時只用學生、科目、科目級別，不用分數。

`WearyDogComputer.FillSemesterSubjectCalcScore` 在 `ruleType == "再次修習"` 時建立 5.3。規則來自「重讀(擇優採計成績)」或「再次修習」。兩個分支都要改，而且只改放入 `repeatScoreList` 的 `Score`、`ScoreP`：

- 本學期科目成績已存在：`currentSubjectScoreList` 有相同科目鍵，`fromPrevSemester` 為真，`repeatRec` 約第 1475 行。同一段稍早用 `studentFinalScoreDict` 把可解析的 `designate_final_score` 以 `GetRoundScore` 寫入「修課直接指定總成績」。
- 本學期科目成績尚未存在：進入新增成績區塊，`repeatRec` 約第 1917 行。指定總成績同樣寫入新節點的「修課直接指定總成績」。

目前這兩筆 5.3.7 都是 `GetRoundScore(sacRecord.FinalScore, decimals, mode)`。5.3.8 使用共用的 `ScoreP`：有 `HasFinalScore` 且未進位的 `FinalScore` 大於等於修課及格標準才是 `1`，否則是 `0`。這個共用 `ScoreP` 也給 4.4 轉學轉科使用，所以 5.3 必須另算，不能改共用變數。

建立條件維持：`fromPrevSemester`、`不計學分` 為「否」，以及後續 `SaveScores53` 的 `checkPass`、`CodePass`。`fromPrevSemester` 來自先前學期成績或封存成績；指定總成績進位後大於等於 0 時，既有程式也可能把這個旗標設成真。5.3.11 仍是原始成績，5.3.9、5.3.10 仍是既有補考截斷與及格判斷。

## Goals / Non-Goals

**Goals:**

- 讓既有 5.3 在沒有補考成績時仍讀得到「修課直接指定總成績」。
- 有效指定總成績先寫入再次修習成績並重算再次修習成績及格，然後既有補考與重讀成績區塊照原順序執行。
- 指定總成績無效時，不改既有 5.3 補考查詢的篩選與最大值計算。
- 計算成績建立 5.3 時，兩個再次修習分支的 5.3.7、5.3.8 與產生學習歷程使用同一套是否採用指定總成績的判斷，但 5.3.7 要先走既有進位，使它等於 `score_info` 的「修課直接指定總成績」。

**Non-Goals:**

- 不放寬 `GetStudentScoreDataAsDictionary` 的補考成績篩選，也不改 `GetStudentScoreDataAsDictionary43`。
- 不修改共用 SQL 範本的新增分支，不改 4.2、4.3、4.4、5.2、6.2、6.4。
- 不補讀目前未映射的「重讀成績」「成績及格」，也不修正既有重讀成績及格在標準解析失敗時變成 0 的寫法。
- 不改擇優、補考截斷、`fromPrevSemester`、不計學分與課程代碼檢查，也不改 `studentFinalScoreDict` 的覆蓋方式。
- 不改 4.4 使用的共用 `ScoreP`，以及 5.3.9 到 5.3.12。

## Decisions

1. **另做只給 5.3 的指定總成績讀取，不改既有補考查詢。**
   既有查詢把沒有補考成績的科目排除。若把篩選拿掉，沒有指定總成績的列也會進入補考與最大值區塊，違反「無效時完全保留原邏輯」。新讀取只取本次學年度學期、同一批已有 5.3 的學生，欄位為科目、科目級別、修課直接指定總成績、修課及格標準。不要求補考成績存在。
   替代方案是加寬既有查詢再在迴圈裡分支。那樣會把兩種入場條件綁在一起，無效指定總成績時也改變可被比對到的列。

2. **先改再次修習成績，再跑既有區塊。**
   對每一筆已讀入的 5.3，用學生、科目、科目級別、本次學年度、學期找科目節點。屬性去掉前後空白後，`decimal.TryParse` 成功才算有效，包含 0、負數與超過 100；寫入去掉空白後的原文字，不再進位。成功時以同一節點的修課及格標準重算再次修習成績及格：解析成功才採用該標準，否則用 60；成績大於等於標準為 `1`，否則為 `0`。這段不產生 `-1`。
   既有補考與最大值迴圈維持原位。再次修習成績已經被指定總成績取代時，最大值的第一個候選就是這個新值。沒有補考列時，該迴圈仍不計算重讀成績。
   替代方案是指定總成績有效時跳過補考與最大值。需求是不受補考「是否存在」影響來源選擇，不是在有補考時停用原計算。

3. **只變更已讀入的 5.3 物件，沿用原清單呼叫 `SaveScores53`。**
   不把新物件加入清單。`SaveScores53` 的 SQL 在鍵對不到時仍可能新增，該範本也供其他名冊使用，所以不改範本。鍵一致時走更新。`RepeatScore` 未被這次設定時維持 null；序列化後 SQL `COALESCE` 會留下資料庫裡的重讀成績，因此沒有進入最大值區塊的列不會被空字串清掉。
   替代方案是改範本禁止新增。那會影響 4.2、4.3、4.4、6.2、6.4。

4. **有效數值沿用這段既有的 `decimal.TryParse`，不加 0 到 100。**
   匯入與計算寫入這個屬性時，條件都是空白或可解析數字，沒有另限範圍。新的及格判斷只在 `TryParse` 成功時覆蓋預設 60，避免沿用重讀成績區塊「解析失敗仍寫入 out 而變成 0」的寫法。

5. **計算成績只替換兩個分支的 5.3.7、5.3.8。**
   在組 `repeatRec` 前，用該分支已經取到的 `studentFinalScoreDict` 列讀 `designate_final_score`。去掉空白後 `TryParse` 成功，包含 0，就令 5.3.7 為 `GetRoundScore(指定值, decimals, mode).ToString()`，與該分支寫入「修課直接指定總成績」的進位結果相同。5.3.8 解析該節點已寫入的「修課及格標準」；只有解析成功才採用，否則用 60，再拿進位後的 5.3.7 比較。這段不看 `HasFinalScore`，也不看補考成績。
   指定總成績不存在、空白或無法解析時，5.3.7 維持 `GetRoundScore(sacRecord.FinalScore, decimals, mode).ToString()`，5.3.8 維持「有 `HasFinalScore` 且未進位的 `FinalScore` 大於等於標準才是 `1`，否則是 `0`」。
   替代方案是直接改共用 `ScoreP`。那會一併改變同區塊的 4.4。

6. **多門修課沿用目前字典，不另選課程。**
   `studentFinalScoreDict` 的鍵是修剪後的科目加底線加級別。同一鍵已存在時，後面的 `DataRow` 直接覆蓋。查詢沒有 `ORDER BY`。`studentHasFinalScoreDict` 卻只保留第一筆可解析的指定總成績，兩者可能不同。5.3.7 必須跟該分支寫入 XML 的那一列相同，也就是 `studentFinalScoreDict` 的最後一列，不改用 `studentHasFinalScoreDict`，也不新增排序。
   查找鍵在再次修習分支是 `sacRecord.Subject + "_" + sacRecord.SubjectLevel`，字典鍵有做 `Trim`。兩邊空白不一致時，目前就不會寫入指定總成績；5.3 沿用同一次查找，不另外放寬。

## Risks / Trade-offs

- [同一學期同一科目與級別有多個 `Subject` 節點] → 與既有 5.3 比對一樣取第一筆，不另訂優先順序。實作時新查詢與舊比對使用同一組鍵。
- [同一學生、科目、級別有多筆修課] → `studentFinalScoreDict` 保留最後一列，而且沒有排序，選到哪一筆不穩定。`studentHasFinalScoreDict` 又保留第一筆可解析值。這次不改字典與查詢，5.3.7 只跟 XML 使用的最後一列。若日後要指定課程或固定排序，會影響所有使用這個字典的成績，需另開變更。
- [只跑其中一個功能] → 只跑計算時，新建立的 5.3.7、5.3.8 直接用進位後的指定總成績。只跑產生學習歷程時，仍讀已存在的 `score_info` 原文，不再進位。兩邊都跑時，計算先寫進位值，產生學習歷程再讀同一個屬性，5.3.7 應維持該進位文字。
- [既有列缺身分證號或課程代碼，或課程代碼不可提交] → `SaveScores53` 目前會跳過。這次不放寬，該列的記憶體修改不會入庫。
- [`SaveScores53` 鍵不一致時仍會新增] → 這次不新增清單項目，也不改共用範本。鍵不一致而新增是既有行為。
- [沒有補考時重讀成績不重算] → 最大值區塊仍只在原補考比對成功時執行。指定總成績只保證改再次修習成績與再次修習成績及格。

## Migration Plan

沒有結構變更，也不做一次性資料修補。下一次產生學習歷程會更新當時已存在的 5.3。回復方式是還原程式；已經寫入的分數留到再跑一次產生學習歷程。

## Open Questions

- 產生學習歷程讀到的同一 `score_info` 若有多筆相同科目與級別，目前程式與這次設計都取第一筆。若日後要改為指定節點或取最大指定總成績，會改變選值結果，需另開變更。
- 計算成績的 `studentFinalScoreDict` 在多門修課時覆蓋規則不穩定。這次不自行加入排序或改其他模組。
- `GetLearningHistoryRetakeDataAsDictionary53` 有查出「重讀成績」「成績及格」，但沒有指定給 `RepeatScore`、`RepeatScoreP`。這次依現況保留，不補映射。
