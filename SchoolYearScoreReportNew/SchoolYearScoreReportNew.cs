using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Windows.Forms;
using System.Xml;
using Aspose.Words;
using Aspose.Words.Reporting;
using SchoolYearScoreReport;
using SmartSchool.Customization.Data;
using SmartSchool.Customization.Data.StudentExtension;
using SmartSchool.Customization.PlugIn;
using SmartSchool.Customization.PlugIn.Report;
using FISCA.Permission;
using FISCA.Presentation;

class SchoolYearScoreReportNew
{
    // Fields
    private ButtonAdapter buttonClass;
    private ButtonAdapter buttonStudent;

    // Methods
    public SchoolYearScoreReportNew()
    {
        string reportName = "學年成績單(新制)";
        string reportPath = "成績相關報表";
        this.buttonStudent = new ButtonAdapter();
        this.buttonStudent.Text = reportName;
        this.buttonStudent.Path = reportPath;
        this.buttonStudent.OnClick += new EventHandler(this.buttonStudent_OnClick);
        this.buttonClass = new ButtonAdapter();
        this.buttonClass.Text = reportName;
        this.buttonClass.Path = reportPath;
        this.buttonClass.OnClick += new EventHandler(this.buttonClass_OnClick);
        StudentReport.AddReport(this.buttonStudent);
        ClassReport.AddReport(this.buttonClass);

        string Student = "SHEvaluation.SchoolYearScoreReportNew.Student";
        string Class = "SHEvaluation.SchoolYearScoreReportNew.Class";
        RibbonBarItem item1 = FISCA.Presentation.MotherForm.RibbonBarItems["學生", "資料統計"];
        item1["報表"][reportPath][reportName].Enable = FISCA.Permission.UserAcl.Current[Student].Executable;
        RibbonBarItem item2 = FISCA.Presentation.MotherForm.RibbonBarItems["班級", "資料統計"];
        item2["報表"][reportPath][reportName].Enable = FISCA.Permission.UserAcl.Current[Class].Executable;

        //權限設定
        Catalog permission1 = RoleAclSource.Instance["學生"]["報表"];
        permission1.Add(new RibbonFeature(Student, reportName));
        Catalog permission2 = RoleAclSource.Instance["班級"]["報表"];
        permission2.Add(new RibbonFeature(Class, reportName));
    }

    private void bkw_DoWork(object sender, DoWorkEventArgs e)
    {
        object[] objects = (object[])e.Argument;
        BackgroundWorker bkw = (BackgroundWorker)objects[0];
        Config config = (Config)objects[1];
        AccessHelper helper = (AccessHelper)objects[2];
        List<StudentRecord> allStudent = (List<StudentRecord>)objects[3];
        bkw.ReportProgress(1);
        allStudent = this.ProcessStudentData(helper, allStudent, bkw);
        int currentStudent = 1;
        int totalStudent = allStudent.Count;
        SmartSchool.Customization.Data.SystemInformation.getField("文字評量對照表");
        Document doc = new Document();
        doc.Sections.Clear();

        List<K12.Data.PeriodMappingInfo> periodMappingInfos = K12.Data.PeriodMapping.SelectAll();
        Dictionary<string, string> dicPeriodMappingType = new Dictionary<string, string>();
        List<string> periodTypes = new List<string>();
        foreach (K12.Data.PeriodMappingInfo periodMappingInfo in periodMappingInfos)
        {
            if (!dicPeriodMappingType.ContainsKey(periodMappingInfo.Name))
                dicPeriodMappingType.Add(periodMappingInfo.Name, periodMappingInfo.Type);

            if (!periodTypes.Contains(periodMappingInfo.Type))
                periodTypes.Add(periodMappingInfo.Type);
        }

        foreach (StudentRecord var in allStudent)
        {
            Document eachStudentDoc = new Document(new MemoryStream(config.Template), "", LoadFormat.Doc, "");
            Dictionary<string, object> mergeKeyValue = new Dictionary<string, object>();
            int currentGradeYear = 0;
            foreach (SemesterSubjectScoreInfo info in var.SemesterSubjectScoreList)
            {
                if (info.SchoolYear == config.SchoolYear)
                {
                    currentGradeYear = info.GradeYear;
                    break;
                }
            }
            mergeKeyValue.Add("學校名稱", SmartSchool.Customization.Data.SystemInformation.SchoolChineseName);
            mergeKeyValue.Add("學校地址", SmartSchool.Customization.Data.SystemInformation.Address);
            mergeKeyValue.Add("學校電話", SmartSchool.Customization.Data.SystemInformation.Telephone);
            if (config.ReceiveNameIndex == 0)
            {
                mergeKeyValue.Add("收件人", var.ParentInfo.CustodianName);
            }
            else if (config.ReceiveNameIndex == 1)
            {
                mergeKeyValue.Add("收件人", var.ParentInfo.FatherName);
            }
            else if (config.ReceiveNameIndex == 2)
            {
                mergeKeyValue.Add("收件人", var.ParentInfo.MotherName);
            }
            else if (config.ReceiveNameIndex == 3)
            {
                mergeKeyValue.Add("收件人", var.StudentName);
            }
            if (config.ReceiveAddressIndex == 0)
            {
                mergeKeyValue.Add("收件人地址", var.ContactInfo.PermanentAddress.FullAddress);
            }
            else if (config.ReceiveAddressIndex == 1)
            {
                mergeKeyValue.Add("收件人地址", var.ContactInfo.MailingAddress.FullAddress);
            }
            mergeKeyValue.Add("學年度", config.SchoolYear.ToString());
            mergeKeyValue.Add("班級科別名稱", (var.RefClass != null) ? var.RefClass.Department : "");
            mergeKeyValue.Add("班級", (var.RefClass != null) ? var.RefClass.ClassName : "");
            mergeKeyValue.Add("學號", var.StudentNumber);
            mergeKeyValue.Add("姓名", var.StudentName);
            mergeKeyValue.Add("座號", var.SeatNo);
            mergeKeyValue.Add("班導師", "");
            if ((var.RefClass != null) && (var.RefClass.RefTeacher != null))
            {
                mergeKeyValue["班導師"] = var.RefClass.RefTeacher.TeacherName;
            }
            List<int> semester1SchoolYears = new List<int>();//上學期的學年度清單，評語及缺曠獎懲統計用，有重讀的情況時，與成績只採最後學期不同，所有學期都統計進去。
            List<int> semester2SchoolYears = new List<int>();//下學期的學年度清單，評語及缺曠獎懲統計用，有重讀的情況時，與成績只採最後學期不同，所有學期都統計進去。
            foreach (var history in var.SemesterHistoryList)
            {
                if (history.GradeYear == currentGradeYear)
                {
                    if (history.Semester == 1)
                        semester1SchoolYears.Add(history.SchoolYear);
                    if (history.Semester == 2)
                        semester2SchoolYears.Add(history.SchoolYear);
                }
            }
            mergeKeyValue.Add("第一學期評語", "");
            mergeKeyValue.Add("第二學期評語", "");
            if (var.SemesterMoralScoreList.Count > 0)
            {
                foreach (SemesterMoralScoreInfo info in var.SemesterMoralScoreList)
                {
                    if (info.Semester == 1 && semester1SchoolYears.Contains(info.SchoolYear))
                    {
                        mergeKeyValue["第一學期評語"] = info.SupervisedByComment;
                    }
                    else if (info.Semester == 2 && semester2SchoolYears.Contains(info.SchoolYear))
                    {
                        mergeKeyValue["第二學期評語"] = info.SupervisedByComment;
                    }
                }
            }
            Dictionary<int, XmlElement> moralDetail = new Dictionary<int, XmlElement>();
            foreach (SemesterMoralScoreInfo info in var.SemesterMoralScoreList)
            {

                if ((info.Semester == 1 && semester1SchoolYears.Contains(info.SchoolYear)) ||
                    (info.Semester == 2 && semester2SchoolYears.Contains(info.SchoolYear)))
                {
                    if (moralDetail.ContainsKey(info.Semester))
                        moralDetail[info.Semester] = info.Detail;
                    else
                        moralDetail.Add(info.Semester, info.Detail);
                }
            }
            if (moralDetail.Count > 0)
            {
                mergeKeyValue.Add("綜合表現", moralDetail);
            }
            else
            {
                mergeKeyValue.Add("綜合表現", null);
            }
            Dictionary<string, Dictionary<string, int[]>> absenceRewardInfos = new Dictionary<string, Dictionary<string, int[]>>();
            absenceRewardInfos.Add("獎勵", new Dictionary<string, int[]>());
            absenceRewardInfos.Add("懲戒", new Dictionary<string, int[]>());
            absenceRewardInfos["獎勵"].Add("大功", new int[2]);
            absenceRewardInfos["獎勵"].Add("小功", new int[2]);
            absenceRewardInfos["獎勵"].Add("嘉獎", new int[2]);
            absenceRewardInfos["懲戒"].Add("大過", new int[2]);
            absenceRewardInfos["懲戒"].Add("小過", new int[2]);
            absenceRewardInfos["懲戒"].Add("警告", new int[2]);
            foreach (RewardInfo info in var.RewardList)
            {
                //if (info.SchoolYear == ((int)config.SchoolYear))
                if ((info.Semester == 1 && semester1SchoolYears.Contains(info.SchoolYear)) ||
                    (info.Semester == 2 && semester2SchoolYears.Contains(info.SchoolYear)))
                {
                    absenceRewardInfos["獎勵"]["大功"][info.Semester - 1] += info.AwardA;
                    absenceRewardInfos["獎勵"]["小功"][info.Semester - 1] += info.AwardB;
                    absenceRewardInfos["獎勵"]["嘉獎"][info.Semester - 1] += info.AwardC;
                    if (!info.Cleared)
                    {
                        absenceRewardInfos["懲戒"]["大過"][info.Semester - 1] += info.FaultA;
                        absenceRewardInfos["懲戒"]["小過"][info.Semester - 1] += info.FaultB;
                        absenceRewardInfos["懲戒"]["警告"][info.Semester - 1] += info.FaultC;
                    }
                }
            }
            mergeKeyValue.Add("缺曠獎懲", absenceRewardInfos);
            foreach (string periodType in config.SelectTypes.Keys)
            {
                if (!absenceRewardInfos.ContainsKey(periodType) && periodTypes.Count > 0 && periodTypes.Contains(periodType))
                {
                    absenceRewardInfos.Add(periodType, new Dictionary<string, int[]>());
                }
                foreach (string absence in config.SelectTypes[periodType])
                {
                    if (!absenceRewardInfos.ContainsKey(periodType))
                        continue;
                    if (!absenceRewardInfos[periodType].ContainsKey(absence) && periodTypes.Count > 0 && periodTypes.Contains(periodType))
                    {
                        absenceRewardInfos[periodType].Add(absence, new int[2]);
                    }
                }
            }
            foreach (AttendanceInfo info in var.AttendanceList)
            {
                string infoType = string.Empty;
                if (dicPeriodMappingType.ContainsKey(info.Period))
                    infoType = dicPeriodMappingType[info.Period];
                else
                    infoType = string.Empty;

                if ((info.Semester == 1 && semester1SchoolYears.Contains(info.SchoolYear)) ||
                    (info.Semester == 2 && semester2SchoolYears.Contains(info.SchoolYear)))
                {
                    if (info.Semester == 1)
                    {
                        if (absenceRewardInfos.ContainsKey(infoType) && absenceRewardInfos[infoType].ContainsKey(info.Absence))
                        {
                            absenceRewardInfos[infoType][info.Absence][0] += 1;
                        }
                    }
                    else if ((info.Semester == 2) && (absenceRewardInfos.ContainsKey(infoType) && absenceRewardInfos[infoType].ContainsKey(info.Absence)))
                    {
                        absenceRewardInfos[infoType][info.Absence][1] += 1;
                    }
                }
            }
            Dictionary<int, decimal> standard = new Dictionary<int, decimal>();
            if (var.Fields.ContainsKey("補考標準") && var.Fields["補考標準"] is Dictionary<int, decimal>)
                standard = var.Fields["補考標準"] as Dictionary<int, decimal>;
            StudentScore studentScore = new StudentScore(config, standard, currentGradeYear);
            foreach (SemesterSubjectScoreInfo info in var.SemesterSubjectScoreList)
            {
                studentScore.AddSubject(info);
            }
            foreach (SchoolYearSubjectScoreInfo info in var.SchoolYearSubjectScoreList)
            {
                studentScore.AddSubject(info);
            }
            foreach (SemesterEntryScoreInfo info in var.SemesterEntryScoreList)
            {
                if (info.Entry != "德行")
                {
                    studentScore.AddEntry(info);
                }
            }
            foreach (SchoolYearEntryScoreInfo info in var.SchoolYearEntryScoreList)
            {
                if (info.Entry != "德行")
                {
                    studentScore.AddEntry(info);
                }
            }
            studentScore.Rating(var, (int)config.SchoolYear);
            studentScore.CreditStatistic();
            mergeKeyValue.Add("科目分項成績", studentScore);
            eachStudentDoc.MailMerge.MergeField += new MergeFieldEventHandler(this.MailMerge_MergeField);
            eachStudentDoc.MailMerge.RemoveEmptyParagraphs = true;
            List<string> keys = new List<string>();
            List<object> values = new List<object>();
            foreach (string key in mergeKeyValue.Keys)
            {
                keys.Add(key);
                values.Add(mergeKeyValue[key]);
            }
            eachStudentDoc.MailMerge.Execute(keys.ToArray(), values.ToArray());
            absenceRewardInfos.Clear();
            studentScore.Clear();
            mergeKeyValue.Clear();
            keys.Clear();
            values.Clear();
            doc.Sections.Add(doc.ImportNode(eachStudentDoc.Sections[0], true));
            bkw.ReportProgress(((int)((currentStudent++ * 30.0) / ((double)totalStudent))) + 70);
        }
        allStudent.Clear();
        e.Result = doc;
    }

    private void bkw_ProgressChanged(object sender, ProgressChangedEventArgs e)
    {
        this.buttonStudent.SetBarMessage("學年成績單產生中...", e.ProgressPercentage);
    }

    private void bkw_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
    {
        this.buttonStudent.SetBarMessage("學年成績單產生完成");
        Common.SaveToFile("學年成績單", e.Result as Document);
    }

    private void buttonClass_OnClick(object sender, EventArgs e)
    {
        AccessHelper helper = new AccessHelper();
        List<StudentRecord> allStudent = new List<StudentRecord>();
        foreach (ClassRecord each_class in helper.ClassHelper.GetSelectedClass())
        {
            allStudent.AddRange(each_class.Students);
        }
        if ((allStudent.Count <= 500) || (Common.Warning(500) != DialogResult.No))
        {
            this.FormStart(helper, allStudent);
        }
    }

    private void buttonStudent_OnClick(object sender, EventArgs e)
    {
        AccessHelper helper = new AccessHelper();
        List<StudentRecord> allStudent = helper.StudentHelper.GetSelectedStudent();
        if ((allStudent.Count <= 500) || (Common.Warning(500) != DialogResult.No))
        {
            this.FormStart(helper, allStudent);
        }
    }

    private void FillStudentData(AccessHelper helper, List<StudentRecord> students)
    {
        helper.StudentHelper.FillAttendance(students);
        helper.StudentHelper.FillReward(students);
        helper.StudentHelper.FillParentInfo(students);
        helper.StudentHelper.FillContactInfo(students);
        helper.StudentHelper.FillSchoolYearEntryScore(true, students);
        helper.StudentHelper.FillSchoolYearSubjectScore(true, students);
        helper.StudentHelper.FillSemesterEntryScore(true, students);
        helper.StudentHelper.FillSemesterSubjectScore(true, students);
        helper.StudentHelper.FillField("SemesterEntryClassRating", students);
        helper.StudentHelper.FillField("SchoolYearEntryClassRating", students);
        helper.StudentHelper.FillField("補考標準", students);
        helper.StudentHelper.FillSemesterMoralScore(true, students);
        helper.StudentHelper.FillSemesterHistory(students);
    }

    private void FormStart(AccessHelper helper, List<StudentRecord> allStudent)
    {
        Config config = new Config();
        SelectSchoolYearForm form = new SelectSchoolYearForm(config);
        if (form.ShowDialog() == DialogResult.OK)
        {
            BackgroundWorker bkw = new BackgroundWorker
            {
                WorkerReportsProgress = true
            };
            bkw.ProgressChanged += new ProgressChangedEventHandler(this.bkw_ProgressChanged);
            bkw.DoWork += new DoWorkEventHandler(this.bkw_DoWork);
            bkw.RunWorkerCompleted += new RunWorkerCompletedEventHandler(this.bkw_RunWorkerCompleted);
            bkw.RunWorkerAsync(new object[] { bkw, config, helper, allStudent });
        }
    }

    private void MailMerge_MergeField(object sender, MergeFieldEventArgs e)
    {
        DocumentBuilder builder;
        Cell cell;
        double width;
        double height;
        int totalRow;
        double microUnit;
        double rowHeight;
        Table table;
        if (e.FieldName == "缺曠獎懲")
        {
            Dictionary<string, Dictionary<string, int[]>> absenceRewardInfos = (Dictionary<string, Dictionary<string, int[]>>)e.FieldValue;
            builder = new DocumentBuilder(e.Document);
            builder.MoveToField(e.Field, true);
            e.Field.Remove();
            cell = builder.CurrentParagraph.ParentNode as Cell;
            width = cell.CellFormat.Width;
            height = (cell.ParentNode as Row).RowFormat.Height;
            totalRow = 0;
            foreach (string var in absenceRewardInfos.Keys)
            {
                totalRow += absenceRewardInfos[var].Count;
            }
            microUnit = width / 7.0;
            rowHeight = (height - 60.0) / ((double)totalRow);
            table = builder.StartTable();
            builder.CellFormat.ClearFormatting();
            builder.CellFormat.Borders.LineWidth = 0.25;
            builder.RowFormat.HeightRule = HeightRule.Exactly;
            builder.RowFormat.Alignment = RowAlignment.Center;
            builder.CellFormat.VerticalAlignment = CellVerticalAlignment.Center;
            builder.CellFormat.LeftPadding = 3.0;
            builder.CellFormat.RightPadding = 3.0;
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
            builder.ParagraphFormat.LineSpacingRule = LineSpacingRule.AtLeast;
            builder.RowFormat.Height = 11.0;
            builder.InsertCell().CellFormat.Width = microUnit * 7.0;
            builder.Write("缺 曠 獎 懲 統 計");
            builder.EndRow();
            builder.RowFormat.Height = 49.0;
            builder.InsertCell().CellFormat.Width = microUnit * 2.0;
            builder.Write("類\n　\n別");
            builder.InsertCell().CellFormat.Width = microUnit * 2.0;
            builder.Write("項\n　\n目");
            builder.InsertCell().CellFormat.Width = microUnit;
            builder.Write("第\n一\n學\n期");
            builder.InsertCell().CellFormat.Width = microUnit;
            builder.Write("第\n二\n學\n期");
            builder.InsertCell().CellFormat.Width = microUnit;
            builder.Write("學\n　\n年");
            builder.EndRow();
            builder.RowFormat.Height = rowHeight;
            foreach (string cata in absenceRewardInfos.Keys)
            {
                int index = 1;
                int count = absenceRewardInfos[cata].Count;
                foreach (string var in absenceRewardInfos[cata].Keys)
                {
                    builder.InsertCell().CellFormat.Width = microUnit * 2.0;
                    if (index == 1)
                    {
                        builder.Write(cata);
                        builder.CellFormat.VerticalMerge = CellMerge.First;
                    }
                    else
                    {
                        builder.CellFormat.VerticalMerge = CellMerge.Previous;
                    }
                    builder.InsertCell().CellFormat.Width = microUnit * 2.0;
                    builder.Write(var);
                    builder.CellFormat.VerticalMerge = CellMerge.None;
                    int first = absenceRewardInfos[cata][var][0];
                    int second = absenceRewardInfos[cata][var][1];
                    int both = first + second;
                    builder.InsertCell().CellFormat.Width = microUnit;
                    if (first > 0)
                    {
                        builder.Write(first + "");
                    }
                    builder.InsertCell().CellFormat.Width = microUnit;
                    if (second > 0)
                    {
                        builder.Write(second + "");
                    }
                    builder.InsertCell().CellFormat.Width = microUnit;
                    if (both > 0)
                    {
                        builder.Write(both + "");
                    }
                    builder.EndRow();
                    index++;
                }
            }
            foreach (Cell frontCell in table.FirstRow.Cells)
            {
                frontCell.CellFormat.Borders.Top.LineStyle = LineStyle.None;
            }
            foreach (Cell frontCell in table.LastRow.Cells)
            {
                frontCell.CellFormat.Borders.Bottom.LineStyle = LineStyle.None;
            }
            foreach (Row row in table.Rows)
            {
                row.FirstCell.CellFormat.Borders.Left.LineStyle = LineStyle.None;
                row.LastCell.CellFormat.Borders.Right.LineStyle = LineStyle.None;
            }
            absenceRewardInfos.Clear();
        }
        if (e.FieldName == "科目分項成績")
        {
            StudentScore studentScore = (StudentScore)e.FieldValue;
            builder = new DocumentBuilder(e.Document);
            builder.MoveToField(e.Field, true);
            e.Field.Remove();
            cell = builder.CurrentParagraph.ParentNode as Cell;
            width = cell.CellFormat.Width;
            height = (cell.ParentNode as Row).RowFormat.Height + cell.ParentRow.ParentTable.Rows[1].RowFormat.Height;
            totalRow = 0;
            totalRow += studentScore.Subjects.Count;
            totalRow += studentScore.Entries.Count;
            if (studentScore.Subjects.Count <= 0)
            {
                return;
            }
            microUnit = width / 17.0;
            rowHeight = (height - 64.0) / ((double)totalRow);
            table = builder.StartTable();
            builder.CellFormat.ClearFormatting();
            builder.CellFormat.Borders.LineWidth = 0.25;
            builder.CellFormat.FitText = true;
            builder.RowFormat.HeightRule = HeightRule.Exactly;
            builder.RowFormat.Alignment = RowAlignment.Center;
            builder.CellFormat.VerticalAlignment = CellVerticalAlignment.Center;
            builder.CellFormat.LeftPadding = 3.0;
            builder.CellFormat.RightPadding = 3.0;
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
            builder.ParagraphFormat.LineSpacingRule = LineSpacingRule.AtLeast;
            builder.RowFormat.Height = 12.0;
            builder.InsertCell().CellFormat.Width = microUnit * 17.0;
            builder.Write("學  業  成  績");
            builder.EndRow();
            builder.RowFormat.Height = 12.0;
            builder.InsertCell().CellFormat.Width = microUnit * 7.0;
            builder.CellFormat.VerticalMerge = CellMerge.First;
            builder.Write("科        目");
            builder.InsertCell().CellFormat.Width = microUnit * 4.0;
            builder.CellFormat.VerticalMerge = CellMerge.None;
            builder.Write("第 一 學 期");
            builder.InsertCell().CellFormat.Width = microUnit * 4.0;
            builder.Write("第 二 學 期");
            builder.InsertCell().CellFormat.Width = microUnit * 2.0;
            builder.Write("學 年");
            builder.EndRow();
            builder.RowFormat.Height = 40.0;
            builder.InsertCell().CellFormat.Width = microUnit * 7.0;
            builder.CellFormat.VerticalMerge = CellMerge.Previous;
            builder.InsertCell().CellFormat.Width = microUnit * 1.0;
            builder.CellFormat.VerticalMerge = CellMerge.None;
            builder.Write("必\n選\n修");
            builder.InsertCell().CellFormat.Width = microUnit * 1.0;
            builder.Write("學\n　\n分");
            builder.InsertCell().CellFormat.Width = microUnit * 2.0;
            builder.Write("成\n　\n績");
            builder.InsertCell().CellFormat.Width = microUnit * 1.0;
            builder.Write("必\n選\n修");
            builder.InsertCell().CellFormat.Width = microUnit * 1.0;
            builder.Write("學\n　\n分");
            builder.InsertCell().CellFormat.Width = microUnit * 2.0;
            builder.Write("成\n　\n績");
            builder.InsertCell().CellFormat.Width = microUnit * 2.0;
            builder.Write("成\n　\n績");
            builder.EndRow();
            builder.RowFormat.Height = rowHeight;
            List<ScoreData> subjectList = new List<ScoreData>(studentScore.Subjects.Values);
            subjectList.Sort(new Common.SubjectComparer());
            foreach (ScoreData data in subjectList)
            {
                builder.InsertCell().CellFormat.Width = microUnit * 7.0;
                builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
                builder.Write(data.Name + " " + data.Level);
                builder.InsertCell().CellFormat.Width = microUnit;
                builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
                builder.Write(data.FirstRequire);
                builder.InsertCell().CellFormat.Width = microUnit;
                if (data.FirstCredit > 0)
                {
                    builder.Write(data.FirstCredit + "");
                }
                builder.InsertCell().CellFormat.Width = microUnit * 2.0;
                if (data.FirstScore > 0M)
                {
                    builder.Write(data.FirstSign + data.FirstScore + "");
                }
                builder.InsertCell().CellFormat.Width = microUnit;
                builder.Write(data.SecondRequire);
                builder.InsertCell().CellFormat.Width = microUnit;
                if (data.SecondCredit > 0)
                {
                    builder.Write(data.SecondCredit + "");
                }
                builder.InsertCell().CellFormat.Width = microUnit * 2.0;
                if (data.SecondScore > 0M)
                {
                    builder.Write(data.SecondSign + data.SecondScore + "");
                }
                builder.InsertCell().CellFormat.Width = microUnit * 2.0;
                if (data.SchoolYearScore > 0M)
                {
                    builder.Write(data.SchoolYearScore + "");
                }
                builder.EndRow();
            }
            List<ScoreData> entryList = new List<ScoreData>(studentScore.Entries.Values);
            entryList.Sort(new Common.EntryCompaper());
            foreach (ScoreData data in entryList)
            {
                builder.InsertCell().CellFormat.Width = microUnit * 7.0;
                builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
                builder.Write(data.Name);
                if (data.Name == "德行")
                {
                    builder.InsertCell().CellFormat.Width = microUnit * 4.0;
                    builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
                    if (data.FirstScore > 0M)
                    {
                        builder.Write(data.FirstScore + "");
                        builder.Write(Common.ParseLevel(data.FirstScore));
                    }
                    builder.InsertCell().CellFormat.Width = microUnit * 4.0;
                    if (data.SecondScore > 0M)
                    {
                        builder.Write(data.SecondScore + "");
                        builder.Write(Common.ParseLevel(data.SecondScore));
                    }
                    builder.InsertCell().CellFormat.Width = microUnit * 2.0;
                    if (data.SchoolYearScore > 0M)
                    {
                        builder.Write(data.SchoolYearScore + "");
                        builder.Write(Common.ParseLevel(data.SchoolYearScore));
                    }
                }
                else if (((data.Name == "學業成績名次") || (data.Name == "實得學分")) || (data.Name == "累計學分"))
                {
                    builder.InsertCell().CellFormat.Width = microUnit * 4.0;
                    builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
                    builder.Write(data.FirstSemesterItem);
                    builder.InsertCell().CellFormat.Width = microUnit * 4.0;
                    builder.Write(data.SecondSemesterItem);
                    builder.InsertCell().CellFormat.Width = microUnit * 2.0;
                    builder.Write(data.SchoolYearItem);
                }
                else
                {
                    builder.InsertCell().CellFormat.Width = microUnit * 4.0;
                    builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
                    if (data.FirstScore > 0M)
                    {
                        builder.Write(data.FirstScore + "");
                    }
                    builder.InsertCell().CellFormat.Width = microUnit * 4.0;
                    if (data.SecondScore > 0M)
                    {
                        builder.Write(data.SecondScore + "");
                    }
                    builder.InsertCell().CellFormat.Width = microUnit * 2.0;
                    if (data.SchoolYearScore > 0M)
                    {
                        builder.Write(data.SchoolYearScore + "");
                    }
                }
                builder.EndRow();
            }
            foreach (Cell frontCell in table.FirstRow.Cells)
            {
                frontCell.CellFormat.Borders.Top.LineStyle = LineStyle.None;
            }
            foreach (Cell frontCell in table.LastRow.Cells)
            {
                frontCell.CellFormat.Borders.Bottom.LineStyle = LineStyle.None;
            }
            foreach (Row row in table.Rows)
            {
                row.FirstCell.CellFormat.Borders.Left.LineStyle = LineStyle.None;
                row.LastCell.CellFormat.Borders.Right.LineStyle = LineStyle.None;
            }
            studentScore.Clear();
        }
        if (e.FieldName == "綜合表現")
        {
            if (e.FieldValue == null)
            {
                e.Field.Remove();
            }
            else
            {
                Dictionary<int, XmlElement> objectValue = (Dictionary<int, XmlElement>)e.FieldValue;
                if (objectValue != null)
                {
                    builder = new DocumentBuilder(e.Document);
                    builder.MoveToField(e.Field, false);
                    width = (builder.CurrentParagraph.ParentNode as Cell).CellFormat.Width;
                    builder.StartTable();
                    Cell temp = builder.InsertCell();
                    temp.CellFormat.Borders.LineWidth = 0.25;
                    temp.CellFormat.LeftPadding = 5.0;
                    temp.CellFormat.Width = 120.0;
                    temp.ParentRow.RowFormat.Alignment = RowAlignment.Left;
                    temp.CellFormat.VerticalAlignment = CellVerticalAlignment.Center;
                    builder.Write("項目");
                    temp = builder.InsertCell();
                    temp.CellFormat.LeftPadding = 5.0;
                    temp.CellFormat.Width = (width - 120.0) / 2.0;
                    temp.ParentRow.RowFormat.Alignment = RowAlignment.Center;
                    temp.CellFormat.VerticalAlignment = CellVerticalAlignment.Center;
                    builder.Write("第一學期");
                    temp = builder.InsertCell();
                    temp.CellFormat.LeftPadding = 5.0;
                    temp.CellFormat.Width = (width - 120.0) / 2.0;
                    temp.ParentRow.RowFormat.Alignment = RowAlignment.Center;
                    temp.CellFormat.VerticalAlignment = CellVerticalAlignment.Center;
                    builder.Write("第二學期");
                    builder.EndRow();
                    Dictionary<string, Dictionary<int, string>> detail = new Dictionary<string, Dictionary<int, string>>();
                    foreach (int semester in objectValue.Keys)
                    {
                        XmlElement xml = objectValue[semester];
                        if (xml != null)
                        {
                            foreach (XmlElement each in xml.SelectNodes("TextScore/Morality"))
                            {
                                string face = each.GetAttribute("Face");
                                if ((SmartSchool.Customization.Data.SystemInformation.Fields["文字評量對照表"] as XmlElement).SelectSingleNode("Content/Morality[@Face='" + face + "']") != null)
                                {
                                    string comment = each.InnerText;
                                    if (!detail.ContainsKey(face))
                                    {
                                        detail.Add(face, new Dictionary<int, string>());
                                    }
                                    if (!detail[face].ContainsKey(semester))
                                    {
                                        detail[face].Add(semester, comment);
                                    }
                                }
                            }
                        }
                    }
                    foreach (string face in detail.Keys)
                    {
                        temp = builder.InsertCell();
                        temp.CellFormat.LeftPadding = 5.0;
                        temp.CellFormat.Width = 120.0;
                        temp.ParentRow.RowFormat.Alignment = RowAlignment.Left;
                        temp.CellFormat.VerticalAlignment = CellVerticalAlignment.Center;
                        builder.Write(face);
                        temp = builder.InsertCell();
                        temp.CellFormat.LeftPadding = 5.0;
                        temp.CellFormat.Width = (width - 120.0) / 2.0;
                        temp.ParentRow.RowFormat.Alignment = RowAlignment.Center;
                        temp.CellFormat.VerticalAlignment = CellVerticalAlignment.Center;
                        if (detail[face].ContainsKey(1))
                        {
                            builder.Write(detail[face][1]);
                        }
                        temp = builder.InsertCell();
                        temp.CellFormat.LeftPadding = 5.0;
                        temp.CellFormat.Width = (width - 120.0) / 2.0;
                        temp.ParentRow.RowFormat.Alignment = RowAlignment.Center;
                        temp.CellFormat.VerticalAlignment = CellVerticalAlignment.Center;
                        if (detail[face].ContainsKey(2))
                        {
                            builder.Write(detail[face][2]);
                        }
                        builder.EndRow();
                    }
                    table = builder.EndTable();
                    if (table.Rows.Count > 0)
                    {
                        foreach (Cell each in table.FirstRow.Cells)
                        {
                            each.CellFormat.Borders.Top.LineStyle = LineStyle.None;
                        }
                        foreach (Cell each in table.LastRow.Cells)
                        {
                            each.CellFormat.Borders.Bottom.LineStyle = LineStyle.None;
                        }
                        foreach (Row each in table.Rows)
                        {
                            each.FirstCell.CellFormat.Borders.Left.LineStyle = LineStyle.None;
                            each.LastCell.CellFormat.Borders.Right.LineStyle = LineStyle.None;
                        }
                    }
                    e.Field.Remove();
                }
            }
        }
    }

    private List<StudentRecord> ProcessStudentData(AccessHelper helper, List<StudentRecord> students, BackgroundWorker bkw)
    {
        int index = 1;
        List<StudentRecord> list = new List<StudentRecord>();
        List<StudentRecord> filled = new List<StudentRecord>();
        foreach (StudentRecord each_stu in students)
        {
            list.Add(each_stu);
            if (list.Count >= 50)
            {
                this.FillStudentData(helper, list);
                filled.AddRange(list);
                list = new List<StudentRecord>();
            }
            bkw.ReportProgress((int)((index++ * 70.0) / ((double)students.Count)));
        }
        if (list.Count > 0)
        {
            this.FillStudentData(helper, list);
            filled.AddRange(list);
            list = new List<StudentRecord>();
        }
        return filled;
    }
}

