using Aspose.Cells;
using SmartSchool.Common;
using SmartSchool.Customization.Data;
using SmartSchool.Customization.Data.StudentExtension;
using SmartSchool.Customization.PlugIn;
using SmartSchool.Customization.PlugIn.Report;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Text;
using System.Xml;

namespace SmartSchool.Evaluation.Reports
{
    class SemesterMoralScoreCalc
    {
        private BackgroundWorker _BGWSemesterMoralScoresCalculate;
        ButtonAdapter button;

        public SemesterMoralScoreCalc()
        {
            button = new SecureButtonAdapter("Report0270");
            button.Text = "�w�榨�Z�պ��(�¨�)";
            button.Path = "�ǰȬ�������";
            button.OnClick += new EventHandler(button_OnClick);
            ClassReport.AddReport(button);
        }

        private void Completed(string inputReportName, Workbook inputWorkbook)
        {
            string reportName = inputReportName;

            string path = Path.Combine(System.Windows.Forms.Application.StartupPath, "Reports");
            if (!Directory.Exists(path))
                Directory.CreateDirectory(path);
            path = Path.Combine(path, reportName + ".xls");

            Workbook wb = inputWorkbook;

            if (File.Exists(path))
            {
                int i = 1;
                while (true)
                {
                    string newPath = Path.GetDirectoryName(path) + "\\" + Path.GetFileNameWithoutExtension(path) + (i++) + Path.GetExtension(path);
                    if (!File.Exists(newPath))
                    {
                        path = newPath;
                        break;
                    }
                }
            }

            try
            {
                wb.Save(path, FileFormatType.Excel2003);
                System.Diagnostics.Process.Start(path);
            }
            catch
            {
                System.Windows.Forms.SaveFileDialog sd = new System.Windows.Forms.SaveFileDialog();
                sd.Title = "�t�s�s��";
                sd.FileName = reportName + ".xls";
                sd.Filter = "Excel�ɮ� (*.xls)|*.xls|�Ҧ��ɮ� (*.*)|*.*";
                if (sd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    try
                    {
                        wb.Save(sd.FileName, FileFormatType.Excel2003);

                    }
                    catch
                    {
                        MsgBox.Show("���w���|�L�k�s���C", "�إ��ɮץ���", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                        return;
                    }
                }
            }
        }

        void button_OnClick(object sender, EventArgs e)
        {
            int schoolyear = 0;
            int semester = 0;
            bool over100 = false;
            int sizeIndex = 0;
            Dictionary<string, List<string>> type = null;

            SemesterMoralScoreCalcForm form = new SemesterMoralScoreCalcForm();
            if (form.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                schoolyear = form.SchoolYear;
                semester = form.Semester;
                over100 = form.AllowMoralScoreOver100;
                sizeIndex = form.PaperSize;
                type = form.UserDefinedType;
            }
            else
                return;

            _BGWSemesterMoralScoresCalculate = new BackgroundWorker();
            _BGWSemesterMoralScoresCalculate.WorkerReportsProgress = true;
            _BGWSemesterMoralScoresCalculate.DoWork += new DoWorkEventHandler(_BGWSemesterMoralScoresCalculate_DoWork);
            _BGWSemesterMoralScoresCalculate.ProgressChanged += new ProgressChangedEventHandler(_BGWSemesterMoralScoresCalculate_ProgressChanged);
            _BGWSemesterMoralScoresCalculate.RunWorkerCompleted += new RunWorkerCompletedEventHandler(_BGWSemesterMoralScoresCalculate_RunWorkerCompleted);
            _BGWSemesterMoralScoresCalculate.RunWorkerAsync(new object[] { schoolyear, semester, over100, sizeIndex, type });
        }

        void _BGWSemesterMoralScoresCalculate_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            button.SetBarMessage("�w�榨�Z�պ���ͧ���");
            Completed("�w�榨�Z�պ��", (Workbook)e.Result);
        }

        void _BGWSemesterMoralScoresCalculate_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            button.SetBarMessage("�w�榨�Z�պ���ͤ�...", e.ProgressPercentage);
        }

        void _BGWSemesterMoralScoresCalculate_DoWork(object sender, DoWorkEventArgs e)
        {
            object[] args = (object[])e.Argument;
            int schoolyear = (int)args[0];
            int semester = (int)args[1];
            bool over100 = (bool)args[2];
            int sizeIndex = (int)args[3];
            Dictionary<string, List<string>> userType = (Dictionary<string, List<string>>)args[4];

            _BGWSemesterMoralScoresCalculate.ReportProgress(1);

            #region ���o���

            AccessHelper dataSeed = new AccessHelper();
            List<ClassRecord> allClasses = dataSeed.ClassHelper.GetSelectedClass();
            List<StudentRecord> allStudents = new List<StudentRecord>();
            Dictionary<string, List<StudentRecord>> classStudents = new Dictionary<string, List<StudentRecord>>();
            AngelDemonComputer computer = new AngelDemonComputer();

            int maxStudents = 0;
            int totalStudent = 0;
            int currentStudent = 1;

            foreach (ClassRecord aClass in allClasses)
            {
                List<StudentRecord> studnetList = aClass.Students;
                if (studnetList.Count > maxStudents)
                    maxStudents = studnetList.Count;
                allStudents.AddRange(studnetList);

                //computer.FillDemonScore(dataSeed, schoolyear, semester, studnetList);

                classStudents.Add(aClass.ClassID, studnetList);
                totalStudent += studnetList.Count;
            }

            computer.FillDemonScore(dataSeed, schoolyear, semester, allStudents);
            if (semester == 2)
            {
                dataSeed.StudentHelper.FillSemesterEntryScore(true, allStudents);
                dataSeed.StudentHelper.FillSchoolYearEntryScore(true, allStudents);
                dataSeed.StudentHelper.FillSemesterHistory(allStudents);
            }
            SystemInformation.getField("DiffItem");
            SystemInformation.getField("Degree");

            Dictionary<string, decimal> degreeList = (Dictionary<string, decimal>)SystemInformation.Fields["Degree"];

            #endregion

            #region ���ͪ��
            Workbook template = new Workbook();
            Workbook prototype = new Workbook();

            //�C�L�ؤo
            if (sizeIndex == 0)
                template.Open(new MemoryStream(Properties.Resources.�w�榨�Z�պ��A3), FileFormatType.Excel2003);
            else if (sizeIndex == 1)
                template.Open(new MemoryStream(Properties.Resources.�w�榨�Z�պ��A4), FileFormatType.Excel2003);
            else if (sizeIndex == 2)
                template.Open(new MemoryStream(Properties.Resources.�w�榨�Z�պ��B4), FileFormatType.Excel2003);

            prototype.Copy(template);

            Worksheet templateSheet = template.Worksheets[0];
            Worksheet prototypeSheet = prototype.Worksheets[0];

            Range tempInfoAndReward = templateSheet.Cells.CreateRange(0, 17, true);
            Range tempAbsence = templateSheet.Cells.CreateRange(17, 2, true);
            Range tempBeforeOtherDiff = templateSheet.Cells.CreateRange(19, 3, true);
            Range tempOtherDiff = templateSheet.Cells.CreateRange(22, 1, true);
            Range tempAfterOtherDiff = templateSheet.Cells.CreateRange(25, 2, true);

            Dictionary<string, int> columnIndexTable = new Dictionary<string, int>();

            Dictionary<string, List<string>> periodAbsence = new Dictionary<string, List<string>>();

            //�������g�� Column Index
            columnIndexTable.Add("�j�\", 4);
            columnIndexTable.Add("�p�\", 6);
            columnIndexTable.Add("�ż�", 8);
            columnIndexTable.Add("�j�L", 10);
            columnIndexTable.Add("�p�L", 12);
            columnIndexTable.Add("ĵ�i", 14);
            columnIndexTable.Add("���g�p�p", 16);

            //���m�[���
            int ptColIndex = 17;
            foreach (SmartSchool.Evaluation.AngelDemonComputer.UsefulPeriodAbsence var in computer.UsefulPeriodAbsences)
            {
                if (!periodAbsence.ContainsKey(var.Period))
                    periodAbsence.Add(var.Period, new List<string>());
                if (!periodAbsence[var.Period].Contains(var.Absence))
                    periodAbsence[var.Period].Add(var.Absence);

                prototypeSheet.Cells.CreateRange(ptColIndex, 2, true).Copy(tempAbsence);
                ptColIndex += 2;
            }
            //foreach (string period in userType.Keys)
            //{
            //    if (!periodAbsence.ContainsKey(period))
            //    {
            //        periodAbsence.Add(period, new List<string>());
            //        foreach (string absence in userType[period])
            //        {
            //            if (!periodAbsence[period].Contains(absence))
            //            {
            //                periodAbsence[period].Add(absence);
            //                prototypeSheet.Cells.CreateRange(ptColIndex, 2, true).Copy(tempAbsence);
            //                ptColIndex += 2;
            //            }
            //        }
            //    }
            //}

            ptColIndex = 17;

            foreach (string period in periodAbsence.Keys)
            {
                prototypeSheet.Cells.CreateRange(2, ptColIndex, 1, periodAbsence[period].Count * 2).Merge();
                prototypeSheet.Cells[2, ptColIndex].PutValue(period);
                foreach (string absence in periodAbsence[period])
                {
                    prototypeSheet.Cells[3, ptColIndex].PutValue(absence);

                    columnIndexTable.Add(period + "_" + absence, ptColIndex);

                    ptColIndex += 2;
                }
            }
            prototypeSheet.Cells.CreateRange(1, 17, 1, ptColIndex - 15).Merge();
            prototypeSheet.Cells[1, 17].PutValue("���m�[���");

            columnIndexTable.Add("����", ptColIndex);
            columnIndexTable.Add("���m�p�p", ptColIndex + 1);

            prototypeSheet.Cells.CreateRange(ptColIndex, 3, true).Copy(tempBeforeOtherDiff);
            ptColIndex += 3;

            //�ɮv�[���
            columnIndexTable.Add("�ɮv�[���", ptColIndex - 1);

            //��L�[�������
            foreach (string var in (List<string>)SystemInformation.Fields["DiffItem"])
            {
                prototypeSheet.Cells.CreateRange(ptColIndex, 1, true).Copy(tempOtherDiff);
                prototypeSheet.Cells[1, ptColIndex].PutValue(var);

                columnIndexTable.Add(var, ptColIndex);

                ptColIndex++;
            }


            if (semester == 2)
            {
                prototypeSheet.Cells.CreateRange(ptColIndex, 1, true).Copy(tempOtherDiff);
                prototypeSheet.Cells[1, ptColIndex].PutValue("�W�Ǵ��w�榨�Z");
                columnIndexTable.Add("�W�Ǵ����Z", ptColIndex++);

                prototypeSheet.Cells.CreateRange(ptColIndex, 1, true).Copy(tempOtherDiff);
                prototypeSheet.Cells[1, ptColIndex].PutValue("�Ǵ��w�榨�Z");
                columnIndexTable.Add("�Ǵ����Z", ptColIndex++);

                prototypeSheet.Cells.CreateRange(ptColIndex, 1, true).Copy(tempOtherDiff);
                prototypeSheet.Cells[1, ptColIndex].PutValue("����");
                columnIndexTable.Add("����", ptColIndex++);

                prototypeSheet.Cells.CreateRange(ptColIndex, 1, true).Copy(tempOtherDiff);
                prototypeSheet.Cells[1, ptColIndex].PutValue("�Ǧ~�w�榨�Z");
                columnIndexTable.Add("�Ǧ~���Z", ptColIndex++);
            }
            else
            {

                prototypeSheet.Cells.CreateRange(ptColIndex, 1, true).Copy(tempOtherDiff);
                prototypeSheet.Cells[1, ptColIndex].PutValue("�Ǵ��w�榨�Z");
                columnIndexTable.Add("�Ǵ����Z", ptColIndex++);

                prototypeSheet.Cells.CreateRange(ptColIndex, 1, true).Copy(tempOtherDiff);
                prototypeSheet.Cells[1, ptColIndex].PutValue("����");
                columnIndexTable.Add("����", ptColIndex++);
            }
            prototypeSheet.Cells.CreateRange(ptColIndex, 2, true).Copy(tempAfterOtherDiff);
            columnIndexTable.Add("���y", ptColIndex++);
            columnIndexTable.Add("��ﵥ��", ptColIndex++);

            //�[�W���u
            prototypeSheet.Cells.CreateRange(maxStudents + 5, 0, 1, ptColIndex).SetOutlineBorder(BorderType.TopBorder, CellBorderType.Medium, System.Drawing.Color.Black);

            //��J�s����
            prototypeSheet.Cells[0, 0].PutValue("�s�����G" + DateTime.Today.ToShortDateString());

            //��J���D
            prototypeSheet.Cells.CreateRange(0, 4, 1, ptColIndex - 4).Merge();
            prototypeSheet.Cells[0, 4].PutValue(SystemInformation.SchoolChineseName + " " + schoolyear + " �Ǧ~�� " + ((semester == 1) ? "�W" : "�U") + " �Ǵ� �w�榨�Z�պ��@�@�@�@�@�@�@�@");

            Range ptEachRow = prototypeSheet.Cells.CreateRange(5, 1, false);

            for (int i = 5; i < maxStudents + 5; i++)
            {
                prototypeSheet.Cells.CreateRange(i, 1, false).Copy(ptEachRow);
            }

            Range pt = prototypeSheet.Cells.CreateRange(0, maxStudents + 5, false);

            #endregion

            #region ��J���
            Workbook wb = new Workbook();
            wb.Copy(prototype);
            Worksheet ws = wb.Worksheets[0];

            int index = 0;
            int dataIndex = 0;
            int classTotalRow = maxStudents + 5;

            foreach (ClassRecord aClass in allClasses)
            {
                //�ƻs�����᪺�˪O
                ws.Cells.CreateRange(index, classTotalRow, false).Copy(pt);

                //��J�Z�ŦW��
                ws.Cells[index + 1, 0].PutValue(aClass.ClassName);

                Dictionary<string, int> degreeCount = new Dictionary<string, int>();
                foreach (string key in degreeList.Keys)
                {
                    degreeCount.Add(key, 0);
                }

                dataIndex = index + 5;

                foreach (StudentRecord aStudent in classStudents[aClass.ClassID])
                {
                    ws.Cells[dataIndex, 0].PutValue(aStudent.SeatNo);
                    ws.Cells[dataIndex, 1].PutValue(aStudent.StudentName);
                    ws.Cells[dataIndex, 2].PutValue(aStudent.StudentNumber);

                    decimal score = 0;
                    decimal rewardScore = 0;
                    decimal absenceScore = 0;
                    decimal diffScore = 0;

                    XmlElement demonScore = (XmlElement)aStudent.Fields["DemonScore"];

                    score = decimal.Parse(demonScore.GetAttribute("Score"));

                    foreach (XmlElement var in demonScore.SelectNodes("SubScore"))
                    {
                        if (var.GetAttribute("Type") == "���")
                        {
                            ws.Cells[dataIndex, 3].PutValue(var.GetAttribute("Score"));
                        }
                        else if (var.GetAttribute("Type") == "���g")
                        {
                            int colIndex = columnIndexTable[var.GetAttribute("Name")];
                            if (decimal.Parse(var.GetAttribute("Count")) != 0)
                                ws.Cells[dataIndex, colIndex].PutValue(var.GetAttribute("Count"));
                            if (decimal.Parse(var.GetAttribute("Score")) != 0)
                                ws.Cells[dataIndex, colIndex + 1].PutValue(var.GetAttribute("Score"));
                            rewardScore += decimal.Parse(var.GetAttribute("Score"));
                        }
                        else if (var.GetAttribute("Type") == "���m")
                        {
                            string pa = var.GetAttribute("PeriodType") + "_" + var.GetAttribute("Absence");
                            if (columnIndexTable.ContainsKey(pa))
                            {
                                int colIndex = columnIndexTable[pa];
                                if (decimal.Parse(var.GetAttribute("Count")) != 0)
                                    ws.Cells[dataIndex, colIndex].PutValue(var.GetAttribute("Count"));
                                if (decimal.Parse(var.GetAttribute("Score")) != 0)
                                    ws.Cells[dataIndex, colIndex + 1].PutValue(var.GetAttribute("Score"));
                            }
                            absenceScore += decimal.Parse(
                                var.GetAttribute("Score"));
                        }
                        else if (var.GetAttribute("Type") == "�[���")
                        {
                            int colIndex = columnIndexTable[var.GetAttribute("DiffItem")];
                            if (decimal.Parse(var.GetAttribute("Score")) != 0)
                                ws.Cells[dataIndex, colIndex].PutValue(var.GetAttribute("Score"));
                            diffScore += decimal.Parse(var.GetAttribute("Score"));
                        }
                        else if (var.GetAttribute("Type") == "����")
                        {
                            int colIndex = columnIndexTable["����"];
                            if (decimal.Parse(var.GetAttribute("Score")) != 0)
                                ws.Cells[dataIndex, colIndex].PutValue(var.GetAttribute("Score"));
                            absenceScore += decimal.Parse(var.GetAttribute("Score"));
                        }
                    }

                    //��J���g�p�p
                    if (rewardScore != 0)
                        ws.Cells[dataIndex, columnIndexTable["���g�p�p"]].PutValue(rewardScore.ToString());
                    //��J���m�p�p
                    if (absenceScore != 0)
                        ws.Cells[dataIndex, columnIndexTable["���m�p�p"]].PutValue(absenceScore.ToString());

                    //��J�Ƿ~���Z�պ�
                    if (!over100 && score > 100)
                        score = 100;
                    ws.Cells[dataIndex, columnIndexTable["�Ǵ����Z"]].PutValue(score.ToString());

                    //��J�W�Ǵ�&�Ǧ~���Z
                    if (semester == 2)
                    {
                        int gradeYear = -1;
                        foreach (SemesterHistory sh in aStudent.SemesterHistoryList)
                        {
                            if (sh.SchoolYear == schoolyear && sh.Semester == semester)
                            {
                                gradeYear = sh.GradeYear;
                                break;
                            }
                        }
                        //�S���Ǵ����{�N�η�U���Ǵ�
                        if (gradeYear == -1 && schoolyear == SystemInformation.SchoolYear)
                        {
                            int.TryParse(aStudent.RefClass.GradeYear, out gradeYear);
                        }
                        //��J�W�Ǵ����Z
                        foreach (SemesterEntryScoreInfo semesterEntryScore in aStudent.SemesterEntryScoreList)
                        {
                            if (semesterEntryScore.Entry == "�w��" && semesterEntryScore.GradeYear == gradeYear && semesterEntryScore.Semester == 1)
                            {
                                if (!over100 && semesterEntryScore.Score > 100)
                                    ws.Cells[dataIndex, columnIndexTable["�W�Ǵ����Z"]].PutValue("100");
                                else
                                    ws.Cells[dataIndex, columnIndexTable["�W�Ǵ����Z"]].PutValue(semesterEntryScore.Score.ToString());
                                break;
                            }
                        }

                        //��J�Ǧ~���Z
                        foreach (SchoolYearEntryScoreInfo schoolyearEntryScore in aStudent.SchoolYearEntryScoreList)
                        {
                            if (schoolyearEntryScore.Entry == "�w��" && schoolyearEntryScore.GradeYear == gradeYear)
                            {
                                if (!over100 && schoolyearEntryScore.Score > 100)
                                    ws.Cells[dataIndex, columnIndexTable["�Ǧ~���Z"]].PutValue("100");
                                else
                                    ws.Cells[dataIndex, columnIndexTable["�Ǧ~���Z"]].PutValue(schoolyearEntryScore.Score.ToString());
                                break;
                            }
                        }
                    }

                    //��J����
                    string degree = computer.ParseLevel(score);
                    ws.Cells[dataIndex, columnIndexTable["����"]].PutValue(degree);

                    //�p�ⵥ�ĥX�{����
                    if (degreeCount.ContainsKey(degree))
                        degreeCount[degree]++;

                    //���y
                    if (demonScore.SelectSingleNode("Others/@Comment") != null)
                        ws.Cells[dataIndex, columnIndexTable["���y"]].PutValue(demonScore.SelectSingleNode("Others/@Comment").InnerText);

                    dataIndex++;

                    //�^���i��
                    _BGWSemesterMoralScoresCalculate.ReportProgress((int)(currentStudent++ * 100.0 / totalStudent));
                }

                ws.Cells.CreateRange(index + classTotalRow + 1, 0, 1, ptColIndex).Merge();
                StringBuilder degreeSumString = new StringBuilder("");
                degreeSumString.Append("�w�浥�Ĳέp    ");
                foreach (string key in degreeCount.Keys)
                {
                    degreeSumString.Append(key + "���G " + degreeCount[key].ToString() + "    ");
                }
                ws.Cells[index + classTotalRow + 1, 0].Style.Font.Size = 12;
                ws.Cells.CreateRange(index + classTotalRow + 1, 1, false).RowHeight = 20;
                ws.Cells[index + classTotalRow + 1, 0].PutValue(degreeSumString.ToString());
                ws.Cells[index + classTotalRow + 1, 0].Style.HorizontalAlignment = TextAlignmentType.Left;

                index += classTotalRow + 3;
                ws.HPageBreaks.Add(index, ptColIndex);
            }

            #endregion

            e.Result = wb;
        }
    }
}