using System;
using System.Collections.Generic;
using System.Text;
using SmartSchool.Customization.PlugIn;
using System.ComponentModel;
using Aspose.Cells;
using System.IO;
using SmartSchool.Customization.Data;
using SmartSchool.Customization.Data.StudentExtension;
using SmartSchool.Customization.PlugIn.Report;
using System.Drawing;
using System.Linq;
using SmartSchool.Common;
using System.Globalization;

namespace SmartSchool.Evaluation.Reports
{
    class ClassSemesterScore
    {
        private ButtonAdapter classButton;
        private BackgroundWorker _BWClassSemesterScore;

        public ClassSemesterScore()
        {
            classButton = new SecureButtonAdapter("Report0160");
            classButton.Text = "�Z�žǥ;Ǵ����Z�@����";
            classButton.Path = "���Z��������";
            classButton.OnClick += new EventHandler(classButton_OnClick);
            ClassReport.AddReport(classButton);
        }

        void classButton_OnClick(object sender, EventArgs e)
        {
            int schoolyear = 0;
            int semester = 0;
            bool over100 = false;
            int papersize = 0;
            bool UseSourceScore = false;

            ClassSemesterScoreForm form = new ClassSemesterScoreForm();
            if (form.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                schoolyear = form.SchoolYear;
                semester = form.Semester;
                over100 = form.AllowMoralScoreOver100;
                papersize = form.PaperSize;
                UseSourceScore = form.UseSourceScore;
            }
            else
                return;

            _BWClassSemesterScore = new BackgroundWorker();
            _BWClassSemesterScore.DoWork += new DoWorkEventHandler(_BWClassSemesterScore_DoWork);
            _BWClassSemesterScore.RunWorkerCompleted += new RunWorkerCompletedEventHandler(_BWClassSemesterScore_RunWorkerCompleted);
            _BWClassSemesterScore.ProgressChanged += new ProgressChangedEventHandler(_BWClassSemesterScore_ProgressChanged);
            _BWClassSemesterScore.WorkerReportsProgress = true;
            _BWClassSemesterScore.RunWorkerAsync(new object[] { schoolyear, semester, over100, papersize, UseSourceScore });
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

        private string GetNumber(int p)
        {
            string levelNumber;
            switch (p)
            {
                #region ����levelNumber
                case 0:
                    levelNumber = "";
                    break;
                case 1:
                    levelNumber = "��";
                    break;
                case 2:
                    levelNumber = "��";
                    break;
                case 3:
                    levelNumber = "��";
                    break;
                case 4:
                    levelNumber = "��";
                    break;
                case 5:
                    levelNumber = "��";
                    break;
                case 6:
                    levelNumber = "��";
                    break;
                case 7:
                    levelNumber = "��";
                    break;
                case 8:
                    levelNumber = "��";
                    break;
                case 9:
                    levelNumber = "��";
                    break;
                case 10:
                    levelNumber = "��";
                    break;
                default:
                    levelNumber = "" + (p);
                    break;
                    #endregion
            }
            return levelNumber;
        }

        //private int SortByTheSameSubjectName(string a, string b)
        //{
        //    string[] asplit = a.Split('_');
        //    string[] bsplit = b.Split('_');
        //    if (asplit[0] == bsplit[0])
        //    {
        //        int c,d;
        //        if (int.TryParse(asplit[1], out c) && int.TryParse(bsplit[1], out d))
        //        {
        //            if (c < d)
        //                return c;
        //            else
        //                return d;
        //        }
        //    }
        //    return SortBySubjectName(a, b);
        //}

        private int SortBySubjectName(string a, string b)
        {
            string a1 = a.Length > 0 ? a.Substring(0, 1) : "";
            string b1 = b.Length > 0 ? b.Substring(0, 1) : "";
            #region �Ĥ@�Ӧr�@�˪��ɭ�
            if (a1 == b1)
            {
                if (a.Length > 1 && b.Length > 1)
                    return SortBySubjectName(a.Substring(1), b.Substring(1));
                else
                    return a.Length.CompareTo(b.Length);
            }
            #endregion
            #region �Ĥ@�Ӧr���P�A���O���o�b�]�w���Ǥ����Ʀr�A�p�G�����b�]�w���Ǥ��N�γ�¦r����
            int ai = getIntForSubject(a1), bi = getIntForSubject(b1);
            if (ai > 0 || bi > 0)
                return ai.CompareTo(bi);
            else
                return a1.CompareTo(b1);
            #endregion
        }

        private int getIntForSubject(string a1)
        {
            int x = 0;
            switch (a1)
            {
                case "��":
                    x = 1;
                    break;
                case "�^":
                    x = 2;
                    break;
                case "��":
                    x = 3;
                    break;
                case "��":
                    x = 4;
                    break;
                case "��":
                    x = 5;
                    break;
                case "��":
                    x = 6;
                    break;
                case "��":
                    x = 7;
                    break;
                case "��":
                    x = 8;
                    break;
                case "�a":
                    x = 9;
                    break;
                case "��":
                    x = 10;
                    break;
                case "��":
                    x = 11;
                    break;
                case "¦":
                    x = 12;
                    break;
                case "��":
                    x = 51; break;
                case "��":
                    x = 52; break;
                case "��":
                    x = 53; break;
                case "��":
                    x = 54; break;
                case "��":
                    x = 55; break;
                case "��":
                    x = 56; break;
                case "��":
                    x = 57; break;
                case "��":
                    x = 58; break;
                case "��":
                    x = 59; break;
                case "��":
                    x = 60; break;
                case "1":
                    x = 61; break;
                case "2":
                    x = 62; break;
                case "3":
                    x = 63; break;
                case "4":
                    x = 64; break;
                case "5":
                    x = 65; break;
                case "6":
                    x = 66; break;
                case "7":
                    x = 67; break;
                case "8":
                    x = 68; break;
                case "9":
                    x = 69; break;
                default:
                    x = 100;
                    break;
            }
            return x;
        }

        private int SortByEntryName(string a, string b)
        {
            int ai = getIntForEntry(a), bi = getIntForEntry(b);
            if (ai > 0 || bi > 0)
                return ai.CompareTo(bi);
            else
                return a.CompareTo(b);
        }

        private int getIntForEntry(string a1)
        {
            int x = 0;
            switch (a1)
            {
                case "�Ƿ~����(��l)":
                    x = 1;
                    break;
                case "�Ƿ~����(���u)":
                    x = 1;
                    break;
                case "�Z�űƦW":
                    x = 2;
                    break;
                case "��߬��":
                    x = 3;
                    break;
                case "��|":
                    x = 4;
                    break;
                case "�꨾�q��":
                    x = 5;
                    break;
                case "���d�P�@�z":
                    x = 6;
                    break;
                case "�w��":
                    x = 7;
                    break;
                default:
                    x = 8;
                    break;
            }
            return x;
        }

        void _BWClassSemesterScore_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            classButton.SetBarMessage("�Z�žǥ;Ǵ����Z�@����", e.ProgressPercentage);
        }

        void _BWClassSemesterScore_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            classButton.SetBarMessage("�Z�žǥ;Ǵ����Z�@�����ͧ���");
            Completed("�Z�žǥ;Ǵ����Z�@����", (Workbook)e.Result);
        }

        void _BWClassSemesterScore_DoWork(object sender, DoWorkEventArgs e)
        {
            object[] objectValue = (object[])e.Argument;
            int schoolyear = (int)objectValue[0];
            int semester = (int)objectValue[1];
            bool over100 = (bool)objectValue[2];
            int papersize = (int)objectValue[3];
            bool UseSourceScore = (bool)objectValue[4];

            _BWClassSemesterScore.ReportProgress(0);

            Workbook template = new Workbook();
            if (papersize == 0)
                template.Open(new MemoryStream(Properties.Resources.�Z�žǥ;Ǵ����Z�@����B4));
            else if (papersize == 1)
                template.Open(new MemoryStream(Properties.Resources.�Z�žǥ;Ǵ����Z�@����A3));

            Worksheet tempws = template.Worksheets[template.Worksheets.Add()];

            // ���Y 0~5 �����ƻs
            Range tempHeader = template.Worksheets[0].Cells.CreateRange(0, 5, false);

            // �Ƿ~
            Range tempEachEntry = template.Worksheets[0].Cells.CreateRange(1, 33, 4, 2);
            // �Ǥ���
            Range tempCreditHeader = template.Worksheets[0].Cells.CreateRange(1, 35, 4, 4);

            Workbook wb = new Workbook();
            wb.Copy(template);
            Worksheet ws = wb.Worksheets[0];

            ScoreAverageMachine machine = new ScoreAverageMachine();

            CacluteCourseGetCreditRate courseGetCreditCount = new CacluteCourseGetCreditRate();

            AccessHelper helper = new AccessHelper();

            List<StudentRecord> allClassStudent = new List<StudentRecord>();
            foreach (ClassRecord aClass in helper.ClassHelper.GetSelectedClass())
            {
                allClassStudent.AddRange(aClass.Students);
            }

            double currentStudent = 1;
            double totalStudent = allClassStudent.Count;

            int rowIndex = 0;
            int headerColIndex = 0;

            foreach (ClassRecord aClass in helper.ClassHelper.GetSelectedClass())
            {
                List<StudentRecord> allStudent = aClass.Students;
                int headerIndex = rowIndex;
                int count = 0;
                if (rowIndex > 0)
                {
                    ws.Cells.CreateRange(headerIndex - 2, 0, 1, headerColIndex + 4).SetOutlineBorder(BorderType.BottomBorder, CellBorderType.Medium, Color.Black);
                    ws.Cells.CreateRange(headerIndex - 1, 0, 1, headerColIndex + 4).SetOutlineBorder(BorderType.BottomBorder, CellBorderType.Medium, Color.Black);
                }

                helper.StudentHelper.FillSchoolYearSubjectScore(true, allStudent);
                helper.StudentHelper.FillSemesterSubjectScore(true, allStudent);
                helper.StudentHelper.FillSemesterEntryScore(true, allStudent);
                helper.StudentHelper.FillField("SemesterEntryClassRating", allStudent); //�Ǵ������Z�ƦW�C

                List<string> subjectHeader = new List<string>();
                Dictionary<string, string> subjectCreditHeader = new Dictionary<string, string>();
                Dictionary<string, string> subjectReqHeader = new Dictionary<string, string>();
                List<string> entryHeader = new List<string>();
                Dictionary<string, int> columnIndexTable = new Dictionary<string, int>();

                foreach (StudentRecord student in allStudent)
                {
                    foreach (SemesterSubjectScoreInfo info in student.SemesterSubjectScoreList)
                    {
                        // ���p�Ǥ��B�����������L  || info.Detail.GetAttribute("���ݵ���") == "�O"
                        if (info.Detail.GetAttribute("���p�Ǥ�") == "�O")
                            continue;

                        if (info.SchoolYear == schoolyear && info.Semester == semester)
                        {
                            int level;
                            string levelString = "";
                            if (int.TryParse(info.Level, out level))
                                levelString = GetNumber(level);
                            //
                            //string header = info.Subject + levelString + "_" + (info.Require ? "�� " : "�� ") + "_" + info.Credit;
                            //string creditHeader = (info.Require ? "�� " : "�� ") + info.Credit;

                            string header = info.Subject + levelString + "_" + info.CreditDec();
                            string creditHeader = info.CreditDec().ToString();
                            string reqHeader = info.Require ? "�� " : "�� ";

                            if (!subjectHeader.Contains(header))
                            {
                                subjectHeader.Add(header);
                                subjectCreditHeader.Add(header, creditHeader);
                                subjectReqHeader.Add(header, reqHeader);
                            }
                        }
                    }

                    foreach (SemesterEntryScoreInfo info in student.SemesterEntryScoreList)
                    {
                        if (info.SchoolYear == schoolyear && info.Semester == semester)
                        {
                            string header = info.Entry;    // ���L�o��l

                            // �ϥέ�l���Z
                            if (UseSourceScore && header == "�Ƿ~(��l)")
                                header = "�Ƿ~����(��l)";

                            if (UseSourceScore && header == "�Ƿ~")
                                continue;

                            if (UseSourceScore == false && header == "�Ƿ~(��l)")
                                continue;

                            if (UseSourceScore == false && header == "�Ƿ~")
                                header = "�Ƿ~����(���u)";

                            if (!entryHeader.Contains(header))
                                entryHeader.Add(header);
                        }
                    }
                }

                entryHeader.Add("�Z�űƦW");
                entryHeader.Sort(SortByEntryName);
                subjectHeader.Sort(SortBySubjectName);

                string sString = "";
                if (UseSourceScore)
                    sString = "(��l)";
                else
                    sString = "(���u)";

                ws.Cells.CreateRange(rowIndex, 4, false).Copy(tempHeader);
                ws.Cells[rowIndex, 0].PutValue(SystemInformation.SchoolChineseName + " " + schoolyear + " �Ǧ~�� �� " + semester + " �Ǵ� �Z�žǥͦ��Z�@����  �Z�šG " + aClass.ClassName + " " + sString);

                headerColIndex = 3;
                foreach (string subject in subjectHeader)
                {
                    columnIndexTable.Add(subject, headerColIndex);
                    machine.AddItem(subject);
                    courseGetCreditCount.AddItem(subject);
                    string sl = subject.Split('_')[0];
                    ws.Cells[rowIndex + 1, headerColIndex].PutValue(sl);
                    ws.Cells[rowIndex + 2, headerColIndex].PutValue(subjectReqHeader[subject]);
                    ws.Cells[rowIndex + 3, headerColIndex].PutValue(subjectCreditHeader[subject]);
                    headerColIndex++;
                }
                headerColIndex = 33;

                foreach (string entry in entryHeader)
                {
                    columnIndexTable.Add(entry, headerColIndex);
                    machine.AddItem(entry);
                    ws.Cells.CreateRange(rowIndex + 1, headerColIndex, 4, 2).Copy(tempEachEntry);
                    ws.Cells[rowIndex + 1, headerColIndex].PutValue(entry);
                    headerColIndex++;
                }

                ws.Cells.CreateRange(rowIndex + 1, headerColIndex, 4, 4).Copy(tempCreditHeader);
                columnIndexTable.Add("���o�Ǥ�", headerColIndex);
                columnIndexTable.Add("��o�Ǥ�", headerColIndex + 1);
                columnIndexTable.Add("���o�Ǥ��֭p", headerColIndex + 2);
                columnIndexTable.Add("��o�Ǥ��֭p", headerColIndex + 3);

                tempws.Cells.CreateRange(0, 1, false).Copy(ws.Cells.CreateRange(rowIndex + 4, 1, false));
                Range eachStudent = tempws.Cells.CreateRange(0, 1, false);
                rowIndex += 4;

                int defSS = schoolyear * 10 + semester;

                foreach (StudentRecord student in allStudent)
                {
                    count++;

                    ws.Cells.CreateRange(rowIndex, 1, false).Copy(eachStudent);
                    ws.Cells[rowIndex, 0].PutValue(student.StudentNumber);
                    ws.Cells[rowIndex, 1].PutValue(student.SeatNo);
                    ws.Cells[rowIndex, 2].PutValue(student.StudentName);

                    decimal shouldGetCredit = 0;
                    decimal gotCredit = 0;
                    decimal shouldGetTotalCredit = 0;
                    decimal gotTotalCredit = 0;

                    foreach (SemesterSubjectScoreInfo info in student.SemesterSubjectScoreList)
                    {
                        // ���p�Ǥ��B�����������L  || info.Detail.GetAttribute("���ݵ���") == "�O"
                        if (info.Detail.GetAttribute("���p�Ǥ�") == "�O")
                            continue;

                        if (info.SchoolYear == schoolyear && info.Semester == semester)
                        {
                            shouldGetCredit += info.CreditDec();
                            if (info.Pass)
                                gotCredit += info.CreditDec();

                            int level;
                            string levelString = "";
                            if (int.TryParse(info.Level, out level))
                                levelString = GetNumber(level);

                            //���e
                            //string key = info.Subject + levelString + "_" + (info.Require ? "�� " : "�� ") + "_" + info.Credit;
                            string key = info.Subject + levelString + "_" + info.CreditDec();

                            if (columnIndexTable.ContainsKey(key))
                            {
                                // �P�_�ϥέ�l/���u
                                decimal iScore = 0;
                                if (UseSourceScore)
                                {
                                    decimal.TryParse(info.Detail.GetAttribute("��l���Z"), out iScore);
                                }
                                else
                                    iScore = info.Score;      // ���u                          

                                //ws.Cells[rowIndex, columnIndexTable[key]].PutValue((info.Pass ? "" : "*") + info.Score);
                                //machine.AddScore(key, info.Score);

                                ws.Cells[rowIndex, columnIndexTable[key]].PutValue((info.Pass ? "" : "*") + iScore);
                                machine.AddScore(key, iScore);

                                courseGetCreditCount.AddCourseCount(key);
                                if (info.Pass)
                                {
                                    courseGetCreditCount.AddGetCreditCount(key);
                                }
                            }
                        }

                        // �֭p���o�B�֭p��o�A�վ�P�_�Ǧ~�׾Ǵ�
                        int iss = info.SchoolYear * 10 + info.Semester;
                        if (iss <= defSS)
                        {
                            shouldGetTotalCredit += info.CreditDec();
                            if (info.Pass)
                                gotTotalCredit += info.CreditDec();
                        }
                    }

                    foreach (SemesterEntryScoreInfo info in student.SemesterEntryScoreList)
                    {
                        if (info.SchoolYear == schoolyear && info.Semester == semester)
                        {

                            if (UseSourceScore && info.Entry == "�Ƿ~(��l)")
                            {
                                decimal score = info.Score;
                                if (!over100 && score > 100)
                                    score = 100;
                                ws.Cells[rowIndex, columnIndexTable["�Ƿ~����(��l)"]].PutValue(score);
                                machine.AddScore("�Ƿ~����(��l)", score);

                            }
                            else if (UseSourceScore == false && info.Entry == "�Ƿ~")
                            {
                                decimal score = info.Score;
                                if (!over100 && score > 100)
                                    score = 100;
                                ws.Cells[rowIndex, columnIndexTable["�Ƿ~����(���u)"]].PutValue(score);
                                machine.AddScore("�Ƿ~����(���u)", score);
                            }
                            else
                            {

                                if (columnIndexTable.ContainsKey(info.Entry))
                                {
                                    decimal score = info.Score;
                                    if (!over100 && score > 100)
                                        score = 100;
                                    ws.Cells[rowIndex, columnIndexTable[info.Entry]].PutValue(score);
                                    machine.AddScore(info.Entry, score);
                                }
                            }


                        }
                    }

                    ws.Cells[rowIndex, columnIndexTable["���o�Ǥ�"]].PutValue(shouldGetCredit.ToString());
                    ws.Cells[rowIndex, columnIndexTable["��o�Ǥ�"]].PutValue(gotCredit.ToString());
                    ws.Cells[rowIndex, columnIndexTable["���o�Ǥ��֭p"]].PutValue(shouldGetTotalCredit.ToString());
                    ws.Cells[rowIndex, columnIndexTable["��o�Ǥ��֭p"]].PutValue(gotTotalCredit.ToString());

                    SemesterEntryRating rating = new SemesterEntryRating(student);
                    ws.Cells[rowIndex, columnIndexTable["�Z�űƦW"]].PutValue(rating.GetPlace(schoolyear, semester));

                    if (count % 5 == 0)
                        ws.Cells.CreateRange(rowIndex, 0, 1, headerColIndex + 4).SetOutlineBorder(BorderType.BottomBorder, CellBorderType.Medium, Color.Black);
                    rowIndex++;

                    _BWClassSemesterScore.ReportProgress((int)(currentStudent++ * 100.0 / totalStudent));
                }

                ws.Cells.CreateRange(headerIndex, 0, 4, headerColIndex + 4).SetOutlineBorder(BorderType.BottomBorder, CellBorderType.Medium, Color.Black);

                ws.Cells.CreateRange(rowIndex, 1, false).Copy(eachStudent);
                //for (int i = 0; i < headerColIndex + 4; i++)
                //{
                //    ws.Cells[rowIndex, i].PutValue("");
                //}
                ws.Cells[rowIndex, 0].PutValue("����");
                foreach (string name in machine.GetAllItemName())
                {
                    ws.Cells[rowIndex, columnIndexTable[name]].PutValue(machine.GetAverage(name).ToString());
                }
                rowIndex++;

                ws.Cells.CreateRange(rowIndex, 1, false).Copy(eachStudent);
                ws.Cells[rowIndex, 0].PutValue("���o�Ǥ���v");
                foreach (string name in courseGetCreditCount.GetAllItemName())
                {
                    ws.Cells[rowIndex, columnIndexTable[name]].PutValue(courseGetCreditCount.GetGetCreditRate(name));
                }
                rowIndex++;

                ws.Cells.CreateRange(rowIndex, 1, false).Copy(eachStudent);
                ws.Cells[rowIndex, 0].PutValue("�����o�Ǥ���v");
                foreach (string name in courseGetCreditCount.GetAllItemName())
                {
                    ws.Cells[rowIndex, columnIndexTable[name]].PutValue(courseGetCreditCount.GetNotGetCreditRate(name));
                }

                machine.Clear();
                rowIndex++;

                ws.HPageBreaks.Add(rowIndex, headerColIndex + 4);
            }
            ws.Cells.CreateRange(rowIndex - 2, 0, 1, headerColIndex + 4).SetOutlineBorder(BorderType.BottomBorder, CellBorderType.Medium, Color.Black);
            ws.Cells.CreateRange(rowIndex - 1, 0, 1, headerColIndex + 4).SetOutlineBorder(BorderType.BottomBorder, CellBorderType.Medium, Color.Black);

            e.Result = wb;
        }
    }

    /// <summary>
    /// �p��U���إ����A�x�Υ����p�����
    /// </summary>
    class ScoreAverageMachine
    {
        Dictionary<string, List<decimal>> _itemScoreList = new Dictionary<string, List<decimal>>();

        public ScoreAverageMachine() { }

        public void AddItem(string item)
        {
            if (!_itemScoreList.ContainsKey(item))
                _itemScoreList.Add(item, new List<decimal>());
        }

        public void AddScore(string item, decimal score)
        {
            if (!_itemScoreList.ContainsKey(item))
                return;
            _itemScoreList[item].Add(score);
        }

        public List<string> GetAllItemName()
        {
            List<string> items = new List<string>();
            foreach (string key in _itemScoreList.Keys)
                items.Add(key);
            return items;
        }

        public object GetAverage(string item)
        {
            if (!_itemScoreList.ContainsKey(item))
                return 0;

            decimal summary = 0;

            foreach (decimal score in _itemScoreList[item])
                summary += score;

            if (_itemScoreList[item].Count > 0)
                return Math.Round((summary / (decimal)_itemScoreList[item].Count), 2);
            else
                return "";
        }

        public void Clear()
        {
            _itemScoreList.Clear();
        }
    }

    class CacluteCourseGetCreditRate
    {
        Dictionary<string, int> _ItemGetCreditDic = new Dictionary<string, int>();
        Dictionary<string, int> _ItemCourseCount = new Dictionary<string, int>();

        public void AddItem(string item)
        {
            if (!_ItemGetCreditDic.ContainsKey(item))
                _ItemGetCreditDic.Add(item, 0);
            if (!_ItemCourseCount.ContainsKey(item))
                _ItemCourseCount.Add(item, 0);
        }

        public void AddGetCreditCount(string item)
        {
            if (!_ItemGetCreditDic.ContainsKey(item))
                return;
            _ItemGetCreditDic[item]++;
        }

        public void AddCourseCount(string item)
        {
            if (!_ItemCourseCount.ContainsKey(item))
                return;
            _ItemCourseCount[item]++;
        }

        public List<string> GetAllItemName()
        {
            List<string> items = new List<string>();
            foreach (string key in _ItemCourseCount.Keys)
                items.Add(key);
            return items;
        }

        public object GetGetCreditRate(string item)
        {
            if (!_ItemCourseCount.ContainsKey(item) || !_ItemGetCreditDic.ContainsKey(item))
                return 0;

            double courseCount = _ItemCourseCount[item]* 1.0;
            double getCreditCount = _ItemGetCreditDic[item] * 1.0;

            if (courseCount > 0)
                return (Math.Round((getCreditCount / courseCount), 4)).ToString("P", CultureInfo.InvariantCulture);
            else
                return "";
        }

        public object GetNotGetCreditRate(string item)
        {
            if (!_ItemCourseCount.ContainsKey(item) || !_ItemGetCreditDic.ContainsKey(item))
                return 0;

            double courseCount = _ItemCourseCount[item] * 1.0;
            double notGetCreditCount = (courseCount - _ItemGetCreditDic[item]) * 1.0;

            if (courseCount > 0)
                return (Math.Round((notGetCreditCount / courseCount), 4)).ToString("P", CultureInfo.InvariantCulture);
            else
                return "";
        }
    }
}
