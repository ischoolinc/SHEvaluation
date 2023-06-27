using Aspose.Cells;
using FISCA.Presentation;
using SmartSchool.Common;
using SmartSchool.Customization.Data;
using SmartSchool.Customization.Data.StudentExtension;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;

namespace SmartSchool.Evaluation.Reports
{
    class ResitListByStudent
    {
        BackgroundWorker _BGWResitList;

        public ResitListByStudent()
        {
            int schoolyear = 0;
            int semester = 0;
            //int gradeYear = 0;
            //bool printAllYear = false;

            SelectSemesterForm form = new SelectSemesterForm("�ɦҦW��-�̾ǥ�");
            if (form.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                schoolyear = form.SchoolYear;
                semester = form.Semester;
                //gradeYear = form.GradeYear;
                //printAllYear = form.IsPrintAllYear;
            }
            else
                return;

            _BGWResitList = new BackgroundWorker();
            _BGWResitList.WorkerReportsProgress = true;
            _BGWResitList.DoWork += new DoWorkEventHandler(_BGWResitList_DoWork);
            _BGWResitList.ProgressChanged += new ProgressChangedEventHandler(_BGWResitList_ProgressChanged);
            _BGWResitList.RunWorkerCompleted += new RunWorkerCompletedEventHandler(_BGWResitList_RunWorkerCompleted);
            //_BGWResitList.RunWorkerAsync(new object[] { schoolyear, semester, gradeYear, printAllYear });
            _BGWResitList.RunWorkerAsync(new object[] { schoolyear, semester });

        }

        private int SortBySemesterSubjectScore(SemesterSubjectScoreInfo a, SemesterSubjectScoreInfo b)
        {
            return SortBySubjectName(a.Subject, b.Subject);
        }

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
                default:
                    x = 12;
                    break;
            }
            return x;
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

        void _BGWResitList_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            MotherForm.SetStatusBarMessage("�ɦҦW�沣�ͧ���");
            Completed("�ɦҦW��", (Workbook)e.Result);
        }

        void _BGWResitList_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            MotherForm.SetStatusBarMessage("�ɦҦW�沣�ͤ�", e.ProgressPercentage);
            Console.WriteLine(e.ProgressPercentage.ToString());
        }

        void _BGWResitList_DoWork(object sender, DoWorkEventArgs e)
        {
            object[] objectValue = (object[])e.Argument;
            int schoolyear = (int)objectValue[0];
            int semester = (int)objectValue[1];

            //int gradeYear = (int)objectValue[2];
            //bool printAllYear = (bool)objectValue[3];

            _BGWResitList.ReportProgress(0);

            #region ���o�Ҧ��ǥͥH�θɦҸ�T

            AccessHelper helper = new AccessHelper();
            List<StudentRecord> allStudents = new List<StudentRecord>();
            List<ClassRecord> allClasses = helper.ClassHelper.GetAllClass();
            WearyDogComputer computer = new WearyDogComputer();

            double currentClass = 1;
            double totalClasses = allClasses.Count;

            foreach (ClassRecord aClass in allClasses)
            {
                List<StudentRecord> classStudents = aClass.Students;
                computer.FillSemesterSubjectScoreInfoWithResit(helper, true, classStudents);
                allStudents.AddRange(classStudents);

                _BGWResitList.ReportProgress((int)(currentClass++ * 90.0 / totalClasses));
            }

            double currentStudent = 1;
            double totalStudents = allStudents.Count;

            #endregion

            #region ���ͪ��ö�J���

            Workbook template = new Workbook();
            template.Open(new MemoryStream(Properties.Resources.�ɦҦW��_�̾ǥ�), FileFormatType.Excel2003);
            Workbook wb = new Workbook();
            wb.Copy(template);
            Worksheet ws = wb.Worksheets[0];

            Range eachRow = template.Worksheets[0].Cells.CreateRange(2, 1, false);

            //if (printAllYear)
            ws.Cells[0, 0].PutValue(SystemInformation.SchoolChineseName + " " + schoolyear + " �Ǧ~�� �� " + semester + " �Ǵ� �ǥ͸ɦҦW��");
            //else
            //    ws.Cells[0, 0].PutValue(SystemInformation.SchoolChineseName + " " + schoolyear + " �Ǧ~�� �� " + semester + " �Ǵ� " + gradeYear + "�~�žǥ͸ɦҦW��");

            int index = 2;

            foreach (StudentRecord aStudent in allStudents)
            {
                //if (!printAllYear && aStudent.RefClass.GradeYear != gradeYear.ToString())
                //    continue;
                string className = aStudent.RefClass.ClassName;
                string seatNo = aStudent.SeatNo;
                string studentNumber = aStudent.StudentNumber;
                string studentName = aStudent.StudentName;

                aStudent.SemesterSubjectScoreList.Sort(SortBySemesterSubjectScore);

                foreach (SemesterSubjectScoreInfo info in aStudent.SemesterSubjectScoreList)
                {
                    if (info.SchoolYear == schoolyear && info.Semester == semester && !info.Pass)
                    {
                        if (info.Detail.GetAttribute("�F�ɦҼз�") == "�O")
                        {
                            string subject = info.Subject;
                            string levelString = "";
                            int level;
                            if (int.TryParse(info.Level, out level))
                                levelString = GetNumber(level);
                            string credit = info.CreditDec().ToString();
                            string score = info.Detail.GetAttribute("��l���Z");
                            // string limit = info.Detail.GetAttribute("�ɦҼз�");
                            string limit = info.Detail.GetAttribute("�׽ҸɦҼз�");
                            ws.Cells.CreateRange(index, 1, false).Copy(eachRow);
                            ws.Cells[index, 0].PutValue(className);
                            ws.Cells[index, 1].PutValue(seatNo);
                            ws.Cells[index, 2].PutValue(studentNumber);
                            ws.Cells[index, 3].PutValue(studentName);
                            ws.Cells[index, 4].PutValue(subject + (string.IsNullOrEmpty(levelString) ? "" : " " + levelString));
                            ws.Cells[index, 5].PutValue(credit);
                            ws.Cells[index, 6].PutValue(score);
                            ws.Cells[index, 7].PutValue(limit);

                            index++;
                        }
                    }
                }

                _BGWResitList.ReportProgress(90 + (int)(currentStudent++ * 10.0 / totalStudents));
            }

            #endregion

            e.Result = wb;
        }
    }
}
