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
    class ResitScoreImport
    {
        private BackgroundWorker _BWResitScoreImport;
        private ResitSubjectMachine machine = new ResitSubjectMachine();

        public ResitScoreImport()
        {
            SelectSemesterForm form = new SelectSemesterForm("�ɦҦ��Z�פJ��");
            if (form.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                _BWResitScoreImport = new BackgroundWorker();
                _BWResitScoreImport.DoWork += new DoWorkEventHandler(_BWResitScoreImport_DoWork);
                _BWResitScoreImport.RunWorkerCompleted += new RunWorkerCompletedEventHandler(_BWResitScoreImport_RunWorkerCompleted);
                _BWResitScoreImport.ProgressChanged += new ProgressChangedEventHandler(_BWResitScoreImport_ProgressChanged);
                _BWResitScoreImport.WorkerReportsProgress = true;
                _BWResitScoreImport.RunWorkerAsync(new object[] { form.SchoolYear, form.Semester });
            }
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

        void _BWResitScoreImport_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            MotherForm.SetStatusBarMessage("�ɦҦ��Z�פJ���ͤ�", e.ProgressPercentage);
        }

        void _BWResitScoreImport_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            MotherForm.SetStatusBarMessage("�ɦҦ��Z�פJ���ͧ���");
            Completed("�ɦҦ��Z�פJ��", (Workbook)e.Result);
        }

        void _BWResitScoreImport_DoWork(object sender, DoWorkEventArgs e)
        {
            object[] objectValue = (object[])e.Argument;
            int schoolyear = (int)objectValue[0];
            int semester = (int)objectValue[1];

            _BWResitScoreImport.ReportProgress(0);

            #region ���o�Ҧ��ǥͥH�θɦҸ�T

            AccessHelper helper = new AccessHelper();
            List<StudentRecord> allStudents = new List<StudentRecord>();
            List<ClassRecord> allClasses = helper.ClassHelper.GetAllClass();
            WearyDogComputer computer = new WearyDogComputer();

            double currentClass = 1;
            double totalClasses = allClasses.Count;

            List<StudentRecord> loadCourseStudents = new List<StudentRecord>();

            foreach (ClassRecord aClass in allClasses)
            {
                List<StudentRecord> classStudents = aClass.Students;
                computer.FillSemesterSubjectScoreInfoWithResit(helper, true, classStudents);
                helper.StudentHelper.FillField("�ή�з�", classStudents);

                allStudents.AddRange(classStudents);
                foreach (StudentRecord aStudent in classStudents)
                {
                    foreach (SemesterSubjectScoreInfo info in aStudent.SemesterSubjectScoreList)
                    {
                        if (info.SchoolYear == schoolyear && info.Semester == semester)
                        {
                            if (!info.Pass && info.Detail.HasAttribute("�F�ɦҼз�") && info.Detail.GetAttribute("�F�ɦҼз�") == "�O")
                            {
                                if (!loadCourseStudents.Contains(aStudent))
                                    loadCourseStudents.Add(aStudent);
                            }
                        }
                    }
                }
                _BWResitScoreImport.ReportProgress((int)(currentClass++ * 90.0 / totalClasses));
            }

            MultiThreadWorker<StudentRecord> multiThreadWorker = new MultiThreadWorker<StudentRecord>();
            multiThreadWorker.MaxThreads = 3;
            multiThreadWorker.PackageSize = 80;
            multiThreadWorker.PackageWorker += new EventHandler<PackageWorkEventArgs<StudentRecord>>(multiThreadWorker_PackageWorker);
            multiThreadWorker.Run(loadCourseStudents, new object[] { schoolyear, semester, helper });

            double currentStudent = 1;
            double totalStudents = allStudents.Count;

            #endregion

            #region ���ͪ��ö�J���

            Workbook template = new Workbook();
            template.Open(new MemoryStream(Properties.Resources.�ɦҦ��Z�פJ), FileFormatType.Excel2003);
            Workbook wb = new Workbook();
            wb.Copy(template);
            Worksheet ws = wb.Worksheets[0];

            Dictionary<string, int> columnIndexTable = new Dictionary<string, int>();

            int headerIndex = 0;
            string headerString = template.Worksheets[0].Cells[0, headerIndex].StringValue;
            while (!string.IsNullOrEmpty(headerString))
            {
                columnIndexTable.Add(headerString, headerIndex);
                headerString = template.Worksheets[0].Cells[0, ++headerIndex].StringValue;
            }

            foreach (StudentRecord aStudent in loadCourseStudents)
            {
                foreach (SemesterSubjectScoreInfo info in aStudent.SemesterSubjectScoreList)
                {
                    if (info.SchoolYear == schoolyear && info.Semester == semester)
                    {
                        if (!info.Pass && info.Detail.HasAttribute("�F�ɦҼз�") && info.Detail.GetAttribute("�F�ɦҼз�") == "�O")
                        {
                            #region �j�M�½ұЮv
                            foreach (StudentAttendCourseRecord attendCourse in aStudent.AttendCourseList)
                            {
                                if (attendCourse.SchoolYear == schoolyear && attendCourse.Semester == semester && attendCourse.Subject == info.Subject && attendCourse.SubjectLevel == info.Level)
                                {
                                    List<TeacherRecord> teachers = helper.TeacherHelper.GetLectureTeacher(attendCourse);
                                    if (teachers.Count > 0)
                                        info.Detail.SetAttribute("�½ұЮv", teachers[0].TeacherName);
                                }
                            }
                            #endregion
                            machine.AddSubject(aStudent, info);
                        }
                    }
                }

                _BWResitScoreImport.ReportProgress(90 + (int)(currentStudent++ * 10.0 / totalStudents));
            }

            machine.Sort();

            int index = 1;
            foreach (Dictionary<string, string> info in machine.GetSubjects())
            {
                foreach (string key in info.Keys)
                {
                    ws.Cells[index, columnIndexTable[key]].PutValue(info[key]);
                }
                index++;
            }

            #endregion

            e.Result = wb;
        }

        void multiThreadWorker_PackageWorker(object sender, PackageWorkEventArgs<StudentRecord> e)
        {
            int schoolyear = (int)((object[])e.Argument)[0];
            int semester = (int)((object[])e.Argument)[1];
            AccessHelper helper = (AccessHelper)((object[])e.Argument)[2];
            helper.StudentHelper.FillAttendCourse(schoolyear, semester, e.List);
        }
    }

    /// <summary>
    /// �ǥ͸�ƻP�ɦҬ��
    /// </summary>
    class ResitSubjectMachine
    {
        List<Dictionary<string, string>> subjectInfo = new List<Dictionary<string, string>>();

        public ResitSubjectMachine()
        {
        }

        private int SortBySemesterSubject(Dictionary<string, string> a, Dictionary<string, string> b)
        {
            if (a["���"] == b["���"])
            {
                if (a["��دŧO"] == b["��دŧO"])
                {
                    return SortBySubjectName(a["�Ǥ���"], b["�Ǥ���"]);
                }
                return SortBySubjectName(a["��دŧO"], b["��دŧO"]);
            }
            return SortBySubjectName(a["���"], b["���"]);
            //return SortBySubjectName(a["���"] + "_" + a["��دŧO"] + "_" + a["�Ǥ���"], b["���"] + "_" + b["��دŧO"] + "_" + b["�Ǥ���"]);
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
                case "¦":
                    x = 12;
                    break;
                case "��":
                    x = 13;
                    break;
                case "��":
                    x = 14;
                    break;
                case "��":
                    x = 15;
                    break;
                case "�q":
                    x = 16;
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
                    x = int.MaxValue;
                    break;
            }
            return x;
        }

        public void AddSubject(StudentRecord student, SemesterSubjectScoreInfo info)
        {
            Dictionary<string, string> aInfo = new Dictionary<string, string>();
            aInfo.Add("�Ǹ�", student.StudentNumber);
            aInfo.Add("�Z��", (student.RefClass != null) ? student.RefClass.ClassName : "");
            aInfo.Add("�y��", student.SeatNo);
            aInfo.Add("��O", student.Department);
            aInfo.Add("�m�W", student.StudentName);
            aInfo.Add("���", info.Subject);
            aInfo.Add("��دŧO", info.Level);
            aInfo.Add("�Ǧ~��", info.SchoolYear.ToString());
            aInfo.Add("�Ǵ�", info.Semester.ToString());
            aInfo.Add("�Ǥ���", info.CreditDec().ToString());
            aInfo.Add("���Z�~��", info.GradeYear.ToString());
            aInfo.Add("�����", info.Require ? "����" : "���");
            aInfo.Add("�ճ��q", info.Detail.HasAttribute("�׽Үճ��q") ? info.Detail.GetAttribute("�׽Үճ��q") : "");
            aInfo.Add("��l���Z", info.Detail.HasAttribute("��l���Z") ? info.Detail.GetAttribute("��l���Z") : "");

            //  aInfo.Add("�ɦҼз�", info.Detail.HasAttribute("�ɦҼз�") ? info.Detail.GetAttribute("�ɦҼз�") : "");

            aInfo.Add("�ɦҼз�", info.Detail.HasAttribute("�׽ҸɦҼз�") ? info.Detail.GetAttribute("�׽ҸɦҼз�") : "");


            aInfo.Add("�ɦҦ��Z", info.Detail.HasAttribute("�ɦҦ��Z") ? info.Detail.GetAttribute("�ɦҦ��Z") : "");

            //       Dictionary<int, decimal> std = student.Fields["�ή�з�"] as Dictionary<int, decimal>;
            //       aInfo.Add("�ή�з�", std[info.GradeYear].ToString());

            aInfo.Add("�ή�з�", info.Detail.HasAttribute("�׽Ҥή�з�") ? info.Detail.GetAttribute("�׽Ҥή�з�") : "");

            aInfo.Add("�½ұЮv", info.Detail.HasAttribute("�½ұЮv") ? info.Detail.GetAttribute("�½ұЮv") : "");
            aInfo.Add("���o�Ǥ�", "");

            subjectInfo.Add(aInfo);
        }

        public List<Dictionary<string, string>> GetSubjects()
        {
            return subjectInfo;
        }

        public void Sort()
        {
            subjectInfo.Sort(SortBySemesterSubject);
        }
    }
}
