using Aspose.Words;
using SmartSchool.Common;
using SmartSchool.Customization.Data;
using SmartSchool.Customization.PlugIn;
using SmartSchool.Customization.PlugIn.Report;
using SmartSchool.Evaluation.Reports.MultiSemesterScore.DataModel;
using SmartSchool.Evaluation.Reports.MultiSemesterScore.Forms;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace SmartSchool.Evaluation.Reports.MultiSemesterScore
{
    class MultiSemesterScore
    {
        private ButtonAdapter _class_button;

        private BackgroundWorker _bgworker;

        //private SemsIndex _sems_index;
        //public SemsIndex SemsIndex
        //{
        //    get { return _sems_index; }
        //}
        private bool IsPrint = true;

        private ReportOptions _option;
        public ReportOptions Option
        {
            get { return _option; }
            set { _option = value; }
        }

        public MultiSemesterScore()
        {
            string report_name = "�h�Ǵ����Z��";
            string report_path = "���Z��������";

            //_student_button = new ButtonAdapter();
            //_student_button.Text = report_name;
            //_student_button.Path = report_path;
            //_student_button.OnClick += new EventHandler(studentButton_OnClick);
            //StudentReport.AddReport(_student_button);

            _class_button = new SecureButtonAdapter("Report0170");
            _class_button.Text = report_name;
            _class_button.Path = report_path;
            _class_button.OnClick += new EventHandler(classButton_OnClick);
            ClassReport.AddReport(_class_button);

            ButtonAdapter _student_button = new SecureButtonAdapter("Report0060");
            _student_button.Text = report_name;
            _student_button.Path = report_path;
            _student_button.OnClick += new EventHandler(_student_button_OnClick);
            StudentReport.AddReport(_student_button);
        }

        void _student_button_OnClick(object sender, EventArgs e)
        {
            //MsgBox.Show("�p���X���ťաA���q�Z��>���Z��������>�h�Ǵ����Z��A���X�Z�Ŧh�Ǵ����Z��A�A���;ǥͪ��h�Ǵ����Z��C���¡C");
            AccessHelper helper = new AccessHelper();
            BackgroundWorker bkw = new BackgroundWorker();
            bkw.DoWork += new DoWorkEventHandler(bkw_DoWork);
            bkw.RunWorkerCompleted += new RunWorkerCompletedEventHandler(_bgworker_RunWorkerCompleted);
            bkw.ProgressChanged += new ProgressChangedEventHandler(_bgworker_ProgressChanged);
            bkw.WorkerReportsProgress = true;
            bkw.RunWorkerAsync(helper.StudentHelper.GetSelectedStudent());
        }

        void bkw_DoWork(object sender, DoWorkEventArgs e)
        {
            BackgroundWorker bkw = (BackgroundWorker)sender;
            List<StudentRecord> students = (List<StudentRecord>)e.Argument;
            bkw.ReportProgress(1);
            MultiThreadWorker<StudentRecord> mWorker = new MultiThreadWorker<StudentRecord>();
            mWorker.MaxThreads = 3;
            mWorker.PackageSize = 150;
            mWorker.PackageWorker += new EventHandler<PackageWorkEventArgs<StudentRecord>>(mWorker_PackageWorker);
            List<PackageWorkEventArgs<StudentRecord>> resp = mWorker.Run(students);

            // �S�����Z�檺�ǥ� (�ϥΪ̿�����ǥͲM��-�����^���Z�檺�ǥ�)
            List<string> noResultStudentList = new List<string>();
            // �����^���Z�檺�ǥ�ID
            List<string> resultStudentList = new List<string>();

            foreach (StudentRecord studentRecord in students)
            {
                noResultStudentList.Add(studentRecord.StudentID);
            }

            List<SmartSchool.Feature.Student.StudentSnapShop> list = new List<SmartSchool.Feature.Student.StudentSnapShop>();
            foreach (PackageWorkEventArgs<StudentRecord> r in resp)
            {
                List<SmartSchool.Feature.Student.StudentSnapShop> snap = (List<SmartSchool.Feature.Student.StudentSnapShop>)r.Result;

                if (snap == null)
                    Console.WriteLine("StudentSnapshop List �O Null");
                else
                    list.AddRange((List<SmartSchool.Feature.Student.StudentSnapShop>)r.Result);

                foreach (var result in snap)
                {
                    resultStudentList.Add(result.RefStudentID);
                }
            }

            foreach (string resultStudentID in resultStudentList)
            {
                noResultStudentList.Remove(resultStudentID);
            }

            bkw.ReportProgress(80);

            StringBuilder s = new StringBuilder();
            foreach (StudentRecord studentRecord in students)
            {
                foreach (string noResultID in noResultStudentList)
                {
                    if (noResultID == studentRecord.StudentID)
                    {
                        s.AppendLine((studentRecord.RefClass == null ? "" : studentRecord.RefClass.ClassName) + "�@" + studentRecord.SeatNo + "�@" + studentRecord.StudentName + "�@" + studentRecord.StudentNumber);
                    }
                }
            }

            if (noResultStudentList.Count > 0)
            {
                MsgBox.Show("�U�C�ǥͨS���i�C�L����ơA�Х��q �Z��>���Z��������>�h�Ǵ����Z��A���X��T��A�~����k���ǥͪ��h�Ǵ����Z��C\r\n" + s.ToString());
            }

            if (students.Count == noResultStudentList.Count)
                IsPrint = false;

            Document doc = new Document();
            doc.Sections.Clear();
            foreach (SmartSchool.Feature.Student.StudentSnapShop var in list)
            {
                Document d = new Document(new MemoryStream(Convert.FromBase64String(var.PresentContent)));

                doc.Sections.Add(doc.ImportNode(d.Sections[0], true));
            }
            bkw.ReportProgress(100);
            e.Result = doc;
        }

        void mWorker_PackageWorker(object sender, PackageWorkEventArgs<StudentRecord> e)
        {
            List<string> list = new List<string>();
            foreach (StudentRecord var in e.List)
            {
                list.Add(var.StudentID);
            }
            Dictionary<string, SmartSchool.Feature.Student.StudentSnapShop> sv = new Dictionary<string, SmartSchool.Feature.Student.StudentSnapShop>();
            foreach (SmartSchool.Feature.Student.StudentSnapShop s in SmartSchool.Feature.Student.StudentSnapShop.GetSnapShop(list.ToArray()))
            {
                if (!sv.ContainsKey(s.RefStudentID))
                    sv.Add(s.RefStudentID, s);
                else
                {
                    if (sv[s.RefStudentID].Version < s.Version)
                        sv[s.RefStudentID] = s;
                }
            }
            List<SmartSchool.Feature.Student.StudentSnapShop> slist = new List<SmartSchool.Feature.Student.StudentSnapShop>();
            slist.AddRange(sv.Values);
            e.Result = slist;
        }


        void classButton_OnClick(object sender, EventArgs e)
        {
            AccessHelper helper = new AccessHelper();
            List<StudentRecord> selected_students = new List<StudentRecord>();
            FormStart(helper, helper.ClassHelper.GetSelectedClass());
        }

        void FormStart(AccessHelper helper, List<ClassRecord> selected_classes)
        {
            Option = new ReportOptions();

            MultiSemesterScoreForm form = new MultiSemesterScoreForm(Option);
            if (form.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                Option.Save();
                if (MsgBox.Show("�N�|���s���ͦ��Z��A�ǥͪ��Z�ƦW���i��|���ܡC\n\n�p�ݦC�L�w���͹L�����Z��A�Шϥξǥͪ�\"�h�Ǵ����Z��\"�\��C\n\n���ͳ���æC�L?", "���s���ͨ��л\�N���h�Ǵ����Z��", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    _bgworker = new BackgroundWorker();
                    _bgworker.DoWork += new DoWorkEventHandler(_bgworker_DoWork);
                    _bgworker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(_bgworker_RunWorkerCompleted);
                    _bgworker.ProgressChanged += new ProgressChangedEventHandler(_bgworker_ProgressChanged);
                    _bgworker.WorkerReportsProgress = true;
                    _bgworker.RunWorkerAsync(new object[] { helper, selected_classes, Option });
                }
            }
            else
            {
                Option.Save();
                return;
            }

        }

        void _bgworker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            SmartSchool.Customization.PlugIn.Global.SetStatusBarMessage("�h�Ǵ����Z�沣�ͤ�...", e.ProgressPercentage);
        }

        void _bgworker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            SmartSchool.Customization.PlugIn.Global.SetStatusBarMessage("�h�Ǵ����Z�沣�ͧ���");

            if (!IsPrint)
            {
                SmartSchool.Customization.PlugIn.Global.SetStatusBarMessage("�S���i�C�L����ơC");
                return;
            }

            Document doc = (Document)e.Result;

            string reportName = "�h�Ǵ����Z��";
            string path = Path.Combine(Application.StartupPath, "Reports");
            if (!Directory.Exists(path))
                Directory.CreateDirectory(path);
            path = Path.Combine(path, reportName + ".doc");

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
                doc.Save(path, SaveFormat.Doc);
                System.Diagnostics.Process.Start(path);
            }
            catch
            {
                SaveFileDialog sd = new SaveFileDialog();
                sd.Title = "�t�s�s��";
                sd.FileName = reportName + ".doc";
                sd.Filter = "Word�ɮ� (*.doc)|*.doc|�Ҧ��ɮ� (*.*)|*.*";
                if (sd.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        doc.Save(sd.FileName, SaveFormat.AsposePdf);
                    }
                    catch
                    {
                        MsgBox.Show("���w���|�L�k�s���C", "�إ��ɮץ���", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }
            }
        }

        void _bgworker_DoWork(object sender, DoWorkEventArgs e)
        {
            object[] obj = (object[])e.Argument;
            AccessHelper helper = (AccessHelper)obj[0];
            List<ClassRecord> selected_classes = (List<ClassRecord>)obj[1];
            ReportOptions option = (ReportOptions)obj[2];


            _bgworker.ReportProgress(0);

            StudentCollection student_collection = new StudentCollection();
            //List<StudentRecord> failed = new List<StudentRecord>();
            double student_counter = 1;
            double student_total = 0;
            List<StudentRecord> selected_students = new List<StudentRecord>();
            foreach (ClassRecord classRec in selected_classes)
            {
                student_total += classRec.Students.Count;
                selected_students.AddRange(classRec.Students);
            }
            #region ��ƨ��o�P��z

            int package_size = 50;
            List<StudentRecord> packaged_students = new List<StudentRecord>();
            foreach (StudentRecord record in selected_students)
            {
                packaged_students.Add(record);
                if (packaged_students.Count >= package_size)
                {
                    helper.StudentHelper.FillSemesterSubjectScore(true, packaged_students);
                    helper.StudentHelper.FillSemesterEntryScore(true, packaged_students);
                    packaged_students.Clear();
                }
                _bgworker.ReportProgress((int)(student_counter++ * 40.0 / student_total));
            }
            if (packaged_students.Count > 0)
            {
                helper.StudentHelper.FillSemesterSubjectScore(true, packaged_students);
                helper.StudentHelper.FillSemesterEntryScore(true, packaged_students);
                packaged_students.Clear();
            }

            foreach (ClassRecord classRec in selected_classes)
            {
                List<Student> list = new List<Student>();
                foreach (StudentRecord each_stu in classRec.Students)
                {
                    Student student = new Student(each_stu, option.ScoreType, option.PrintSemester, option.PrintEntries);
                    list.Add(student);
                    if (!student_collection.ContainsKey(each_stu.StudentID))
                        student_collection.Add(each_stu.StudentID, student);
                }
                #region ��دZ�ƦW
                List<string> rankedSubject = new List<string>();
                foreach (Student stu in list)
                {
                    foreach (string subject in stu.SubjectCollection.Keys)
                    {
                        if (rankedSubject.Contains(subject))
                            continue;
                        List<decimal> scores = new List<decimal>();
                        Dictionary<decimal, List<SubjectInfo>> rank = new Dictionary<decimal, List<SubjectInfo>>();
                        foreach (Student student in list)
                        {
                            if (student.SubjectCollection.ContainsKey(subject))
                            {
                                decimal score = student.SubjectCollection[subject].GetAverange();
                                scores.Add(score);
                                if (!rank.ContainsKey(score))
                                    rank.Add(score, new List<SubjectInfo>());
                                rank[score].Add(student.SubjectCollection[subject]);
                            }
                        }
                        scores.Sort();
                        scores.Reverse();

                        decimal lastscore = decimal.MinValue;
                        int count = scores.Count;
                        foreach (decimal score in scores)
                        {
                            if (lastscore != score)
                            {
                                int place = scores.IndexOf(score) + 1;
                                lastscore = score;
                                foreach (SubjectInfo s in rank[lastscore])
                                {
                                    s.Place = place;
                                    s.Radix = count;
                                }
                            }
                        }

                        rankedSubject.Add(subject);
                    }
                }
                #endregion
                #region �����Z�ƦW
                List<string> rankedEntrys = new List<string>();
                foreach (Student stu in list)
                {
                    foreach (string entry in stu.EntryCollection.Keys)
                    {
                        if (rankedEntrys.Contains(entry))
                            continue;
                        List<decimal> scores = new List<decimal>();
                        Dictionary<decimal, List<EntryInfo>> rank = new Dictionary<decimal, List<EntryInfo>>();
                        foreach (Student student in list)
                        {
                            if (student.EntryCollection.ContainsKey(entry))
                            {
                                decimal score = student.EntryCollection[entry].GetAverange();
                                scores.Add(score);
                                if (!rank.ContainsKey(score))
                                    rank.Add(score, new List<EntryInfo>());
                                rank[score].Add(student.EntryCollection[entry]);
                            }
                        }
                        scores.Sort();
                        scores.Reverse();

                        decimal lastscore = decimal.MinValue;
                        int count = scores.Count;
                        //��ƦW��^�C�@�Ӧ��Z��Ƥ�
                        foreach (decimal score in scores)
                        {
                            if (lastscore != score)
                            {
                                int place = scores.IndexOf(score) + 1;
                                lastscore = score;
                                foreach (EntryInfo s in rank[lastscore])
                                {
                                    s.Place = place;
                                    s.Radix = count;
                                    //�B�z�w�椣���\�}��
                                    if (s.Name == "�w��" && option.FixMoralScore)
                                    {
                                        s.FixToLimit(100);
                                    }
                                }
                            }
                        }

                        rankedEntrys.Add(entry);
                    }
                }
                #endregion
            }
            #endregion
            student_counter = 1;
            #region ���ͳ����SnapShop
            MemoryStream template = GetTemplate(option);
            Document doc = new Document();
            doc.Sections.Clear();
            List<SmartSchool.Feature.Student.StudentSnapShop> snapShops = new List<SmartSchool.Feature.Student.StudentSnapShop>();
            MultiThreadWorker<SmartSchool.Feature.Student.StudentSnapShop> mWorker = new MultiThreadWorker<SmartSchool.Feature.Student.StudentSnapShop>();
            mWorker.MaxThreads = 3;
            mWorker.PackageSize = 50;
            mWorker.PackageWorker += new EventHandler<PackageWorkEventArgs<SmartSchool.Feature.Student.StudentSnapShop>>(mWorker_PackageWorker);

            foreach (string stuid in student_collection.Keys)
            {
                Student each_stu = student_collection[stuid];
                #region �@�ӾǥͲ��ͤ@�ӷs��Document
                //�@�ӾǥͲ��ͤ@�ӷs��Document
                Document each_page = new Document(template, "", LoadFormat.Doc, "");
                #region �إߦ��ǥͪ����Z��
                //�X�ְ򥻸��
                List<string> merge_keys = new List<string>();
                List<object> merge_values = new List<object>();
                merge_keys.AddRange(new string[] { "�ǮզW��", "��O�W��", "�Z�ŦW��", "�y��", "�Ǹ�", "�m�W", "�ƦW", "�U�Ǵ����Z" });
                merge_values.AddRange(new object[] {
                    SmartSchool.Customization.Data.SystemInformation.SchoolChineseName,
                    each_stu.Department ,
                    each_stu.ClassName,
                    each_stu.SeatNo,
                    each_stu.StudentNumber,
                    each_stu.StudentName,
                    option.RatingMethod.ToString(),
                    each_stu
                });
                each_page.MailMerge.MergeField += new Aspose.Words.Reporting.MergeFieldEventHandler(MailMerge_MergeField);
                each_page.MailMerge.Execute(merge_keys.ToArray(), merge_values.ToArray());
                #endregion
                #endregion
                #region �[�J�s��SnapShop
                SmartSchool.Feature.Student.StudentSnapShop newSnapShop = new SmartSchool.Feature.Student.StudentSnapShop();
                newSnapShop.SnapshopName = "�h�Ǵ����Z��";
                newSnapShop.RefStudentID = stuid;
                newSnapShop.PresentType = ".Doc";
                MemoryStream stream = new MemoryStream();
                each_page.Save(stream, SaveFormat.Doc);
                newSnapShop.PresentContent = Convert.ToBase64String(stream.ToArray());
                stream.Close();
                snapShops.Add(newSnapShop);
                #endregion
                //�C300�H���N�����W�Ǥ@��snapShops
                #region �C300�H���N�����W�Ǥ@��
                if (snapShops.Count == 300)
                {
                    mWorker.Run(snapShops);
                    snapShops.Clear();
                }
                #endregion
                //�X�֦��Z��ܤ@���ɮ�
                doc.Sections.Add(doc.ImportNode(each_page.Sections[0], true));
                //�^���i��
                _bgworker.ReportProgress((int)(student_counter++ * 60.0 / student_total) + 40);
            }
            //�N�̫�@�q��SnapShop�W��
            if (snapShops.Count > 0)
            {
                mWorker.Run(snapShops);
                snapShops.Clear();
            }
            #endregion

            e.Result = doc;
        }

        void mWorker_PackageWorker(object sender, PackageWorkEventArgs<SmartSchool.Feature.Student.StudentSnapShop> e)
        {
            SmartSchool.Feature.Student.StudentSnapShop.AddSnapShop(e.List.ToArray());
        }

        void MailMerge_MergeField(object sender, Aspose.Words.Reporting.MergeFieldEventArgs e)
        {

            if (e.FieldName == "�U�Ǵ����Z")
            {
                e.Text = string.Empty;

                Student stu = (Student)e.FieldValue;
                DocumentBuilder builder = new DocumentBuilder(e.Document);
                //�����Z���Ǵ�
                List<int> semesters = new List<int>();
                #region �`�p���ǥͦ��Ǵ����Z���Ǵ�
                foreach (SubjectInfo subject in stu.SubjectCollection.Values)
                {
                    foreach (int sem in subject.SemsScores.Keys)
                    {
                        if (!semesters.Contains(sem))
                            semesters.Add(sem);
                    }
                }
                foreach (EntryInfo entry in stu.EntryCollection.Values)
                {
                    foreach (int sem in entry.SemsScores.Keys)
                    {
                        if (!semesters.Contains(sem))
                            semesters.Add(sem);
                    }
                }
                #endregion
                semesters.Sort();

                builder.MoveToField(e.Field, false);
                #region ���o�~�ؼe�רíp����e
                Cell SCell = (Cell)builder.CurrentParagraph.ParentNode;
                double Swidth = SCell.CellFormat.Width;
                double microUnit = Swidth / (semesters.Count + 6); //�C�Ǵ����@���A�`�������@���A�ʤ��񵹤@���A�Z�űƦW�ά�ئU�����
                #endregion
                Table table = builder.StartTable();

                builder.CellFormat.ClearFormatting();
                builder.CellFormat.Borders.LineWidth = 0.5;

                builder.RowFormat.HeightRule = HeightRule.Auto;
                builder.RowFormat.Height = builder.Font.Size * 1.2d;
                builder.RowFormat.Alignment = RowAlignment.Center;
                builder.CellFormat.VerticalAlignment = CellVerticalAlignment.Center;
                builder.CellFormat.LeftPadding = 3.0;
                builder.CellFormat.RightPadding = 3.0;
                builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
                builder.ParagraphFormat.LineSpacingRule = LineSpacingRule.Exactly;
                builder.ParagraphFormat.LineSpacing = 10;
                //builder.Font.Size = 8;

                List<EntryInfo> entries = new List<EntryInfo>();
                entries.AddRange(stu.EntryCollection.Values);
                entries.Sort(new EntryComparer());
                if (entries.Count > 0)
                {
                    #region ����Y
                    builder.InsertCell().CellFormat.Width = microUnit * 2;
                    builder.CellFormat.Borders.Right.LineWidth = 0.25;
                    builder.Write("�����`���Z");
                    foreach (int sem in semesters)
                    {
                        #region �C�Ǵ����@��
                        builder.InsertCell().CellFormat.Width = microUnit;
                        builder.CellFormat.Borders.Right.LineWidth = 0.25;
                        switch (sem)
                        {
                            case 1:
                                builder.Write("�@�W");
                                break;
                            case 2:
                                builder.Write("�@�U");
                                break;
                            case 3:
                                builder.Write("�G�W");
                                break;
                            case 4:
                                builder.Write("�G�U");
                                break;
                            case 5:
                                builder.Write("�T�W");
                                break;
                            case 6:
                                builder.Write("�T�U");
                                break;
                            case 7:
                                builder.Write("�|�W");
                                break;
                            case 8:
                                builder.Write("�|�U");
                                break;
                            default:
                                builder.Write("��" + sem + "�Ǵ�");
                                break;
                        }
                        #endregion
                    }
                    builder.InsertCell().CellFormat.Width = microUnit * 1.5;
                    builder.CellFormat.Borders.Right.LineWidth = 0.25;
                    builder.Write("�`����");
                    builder.InsertCell().CellFormat.Width = microUnit * 2;
                    builder.CellFormat.Borders.Right.LineWidth = 0.25;
                    builder.Write("�Z�űƦW");
                    builder.InsertCell().CellFormat.Width = microUnit;
                    builder.Write("�ʤ���");
                    builder.EndRow();
                    #endregion
                    //�������ئ��Z���������u
                    foreach (Cell cell in table.LastRow.Cells)
                        cell.CellFormat.Shading.BackgroundPatternColor = Color.LightGray;
                }
                //�������ئ��Z���������u
                //foreach ( Cell cell in table.LastRow.Cells )
                //    cell.CellFormat.Borders.Bottom.LineStyle = LineStyle.Double;
                #region ��������Z
                foreach (EntryInfo entryInfo in entries)
                {
                    //��ئW��
                    builder.InsertCell().CellFormat.Width = microUnit * 2;
                    builder.CellFormat.Borders.Right.LineWidth = 0.25;
                    builder.Write(entryInfo.Name);
                    #region �C�Ǵ����Z
                    foreach (int sem in semesters)
                    {
                        builder.InsertCell().CellFormat.Width = microUnit;
                        builder.CellFormat.Borders.Right.LineWidth = 0.25;
                        if (entryInfo.SemsScores.ContainsKey(sem))
                        {
                            builder.Write("" + entryInfo.SemsScores[sem]);
                        }
                    }
                    #endregion
                    //�`����
                    builder.InsertCell().CellFormat.Width = microUnit * 1.5;
                    builder.CellFormat.Borders.Right.LineWidth = 0.25;
                    if (entryInfo.Name == "�w��")
                    {
                        decimal moral_score = Math.Round(entryInfo.GetAverange());
                        builder.Write("" + moral_score + " (" + GetMoralLevel(moral_score) + "�Ť�)");
                    }
                    else
                        builder.Write("" + entryInfo.GetAverange());
                    //�Z�űƦW
                    builder.InsertCell().CellFormat.Width = microUnit * 2;
                    builder.CellFormat.Borders.Right.LineWidth = 0.25;
                    builder.Write("" + entryInfo.Place + " / " + entryInfo.Radix);
                    //�ʤ���
                    builder.InsertCell().CellFormat.Width = microUnit;
                    builder.Write("" + entryInfo.GetPercentage() + "%");
                    builder.EndRow();
                }
                #endregion

                List<SubjectInfo> subjects = new List<SubjectInfo>();
                subjects.AddRange(stu.SubjectCollection.Values);
                subjects.Sort(new SubjectComparer());
                if (subjects.Count > 0)
                {
                    #region ����Y
                    builder.InsertCell().CellFormat.Width = microUnit * 2;
                    builder.CellFormat.Borders.Right.LineWidth = 0.25;
                    builder.Write("��ئ��Z");
                    foreach (int sem in semesters)
                    {
                        #region �C�Ǵ����@��
                        builder.InsertCell().CellFormat.Width = microUnit;
                        builder.CellFormat.Borders.Right.LineWidth = 0.25;
                        switch (sem)
                        {
                            case 1:
                                builder.Write("�@�W");
                                break;
                            case 2:
                                builder.Write("�@�U");
                                break;
                            case 3:
                                builder.Write("�G�W");
                                break;
                            case 4:
                                builder.Write("�G�U");
                                break;
                            case 5:
                                builder.Write("�T�W");
                                break;
                            case 6:
                                builder.Write("�T�U");
                                break;
                            case 7:
                                builder.Write("�|�W");
                                break;
                            case 8:
                                builder.Write("�|�U");
                                break;
                            default:
                                builder.Write("��" + sem + "�Ǵ�");
                                break;
                        }
                        #endregion
                    }
                    builder.InsertCell().CellFormat.Width = microUnit * 1.5;
                    builder.CellFormat.Borders.Right.LineWidth = 0.25;
                    builder.Write("�`����");
                    builder.InsertCell().CellFormat.Width = microUnit * 2;
                    builder.CellFormat.Borders.Right.LineWidth = 0.25;
                    builder.Write("�Z�űƦW");
                    builder.InsertCell().CellFormat.Width = microUnit;
                    builder.Write("�ʤ���");
                    builder.EndRow();
                    #endregion
                    //�������ئ��Z���������u
                    foreach (Cell cell in table.LastRow.Cells)
                        cell.CellFormat.Shading.BackgroundPatternColor = Color.LightGray;
                    //�������ئ��Z���������u
                    //foreach ( Cell cell in table.LastRow.Cells )
                    //    cell.CellFormat.Borders.Bottom.LineStyle = LineStyle.Double;
                }
                #region ���ئ��Z

                foreach (SubjectInfo subjectInfo in subjects)
                {
                    //��ئW��
                    builder.InsertCell().CellFormat.Width = microUnit * 2;
                    builder.CellFormat.Borders.Right.LineWidth = 0.25;
                    builder.Write(subjectInfo.Name);
                    #region �C�Ǵ����Z
                    foreach (int sem in semesters)
                    {
                        builder.InsertCell().CellFormat.Width = microUnit;
                        builder.CellFormat.Borders.Right.LineWidth = 0.25;
                        if (subjectInfo.SemsScores.ContainsKey(sem))
                        {
                            builder.Write("" + subjectInfo.SemsScores[sem]);
                        }
                    }
                    #endregion
                    //�`����
                    builder.InsertCell().CellFormat.Width = microUnit * 1.5;
                    builder.CellFormat.Borders.Right.LineWidth = 0.25;
                    builder.Write("" + subjectInfo.GetAverange());
                    //�Z�űƦW
                    builder.InsertCell().CellFormat.Width = microUnit * 2;
                    builder.CellFormat.Borders.Right.LineWidth = 0.25;
                    builder.Write("" + subjectInfo.Place + " / " + subjectInfo.Radix);
                    //�ʤ���
                    builder.InsertCell().CellFormat.Width = microUnit;
                    builder.Write("" + subjectInfo.GetPercentage() + "%");
                    builder.EndRow();
                }
                #endregion
                #region �h�����|�䪺�u
                if (table.FirstRow != null)
                {
                    foreach (Cell cell in table.FirstRow.Cells)
                        cell.CellFormat.Borders.Top.LineStyle = LineStyle.None;
                }

                if (table.LastRow != null)
                {
                    foreach (Cell cell in table.LastRow.Cells)
                        cell.CellFormat.Borders.Bottom.LineStyle = LineStyle.None;
                }

                foreach (Row row in table.Rows)
                {
                    row.FirstCell.CellFormat.Borders.Left.LineStyle = LineStyle.None;
                    row.LastCell.CellFormat.Borders.Right.LineStyle = LineStyle.None;
                }
                #endregion
            }
            //int row_index = 1;

            ////�ƧǬ��
            //List<SubjectInfo> subject_list = new List<SubjectInfo>();
            //subject_list.AddRange(each_stu.SubjectCollection.Values);
            //subject_list.Sort(new SubjectComparer());

            ////�ƧǤ���
            //List<EntryInfo> entry_list = new List<EntryInfo>();
            //entry_list.AddRange(each_stu.EntryCollection.Values);
            //entry_list.Sort(new EntryComparer());

            ////��J���
            //foreach ( SubjectInfo info in subject_list )
            //{
            //    run.Text = info.Name;

            //    subject_table.Rows[row_index].Cells[0].Paragraphs[0].Runs.Add(run.Clone(true));

            //    foreach ( int sems_index in info.SemsScores.Keys )
            //    {
            //        run.Text = info.SemsScores[sems_index].ToString();
            //        subject_table.Rows[row_index].Cells[sems_index].Paragraphs[0].Runs.Add(run.Clone(true));
            //    }
            //    //����
            //    run.Text = info.GetAverange().ToString();
            //    subject_table.Rows[row_index].Cells[9].Paragraphs[0].Runs.Add(run.Clone(true));
            //    //�ƦW
            //    run.Text = "" + info.Place + " / " + info.Radix;
            //    subject_table.Rows[row_index].Cells[10].Paragraphs[0].Runs.Add(run.Clone(true));
            //    //�ʤ���
            //    run.Text = "" + info.GetPercentage() + "%";
            //    subject_table.Rows[row_index].Cells[11].Paragraphs[0].Runs.Add(run.Clone(true));

            //    row_index++;
            //}

            //row_index = 0;

            ////��J����
            //foreach ( EntryInfo info in entry_list )
            //{
            //    run.Text = info.Name;

            //    entry_table.Rows[row_index].Cells[0].Paragraphs[0].Runs.Add(run.Clone(true));

            //    foreach ( int sems_index in info.SemsScores.Keys )
            //    {
            //        run.Text = info.SemsScores[sems_index].ToString();
            //        entry_table.Rows[row_index].Cells[sems_index].Paragraphs[0].Runs.Add(run.Clone(true));
            //    }
            //    //����
            //    run.Text = info.GetAverange().ToString();
            //    entry_table.Rows[row_index].Cells[9].Paragraphs[0].Runs.Add(run.Clone(true));
            //    //�ƦW
            //    run.Text = "" + info.Place + " / " + info.Radix;
            //    entry_table.Rows[row_index].Cells[10].Paragraphs[0].Runs.Add(run.Clone(true));
            //    //�ʤ���
            //    run.Text = "" + info.GetPercentage() + "%";
            //    entry_table.Rows[row_index].Cells[11].Paragraphs[0].Runs.Add(run.Clone(true));

            //    row_index++;
            //}
        }

        MemoryStream GetTemplate(ReportOptions option)
        {
            if (option.IsDefaultTemplate)
                return new MemoryStream(option.DefaultTemplate);
            return new MemoryStream(option.CustomizeTemplate);
        }

        int GetMoralLevel(decimal moral_score)
        {
            //moral_score = Math.Round(moral_score);
            if (moral_score >= 90)
                return 4;
            else if (moral_score >= 80)
                return 3;
            else if (moral_score >= 70)
                return 2;
            else if (moral_score >= 60)
                return 1;
            else
                return 0;
        }
    }

    class EntryComparer : IComparer<EntryInfo>
    {
        #region IComparer<string> ����

        public int Compare(EntryInfo x, EntryInfo y)
        {
            return SortByEntryName(x.Name, y.Name);
        }

        #endregion

        public int SortByEntryName(string a, string b)
        {
            int ai = getIntForEntry(a), bi = getIntForEntry(b);
            if (ai > 0 || bi > 0)
                return ai.CompareTo(bi);
            else
                return a.CompareTo(b);
        }

        private int getIntForEntry(string a1)
        {
            List<string> list = new List<string>();
            list.AddRange(new string[] { "�Ƿ~", "��߬��", "��|", "�꨾�q��", "���d�P�@�z", "�w��" });

            int x = list.IndexOf(a1);
            if (x < 0)
                return list.Count;
            else
                return x;
        }
    }

    class SubjectComparer : IComparer<SubjectInfo>
    {
        #region IComparer<SemesterSubjectScoreInfo> ����

        public int Compare(SubjectInfo x, SubjectInfo y)
        {
            return SortBySubjectName(x.Name, y.Name);
        }

        #endregion

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
            List<string> list = new List<string>();
            list.AddRange(new string[] { "��", "�^", "��", "��", "��", "��", "��", "��", "�a", "��", "��", "¦", "�@" });

            int x = list.IndexOf(a1);
            if (x < 0)
                return list.Count;
            else
                return x;
        }
    }
}