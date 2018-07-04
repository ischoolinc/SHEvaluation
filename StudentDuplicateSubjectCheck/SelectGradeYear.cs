using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using FISCA.Presentation.Controls;
using K12.Data;
using FISCA.Data;
using SHSchool.Data;

namespace StudentDuplicateSubjectCheck
{
    public partial class SelectGradeYear : BaseForm
    {
        string selectGradeYear;
        string targetGradeYear;
        string schoolYear = School.DefaultSchoolYear;
        string semester = School.DefaultSemester;
        private BackgroundWorker _backgroundWorker;

        List<SCAttendRecord> scaDuplicateList = new List<SCAttendRecord>();

        Dictionary<string, List<string>> dataCompareDict = new Dictionary<string, List<string>>(); // <sid,<科目名稱 + _ + 級別>>

        public SelectGradeYear()
        {
            InitializeComponent();

            
            this.numericUpDown1.Minimum = 1;
            this.numericUpDown1.Maximum = 5;
            this.numericUpDown1.Value = 1;

            labelX2.Text = "本功能將會對本學期學生的修課紀錄檢查，" + Environment.NewLine + "假若與先前的學期成績有重覆的科目級別則會列出，" + Environment.NewLine + "由人工設定計算方式。";

            _backgroundWorker = new BackgroundWorker();
            _backgroundWorker.DoWork += new DoWorkEventHandler(_backgroundWorker_DoWork);
            _backgroundWorker.ProgressChanged += new ProgressChangedEventHandler(_backgroundWorker_ProgressChanged);
            _backgroundWorker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(_backgroundWorker_RunWorkerCompleted);
            _backgroundWorker.WorkerReportsProgress = true;

        }

        private void buttonX1_Click(object sender, EventArgs e)
        {
            
            targetGradeYear = ""+this.numericUpDown1.Value;
            FreezeUI();
            _backgroundWorker.RunWorkerAsync();
        }

        private void _backgroundWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            ActivateUI();

            FISCA.Presentation.MotherForm.SetStatusBarMessage("檢查本學期重覆修課完畢。");

            // 有資料才show
            if (scaDuplicateList.Count != 0)
            {
                CheckCalculatedLogicForm cclf = new CheckCalculatedLogicForm(scaDuplicateList);
                cclf.ShowDialog();
            }
            else
            {
                MsgBox.Show("本年級本學期並無重覆修課的紀錄。");
            }
            
        }

        private void _backgroundWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            FISCA.Presentation.MotherForm.SetStatusBarMessage("取得重覆級別科目中...", e.ProgressPercentage);
        }

        private void _backgroundWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            _backgroundWorker.ReportProgress(10);

            //清空上次資料
            scaDuplicateList.Clear();



            List<string> sidList = new List<string>();

            // 學生班級基本資料(學生狀態為一般)
            QueryHelper qh1 = new QueryHelper();
            string query1 = "select student.id,class.grade_year,student_number,class.class_name,seat_no,student.name from student inner join class on student.ref_class_id=class.id where student.status=1 and class.grade_year ='" + targetGradeYear + "' order by class.grade_year,student_number;";
            DataTable dt1 = qh1.Select(query1);
            foreach (DataRow dr in dt1.Rows)
            {
                sidList.Add("" + dr["id"]);
            }

            _backgroundWorker.ReportProgress(30);

            // 2018/7/4 穎驊協助 嘉詮處理客服 https://ischool.zendesk.com/agent/tickets/6133 ，發現 ischool API  SHSemesterScore.SelectByStudentID 
            // 預設會將 學期歷程重覆的 重讀學期資料濾掉， 需要設定為 false，才會有。
            List<SHSchool.Data.SHSemesterScoreRecord> ssList = SHSchool.Data.SHSemesterScore.SelectByStudentIDs(sidList,false);

            foreach (SHSemesterScoreRecord data in ssList)
            {
                // 2018/6/12 穎驊筆記，佳樺測出 會抓到同學期的成績比對，在此濾掉， 同學期的重覆不算，因為學校可能會先算過本學期的 成績， 那本學期成績 與本學期修課重覆 是很正常的事
                if ("" + data.SchoolYear == schoolYear && "" + data.Semester == semester)
                {
                    continue;
                }
                // 以前學期，學期科目成績紀錄(題外話 如果未來覺得 scaList 使用ischool API的抓法太沒效率，可以考慮直接寫一串SQL)
                if (!dataCompareDict.ContainsKey(data.RefStudentID))
                {
                    dataCompareDict.Add(data.RefStudentID, new List<string>());

                    foreach (var subjectData in data.Subjects.Values)
                    {
                        if (!dataCompareDict[data.RefStudentID].Contains(subjectData.Subject + "_" + subjectData.Level))
                        {
                            dataCompareDict[data.RefStudentID].Add(subjectData.Subject + "_" + subjectData.Level);
                        }
                    }
                }
                else
                {
                    foreach (var subjectData in data.Subjects.Values)
                    {
                        if (!dataCompareDict[data.RefStudentID].Contains(subjectData.Subject + "_" + subjectData.Level))
                        {
                            dataCompareDict[data.RefStudentID].Add(subjectData.Subject + "_" + subjectData.Level);
                        }
                    }
                }
            }

            _backgroundWorker.ReportProgress(60);

            #region 舊寫法
            // 2018/5/23 穎驊 筆記 使用ischool API 抓不到 scattend 中的 extensions 資料，故採用SQL 來抓
            //List<SHSchool.Data.SHSCAttendRecord> scaList = SHSchool.Data.SHSCAttend.SelectByStudentIDs(sidList);

            //// 是本學期的修課紀錄(題外話 如果未來覺得 scaList 使用ischool API的抓法太沒效率，可以考慮直接寫一串SQL)
            //foreach (SHSCAttendRecord data in scaList)
            //{
            //    // 本學期的修課紀錄(題外話 如果未來覺得 scaList 使用ischool API的抓法太沒效率，可以考慮直接寫一串SQL)
            //    if (data.Course.SchoolYear == int.Parse(schoolYear) && data.Course.Semester == int.Parse(semester))
            //    {
            //        if (dataCompareDict.ContainsKey(data.RefStudentID))
            //        {
            //            if (dataCompareDict[data.RefStudentID].Contains(data.Course.Subject + "_" + data.Course.Level))
            //            {
            //                scaDuplicateList.Add(data); // 假如有找到重覆的科目級別資料，將 本學期修課紀錄加入List 處理
            //            }
            //        }
            //    }
            //} 
            #endregion

            //取得學生的修課紀錄、學生基本資料、課程名稱、級別

            string query2 = string.Format(@"SELECT sc_attend.id
	,sc_attend.extensions AS extensions
	,student.id AS refStudentID	
	,student.name AS studentName
	,student.student_number AS studentNumber
	,student.seat_no AS seatNo  
	,class.class_name AS className
	,class.grade_year AS gradeYear
	,course.id AS refCourseID
    ,course.course_name AS courseName
	,course.subject AS subjectName
	,course.subj_level AS subjectLevel
	FROM sc_attend 
	LEFT JOIN student ON sc_attend.ref_student_id =student.id 
	LEFT JOIN class ON student.ref_class_id =class.id  
	LEFT JOIN course ON sc_attend.ref_course_id =course.id  
	WHERE 
	student.status ='1' 
	AND course.school_year = '{0}'
    AND course.semester = '{1}'
    AND class.grade_year = '{2}'
    ORDER BY courseName,className, seatNo ASC", schoolYear, semester, targetGradeYear);

            DataTable dt_SCAttend = qh1.Select(query2);

            _backgroundWorker.ReportProgress(90);

            foreach (DataRow dr in dt_SCAttend.Rows)
            {
                if (dataCompareDict.ContainsKey("" + dr["refStudentID"]))
                {
                    if (dataCompareDict["" + dr["refStudentID"]].Contains(dr["subjectName"] + "_" + dr["subjectLevel"]))
                    {
                        SCAttendRecord scar = new SCAttendRecord();

                        scar.ID = "" + dr["id"];
                        scar.Extensions = "" + dr["extensions"];
                        scar.RefStudentID = "" + dr["refStudentID"];
                        scar.Name = "" + dr["studentName"];
                        scar.StudentNumber = "" + dr["studentNumber"];
                        scar.SeatNo = "" + dr["seatNo"];
                        scar.RefCourseID = "" + dr["refStudentID"];
                        scar.ClassName = "" + dr["className"];
                        scar.GradeYear = "" + dr["gradeYear"];
                        scar.RefCourseID = "" + dr["refCourseID"];
                        scar.CourseName = "" + dr["courseName"];
                        scar.SubjectName = "" + dr["subjectName"];
                        scar.SubjectLevel = "" + dr["subjectLevel"];

                        scaDuplicateList.Add(scar); // 假如有找到重覆的科目級別資料，將 本學期修課紀錄加入List 處理
                    }
                }
            }
            _backgroundWorker.ReportProgress(100);
        }


        //關閉
        private void buttonX2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void FreezeUI()
        {            
            buttonX1.Enabled = false;
            buttonX2.Enabled = false;
        }

        private void ActivateUI()
        {            
            buttonX1.Enabled = true;
            buttonX2.Enabled = true;
        }
    }
}

