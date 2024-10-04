using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;
using Aspose.Cells;
using FISCA.Presentation.Controls;
using SmartSchool.ApplicationLog;

namespace SmartSchool.Evaluation.Content.ChangeSchoolYear
{
    public partial class frmChangeSemesterSchoolYearMsg : BaseForm
    {
        private string SourceMessage = "";
        private string ChangeMessage = "";
        private StudentInfo studentInfo = null;

        SemesterScoreData semsSourceData = null;
        SemesterScoreData semsChangeData = null;
        SemesterScoreDataTransfer semsTransfer = null;

        private string SourceSchoolYear = "", ChangeSchoolYear = "", Semester = "", GradeYear = "";

        public frmChangeSemesterSchoolYearMsg()
        {
            InitializeComponent();
            semsTransfer = new SemesterScoreDataTransfer();
            studentInfo = new StudentInfo();
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
        }

        private void btnChange_Click(object sender, EventArgs e)
        {
            btnChange.Enabled = false;

            try
            {
                // 取得來源資料
                semsSourceData = new SemesterScoreData();
                semsSourceData.StudentID = studentInfo.StudentID;
                semsSourceData.SchoolYear = SourceSchoolYear;
                semsSourceData.Semester = Semester;
                semsSourceData.GradeYear = GradeYear;
                semsSourceData = semsTransfer.GetStudentScoreDataByIDSchoolYear(semsSourceData);

                bool chkSourceHasData = true;
                // 檢查來源資料是否有資料
                if (string.IsNullOrEmpty(semsSourceData.ScoreInfo))
                {
                    chkSourceHasData = false;
                }
                else
                {
                    try
                    {
                        XElement elmRoot = XElement.Parse(semsSourceData.ScoreInfo);
                        if (elmRoot.Elements("Subject").Count() == 0)
                            chkSourceHasData = false;
                    }
                    catch (Exception ex)
                    {
                        chkSourceHasData = false;
                        Console.WriteLine(ex.Message);
                    }

                }

                if (chkSourceHasData == false)
                {
                    MsgBox.Show("沒有學期科目成績。");
                    btnChange.Enabled = true;
                    return;
                }


                // 預計寫入資料
                semsChangeData = new SemesterScoreData();
                semsChangeData.StudentID = studentInfo.StudentID;
                semsChangeData.SchoolYear = ChangeSchoolYear;
                semsChangeData.Semester = Semester;
                semsChangeData.GradeYear = GradeYear;
                semsChangeData = semsTransfer.GetStudentScoreDataByIDSchoolYear(semsChangeData);

                // CT,檢查資料庫 sems_subj_scoree 有資料庫限制，StudentID+SchoolYear+Semester 是為唯一值

                // 檢查要寫入是否有資料
                if (semsSourceData.SchoolYear != semsChangeData.SchoolYear && semsSourceData.Semester != semsChangeData.Semester)
                {
                    // 更新資料                    
                    int result = semsTransfer.UpdateStudentSemesterScoreSchoolYearBySemsID(semsSourceData.ID, ChangeSchoolYear);
                    if (result > 0)
                    {
                        // log 
                        StringBuilder updateDesc = new StringBuilder("");
                        updateDesc.AppendLine("學號：" + studentInfo.StudentNumber + ",班級：" + studentInfo.ClassName + ",座號：" + studentInfo.SeatNo + ",姓名：" + studentInfo.StudentName);
                        updateDesc.AppendLine("更新學期成績，學年度由「" + semsSourceData.SchoolYear + "」改成「" + ChangeSchoolYear + "」。");
                        updateDesc.AppendLine("學期成績內，科目與科目級別：");
                        try
                        {
                            XElement elmRoot = XElement.Parse(semsSourceData.ScoreInfo);
                            foreach (XElement elm in elmRoot.Elements("Subject"))
                            {
                                try
                                {
                                    string subjectName = elm.Attribute("科目").Value;
                                    string level = elm.Attribute("科目級別").Value;
                                    updateDesc.AppendLine("科目：" + subjectName + " ,級別" + level);
                                }
                                catch (Exception ex)
                                { Console.WriteLine(ex.Message); }

                            }

                            CurrentUser.Instance.AppLog.Write(EntityType.Student, EntityAction.Update, semsSourceData.StudentID, updateDesc.ToString(), "學期成績", "");

                            // 檢查學業分項成績學年度、學期是否有資料，如果沒有相同就搬過去
                            int resultEntry = semsTransfer.CheckAndUpdateStudentSemesterScoreSchoolYear(semsSourceData, ChangeSchoolYear);

                            if (resultEntry > 0)
                            {
                                EventHub.Instance.InvokScoreChanged(semsSourceData.StudentID);
                                MsgBox.Show("更新資料成功");
                                this.DialogResult = DialogResult.Yes;
                            }
                            else
                            {
                                MsgBox.Show("更新學期成績成功，請重新計算學期分項成績。");
                                EventHub.Instance.InvokScoreChanged(semsSourceData.StudentID);
                                this.DialogResult = DialogResult.Yes;
                            }
                        }
                        catch (Exception ex)
                        { Console.WriteLine(ex.Message); }

                    }
                    else
                    {
                        this.DialogResult = DialogResult.No;
                        MsgBox.Show("更新學生學期科目成績學年度資料失敗。");
                    }
                }
                else
                {
                    // 已有資料需要比對
                    List<SubjectScoreInfo> CheckDataList = semsTransfer.CheckHasSubjectNameLevel(semsSourceData.ScoreInfo, semsChangeData.ScoreInfo);

                    // 有重複資料，提示請使用者修改
                    if (CheckDataList.Count > 0)
                    {
                        frmChangeSemesterSchoolYearHasData fHasData = new frmChangeSemesterSchoolYearHasData();
                        fHasData.SetMessage(semsChangeData.SchoolYear, semsChangeData.Semester, semsChangeData.GradeYear);

                        // 填入 CheckDataList 學年度、學期、年級
                        foreach (SubjectScoreInfo ssi in CheckDataList)
                        {
                            ssi.SchoolYear = semsChangeData.SchoolYear;
                            ssi.Semester = semsChangeData.Semester;
                            ssi.GradeYear = semsChangeData.GradeYear;
                        }

                        fHasData.SetStudentInfo(studentInfo);
                        fHasData.SetHasDataList(CheckDataList);
                        if (fHasData.ShowDialog() == DialogResult.Cancel)
                            this.DialogResult = DialogResult.Cancel;
                    }
                    else
                    {
                        // 沒有重複資料，直接接續寫入
                        semsChangeData.ScoreInfo = semsTransfer.AppendScoreData(semsSourceData.ScoreInfo, semsChangeData.ScoreInfo);
                        int result = 0;
                        // 判斷id是否存在
                        if (semsChangeData.ID == null)
                        {
                            result = semsTransfer.InsertData(semsChangeData);
                        }
                        else
                        {
                            result = semsTransfer.UpdateScoreDataBySemesScoreID(semsChangeData);
                        }

                        if (result > 0)
                        {
                            // log 
                            StringBuilder updateDesc = new StringBuilder("");
                            updateDesc.AppendLine("學號：" + studentInfo.StudentNumber + ",班級：" + studentInfo.ClassName + ",座號：" + studentInfo.SeatNo + ",姓名：" + studentInfo.StudentName);
                            updateDesc.AppendLine("更新學期成績，學年度由「" + semsSourceData.SchoolYear + "」改成「" + semsChangeData.SchoolYear + "」。");
                            updateDesc.AppendLine("調整後學期成績內，科目與科目級別：");
                            try
                            {
                                XElement elmRoot = XElement.Parse(semsChangeData.ScoreInfo);
                                foreach (XElement elm in elmRoot.Elements("Subject"))
                                {
                                    try
                                    {
                                        string subjectName = elm.Attribute("科目").Value;
                                        string level = elm.Attribute("科目級別").Value;
                                        updateDesc.AppendLine("科目：" + subjectName + " ,級別" + level);
                                    }
                                    catch (Exception ex)
                                    { Console.WriteLine(ex.Message); }

                                }

                                CurrentUser.Instance.AppLog.Write(EntityType.Student, EntityAction.Update, semsSourceData.StudentID, updateDesc.ToString(), "學期成績", "");

                                // 刪除原本來源
                                int resultDel = semsTransfer.DeleteSemesScoreBySemsID(semsSourceData.ID);
                                EventHub.Instance.InvokScoreChanged(semsChangeData.StudentID);

                                if (resultDel > 0)
                                {
                                    MsgBox.Show("更新學期成績成功，請重新計算學期分項成績。");
                                    this.Close();
                                }
                            }
                            catch (Exception ex) { Console.WriteLine(ex.Message); }
                        }
                        else
                        {
                            MsgBox.Show("更新資料失敗");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MsgBox.Show(ex.Message);
            }


            btnChange.Enabled = true;
        }

        private void frmChangeSemesterSchoolYearMsg_Load(object sender, EventArgs e)
        {
            this.MinimumSize = this.MaximumSize = this.Size;
            labelSourceMsg.Text = SourceMessage;
            labelChangeMsg.Text = ChangeMessage;
        }

        public void SetSourceMessage(string SchoolYear, string Semester, string GradeYear)
        {
            SourceSchoolYear = SchoolYear;
            this.Semester = Semester;
            this.GradeYear = GradeYear;

            SourceMessage = "將學年度：" + SourceSchoolYear + " 學期：" + Semester + " 年級：" + GradeYear;
        }


        public void SetChangeMessage(string SchoolYear, string Semester, string GradeYear)
        {
            ChangeSchoolYear = SchoolYear;
            this.Semester = Semester;
            this.GradeYear = GradeYear;

            ChangeMessage = "調整為學年度：" + ChangeSchoolYear + " 學期：" + Semester + " 年級：" + GradeYear;
        }
        public void SetStudentInfo(StudentInfo studentInfo)
        {
            this.studentInfo = studentInfo;
        }
    }
}
