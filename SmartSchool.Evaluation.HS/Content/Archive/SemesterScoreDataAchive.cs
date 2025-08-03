using FISCA.Data;
using FISCA.Permission;
using FISCA.Presentation;
using FISCA.Presentation.Controls;
using FISCA.UDT;
using SmartSchool.Evaluation.Content.Model;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Text;
using System.Windows.Forms;

namespace SmartSchool.Evaluation.Content
{
    [FeatureCode("SemesterScoreDataAchive", "學期成績(封存)")]
    public partial class SemesterScoreDataAchive : DetailContent
    {
        AccessHelper accessHelper = new AccessHelper();
        private BackgroundWorker BGW = new BackgroundWorker();
        private bool BkWBool = false;
        internal static FeatureAce UserPermission;
        bool _ReloadArchiveData = true;  //Cyn
        public SemesterScoreDataAchive()
        {
            InitializeComponent();

            #region UDT
            AccessHelper _a = new AccessHelper();
            //List<SemesterSubjectScoreArchive> sssa = new List<SemesterSubjectScoreArchive>();
            //List<SemesterEntryScoreArchive> sesa = new List<SemesterEntryScoreArchive>();
            //SchemaManager schema = new SchemaManager(FISCA.Authentication.DSAServices.DefaultConnection);
            //schema.SyncSchema(new SemesterSubjectScoreArchive());
            //schema.SyncSchema(new SemesterEntryScoreArchive());
            #endregion

            UserPermission = UserAcl.Current[FISCA.Permission.FeatureCodeAttribute.GetCode(GetType())];
            Group = "學期成績(封存)";
            btnDelete.Visible = FISCA.Permission.UserAcl.Current[Permissions.學期成績封存].Editable;
            BGW.DoWork += BGW_DoWork;
            BGW.RunWorkerCompleted += BGW_RunWorkerCompleted;

            EventHub.ArchiveChanged += new EventHandler(EventHub_ArchiveChanged);  //Cyn
        }

        void EventHub_ArchiveChanged(object sender, EventArgs e)  //Cyn
        {
            _ReloadArchiveData = true;
        }
        private void BGW_DoWork(object sender, DoWorkEventArgs e)
        {
            List<SemesterEntryScoreArchive> semesterEntryScoreArchives = accessHelper.Select<SemesterEntryScoreArchive>("ref_student_id=" + this.PrimaryKey);
            List<SemesterSubjectScoreArchive> semesterSubjectScoreArchives = accessHelper.Select<SemesterSubjectScoreArchive>("ref_student_id=" + this.PrimaryKey);

            e.Result = semesterEntryScoreArchives;
            e.Result = semesterSubjectScoreArchives;

        }

        private void BGW_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (BkWBool) //如果有其他的更新事件
            {
                BkWBool = false;
                BGW.RunWorkerAsync();
                return;
            }

            FillListView();

            this.Loading = false;
            btnDelete.Enabled = false;
            btnView.Enabled = false;
        }

        protected override void OnPrimaryKeyChanged(EventArgs e)
        {
            Changed();
        }

        public void Changed()
        {
            //#region 更新時
            if (this.PrimaryKey != "")
            {
                this.Loading = true;
                if (BGW.IsBusy)
                {
                    BkWBool = true;
                }
                else
                {
                    listView1.Items.Clear();
                    BGW.RunWorkerAsync();
                }
            }
        }

        private void listView1_SelectedIndexChanged(object sender, EventArgs e)
        {
            btnView.Enabled = (listView1.SelectedIndices.Count == 1 && (FISCA.Permission.UserAcl.Current[Permissions.學期成績封存].Viewable || FISCA.Permission.UserAcl.Current[Permissions.學期成績封存].Editable));
            btnDelete.Enabled = (listView1.SelectedIndices.Count == 1 && FISCA.Permission.UserAcl.Current[Permissions.學期成績封存].Editable);
        }

        private void btnDelete_Click(object sender, EventArgs e)
        {
            List<K12.Data.StudentRecord> studentList = K12.Data.Student.SelectByIDs(K12.Presentation.NLDPanels.Student.SelectedSource);
            Dictionary<string, K12.Data.StudentRecord> studentDic = new Dictionary<string, K12.Data.StudentRecord>();
            foreach (K12.Data.StudentRecord each in studentList)
            {
                if (!studentDic.ContainsKey(each.ID))
                {
                    studentDic.Add(each.ID, each);
                }
            }

            UpdateHelper updateHelper = new UpdateHelper();
            // QueryHelper qh = new QueryHelper();

            if (MsgBox.Show("您確定要刪除此學期成績(封存)？", "", MessageBoxButtons.YesNo) == DialogResult.No) return;

            foreach (string studentID in K12.Presentation.NLDPanels.Student.SelectedSource)
            {
                StudentSemsSocreArchive sssa = listView1.SelectedItems[0].Tag as StudentSemsSocreArchive;
                string sql = "DELETE  FROM $semester_entry_score_archive WHERE uid={0} ;";
                string sql2 = "DELETE FROM $semester_subject_score_archive WHERE uid={0}";

                // 刪除分項成績
                sql = string.Format(sql, sssa.RefEntryUid);
                //刪除科目成績
                sql2 = string.Format(sql2, sssa.Uid);

                List<string> sqls = new List<string>();
                sqls.Add(sql);
                sqls.Add(sql2);
                //string sqls = sql + sql2;
                //qh.Select(sqls);

                updateHelper.Execute(sqls);

                // 刪除學期成績(封存) +封存時間log
                StringBuilder deleteDesc = new StringBuilder("");
                deleteDesc.AppendLine("學生姓名：" + studentDic[studentID].Name + " ");
                deleteDesc.AppendLine("刪除 " + listView1.SelectedItems[0].SubItems[0].Text + " 學年度 第 " + listView1.SelectedItems[0].SubItems[1].Text + " 學期 學期成績(封存)");
                FISCA.LogAgent.ApplicationLog.Log("學期成績(封存)", "刪除", "學生", this.PrimaryKey, deleteDesc.ToString());

                MsgBox.Show("刪除完成。");
                Changed();
            }
        }


        private void btnView_Click(object sender, EventArgs e)
        {
            StudentSemsSocreArchive sssa = listView1.SelectedItems[0].Tag as StudentSemsSocreArchive;
            ScoreEditor.SemesterScoreArchiveDetails viewer = new ScoreEditor.SemesterScoreArchiveDetails(this.PrimaryKey, sssa.Uid, sssa.RefEntryUid);
            if (viewer.ShowDialog() == DialogResult.OK)
            {
                Changed();
            }

        }

        private void listView1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            if (listView1.SelectedIndices.Count == 1)
            {
                StudentSemsSocreArchive sssa = listView1.SelectedItems[0].Tag as StudentSemsSocreArchive;
                ScoreEditor.SemesterScoreArchiveDetails viewer = new ScoreEditor.SemesterScoreArchiveDetails(this.PrimaryKey, sssa.Uid, sssa.RefEntryUid);
                viewer.ShowDialog();
            }
        }

        public void FillListView()
        {
            QueryHelper qh = new QueryHelper();
            string sql = @"SELECT 
a.uid
, a.ref_student_id
, a.school_year
, a.semester
, a.grade_year, array_to_string(xpath('//SemesterEntryScore/Entry[@分項=''學業'']/@成績', xmlparse(content b.score_info)), '')::text AS 學業成績
, a.last_update, a.ref_sems_entry_uid
FROM $semester_subject_score_archive a
LEFT JOIN $semester_entry_score_archive b
ON a.ref_sems_entry_uid=b.uid
WHERE a.ref_student_id={0}";
            sql = string.Format(sql, this.PrimaryKey);

            try
            {
                DataTable dt = qh.Select(sql);

                int i = 0;
                foreach (DataRow dr in dt.Rows)
                {
                    StudentSemsSocreArchive sssa = new StudentSemsSocreArchive();
                    sssa.Uid = dr["uid"].ToString();
                    sssa.RefEntryUid = dr["ref_sems_entry_uid"].ToString();

                    listView1.Items.Add(dr["school_year"].ToString());
                    listView1.Items[i].Tag = sssa;
                    listView1.Items[i].SubItems.Add(dr["semester"].ToString());
                    listView1.Items[i].SubItems.Add(dr["grade_year"].ToString());
                    listView1.Items[i].SubItems.Add(dr["學業成績"].ToString());
                    listView1.Items[i].SubItems.Add(dr["last_update"].ToString());
                    i++;
                }
            }
            catch (Exception ex)
            {
                MsgBox.Show("取得學生學期成績(封存)發生錯誤。");
            }
            _ReloadArchiveData = false;    //Cyn
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            MsgBox.Show("\r\n" +
                "1. 需要有「學期成績(封存)」的編輯權限才能使用。\r\n\r\n" +
                "2. 此功能運用在「校內轉科」，用來保留學生當時的學期科目成績及課程代碼。\r\n\r\n" +
                "3. 封存的成績修改儲存會產生新的封存資料。\r\n\r\n" +
                "4. 德行成績已過時，故不在封存範圍。\r\n\r\n" +
                "5. 流程：\r\n" +
                "　　(1) 請先封存學生當前所有的學期成績至「學期成績(封存)」。\r\n" +
                "　　(2) 進行轉科異動作業。\r\n" +
                "　　(3) 將不可抵免的科目自學期成績中刪除。", "學期成績(封存)說明", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            ScoreEditor.SemesterScoreArchiveDetails Add = new ScoreEditor.SemesterScoreArchiveDetails(this.PrimaryKey, "", "");
            if (Add.ShowDialog() == DialogResult.OK)
            {
                Changed();
            }
        }
    }
}
