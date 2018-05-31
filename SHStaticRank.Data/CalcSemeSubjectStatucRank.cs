using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using FISCA.Presentation.Controls;
using K12.Data;

namespace SHStaticRank.Data
{
    public partial class CalcSemeSubjectStatucRank : BaseForm
    {
        
        private string _DefalutSchoolYear = "";
        private string _DefaultSemester = "";
        private List<TagConfigRecord> _TagConfigRecords = new List<TagConfigRecord>();

        public CalcSemeSubjectStatucRank():this(false)
        {
            
        }

        public CalcSemeSubjectStatucRank(bool fixSchoolyearSemester)
        {
            _TagConfigRecords = K12.Data.TagConfig.SelectByCategory(TagCategory.Student);
            _DefalutSchoolYear = "" + K12.Data.School.DefaultSchoolYear;
            _DefaultSemester = "" + K12.Data.School.DefaultSemester;


            InitializeComponent();

            int i;
            if (int.TryParse(_DefalutSchoolYear, out i))
            {
                for (int j = 0; j < 5; j++)
                {
                    cboSchoolYear.Items.Add("" + (i - j));
                }
            }
            cboSemester.Items.Add("1");
            cboSemester.Items.Add("2");

            cboSchoolYear.Text = "" + _DefalutSchoolYear;
            cboSemester.Text = "" + _DefaultSemester;

            cboSchoolYear.Enabled = !fixSchoolyearSemester; //// 2018/5/31 穎驊 因應[H成績][03] 學期成績計算相關流程操作文字提醒 調整，將原本的計算排名功能關閉，一率引導到固定排名功能
            cboSemester.Enabled = !fixSchoolyearSemester;   // 假若是自 教務作業 批次作業而來，則不給其調整學年度學期，只能在原學期動作

            List<string> prefix = new List<string>();
            List<string> tag = new List<string>();
            foreach (var item in _TagConfigRecords)
            {
                if (item.Prefix != "")
                {
                    if (!prefix.Contains(item.Prefix))
                        prefix.Add(item.Prefix);
                }
                else
                {
                    tag.Add(item.Name);
                }
            }
            cboRankRilter.Items.Clear();
            cboTagRank1.Items.Clear();
            cboTagRank2.Items.Clear();
            cboRankRilter.Items.Add("");
            cboTagRank1.Items.Add("");
            cboTagRank2.Items.Add("");
            foreach (var s in prefix)
            {
                cboRankRilter.Items.Add("[" + s + "]");
                cboTagRank1.Items.Add("[" + s + "]");
                cboTagRank2.Items.Add("[" + s + "]");
            }
            foreach (var s in tag)
            {
                cboRankRilter.Items.Add(s);
                cboTagRank1.Items.Add(s);
                cboTagRank2.Items.Add(s);
            }
        }



        private void CalcSemeSubjectStatucRank_Load(object sender, EventArgs e)
        {
            MessageBox.Show(
@"功能說明：
本功能以原有固定功能為基礎，加上類別篩選及類組排名。
排名完將會顯示排名細節，此資料目前暫不儲存，請自行留存。
1.排名對象為目前為該年級，扣除標註有不排名類別以及非一般狀態學生(非在選取的學期中就讀該年級的學生)。
2.計算過程中所有學生年級、班級、所屬類別等皆以目前的狀態進行計算，不會因為選取以前的學年度而往回推。
3.排名所採用的目前學期科目成績分數為勾選成績項目中選最高值。
4.若勾選'僅預覽，不儲存結果'項目，排名資料將不會寫入資料庫，僅有排名細節供預覽用。");
        }


        public Configure Setting { get; private set; }
        private void buttonX1_Click(object sender, EventArgs e)
        {
            Setting = new Configure();
            Setting.SchoolYear = cboSchoolYear.Text;
            Setting.Semester = cboSemester.Text;
            Setting.CalcGradeYear1 = chkGrade1.Checked;
            Setting.CalcGradeYear2 = chkGrade2.Checked;
            Setting.CalcGradeYear3 = chkGrade3.Checked;
            Setting.DoNotSaveIt = chkDoNotSaveIt.Checked;
            Setting.NotRankTag = cboRankRilter.Text;
            Setting.use原始成績 = chk原始成績.Checked;
            Setting.use補考成績 = chk補考成績.Checked;
            Setting.use重修成績 = chk重修成績.Checked;
            Setting.use手動調整成績 = chk手動調整成績.Checked;
            Setting.use學年調整成績 = chk學年調整成績.Checked;
            Setting.Rank1Tag = cboTagRank1.Text;
            Setting.Rank2Tag = cboTagRank2.Text;
            Setting.計算學業成績排名 = chk計算學業成績排名.Checked;
            Setting.清除不排名學生排名資料 = chkClearNoRankData.Checked;
            DialogResult = System.Windows.Forms.DialogResult.OK;
            this.Close();
        }
    }
}
