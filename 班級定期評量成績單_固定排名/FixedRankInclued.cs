using FISCA.Presentation.Controls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace 班級定期評量成績單_固定排名
{
    public partial class FixedRankInclued : BaseForm
    {
        public List<string> SelectSubjucts = new List<string>();
        public Boolean BringSelectedSubj = false;
        public Dictionary<string, List<string>> FixedRankSubjInclude;
        public FixedRankInclued()
        {
            InitializeComponent();

        }

        /// <summary>
        /// 建構子
        /// </summary>
        /// <param name="fixedRankSubjInclude"> 某次考試 有結算的科目</param>
        public FixedRankInclued(Dictionary<string, List<string>> fixedRankSubjInclude, List<string> gradeYear)
        {
            InitializeComponent();
            FixedRankSubjInclude = fixedRankSubjInclude;
            //   this.labelX1.Text =$"{string.Join("、", gradeYear)} 年級" ;
            // dataGridViewX1.DataSource = fixedRankSubjInclude.ToArray();
            ReloadDataGridView();
        }

        private void btncheckSubj_Click(object sender, EventArgs e)
        {
            if (DialogResult.Yes == MsgBox.Show("原本勾選之科目有可能消失，確定帶入?", MessageBoxButtons.YesNo))
            {
                this.BringSelectedSubj = true;
                foreach (DataGridViewCell cell in dataGridViewX1.SelectedCells)
                {
                    if(cell.Value!=null)
                    SelectSubjucts.Add(cell.Value.ToString());
                }
                this.Close();
            }
            else
            {
                return;
            }
        }



        private void FixedRankInclued_Load(object sender, EventArgs e)
        {
            this.cboRankType.Items.Add("年、科、班排名");
            this.cboRankType.Items.Add("類別1排名");
            this.cboRankType.Items.Add("類別2排名");
            this.cboRankType.SelectedIndex = 0;
        }

        private void cboRankType_SelectionChangeCommitted(object sender, EventArgs e)
        {
            string rankType = this.cboRankType.Text;
            ReloadDataGridView(rankType);
        }



        private void ReloadDataGridView(string rankType = "年、科、班排名")
        {
            this.dataGridViewX1.Rows.Clear();
            if (!FixedRankSubjInclude.ContainsKey(rankType))
            {

                return;
            }
                FixedRankSubjInclude[rankType].Sort(new StringComparer("國文"
                                              , "英文"
                                              , "數學"
                                              , "理化"
                                              , "生物"
                                              , "社會"
                                              , "物理"
                                              , "化學"
                                              , "歷史"
                                              , "地理"
                                              , "公民"));


          
                foreach (string dr in FixedRankSubjInclude[rankType])
                {
                    DataGridViewRow dgrow = new DataGridViewRow();
                    dgrow.CreateCells(dataGridViewX1);
                    dgrow.Cells[0].Value = dr;
                    this.dataGridViewX1.Rows.Add(dgrow);
                }
        }
    }
}
