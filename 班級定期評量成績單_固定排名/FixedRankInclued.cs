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
        public FixedRankInclued()
        {
            InitializeComponent();
        }

        /// <summary>
        /// 建構子
        /// </summary>
        /// <param name="fixedRankSubjInclude"> 某次考試 有結算的科目</param>
        public FixedRankInclued(List<string> fixedRankSubjInclude)
        {
            InitializeComponent();
           // dataGridViewX1.DataSource = fixedRankSubjInclude.ToArray();
            foreach (string dr in fixedRankSubjInclude)
            {
                DataGridViewRow dgrow = new DataGridViewRow();
                dgrow.CreateCells(dataGridViewX1);
                dgrow.Cells[0].Value = dr;
                this.dataGridViewX1.Rows.Add(dgrow);
            }
        }
    }
}
