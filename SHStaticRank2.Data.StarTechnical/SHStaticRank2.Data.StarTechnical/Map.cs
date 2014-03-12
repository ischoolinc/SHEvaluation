using FISCA.Presentation.Controls;
using K12.Data;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace SHStaticRank2.Data.StarTechnical
{
    public partial class Map : BaseForm
    {
        private List<TagConfigRecord> _TagConfigRecords = new List<TagConfigRecord>();

        private FISCA.UDT.AccessHelper _AccessHelper = new FISCA.UDT.AccessHelper();
        private List<MapRecord> _MapRecords = new List<MapRecord>();
        public Map()
        {
            InitializeComponent();
            _TagConfigRecords = K12.Data.TagConfig.SelectByCategory(TagCategory.Student);
            List<string> prefix = new List<string>();
            List<string> tag = new List<string>();
            
            student_tag.Items.Clear();
            student_tag.Items.Add("");
            foreach (var item in _TagConfigRecords)
            {
                if (item.Prefix != "")
                    student_tag.Items.Add(item.Prefix + ":" + item.Name);
                else
                    student_tag.Items.Add(item.Name);
            }
            _MapRecords =  _AccessHelper.Select<MapRecord>();
            DataGridViewRow row ;
            foreach (MapRecord item in _MapRecords)
            {
                if (string.IsNullOrWhiteSpace(item.student_tag))
                    continue;
                row = new DataGridViewRow();
                row.CreateCells(dataGridView1);
                row.Cells[0].Value = item.student_tag;
                row.Cells[1].Value = item.code1;
                row.Cells[2].Value = item.code2;
                row.Cells[3].Value = item.note;
                dataGridView1.Rows.Add(row);
            }
        }
        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }
        private void buttonX1_Click(object sender, EventArgs e)
        {
            _AccessHelper.DeletedValues(_MapRecords);
            _MapRecords.Clear();
            MapRecord mr;
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                if (string.IsNullOrWhiteSpace("" + row.Cells["student_tag"].Value))
                    continue;
                mr = new MapRecord();
                mr.student_tag = "" + row.Cells["student_tag"].Value;
                mr.code1 = "" + row.Cells["code1"].Value;
                mr.code2 = "" + row.Cells["code2"].Value;
                mr.note = "" + row.Cells["note"].Value;
                _MapRecords.Add(mr);
            }
            _MapRecords.SaveAll();
            this.Close();
        }
    }
}
