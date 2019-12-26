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

namespace StudentDuplicateSubjectCheck
{
    public partial class ErrorMessageForm : BaseForm
    {
        List<string> studentIDList = new List<string>();

        public ErrorMessageForm()
        {
            InitializeComponent();
        }

        private void btnOk_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void ErrorMessageForm_Load(object sender, EventArgs e)
        {


        }

        public void SetDataTable(DataTable dt)
        {
            dgData.Rows.Clear();
            studentIDList.Clear();
            foreach (DataRow dr in dt.Rows)
            {
                int rowIdx = dgData.Rows.Add();
                string sid = dr["student_id"].ToString();
                dgData.Rows[rowIdx].Tag = sid;
                if (!studentIDList.Contains(sid))
                  studentIDList.Add(sid);
                dgData.Rows[rowIdx].Cells[colStudentNumber.Index].Value = dr["student_number"].ToString();
                dgData.Rows[rowIdx].Cells[colClassName.Index].Value = dr["class_name"].ToString();
                dgData.Rows[rowIdx].Cells[colSeatNo.Index].Value = dr["seat_no"].ToString();
                dgData.Rows[rowIdx].Cells[colStudentName.Index].Value = dr["student_name"].ToString();
            }
        }

        public void SetMessage(string msg)
        {
            this.lblMsg.Text = msg;
        }

        private void btnAddTemp_Click(object sender, EventArgs e)
        {
            List<string> addIDList = new List<string>();

            if (studentIDList.Count> 0)
            {
                foreach (string id in studentIDList)
                {
                    if (!K12.Presentation.NLDPanels.Student.TempSource.Contains(id))
                        addIDList.Add(id);
                }

                K12.Presentation.NLDPanels.Student.AddToTemp(addIDList);
                MsgBox.Show("已於[學生待處理]加入" + studentIDList.Count + "名學生");
            }
        }
    }
}
