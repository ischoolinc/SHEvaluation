using SmartSchool.Common;
using System;
using System.Windows.Forms;

namespace SmartSchool.Evaluation.Configuration
{
    public partial class SubjectTableConfiguration : SmartSchool.Customization.PlugIn.Configure.ConfigurationItem
    {
        private string _Catalog;
        public SubjectTableConfiguration(string catalog)
        {
            InitializeComponent();
            this.expandablePanel1.TitleText = this.Caption = _Catalog = catalog;
            if (catalog == "�ǵ{��ت�")
            {
                this.subjectTableEditor1.ProgramTable = true;
            }
            else if (catalog == "�֤߬�ت�")
            {
                this.expandablePanel1.TitleText = "�ۭq���~���פή��ت�";
                this.Caption = "�ۭq���~���פή��ت�";
            }

        }
        protected override void OnActive()
        {
            RefillSubjectTables();
        }

        private void buttonX3_Click(object sender, EventArgs e)
        {
            new SubjectTableCreator(_Catalog).ShowDialog();
            RefillSubjectTables();
        }

        private void RefillSubjectTables()
        {
            string selectedName = "";
            if (dataGridViewX1.SelectedRows.Count == 1)
                selectedName = dataGridViewX1.SelectedRows[0].Cells[0].Value.ToString();

            dataGridViewX1.Rows.Clear();
            foreach (SubjectTableItem item in SubjectTable.Items[_Catalog])
            {
                if (item.Name == selectedName)
                {
                    dataGridViewX1.Rows[dataGridViewX1.Rows.Add(item)].Selected = true;
                }
                else
                    dataGridViewX1.Rows.Add(item);
            }
        }

        private void buttonX1_Click(object sender, EventArgs e)
        {
            if (dataGridViewX1.SelectedRows.Count == 1)
            {
                if (MsgBox.Show("�T�w�n�R�� '" + dataGridViewX1.SelectedRows[0].Cells[0].Value + "' �H", "�T�w", MessageBoxButtons.OKCancel) == DialogResult.OK)
                {
                    SmartSchool.Feature.SubjectTable.RemoveSubejctTable.Delete(((SubjectTableItem)dataGridViewX1.SelectedRows[0].Cells[0].Value).ID);
                    SubjectTable.Items[_Catalog].Reflash();
                    RefillSubjectTables();
                }
            }
        }

        private void dataGridViewX1_SelectionChanged(object sender, EventArgs e)
        {
            this.subjectTableEditor1.Visible = (dataGridViewX1.SelectedRows.Count == 1);
            if (dataGridViewX1.SelectedRows.Count == 1)
            {
                subjectTableEditor1.Content = ((SubjectTableItem)dataGridViewX1.SelectedRows[0].Cells[0].Value).Content;
            }
        }

        private void buttonX2_Click(object sender, EventArgs e)
        {
            if (dataGridViewX1.SelectedRows.Count == 1)
            {
                if (subjectTableEditor1.IsValidated())
                {
                    SmartSchool.Feature.SubjectTable.EditSubejctTable.UpdateSubject(((SubjectTableItem)dataGridViewX1.SelectedRows[0].Cells[0].Value).ID, this.subjectTableEditor1.Content);
                    SubjectTable.Items[_Catalog].Reflash();
                    RefillSubjectTables();
                }
                else
                {
                    MsgBox.Show("��J��Ʀ��~�A�L�k�x�s�C\n���ˬd��J��ơC");
                }
            }
        }
    }
}

