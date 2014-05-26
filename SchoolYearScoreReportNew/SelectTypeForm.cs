using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Collections;

namespace SchoolYearScoreReport
{
    public partial class SelectTypeForm : FISCA.Presentation.Controls.BaseForm
    {
        private BackgroundWorker _BGWAbsenceAndPeriodList;
        private SchoolYearScoreReport.Config _config;
        private string _preferenceElementName;
        private List<string> absenceList = new List<string>();
        private List<string> typeList = new List<string>();
        
        public SelectTypeForm(string name, SchoolYearScoreReport.Config config)
        {
            this.InitializeComponent();

            this.buttonX1.AccessibleRole = AccessibleRole.PushButton;

            this._config = config;
            this._preferenceElementName = name;
            this._BGWAbsenceAndPeriodList = new BackgroundWorker();
            this._BGWAbsenceAndPeriodList.DoWork += new DoWorkEventHandler(this._BGWAbsenceAndPeriodList_DoWork);
            this._BGWAbsenceAndPeriodList.RunWorkerCompleted += new RunWorkerCompletedEventHandler(this._BGWAbsenceAndPeriodList_RunWorkerCompleted);
            this._BGWAbsenceAndPeriodList.RunWorkerAsync();
        }

        private void _BGWAbsenceAndPeriodList_DoWork(object sender, DoWorkEventArgs e)
        {
            List<K12.Data.PeriodMappingInfo> periodMappingInfos = K12.Data.PeriodMapping.SelectAll();
            List<K12.Data.AbsenceMappingInfo> absenceMappingInfos = K12.Data.AbsenceMapping.SelectAll();

            //SmartSchool.Customization.Data.SystemInformation.getField("Period");
            //SmartSchool.Customization.Data.SystemInformation.getField("Absence");
            //XmlHelper periodHelper = new XmlHelper(SmartSchool.Customization.Data.SystemInformation.Fields["Period"] as XmlElement);
            //XmlHelper absenceHelper = new XmlHelper(SmartSchool.Customization.Data.SystemInformation.Fields["Absence"] as XmlElement);
            //foreach (XmlElement var in periodHelper.GetElements("Period"))
            //{
            //    if (!this.typeList.Contains(var.GetAttribute("Type")))
            //    {
            //        this.typeList.Add(var.GetAttribute("Type"));
            //    }
            //}
            //foreach (XmlElement var in absenceHelper.GetElements("Absence"))
            //{
            //    if (!this.absenceList.Contains(var.GetAttribute("Name")))
            //    {
            //        this.absenceList.Add(var.GetAttribute("Name"));
            //    }
            //}
            foreach (K12.Data.PeriodMappingInfo periodMappingInfo in periodMappingInfos)
            {
                if (!this.typeList.Contains(periodMappingInfo.Type))
                    this.typeList.Add(periodMappingInfo.Type);
            }
            foreach (K12.Data.AbsenceMappingInfo absenceMappingInfo in absenceMappingInfos)
            {
                if (!this.absenceList.Contains(absenceMappingInfo.Name))
                    this.absenceList.Add(absenceMappingInfo.Name);
            }
        }

        private void _BGWAbsenceAndPeriodList_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            DataGridViewTextBoxColumn colName = new DataGridViewTextBoxColumn {
                HeaderText = "節次分類",
                MinimumWidth = 70,
                Name = "colName",
                ReadOnly = true,
                SortMode = DataGridViewColumnSortMode.NotSortable,
                Width = 70
            };
            this.dataGridViewX1.Columns.Add(colName);
            foreach (string absence in this.absenceList)
            {
                DataGridViewCheckBoxColumn newCol = new DataGridViewCheckBoxColumn {
                    HeaderText = absence,
                    Width = 0x37,
                    ReadOnly = false,
                    SortMode = DataGridViewColumnSortMode.NotSortable,
                    Tag = absence,
                    ValueType = typeof(bool)
                };
                this.dataGridViewX1.Columns.Add(newCol);
            }
            foreach (string type in this.typeList)
            {
                DataGridViewRow row = new DataGridViewRow();
                row.CreateCells(this.dataGridViewX1, new object[] { type });
                row.Tag = type;
                this.dataGridViewX1.Rows.Add(row);
            }
            foreach (string type in this.Config.SelectTypes.Keys)
            {
                foreach (DataGridViewRow row in (IEnumerable) this.dataGridViewX1.Rows)
                {
                    if (type == ("" + row.Tag))
                    {
                        foreach (string absence in this.Config.SelectTypes[type])
                        {
                            foreach (DataGridViewCell cell in row.Cells)
                            {
                                if ((cell.OwningColumn is DataGridViewCheckBoxColumn) && (("" + cell.OwningColumn.Tag) == absence))
                                {
                                    cell.Value = true;
                                }
                            }
                        }
                        break;
                    }
                }
            }
        }

        private void buttonX1_Click(object sender, EventArgs e)
        {
            Dictionary<string, List<string>> types = new Dictionary<string, List<string>>();
            foreach (DataGridViewRow row in (IEnumerable) this.dataGridViewX1.Rows)
            {
                string type = row.Tag.ToString();
                if (!types.ContainsKey(type))
                {
                    types.Add(type, new List<string>());
                }
                foreach (DataGridViewCell cell in row.Cells)
                {
                    if ((cell.Value is bool) && ((bool) cell.Value))
                    {
                        string absence = cell.OwningColumn.Tag.ToString();
                        if (!types[type].Contains(absence))
                        {
                            types[type].Add(absence);
                        }
                    }
                }
            }
            this.Config.SetTypes(types);
            this.Config.Save();
            base.DialogResult = DialogResult.OK;
            base.Close();
        }

        private SchoolYearScoreReport.Config Config
        {
            get
            {
                return this._config;
            }
        }    
    }
}
