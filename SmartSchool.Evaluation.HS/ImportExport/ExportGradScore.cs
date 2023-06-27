using SmartSchool.AccessControl;
using SmartSchool.Customization.PlugIn.ImportExport;
using System;
using System.Xml;

namespace SmartSchool.Evaluation.ImportExport
{
    [FeatureCode("Button0175")]
    class ExportGradScore : ExportProcess
    {
        public ExportGradScore()
        {
            this.Image = null;
            this.Title = "匯出畢業成績";
            this.Group = "畢業成績";
            // 2022-01 Cynthia 移除"體育畢業成績", "國防通識畢業成績", "健康與護理畢業成績"，新增"專業科目畢業成績"
            foreach (string var in new string[] { "學業畢業成績", "專業科目畢業成績", "實習科目畢業成績", "德行畢業成績" })
            {
                this.ExportableFields.Add(var);
            }
            this.ExportPackage += new EventHandler<ExportPackageEventArgs>(ExportGradScore_ExportPackage);
        }

        void ExportGradScore_ExportPackage(object sender, ExportPackageEventArgs e)
        {
            foreach (XmlElement studentElement in SmartSchool.Feature.QueryStudent.GetDetailList(new string[] { "ID", "GradScore" }, e.List.ToArray()).GetContent().GetElements("Student"))
            {
                RowData row = new RowData();
                row.ID = studentElement.GetAttribute("ID");
                foreach (XmlElement scoreElement in studentElement.SelectNodes("GradScore/GradScore/EntryScore"))
                {
                    string entry = scoreElement.GetAttribute("Entry") + "畢業成績";
                    if (ExportableFields.Contains(entry))
                    {
                        row.Add(entry, scoreElement.GetAttribute("Score"));
                    }
                }
                e.Items.Add(row);
            }
        }
    }
}
