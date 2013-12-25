using System;
using System.Collections.Generic;
using System.Text;
//using SmartSchool.Customization.PlugIn.ImportExport;
using SmartSchool.Customization.Data;
using SmartSchool.Customization.Data.StudentExtension;
using SmartSchool.AccessControl;
using SmartSchool.API.PlugIn;

namespace SmartSchool.Evaluation.ImportExport
{
    [FeatureCode("Button0150")]
    class ExportSemesterEntryScore : SmartSchool.API.PlugIn.Export.Exporter
    {

        public ExportSemesterEntryScore()
        {
            this.Image = null;
            this.Text = "匯出學期分項成績";
        }

        public override void InitializeExport(SmartSchool.API.PlugIn.Export.ExportWizard wizard)
        {
            SmartSchool.API.PlugIn.VirtualCheckBox filterRepeat = new SmartSchool.API.PlugIn.VirtualCheckBox("自動略過重讀成績",true);
            wizard.Options.Add(filterRepeat);
            wizard.ExportableFields.AddRange("學年度", "學期", "成績年級", "學業", "體育", "國防通識", "健康與護理", "實習科目", "專業科目", "學業(原始)", "體育(原始)", "國防通識(原始)", "健康與護理(原始)", "實習科目(原始)", "專業科目(原始)", "德行");
            AccessHelper _AccessHelper=new AccessHelper();
            wizard.ExportPackage += delegate(object sender, SmartSchool.API.PlugIn.Export.ExportPackageEventArgs e)
            {
                List<StudentRecord> students = _AccessHelper.StudentHelper.GetStudents(e.List);
                _AccessHelper.StudentHelper.FillSemesterEntryScore(filterRepeat.Checked, students);
                foreach ( StudentRecord stu in students )
                {
                    Dictionary<int, Dictionary<int, List<SemesterEntryScoreInfo>>> semesterEntryScoreList = new Dictionary<int, Dictionary<int, List<SemesterEntryScoreInfo>>>();
                    foreach ( SemesterEntryScoreInfo var in stu.SemesterEntryScoreList )
                    {
                        if ( !semesterEntryScoreList.ContainsKey(var.SchoolYear) )
                            semesterEntryScoreList.Add(var.SchoolYear, new Dictionary<int, List<SemesterEntryScoreInfo>>());
                        if ( !semesterEntryScoreList[var.SchoolYear].ContainsKey(var.Semester) )
                            semesterEntryScoreList[var.SchoolYear].Add(var.Semester, new List<SemesterEntryScoreInfo>());
                        semesterEntryScoreList[var.SchoolYear][var.Semester].Add(var);
                    }
                    foreach ( int sy in semesterEntryScoreList.Keys )
                    {
                        foreach ( int se in semesterEntryScoreList[sy].Keys )
                        {
                            RowData row = new RowData();
                            row.ID = stu.StudentID;
                            row.Add("學年度", "" + sy);
                            row.Add("學期", "" + se);
                            foreach ( SemesterEntryScoreInfo var in semesterEntryScoreList[sy][se] )
                            {
                                if ( !row.ContainsKey("成績年級") )
                                    row.Add("成績年級", "" + var.GradeYear);
                                if ( e.ExportFields.Contains(var.Entry) )
                                {
                                    row.Add(var.Entry, "" + var.Score);
                                }
                            }
                            e.Items.Add(row);
                        }
                    }
                }
            };
        }
    }
}
