using System;
using System.Collections.Generic;
using System.Text;
//using SmartSchool.Customization.PlugIn.ImportExport;
using SmartSchool.Customization.Data;
using SmartSchool.Customization.Data.StudentExtension;
using SmartSchool.AccessControl;
using SmartSchool.API.PlugIn;
using FISCA.Data;
using System.Data;
using System.Xml.Linq;

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
            SmartSchool.API.PlugIn.VirtualCheckBox filterRepeat = new SmartSchool.API.PlugIn.VirtualCheckBox("自動略過重讀成績", true);
            wizard.Options.Add(filterRepeat);
            wizard.ExportableFields.AddRange("學年度", "學期", "成績年級", "學業", "體育", "國防通識", "健康與護理", "實習科目", "專業科目", "學業(原始)", "體育(原始)", "國防通識(原始)", "健康與護理(原始)", "實習科目(原始)", "專業科目(原始)", "德行", "學業成績班排名", "學業成績班排名母數", "學業成績科排名", "學業成績科排名母數", "學業成績校排名", "學業成績校排名母數", "學業成績排名類別1", "學業成績類1排名", "學業成績類1排名母數", "學業成績排名類別2", "學業成績類2排名", "學業成績類2排名母數");
            AccessHelper _AccessHelper = new AccessHelper();
            wizard.ExportPackage += delegate (object sender, SmartSchool.API.PlugIn.Export.ExportPackageEventArgs e)
            {
                List<StudentRecord> students = _AccessHelper.StudentHelper.GetStudents(e.List);

           

                Dictionary<string, List<StudSemsEntryRating>> StudSemsEntryRatingDict = new Dictionary<string, List<StudSemsEntryRating>>();
                StudSemsEntryRatingDict = GetStudSemsEntryRatingByStudentID(e.List);

                _AccessHelper.StudentHelper.FillSemesterEntryScore(filterRepeat.Checked, students);
                foreach (StudentRecord stu in students)
                {
                    Dictionary<int, Dictionary<int, List<SemesterEntryScoreInfo>>> semesterEntryScoreList = new Dictionary<int, Dictionary<int, List<SemesterEntryScoreInfo>>>();
                    foreach (SemesterEntryScoreInfo var in stu.SemesterEntryScoreList)
                    {
                        if (!semesterEntryScoreList.ContainsKey(var.SchoolYear))
                            semesterEntryScoreList.Add(var.SchoolYear, new Dictionary<int, List<SemesterEntryScoreInfo>>());
                        if (!semesterEntryScoreList[var.SchoolYear].ContainsKey(var.Semester))
                            semesterEntryScoreList[var.SchoolYear].Add(var.Semester, new List<SemesterEntryScoreInfo>());
                        semesterEntryScoreList[var.SchoolYear][var.Semester].Add(var);
                    }
                    foreach (int sy in semesterEntryScoreList.Keys)
                    {
                        foreach (int se in semesterEntryScoreList[sy].Keys)
                        {
                            RowData row = new RowData();
                            row.ID = stu.StudentID;
                            row.Add("學年度", "" + sy);
                            row.Add("學期", "" + se);
                            foreach (SemesterEntryScoreInfo var in semesterEntryScoreList[sy][se])
                            {
                                if (!row.ContainsKey("成績年級"))
                                    row.Add("成績年級", "" + var.GradeYear);
                                if (e.ExportFields.Contains(var.Entry))
                                {
                                    row.Add(var.Entry, "" + var.Score);
                                }
                            }

                            //處理學業成績排名資料

                            if (StudSemsEntryRatingDict.ContainsKey(stu.StudentID))
                            {
                                foreach (var record in StudSemsEntryRatingDict[stu.StudentID])
                                {
                                    if (record.SchoolYear == "" + sy && record.Semester == "" + se)
                                    {
                                        row.Add("學業成績班排名", "" + record.ClassRank);
                                        row.Add("學業成績班排名母數", "" + record.ClassCount);
                                        row.Add("學業成績科排名", "" + record.DeptRank);
                                        row.Add("學業成績科排名母數", "" + record.DeptCount);
                                        row.Add("學業成績校排名", "" + record.YearRank);
                                        row.Add("學業成績校排名母數", "" + record.YearCount);

                                        row.Add("學業成績排名類別1", "" + record.Group1);
                                        row.Add("學業成績類1排名", "" + record.Group1Rank);
                                        row.Add("學業成績類1排名母數", "" + record.Group1Count);
                                        row.Add("學業成績排名類別2", "" + record.Group2);
                                        row.Add("學業成績類2排名", "" + record.Group2Rank);
                                        row.Add("學業成績類2排名母數", "" + record.Group2Count);
                                    }

                                }                                
                            }


                            e.Items.Add(row);
                        }
                    }
                }
            };
        }


        public static Dictionary<string, List<StudSemsEntryRating>> GetStudSemsEntryRatingByStudentID(List<string> StudentIDList)
        {
            Dictionary<string, List<StudSemsEntryRating>> retValue = new Dictionary<string, List<StudSemsEntryRating>>();
            if (StudentIDList.Count > 0)
            {
                QueryHelper qh = new QueryHelper();
                string query = "select ref_student_id as sid,school_year,semester,grade_year,class_rating,dept_rating,year_rating,group_rating from sems_entry_score where ref_student_id in (" + string.Join(",", StudentIDList.ToArray()) + ") and entry_group=1  order by ref_student_id,grade_year,semester";

                DataTable dt = qh.Select(query);
                foreach (DataRow dr in dt.Rows)
                {
                    string sid = dr["sid"].ToString();

                    if (!retValue.ContainsKey(sid))
                        retValue.Add(sid, new List<StudSemsEntryRating>());

                    StudSemsEntryRating sser = new StudSemsEntryRating();
                    sser.StudentID = sid;
                    sser.SchoolYear = dr["school_year"].ToString();
                    sser.Semester = dr["semester"].ToString();
                    sser.GradeYear = dr["grade_year"].ToString();
                    sser.EntryName = "學業";
                    // Parse XML
                    try
                    {
                        string cStr = dr["class_rating"].ToString();
                        string dStr = dr["dept_rating"].ToString();
                        string yStr = dr["year_rating"].ToString();
                        string g1Str = dr["group_rating"].ToString();

                        // 班排
                        if (!string.IsNullOrEmpty(cStr))
                        {
                            XElement elmC = XElement.Parse(cStr);
                            foreach (XElement elm in elmC.Elements("Item"))
                            {
                                if (elm.Attribute("分項") != null && elm.Attribute("分項").Value == "學業")
                                {
                                    if (elm.Attribute("成績人數") != null && elm.Attribute("成績人數").Value != "")
                                    {
                                        int x;
                                        if (int.TryParse(elm.Attribute("成績人數").Value, out x))
                                            sser.ClassCount = x;

                                    }
                                    if (elm.Attribute("排名") != null && elm.Attribute("排名").Value != "")
                                    {
                                        int x;
                                        if (int.TryParse(elm.Attribute("排名").Value, out x))
                                            sser.ClassRank = x;
                                    }
                                }
                            }
                        }

                        // 科排
                        if (!string.IsNullOrEmpty(dStr))
                        {
                            XElement elmD = XElement.Parse(dStr);
                            foreach (XElement elm in elmD.Elements("Item"))
                            {
                                if (elm.Attribute("分項") != null && elm.Attribute("分項").Value == "學業")
                                {
                                    if (elm.Attribute("成績人數") != null && elm.Attribute("成績人數").Value != "")
                                    {
                                        int x;
                                        if (int.TryParse(elm.Attribute("成績人數").Value, out x))
                                            sser.DeptCount = x;
                                    }
                                    if (elm.Attribute("排名") != null && elm.Attribute("排名").Value != "")
                                    {
                                        int x;
                                        if (int.TryParse(elm.Attribute("排名").Value, out x))
                                            sser.DeptRank = x;
                                    }
                                }
                            }
                        }

                        // 校排
                        if (!string.IsNullOrEmpty(yStr))
                        {
                            XElement elmY = XElement.Parse(yStr);
                            foreach (XElement elm in elmY.Elements("Item"))
                            {
                                if (elm.Attribute("分項") != null && elm.Attribute("分項").Value == "學業")
                                {
                                    if (elm.Attribute("成績人數") != null && elm.Attribute("成績人數").Value != "")
                                    {
                                        int x;
                                        if (int.TryParse(elm.Attribute("成績人數").Value, out x))
                                            sser.YearCount = x;
                                    }

                                    if (elm.Attribute("排名") != null && elm.Attribute("排名").Value != "")
                                    {
                                        int x;
                                        if (int.TryParse(elm.Attribute("排名").Value, out x))
                                            sser.YearRank = x;
                                    }
                                }
                            }
                        }

                        // 類排  ，穎驊新增 類1 類2 、類別種類邏輯 
                        if (!string.IsNullOrEmpty(g1Str))
                        {
                            int groupCount = 1;

                            XElement elmG1 = XElement.Parse(g1Str);
                            foreach (XElement elmR in elmG1.Elements("Rating"))
                            {
                                //類別1
                                if (groupCount == 1)
                                {
                                    if (elmR.Attribute("類別") != null && elmR.Attribute("類別").Value != "")
                                    {
                                        sser.Group1 = elmR.Attribute("類別").Value;
                                    }

                                    foreach (XElement elm in elmR.Elements("Item"))
                                    {
                                        if (elm.Attribute("分項") != null && elm.Attribute("分項").Value == "學業")
                                        {
                                            if (elm.Attribute("成績人數") != null && elm.Attribute("成績人數").Value != "")
                                            {
                                                int x;
                                                if (int.TryParse(elm.Attribute("成績人數").Value, out x))
                                                    sser.Group1Count = x;
                                            }

                                            if (elm.Attribute("排名") != null && elm.Attribute("排名").Value != "")
                                            {
                                                int x;
                                                if (int.TryParse(elm.Attribute("排名").Value, out x))
                                                    sser.Group1Rank = x;
                                            }
                                        }
                                    }
                                }
                                else
                                {
                                    if (elmR.Attribute("類別") != null && elmR.Attribute("類別").Value != "")
                                    {
                                        sser.Group2 = elmR.Attribute("類別").Value;
                                    }

                                    foreach (XElement elm in elmR.Elements("Item"))
                                    {
                                        if (elm.Attribute("分項") != null && elm.Attribute("分項").Value == "學業")
                                        {
                                            if (elm.Attribute("成績人數") != null && elm.Attribute("成績人數").Value != "")
                                            {
                                                int x;
                                                if (int.TryParse(elm.Attribute("成績人數").Value, out x))
                                                    sser.Group2Count = x;
                                            }

                                            if (elm.Attribute("排名") != null && elm.Attribute("排名").Value != "")
                                            {
                                                int x;
                                                if (int.TryParse(elm.Attribute("排名").Value, out x))
                                                    sser.Group2Rank = x;
                                            }

                                            if (elm.Attribute("成績") != null && elm.Attribute("成績").Value != "")
                                            {
                                                sser.Score = "" + elm.Attribute("成績").Value;
                                            }
                                        }
                                    }

                                }

                                groupCount++;


                            }
                        }
                    }
                    catch (Exception ex) { }

                    retValue[sid].Add(sser);
                }
            }
            return retValue;
        }




    }
}
