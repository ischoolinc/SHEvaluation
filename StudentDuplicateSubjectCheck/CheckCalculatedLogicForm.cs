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
using System.Xml.Linq;
using FISCA.Data;
using FISCA.Permission;
using K12.Data;
using FISCA.Authentication;
using FISCA.LogAgent;
using System.IO;
using System.Xml;

namespace StudentDuplicateSubjectCheck
{
    public partial class CheckCalculatedLogicForm : BaseForm
    {
        string school_year = School.DefaultSchoolYear;
        string semster = School.DefaultSemester;

        List<SCAttendRecord> _scaDuplicateList;

        //紀錄 每一筆 SCAttendRecord extension 裏頭的 XML 格式 使用，之所以會這樣做是因為
        // extension 在本設定重覆科目級別計算方式功能之前
        // 是用來儲存老師的平時評量 小考資料(Web)
        // 其格式如下(GradeBook的部分，DuplicatedLevelSubjectCalRule為本次所新增):
        // 要判斷是否有老師的設定資料，保留資料 一起更新
        // <Extensions>
        //<Extension Name = "GradeBook" >
        //< Exam ExamID="3" Score=""/>
        //<Exam ExamID = "4" Score="87.5">
        //<Item Score = "90" SubExamID="43780E9A-8DC4-18AA-15D2-020B5EF23F67"/>
        //<Item Score = "85" SubExamID="7B075D5D-C479-93BB-CBF8-020B60DB003B"/>
        //</Exam>
        //</Extension>
        //<Extension Name = "DuplicatedLevelSubjectCalRule" >
        //< Rule > 視為一般修課 </ Rule >
        //</ Extension >
        //</ Extensions >
        Dictionary<string, string> xmlRecordDict = new Dictionary<string, string>(); // <scattendID,xml>


        public CheckCalculatedLogicForm(List<SCAttendRecord> scaDuplicateList)
        {
            InitializeComponent();

            _scaDuplicateList = scaDuplicateList;

            ////計算方式設定 的選項
            //Column8.Items.Add("重修(寫回原學期)");
            //Column8.Items.Add("重讀(擇優採計成績)");
            //Column8.Items.Add("抵免");
            //Column8.Items.Add("視為一般修課");
            //Column8.DropDownStyle = ComboBoxStyle.DropDownList; // 讓使用者不要亂輸入資料

            comboBoxEx1.Items.Clear();
            comboBoxEx1.Items.Add("視為一般修課");
            comboBoxEx1.Items.Add("重修成績");
            comboBoxEx1.Items.Add("再次修習");
            comboBoxEx1.Items.Add("補修成績");

            Column8.Items.Clear();
            Column8.Items.Add("視為一般修課");
            Column8.Items.Add("重修成績");
            Column8.Items.Add("再次修習");
            Column8.Items.Add("補修成績");
            //Column8.DropDownStyle = ComboBoxStyle.DropDownList;  // 為支援舊資料先註解


            labelX3.Text = school_year + "學年度" + "第" + semster + "學期";

            FillUI();
        }

        private void FillUI()
        {
            foreach (SCAttendRecord record in _scaDuplicateList)
            {
                DataGridViewRow row = new DataGridViewRow();
                row.CreateCells(dataGridViewX1);

                row.Tag = record.ID; // 每一條的 Tag 為 scattend 的 ID 方便對應上傳

                row.Cells[0].Value = record.CourseName;
                row.Cells[1].Value = record.SubjectName;
                row.Cells[2].Value = record.SubjectLevel;
                row.Cells[3].Value = record.ClassName;  
                row.Cells[4].Value = record.StudentNumber;
                row.Cells[5].Value = record.SeatNo;                
                row.Cells[6].Value = record.Name;
                if (record.ScoreSource.Contains("學期成績"))
                {
                    row.Cells[7].Value = "√";
                }
                if (record.ScoreSource.Contains("封存成績"))
                {
                    row.Cells[8].Value = "√";
                }
                
                // 學生姓名
                //                row.Cells[7].Value = ... // 成績來源學期成績（欄位未在程式片段中顯示詳細內容）
                //row.Cells[8].Value = ... // 成績來源封存成績（同上）
                //row.Cells[9].Value = ... // 成績認定方式

                //row.Cells[8].Value = "視為一般修課"; // 如果有抓到重覆，先通預設為一般修課  ， 2018/5/24取消設定

                string xmlStr = "<root>" + record.Extensions + "</root>";

                XElement elmRoot = XElement.Parse(xmlStr);

                if (elmRoot != null)
                {
                    if (elmRoot.Element("Extensions") != null)
                    {
                        foreach (XElement ex in elmRoot.Element("Extensions").Elements("Extension"))
                        {
                            if (ex.Attribute("Name").Value == "DuplicatedLevelSubjectCalRule")
                            {
                                row.Cells[9].Value = ex.Element("Rule").Value;
                            }
                        }
                    }
                }
                xmlRecordDict.Add(record.ID, record.Extensions);
                dataGridViewX1.Rows.Add(row);
            }
        }

        // 儲存
        private void buttonX1_Click(object sender, EventArgs e)
        {
            string _actor = DSAServices.UserAccount; ;
            string _client_info = ClientInfo.GetCurrentClientInfo().OutputResult().OuterXml;

            // 兜資料
            List<string> dataList = new List<string>();
            foreach (SCAttendRecord record in _scaDuplicateList)
            {

                string data = string.Format(@"
                SELECT
                    {0}::BIGINT AS sc_attend_id                    
                    , {1}::BIGINT AS ref_student_id   
                    , '{2}' ::TEXT AS school_year
                    , '{3}' ::TEXT AS semester
                    , '{4}' ::TEXT AS course_name
                    , '{5}' ::TEXT AS subject_name
                    , '{6}' ::TEXT AS subject_level
                    , '{7}' ::TEXT AS student_number                    
                    , '{8}' ::TEXT AS student_name                 
                    , '{9}'::TEXT AS extensions
                    , '{10}'::TEXT AS extensions_string
                ", record.ID, record.RefStudentID, school_year, semster, record.CourseName, record.SubjectName, record.SubjectLevel, record.StudentNumber, record.Name, GetExtensions(record.ID),GetExtensions_string(record.ID));

                dataList.Add(data);
            }

            string Data = string.Join(" UNION ALL", dataList);


            string sql = string.Format(@" WITH data_row AS( 
			 {0}),update_data AS(
	UPDATE sc_attend
	SET
		extensions = data_row.extensions
	FROM
		data_row
	WHERE
		sc_attend.id = data_row.sc_attend_id
	RETURNING sc_attend.*
) 
-- 新增 LOG
INSERT INTO log(
	actor
	, action_type
	, action
	, target_category
	, target_id
	, server_time
	, client_info
	, action_by
	, description
)
SELECT 
	'{1}'::TEXT AS actor
	, 'Record' AS action_type
	, '設定重覆級別科目學期成績計算方式' AS action
	, 'student'::TEXT AS target_category
	, data_row.ref_student_id AS target_id
	, now() AS server_time
	, '{2}' AS client_info
	, '成績計算'AS action_by   
	, '設定學生「'|| data_row.student_name || '」 學號 「'|| data_row.student_number || '」 「'|| data_row.school_year || '」 學年  第「'|| data_row.semester || '」 學期 課程「' || data_row.course_name || '」 科目「' || data_row.subject_name || '」 級別 「'|| data_row. subject_level || '」,計算方式為「'|| data_row. extensions_string || '」' AS description 
FROM
	data_row
                ", Data, _actor, _client_info);

            UpdateHelper uh = new UpdateHelper();

            //執行sql
            uh.Execute(sql);

            MsgBox.Show("儲存設定成功，可至系統歷程檢查更動。");
        }

        // 取消
        private void buttonX2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private string GetExtensions(string scattendID)
        {
            string Extension = "";

            // 沒東西，做新的給它
            if (xmlRecordDict[scattendID] == "")
            {
                foreach (DataGridViewRow r in dataGridViewX1.Rows)
                {
                    if ("" + r.Tag == scattendID)
                    {
                        XmlDocument doc = new XmlDocument();

                        XmlElement element_Rule = doc.CreateElement(string.Empty, "Rule", string.Empty);

                        element_Rule.InnerXml = "" + r.Cells[9].Value; ;

                        XmlElement element_Extension = doc.CreateElement(string.Empty, "Extension", string.Empty);

                        element_Extension.SetAttribute("Name", "DuplicatedLevelSubjectCalRule");

                        XmlElement element_Extensions = doc.CreateElement(string.Empty, "Extensions", string.Empty);

                        element_Extension.AppendChild(element_Rule);

                        element_Extensions.AppendChild(element_Extension);

                        doc.AppendChild(element_Extensions);

                        Extension = doc.OuterXml;
                    }
                }

            }
            // 原本有東西 舊的保留 再另外加
            else
            {
                foreach (DataGridViewRow r in dataGridViewX1.Rows)
                {
                    if ("" + r.Tag == scattendID)
                    {
                        XmlDocument doc = new XmlDocument();

                        doc.LoadXml(xmlRecordDict[scattendID]);

                        XmlNodeList nodeList = doc.GetElementsByTagName("Rule");

                        // 第一次檢查 是沒有 Rule element 的
                        if (nodeList.Count == 0)
                        {
                            XmlElement element_Rule = doc.CreateElement(string.Empty, "Rule", string.Empty);

                            element_Rule.InnerXml = "" + r.Cells[9].Value; 

                            XmlElement element_Extension = doc.CreateElement(string.Empty, "Extension", string.Empty);

                            element_Extension.SetAttribute("Name", "DuplicatedLevelSubjectCalRule");

                            element_Extension.AppendChild(element_Rule);

                            doc.DocumentElement.AppendChild(element_Extension);
                        }
                        else // 如果有的話 直接更新
                        {
                            nodeList[0].InnerXml = "" + r.Cells[9].Value;
                        }

                        Extension = doc.OuterXml;

                    }
                }
            }
            return Extension;
        }


        private string GetExtensions_string(string scattendID)
        {
            string Extension_string = "";

            foreach (DataGridViewRow r in dataGridViewX1.Rows)
            {
                if ("" + r.Tag == scattendID)
                {
                    Extension_string = ""+ r.Cells[9].Value;
                }
            }
            return Extension_string;
        }


        // 設定 計算方式
        private void comboBoxEx1_SelectedIndexChanged(object sender, EventArgs e)
        {
            foreach (DataGridViewRow r in dataGridViewX1.Rows)
            {
                if (r.Selected)
                {
                    r.Cells[9].Value = comboBoxEx1.SelectedItem;
                }
            }
        }

        private void buttonX3_Click(object sender, EventArgs e)
        {
            List<string> StudentIDList = new List<string>();

            foreach (DataGridViewRow r in dataGridViewX1.Rows)
            {
                if (r.Selected)
                {
                    SCAttendRecord record = _scaDuplicateList.Find(x => x.ID == ""+r.Tag);

                    if (!StudentIDList.Contains(record.RefStudentID))
                    {
                        StudentIDList.Add(record.RefStudentID);

                    }
                }
            }

            K12.Presentation.NLDPanels.Student.AddToTemp(StudentIDList);

            MsgBox.Show("已於[學生待處理]加入" + StudentIDList.Count + "名學生");
        }

        //匯出
        private void buttonX4_Click(object sender, EventArgs e)
        {
            #region 匯出
            if (saveFileDialog1.ShowDialog() != DialogResult.OK) return;

            DataGridViewExport export = new DataGridViewExport(dataGridViewX1);
            export.Save(saveFileDialog1.FileName);

            if (new CompleteForm().ShowDialog() == DialogResult.Yes)
                System.Diagnostics.Process.Start(saveFileDialog1.FileName);
            #endregion
        }
    }
}




