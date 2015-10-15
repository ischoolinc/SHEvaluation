using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using FISCA.Presentation.Controls;
using K12.Data;
using FISCA.Data;
using Aspose.Words;
using System.IO;
using DevComponents.DotNetBar.Controls;
using System.Xml.Linq;

namespace SHStaticRank2.Data.StarUniversity
{
    public partial class StarUniversity : BaseForm
    {
        private List<TagConfigRecord> _TagConfigRecords = new List<TagConfigRecord>();
        private FISCA.UDT.AccessHelper _AccessHelper = new FISCA.UDT.AccessHelper();
        List<string> SubjectNameList = new List<string>();
        public Configure Configure { get; private set; }
        private Dictionary<string, int> _SpecialListViewItem = new Dictionary<string,int>();
        public StarUniversity()
        {
            InitializeComponent();
            buttonX1.Enabled = false;
            _TagConfigRecords = K12.Data.TagConfig.SelectByCategory(TagCategory.Student);
            List<Configure> lc = _AccessHelper.Select<Configure>("Name = '大學繁星'");
            this.Configure = (lc.Count >= 1)?lc[0]:new Configure() { Name = "大學繁星" };
            if (lc.Count >= 1)
                this.Configure.Decode();
            #region 填入類別

            List<string> prefix = new List<string>();
            List<string> tag = new List<string>();
            foreach (var item in _TagConfigRecords)
            {
                if (item.Prefix != "")
                {
                    if (!prefix.Contains(item.Prefix))
                        prefix.Add(item.Prefix);
                }
                else
                {
                    tag.Add(item.Name);
                }
            }
            // 不排名學生類別
            cboRankRilter.Items.Clear();
            // 類別排名1
            cboTagRank1.Items.Clear();
            // 類別排名2
            cboTagRank2.Items.Clear();

            cboRankRilter.Items.Add("");
            cboTagRank1.Items.Add("");
            cboTagRank2.Items.Add("");
            foreach (var s in prefix)
            {
                cboRankRilter.Items.Add("[" + s + "]");
                cboTagRank1.Items.Add("[" + s + "]");
                cboTagRank2.Items.Add("[" + s + "]");
            }
            foreach (var s in tag)
            {
                cboRankRilter.Items.Add(s);
                cboTagRank1.Items.Add(s);
                cboTagRank2.Items.Add(s);
            }

            if (this.Configure.CheckExportPDF)
                cbxExportPDF.Checked = true;
            else
                cbxExportWord.Checked = true;
            
            cbxIDNumber.Checked = this.Configure.CheckUseIDNumber;
            #endregion
            buttonX1.Enabled = true;
        }
        private void buttonX1_Click(object sender, EventArgs e)
        {
            FISCA.LogAgent.ApplicationLog.Log("成績", "計算", "計算大學繁星多學期成績單。");
            #region 固定的setting
            List<string> studIDList = new List<string>();
            QueryHelper qh = new QueryHelper();

            string strSQ = @"SELECT DISTINCT tmp.subject
FROM xpath_table( 'id',
				'''<root>''||score_info||''</root>''',
				'sems_subj_score',
				'/root/SemesterSubjectScoreInfo/Subject/@科目',
				'ref_student_id IN ( select student.id from student 
									INNER JOIN class ON student.ref_class_id = class.id 
									WHERE student.status=1 AND class.grade_year = 3 )'
				) 
AS tmp(id int, subject varchar(200))";
            DataTable dt = qh.Select(strSQ);
            foreach (DataRow dr in dt.Rows)
                SubjectNameList.Add(dr[0].ToString());
            SubjectNameList.Sort(new StringComparer("國文"
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
            Configure.CalcGradeYear1 = false;
            Configure.CalcGradeYear2 = false;
            Configure.CalcGradeYear3 = true; //三年級
            Configure.CalcGradeYear4 = false;
            Configure.DoNotSaveIt = true;

            if (cbxScoreType.Text == "擇優成績")
            {
                Configure.use原始成績 = true;//原始成績
                Configure.use補考成績 = true;
                Configure.use重修成績 = true;
                Configure.use手動調整成績 = true;
                Configure.use學年調整成績 = true;
                Configure.RankFilterUseScoreList.Add("原始成績");
                Configure.RankFilterUseScoreList.Add("補考成績");
                Configure.RankFilterUseScoreList.Add("重修成績");
                Configure.RankFilterUseScoreList.Add("手動調整成績");
                Configure.RankFilterUseScoreList.Add("學年調整成績");
            }
            else
            {
                Configure.use原始成績 = true;//原始成績
                Configure.use補考成績 = false;
                Configure.use重修成績 = false;
                Configure.use手動調整成績 = false;
                Configure.use學年調整成績 = false;
                Configure.RankFilterUseScoreList.Add("原始成績");
            }
            
            Configure.計算學業成績排名 = true;
            Configure.WithCalSemesterScoreRank = true;

            foreach (string SubjectName in SubjectNameList)//所有科目
            {
                Configure.useSubjectPrintList.Add(SubjectName);
                Configure.useSubjecOrder1List.Add(SubjectName);
                Configure.useSubjecOrder2List.Add(SubjectName);
            }
            Configure.Name = "大學繁星";
            Configure.SortGradeYear = "三年級";

            if (chk4Grade.Checked)
            {
                Configure.useGradeSemesterList.Add("11");
                Configure.useGradeSemesterList.Add("12");
                Configure.useGradeSemesterList.Add("21");
                Configure.useGradeSemesterList.Add("22");
                Configure.RankFilterGradeSemeterList.Add("一上");
                Configure.RankFilterGradeSemeterList.Add("一下");
                Configure.RankFilterGradeSemeterList.Add("二上");
                Configure.RankFilterGradeSemeterList.Add("二下");            
            }            
            #endregion
            
            if (chk5Grade.Checked)
            {
                Configure.useGradeSemesterList.Add("11");
                Configure.useGradeSemesterList.Add("12");
                Configure.useGradeSemesterList.Add("21");
                Configure.useGradeSemesterList.Add("22");
                Configure.RankFilterGradeSemeterList.Add("一上");
                Configure.RankFilterGradeSemeterList.Add("一下");
                Configure.RankFilterGradeSemeterList.Add("二上");
                Configure.RankFilterGradeSemeterList.Add("二下");
                Configure.useGradeSemesterList.Add("31");
                Configure.RankFilterGradeSemeterList.Add("三上");
            }

            if (chk6Grade.Checked)
            {
                Configure.useGradeSemesterList.Add("32");
                Configure.RankFilterGradeSemeterList.Add("三下");
            }

            Configure.NotRankTag = cboRankRilter.Text;
            Configure.Rank1Tag = cboTagRank1.Text;
            Configure.Rank2Tag = cboTagRank2.Text;
            Configure.RankFilterTagName = cboRankRilter.Text;


            Document docTemp = null;

            if (chk4Grade.Checked)
            {
                // 檢查沒有樣版
                if (this.Configure.Template2 == null)
                    Configure.Template2 = new Document(new MemoryStream(Properties.Resources.多學期成績單_大學4學期));             

                // 檢查合併欄位
                if(this.Configure.Template2.MailMerge.GetFieldNames().Count()==0)
                    docTemp = new Document(new MemoryStream(Properties.Resources.多學期成績單_大學4學期));
             
                 docTemp = this.Configure.Template2.Clone();                    
             
            }

            if (chk5Grade.Checked)
            {
                // 檢查沒有樣版
                if (this.Configure.Template1 == null)
                    Configure.Template1 = new Document(new MemoryStream(Properties.Resources.多學期成績單_5學期));

                // 檢查合併欄位
                if(Configure.Template1.MailMerge.GetFieldNames().Count()==0)
                        docTemp = new Document(new MemoryStream(Properties.Resources.多學期成績單_5學期));
                
                 docTemp = this.Configure.Template1.Clone();                    
                
            }

            if (chk6Grade.Checked)
            {
                // 檢查沒有樣版
                if (this.Configure.Template3 == null)
                    Configure.Template3 = new Document(new MemoryStream(Properties.Resources.多學期成績單_第6學期));

                // 檢查合併欄位
                if(Configure.Template3.MailMerge.GetFieldNames().Count()==0)
                        docTemp = new Document(new MemoryStream(Properties.Resources.多學期成績單_第6學期));
                
                docTemp = this.Configure.Template3.Clone();                        
                
            }

            this.Configure.Template = docTemp.Clone();

            Configure.CheckExportPDF = cbxExportPDF.Checked;
            if (cbxIDNumber.Checked)
                Configure.CheckUseIDNumber = true;

            if (cbxSeNo.Checked)
                Configure.CheckUseIDNumber = false;

            Configure.CheckExportStudent = true;

            // 儲存設定
            SetUIConfig();

            Configure.Save();
            
            DialogResult = System.Windows.Forms.DialogResult.OK;            
            this.Close();
        }
        private void DownloadDefaultTemplate()
        {
            if (this.Configure == null) return;
            #region 儲存檔案
            string inputReportName = "多學期成績單樣板(" + this.Configure.Name + ")";
            string reportName = inputReportName;

            string path = Path.Combine(System.Windows.Forms.Application.StartupPath, "Reports");
            if (!Directory.Exists(path))
                Directory.CreateDirectory(path);
            path = Path.Combine(path, reportName + ".doc");

            if (File.Exists(path))
            {
                int i = 1;
                while (true)
                {
                    string newPath = Path.GetDirectoryName(path) + "\\" + Path.GetFileNameWithoutExtension(path) + (i++) + Path.GetExtension(path);
                    if (!File.Exists(newPath))
                    {
                        path = newPath;
                        break;
                    }
                }
            }



            Document docTemp = null;

            try
            {            

            //計算檔案大小
            MemoryStream ms = new MemoryStream();

            // 四學期
            if (chk4Grade.Checked)
            {
                if (this.Configure.Template2 == null)
                    docTemp = new Document(new MemoryStream(Properties.Resources.多學期成績單_大學4學期));
                else
                {
                    if(Configure.Template2.MailMerge.GetFieldNames().Count()==0)
                        docTemp = new Document(new MemoryStream(Properties.Resources.多學期成績單_大學4學期));
                    else
                        docTemp = this.Configure.Template2.Clone();                        
                }                    
            }

            // 五學期
            if (chk5Grade.Checked)
            {
                if(this.Configure.Template1==null)
                    docTemp = new Document(new MemoryStream(Properties.Resources.多學期成績單_5學期));
                else
                {
                    if(Configure.Template1.MailMerge.GetFieldNames().Count()==0)
                        docTemp = new Document(new MemoryStream(Properties.Resources.多學期成績單_5學期));
                    else
                        docTemp = this.Configure.Template1.Clone();                       
                }
                    
            }                

            // 第六學期
            if (chk6Grade.Checked)
            {             
                if (this.Configure.Template3 == null)
                    docTemp = new Document(new MemoryStream(Properties.Resources.多學期成績單_第6學期));
                else
                {
                    if(Configure.Template3.MailMerge.GetFieldNames().Count()==0)
                        docTemp = new Document(new MemoryStream(Properties.Resources.多學期成績單_第6學期));
                    else
                        docTemp = this.Configure.Template3.Clone();                        
                }
            }                                    

                  docTemp.Save(path, SaveFormat.Doc);

                System.Diagnostics.Process.Start(path);
            }
            catch
            {
                System.Windows.Forms.SaveFileDialog sd = new System.Windows.Forms.SaveFileDialog();
                sd.Title = "另存新檔";
                sd.FileName = reportName + ".doc";
                sd.Filter = "Word檔案 (*.doc)|*.doc|所有檔案 (*.*)|*.*";
                if (sd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    try
                    {
                        docTemp.Save(sd.FileName, Aspose.Words.SaveFormat.Doc);
                    }
                    catch
                    {
                        FISCA.Presentation.Controls.MsgBox.Show("指定路徑無法存取。", "建立檔案失敗", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                        return;
                    }
                }
            }
            #endregion
        }
        private void UploadUserDefTemplate()
        {
            if (Configure == null) return;
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Title = "上傳樣板";
            dialog.Filter = "Word檔案 (*.doc)|*.doc|所有檔案 (*.*)|*.*";
            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                Document uploadDocTemp = null;

                try
                {
                    uploadDocTemp = new Aspose.Words.Document(dialog.FileName);
                    // 計算檔案大小
                    MemoryStream ms = new MemoryStream();
                    uploadDocTemp.Save(ms, SaveFormat.Pdf);
                    byte[] bb = ms.ToArray();
                    uploadDocTemp.Save(ms, SaveFormat.Doc);
                    byte[] bbw = ms.ToArray();
                    double bbSize = (bb.Count() / 1024);
                    double bbwSize = (bbw.Count() / 1024);
                    if (bbSize >= 200)
                        MsgBox.Show("上傳範本檔案 "+bbwSize+" K ，產生PDF檔案大小約 "+bbSize+"K 超過200K 無法上傳，請調整範本大小再次上傳。");
                    else
                    {
                        List<string> fields = new List<string>(uploadDocTemp.MailMerge.GetFieldNames());
                        this.Configure.SubjectLimit = 0;
                        while (fields.Contains("科目名稱" + (this.Configure.SubjectLimit + 1)))
                        {
                            this.Configure.SubjectLimit++;
                        }

                        if (chk4Grade.Checked)
                            this.Configure.Template2 = uploadDocTemp.Clone();

                        if (chk5Grade.Checked)
                            this.Configure.Template1 = uploadDocTemp.Clone();

                        if (chk6Grade.Checked)
                            this.Configure.Template3 = uploadDocTemp.Clone();                        

                        Configure.Encode();

                        Configure.Save();
                        MsgBox.Show("上傳完成.");
                    }                    
                }
                catch
                {
                    MessageBox.Show("樣板開啟失敗");
                }
            }
        }

        private void lnkDownload_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            DownloadDefaultTemplate();
        }
        private void lnkUpload_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            buttonX1.Enabled = false;
            UploadUserDefTemplate();
            buttonX1.Enabled = true;
        }
        private void lblMappingTemp_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Document doc = CalcMutilSemeSubjectRank.getMergeTable();

            SaveFileDialog saveDialog = new SaveFileDialog();
            saveDialog.Filter = "Word (*.doc)|*.doc";
            saveDialog.FileName = "合併欄位總表";
            if (saveDialog.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    doc.Save(saveDialog.FileName); 
                }
                catch (Exception ex)
                {
                    FISCA.Presentation.Controls.MsgBox.Show("儲存失敗。" + ex.Message);
                    return;
                }

                try
                {
                    System.Diagnostics.Process.Start(saveDialog.FileName);
                }
                catch (Exception ex)
                {
                    FISCA.Presentation.Controls.MsgBox.Show("開啟失敗。" + ex.Message);
                    return;
                }
            }
        }
        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void StarUniversity_Load(object sender, EventArgs e)
        {
            cbxScoreType.Items.Add("擇優成績");
            cbxScoreType.Items.Add("原始成績");
            cbxScoreType.Text = "擇優成績";
            cbxScoreType.DropDownStyle = ComboBoxStyle.DropDownList;

            // 當三個樣版都空白時，將預設樣版放入(支援相容)
            if (Configure.Template1 == null && Configure.Template2 == null && Configure.Template3 == null)
            {
                if(Configure.Template !=null)
                {
                    Configure.Template1 = Configure.Template.Clone();
                    Configure.Encode();
                    Configure.Save();
                }                
            }

            // 載入設定
            GetUIConfig();
        }

        /// <summary>
        /// 載入畫面設定
        /// </summary>
        /// <param name="conf"></param>
        private void GetUIConfig()
        {
            try
            {
                XElement elmRoot = XElement.Parse(Configure.UIConfigString);
                foreach (XElement elm in elmRoot.Elements("Item"))
                {
                    string name = "";
                    string value="";
                    if (elm.Attribute("Name") != null)
                    {
                        name = elm.Attribute("Name").Value;
                        value=elm.Attribute("Value").Value;
                    }
                    if (name == "chk4Grade")
                        chk4Grade.Checked = ParseBool(value);

                    if (name == "chk5Grade")
                        chk5Grade.Checked = ParseBool(value);

                    if (name == "chk6Grade")
                        chk6Grade.Checked = ParseBool(value);

                    if (name == "cboRankRilter")
                        cboRankRilter.Text = value;

                    if (name == "cboTagRank1")
                        cboTagRank1.Text = value;

                    if (name == "cboTagRank2")
                        cboTagRank2.Text = value;

                    if (name == "cbxExportPDF")
                        cbxExportPDF.Checked = ParseBool(value);

                    if (name == "cbxExportWord")
                        cbxExportWord.Checked = ParseBool(value);

                    if (name == "cbxSeNo")
                        cbxSeNo.Checked = ParseBool(value);

                    if (name == "cbxIDNumber")
                        cbxIDNumber.Checked = ParseBool(value);

                    if (name == "cbxScoreType")
                        cbxScoreType.Text = value;
                }
                
            }
            catch (Exception ex)
            {
                //MessageBox.Show("解析設定檔失敗," + ex.Message);
            }
            
        }

        private void SetUIConfig()
        {
            try
            {
                XElement elmRoot = new XElement("Items");
                elmRoot.Add(AddElement("chk4Grade", chk4Grade.Checked.ToString()));
                elmRoot.Add(AddElement("chk5Grade", chk5Grade.Checked.ToString()));
                elmRoot.Add(AddElement("chk6Grade", chk6Grade.Checked.ToString()));
                elmRoot.Add(AddElement("cboRankRilter", cboRankRilter.Text));
                elmRoot.Add(AddElement("cboTagRank1", cboTagRank1.Text));
                elmRoot.Add(AddElement("cboTagRank2", cboTagRank2.Text));
                elmRoot.Add(AddElement("cbxExportPDF", cbxExportPDF.Checked.ToString()));
                elmRoot.Add(AddElement("cbxExportWord", cbxExportWord.Checked.ToString()));
                elmRoot.Add(AddElement("cbxSeNo", cbxSeNo.Checked.ToString()));
                elmRoot.Add(AddElement("cbxIDNumber", cbxIDNumber.Checked.ToString()));
                elmRoot.Add(AddElement("cbxScoreType", cbxScoreType.Text));

                Configure.UIConfigString = elmRoot.ToString();
            }
            catch (Exception ex)
            {
                MessageBox.Show("寫入設定檔失敗," + ex.Message);
            }
        
        }

        private XElement AddElement(string Name, string Value)
        {
            XElement elm = new XElement("Item");
            elm.SetAttributeValue("Name", Name);
            elm.SetAttributeValue("Value", Value);
            return elm;
        }

        private bool ParseBool(string str)
        {
            bool retVal = false;
            bool.TryParse(str, out retVal);
            return retVal;
        }
    }
}
