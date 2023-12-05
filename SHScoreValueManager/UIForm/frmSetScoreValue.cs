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
using System.Web.Script.Serialization;
using SHScoreValueManager.DAO;

namespace SHScoreValueManager.UIForm
{
    public partial class frmSetScoreValue : BaseForm
    {
        K12.Data.Configuration.ConfigData cd = K12.Data.School.Configuration["評量成績缺考設定"];
        // 設定 XML
        XElement SettingXML = null;

        // 設定資料內容
        List<ScoreSettingConfig> ScoreSettingConfigList;


        public frmSetScoreValue()
        {
            InitializeComponent();
            ScoreSettingConfigList = new List<ScoreSettingConfig>();
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            btnSave.Enabled = false;
            if (CheckData())
            {
                SaveData();
                MsgBox.Show("儲存成功");
                this.Close();
            }
            btnSave.Enabled = true;
        }

        private void frmSetScoreValue_Load(object sender, EventArgs e)
        {
            // 載入欄位
            LoadDataGridViewColumns();

            //載入欄位與預設值
            LoadData();
            LoadDataToDataGridView();
        }

        // 載入資料至畫面DataGridView
        private void LoadDataToDataGridView()
        {
            dgData.Rows.Clear();
            if (SettingXML != null)
            {
                foreach (ScoreSettingConfig ss in ScoreSettingConfigList)
                {
                    int rowIdx = dgData.Rows.Add();
                    dgData.Rows[rowIdx].Tag = ss;
                    dgData.Rows[rowIdx].Cells["缺考原因"].Value = ss.ReportValue;
                    dgData.Rows[rowIdx].Cells["輸入內容"].Value = ss.UseText; dgData.Rows[rowIdx].Cells["分數認定"].Value = ss.ScoreType;
                }
            }
        }

        // 取得XML屬性
        private string GetXMLAttribute(XElement elm, string attrName)
        {
            if (elm.Attribute(attrName) != null)
                return elm.Attribute(attrName).Value;
            else
                return "";
        }

        // 取得XMLElm
        private string GetXMLElement(XElement elm, string elmName)
        {
            if (elm.Element(elmName) != null)
                return elm.Element(elmName).Value;
            else
                return "";
        }

        // 載入畫面欄位設定
        private void LoadDataGridViewColumns()
        {
            dgData.Columns.Clear();
            try
            {
                string textColumnStrig = @"
		[
		{
			""HeaderText"": ""缺考原因"",
			""Name"": ""缺考原因"",
			""Width"": 160,
			""ReadOnly"": false
		},
		{
			""HeaderText"": ""輸入內容"",
			""Name"": ""輸入內容"",
			""Width"": 90,
			""ReadOnly"": false
		}		
		]   
";

                JavaScriptSerializer serializer = new JavaScriptSerializer();
                List<DataGridViewTextBoxColumnInfo> jsonObjArray = serializer.Deserialize<List<DataGridViewTextBoxColumnInfo>>(textColumnStrig);
                foreach (DataGridViewTextBoxColumnInfo jObj in jsonObjArray)
                {
                    DataGridViewTextBoxColumn dgt = new DataGridViewTextBoxColumn();
                    dgt.Name = jObj.Name;
                    dgt.Width = jObj.Width;
                    dgt.HeaderText = jObj.HeaderText;
                    dgt.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    dgt.ReadOnly = jObj.ReadOnly;
                    dgData.Columns.Add(dgt);
                }

                // 分數認定
                Dictionary<string, string> ScoreTypeDict = new Dictionary<string, string>();
                ScoreTypeDict.Add("0分", "-1");
                ScoreTypeDict.Add("免試", "-2");

                DataGridViewComboBoxColumn cbScoreType = new DataGridViewComboBoxColumn();
                cbScoreType.Name = "分數認定";
                cbScoreType.Width = 90;
                cbScoreType.DropDownWidth = 90;
                cbScoreType.HeaderText = "分數認定";
                cbScoreType.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;

                DataTable dtScoreType = new DataTable();
                dtScoreType.Columns.Add("VALUE");
                dtScoreType.Columns.Add("ITEM");

                foreach (string str in ScoreTypeDict.Keys)
                {
                    DataRow dr = dtScoreType.NewRow();
                    dr["VALUE"] = str;
                    dr["ITEM"] = str;
                    dtScoreType.Rows.Add(dr);
                }
                cbScoreType.DataSource = dtScoreType;
                cbScoreType.DisplayMember = "ITEM";
                cbScoreType.ValueMember = "VALUE";

                dgData.Columns.Add(cbScoreType);

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

        }

        // 載入設定資料
        private void LoadData()
        {
            // 讀取設定資料
            if (cd.Contains("Settings"))
            {
                try
                {
                    SettingXML = XElement.Parse(cd["Settings"]);
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Settings:", ex.Message);
                }
            }
            if (SettingXML == null)
            {
                SettingXML = new XElement("Settings");
            }

            // 解析上XML
            ScoreSettingConfigList.Clear();
            foreach (XElement elm in SettingXML.Elements("Setting"))
            {
                ScoreSettingConfig ss = new ScoreSettingConfig();
                ss.UseText = GetXMLElement(elm, "UseText");
                ss.UseValue = GetXMLElement(elm, "UseValue");
                ss.ScoreType = GetXMLElement(elm, "ScoreType");
                ss.ReportValue = GetXMLElement(elm, "ReportValue");
                ScoreSettingConfigList.Add(ss);
            }

        }

        // 儲存設定資料
        private void SaveData()
        {
            try
            {
                // 重新讀取資料轉成物件
                XElement elmRoot = new XElement("Settings");
                List<string> logStr = new List<string>();
                StringBuilder sb = new StringBuilder();
                sb.AppendLine("== 缺考設定 ==");
                foreach (DataGridViewRow row in dgData.Rows)
                {
                    if (row.IsNewRow)
                        continue;

                    logStr.Clear();
                    XElement elm = new XElement("Setting");
                    string ScoreType = "" + row.Cells["分數認定"].Value;
                    string UseText = "" + row.Cells["輸入內容"].Value;
                    string ReportValue = "" + row.Cells["缺考原因"].Value;
                    elm.SetElementValue("UseText", UseText);
                    elm.SetElementValue("ScoreType", ScoreType);
                    elm.SetElementValue("ReportValue", ReportValue);

                    // log
                    logStr.Add("缺考原因：" + ReportValue);
                    logStr.Add("輸入內容：" + UseText);
                    logStr.Add("分數認定：" + ScoreType);
                    sb.AppendLine(string.Join(",", logStr.ToArray()));

                    if (ScoreType == "0分")
                        elm.SetElementValue("UseValue", "-1");
                    else if (ScoreType == "免試")
                        elm.SetElementValue("UseValue", "-2");
                    else elm.SetElementValue("UseValue", "0");

                    elmRoot.Add(elm);
                }
                cd["Settings"] = elmRoot.ToString();
                cd.Save();

                // 紀錄log
                FISCA.LogAgent.ApplicationLog.Log("缺考設定", "儲存", sb.ToString());
            }
            catch (Exception ex)
            {
                MsgBox.Show("儲存過程發生錯誤," + ex.Message);
                return;
            }
        }

        // 檢查資料
        private bool CheckData()
        {
            bool value = true;
            int rowCount = 0;
            List<string> chkUseTextSame = new List<string>();

            // 檢查 DataGridView 資料
            foreach (DataGridViewRow row in dgData.Rows)
            {
                if (row.IsNewRow)
                    continue;

                if ("" + row.Cells["缺考原因"].Value == "")
                {
                    row.Cells["缺考原因"].ErrorText = "必填";
                    value = false;
                }
                else
                    row.Cells["缺考原因"].ErrorText = "";

                string UseText = "" + row.Cells["輸入內容"].Value;
                if (UseText == "")
                {
                    row.Cells["輸入內容"].ErrorText = "必填";
                    value = false;
                }
                else
                {
                    row.Cells["輸入內容"].ErrorText = "";
                    // 檢查輸入內容是否重複
                    if (chkUseTextSame.Contains(UseText))
                    {
                        row.Cells["輸入內容"].ErrorText = UseText + " 重複!";
                        value = false;
                    }
                    else
                    {
                        chkUseTextSame.Add(UseText);
                    }

                }


                if ("" + row.Cells["分數認定"].Value == "")
                {
                    row.Cells["分數認定"].ErrorText = "必填";
                    value = false;
                }
                else
                    row.Cells["分數認定"].ErrorText = "";

                rowCount++;
            }

            if (rowCount == 0)
            {
                MessageBox.Show("請輸入資料");
                value = false;
            }

            return value;
        }

        private void dgData_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex > -1 && e.ColumnIndex > -1)
                dgData.Rows[e.RowIndex].Cells[e.ColumnIndex].ErrorText = "";
        }
    }
}
