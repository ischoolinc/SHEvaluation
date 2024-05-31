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
using System.Text.RegularExpressions;
using Aspose.Cells;

namespace SHScoreValueManager.UIForm
{
    public partial class frmSetScoreValue : BaseForm
    {
        K12.Data.Configuration.ConfigData cd = K12.Data.School.Configuration["評量成績缺考設定"];
        // 設定 XML
        XElement SettingXML = null;

        // 設定資料內容
        List<ScoreSettingConfig> ScoreSettingConfigList;
        List<string> chkUseTextSame = new List<string>();

        // 紀錄修改前
        StringBuilder CheckOldSB;

        // 紀錄修改後
        StringBuilder CheckNewSB;

        public frmSetScoreValue()
        {
            InitializeComponent();
            ScoreSettingConfigList = new List<ScoreSettingConfig>();
            CheckOldSB = new StringBuilder();
            CheckNewSB = new StringBuilder();
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            if (CheckDataSB())
                this.Close();
            else
            {
                DialogResult dr = MsgBox.Show("資料已修改，尚未儲存，確定離開?", MessageBoxButtons.YesNo, MessageBoxDefaultButton.Button2);
                if (dr == DialogResult.No)
                    return;
                else
                    this.Close();
            }
        }

        // 離開提示
        private void frmSetScoreValue_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (CheckDataSB())
                e.Cancel = false;
            else
            {
                DialogResult dr = MsgBox.Show("資料已修改，尚未儲存，確定離開?", MessageBoxButtons.YesNo, MessageBoxDefaultButton.Button2);
                if (dr == DialogResult.No)
                    e.Cancel = true;
            }
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
            else
            {
                MsgBox.Show("資料有誤，請檢查。");
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

            // 檢查資料
            CheckData();
            dgData.ImeMode = ImeMode.OnHalf;

            // 清除修改前紀錄
            CheckOldSB.Clear();
            RecordChangeData(CheckOldSB);
        }

        // 記錄修改資料
        private void RecordChangeData(StringBuilder sb)
        {
            sb.Clear();

            foreach (DataGridViewRow dr in dgData.Rows)
            {
                if (dr.IsNewRow) continue;
                foreach (DataGridViewCell cell in dr.Cells)
                {
                    if (cell != null)
                    {
                        sb.Append(cell.Value);
                    }
                }
            }
        }

        // 比對新舊記錄是否相同
        private bool CheckDataSB()
        {
            // 整理目前資料
            RecordChangeData(CheckNewSB);

            if (CheckOldSB.ToString() == CheckNewSB.ToString())
                return true;
            else
                return false;
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
                // 分數認定
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
            chkUseTextSame.Clear();

            // 檢查 DataGridView 資料
            foreach (DataGridViewRow row in dgData.Rows)
            {
                if (row.IsNewRow)
                    continue;

                // 檢查缺考原因
                string ReportValue = ("" + row.Cells["缺考原因"].Value).Trim();
                if (ReportValue == "")
                {
                    row.Cells["缺考原因"].ErrorText = "必填";
                    value = false;
                }
                else
                {
                    row.Cells["缺考原因"].ErrorText = "";
                    if (ContainsSpecialCharacters(ReportValue))
                    {
                        row.Cells["缺考原因"].ErrorText = "不可包含特殊字元";
                        value = false;
                    }
                }


                // 檢查輸入內容
                string UseText = ("" + row.Cells["輸入內容"].Value).Trim();
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

                    if (ContainsSpecialCharacters(UseText))
                    {
                        row.Cells["輸入內容"].ErrorText = "不可包含特殊字元";
                        value = false;
                    }

                }

                // 檢查分數認定
                string ScoreType = ("" + row.Cells["分數認定"].Value).Trim();
                if (ScoreType == "")
                {
                    row.Cells["分數認定"].ErrorText = "必填";
                    value = false;
                }
                else
                    row.Cells["分數認定"].ErrorText = "";

                rowCount++;
            }

            return value;
        }

        // 檢查特殊自字元
        private bool ContainsSpecialCharacters(string input)
        {
            // 定義一個正則表達式來匹配特殊字元
            string pattern = @"[!@#$%^&*?""{}|<>/\\]";
            Regex regex = new Regex(pattern);

            return regex.IsMatch(input);
        }

        private void dgData_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            //if (e.RowIndex > -1 && e.ColumnIndex > -1)
            //    dgData.Rows[e.RowIndex].Cells[e.ColumnIndex].ErrorText = "";
        }

        private void dgData_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex > -1 && e.ColumnIndex > -1)
            {
                // 判斷是否是新列
                if (dgData.Rows[e.RowIndex].IsNewRow)
                    return;

                if (dgData.Columns[e.ColumnIndex].Name == "缺考原因")
                {
                    dgData.Rows[e.RowIndex].Cells[e.ColumnIndex].ErrorText = "";
                    string ReportValue = ("" + dgData.Rows[e.RowIndex].Cells["缺考原因"].FormattedValue).Trim();
                    if (ReportValue == "")
                    {
                        dgData.Rows[e.RowIndex].Cells["缺考原因"].ErrorText = "必填";
                    }
                    else
                    {
                        dgData.Rows[e.RowIndex].Cells["缺考原因"].ErrorText = "";
                        // 檢查輸入內容是否包含特殊字元
                        if (ContainsSpecialCharacters(ReportValue))
                        {
                            dgData.Rows[e.RowIndex].Cells["缺考原因"].ErrorText = "不可包含特殊字元";
                        }

                    }
                }

                if (dgData.Columns[e.ColumnIndex].Name == "輸入內容")
                {
                    dgData.Rows[e.RowIndex].Cells[e.ColumnIndex].ErrorText = "";
                    string UseText = ("" + dgData.Rows[e.RowIndex].Cells["輸入內容"].FormattedValue).Trim();
                    if (UseText == "")
                    {
                        dgData.Rows[e.RowIndex].Cells["輸入內容"].ErrorText = "必填";
                    }
                    else
                    {

                        chkUseTextSame.Clear();
                        dgData.Rows[e.RowIndex].Cells["輸入內容"].ErrorText = "";
                        // 檢查輸入內容是否重複
                        int rowIdx = 0;
                        foreach (DataGridViewRow drv in dgData.Rows)
                        {
                            if (drv.IsNewRow || rowIdx == e.RowIndex)
                                continue;

                            string chkStr = drv.Cells["輸入內容"].Value + "";
                            if (chkStr != "")
                            {
                                if (!chkUseTextSame.Contains(chkStr))
                                    chkUseTextSame.Add(chkStr);
                            }

                            rowIdx++;
                        }


                        if (chkUseTextSame.Contains(UseText))
                        {
                            dgData.Rows[e.RowIndex].Cells["輸入內容"].ErrorText = UseText + " 重複!";
                        }

                        // 檢查輸入內容是否包含特殊字元
                        if (ContainsSpecialCharacters(UseText))
                        {
                            dgData.Rows[e.RowIndex].Cells["輸入內容"].ErrorText = "不可包含特殊字元";
                        }
                    }
                }

                if (dgData.Columns[e.ColumnIndex].Name == "分數認定")
                {
                    dgData.Rows[e.RowIndex].Cells[e.ColumnIndex].ErrorText = "";
                    if ("" + dgData.Rows[e.RowIndex].Cells["分數認定"].FormattedValue == "")
                    {
                        dgData.Rows[e.RowIndex].Cells["分數認定"].ErrorText = "必填";
                    }
                    else
                        dgData.Rows[e.RowIndex].Cells["分數認定"].ErrorText = "";
                }

            }
        }

        private void dgData_CellValidating(object sender, DataGridViewCellValidatingEventArgs e)
        {


        }

        private void dgData_CellEnter(object sender, DataGridViewCellEventArgs e)
        {
            dgData.ImeMode = ImeMode.OnHalf;
            dgData.ImeMode = ImeMode.Off;
        }

        private void dgData_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex > -1 && e.ColumnIndex > -1)
            {
                // 判斷是否是新列
                if (dgData.Rows[e.RowIndex].IsNewRow)
                    return;


                if (dgData.Columns[e.ColumnIndex].Name == "分數認定")
                {
                    dgData.Rows[e.RowIndex].Cells[e.ColumnIndex].ErrorText = "";

                    // 如果是空值，則預設為0分
                    if ("" + dgData.Rows[e.RowIndex].Cells["分數認定"].Value == "")
                        dgData.Rows[e.RowIndex].Cells["分數認定"].Value = "0分";

                    if ("" + dgData.Rows[e.RowIndex].Cells["分數認定"].FormattedValue == "")
                    {
                        dgData.Rows[e.RowIndex].Cells["分數認定"].ErrorText = "必填";
                    }
                    else
                        dgData.Rows[e.RowIndex].Cells["分數認定"].ErrorText = "";
                }
            }
        }
    }
}
