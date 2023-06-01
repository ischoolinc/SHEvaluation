using System;
using System.ComponentModel;
using System.Xml;
using DevComponents.Editors;
using SmartSchool.Evaluation.GraduationPlan;
using FISCA.Data;
using SmartSchool.Feature.GraduationPlan;
using System.Data;
using System.Collections.Generic;
using DevComponents.DotNetBar.Controls;
using SmartSchool.ClassRelated.RibbonBars;
using SmartSchool.Evaluation.WearyDogComputerHelper;
using static System.ComponentModel.Design.ObjectSelectorEditor;
using System.IO;
using System.Web.UI.WebControls;
using System.Windows.Forms;
using System.Xml.Linq;
using System.Linq;
using System.Drawing.Design;

namespace SmartSchool.Evaluation.Configuration
{
    public partial class GraduationPlanSimplePicker : FISCA.Presentation.Controls.BaseForm
    {
        private List<GraduationPlanSimple> gpSimpleCollection=new List<GraduationPlanSimple>();

        private List<GraduationPlanSimple> gpTargetCollection = new List<GraduationPlanSimple>();

        private List<GraduationPlanSimple> gpSelectedCollection = new List<GraduationPlanSimple>();

        private string _Catalog; // 科目表類別

        public GraduationPlanSimplePicker(string cata)
        {
            InitializeComponent();

            // 預設為目前的學年度
            _Catalog = cata;
            iiSchoolYear.Text = K12.Data.School.DefaultSchoolYear;
            loadGPCbo(iiSchoolYear.Text);
            loadDomainCbo(cboDomain);
            loadEntryCbo(cboEntry);
            loadRequiredByCbo(cboRequiredBy);
            loadRequiredCbo(cboRequired);
            loadSubjAttrib(cboSubjAttrib);
        }
        public List<GraduationPlanSimple> getGPSelection()
        {
            return gpSelectedCollection;
        }

        private void buttonX2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void buttonX1_Click(object sender, EventArgs e)
        {
            if(dgGraduationPlan.SelectedRows.Count>0)
            {
                gpSelectedCollection.Clear();
                foreach(DataGridViewRow row in dgGraduationPlan.SelectedRows)
                {
                    gpSelectedCollection.Add((GraduationPlanSimple)row.Tag);
                }

                this.Close();
            }
            else
            {
                MessageBox.Show("請選擇科目");
            }
        }


        private void iiSchoolYear_ValueChanged(object sender, EventArgs e)
        {
            loadGPCbo(iiSchoolYear.Text);

        }

        private void loadGPCbo(string schoolYear)
        {
            cboGPlan.Items.Clear();
            string query = "SELECT id,name, array_to_string(xpath('//GraduationPlan/@SchoolYear', xmlparse(content content)), '')::text as SchoolYear FROM graduation_plan WHERE array_to_string(xpath('//GraduationPlan/@SchoolYear', xmlparse(content content)), '')::text ='" + schoolYear + "'";            
            QueryHelper qh = new QueryHelper();
            DataTable dt = qh.Select(query);
            ComboItem leadItem=new ComboItem();
            if (dt.Rows.Count > 0)
            {
                leadItem.Text = "請選擇課程規劃表";
                cboGPlan.Items.Add(leadItem);
                foreach (DataRow dr in dt.Rows)
                {
                    ComboItem item = new DevComponents.Editors.ComboItem();
                    item.Text = dr["name"].ToString();
                    item.Tag = dr["id"].ToString();
                    cboGPlan.Items.Add(item);
                }
                gpTargetCollection.Clear();
                gpTargetCollection = gpSimpleCollection.ToList();
            }
            else
            {
                leadItem.Text = "本學年度尚無課程規劃表";
                cboGPlan.Items.Add(leadItem);
                gpTargetCollection.Clear();
                gpSimpleCollection.Clear();

            }
            resetFilter();
            cboGPlan.SelectedIndex = 0;
            dgGraduationPlan.Rows.Clear();
        }

        public void resetFilter()
        {
            cboDomain.SelectedIndex = 0;
            cboEntry.SelectedIndex = 0;
            cboRequired.SelectedIndex = 0;
            cboRequiredBy.SelectedIndex = 0;
        }
        

        /// <summary>
        /// 取得領域對照
        /// </summary>
        public void loadDomainCbo(ComboBoxEx cbox)
        {
            Dictionary<string, string> value = new Dictionary<string, string>();
            value.Add("00", "不分");
            value.Add("01", "語文");
            value.Add("02", "數學");
            value.Add("03", "社會");
            value.Add("04", "自然科學");
            value.Add("05", "藝術");
            value.Add("06", "綜合活動");
            value.Add("07", "科技");
            value.Add("08", "健康與體育");
            value.Add("09", "全民國防教育");
            value.Add("0A", "跨領域科目專題");
            value.Add("0B", "實作(實驗)及探索體驗");
            value.Add("0C", "職涯試探");
            value.Add("0D", "專題探究");
            value.Add("0E", "跨領域/科目統整");
            value.Add("0F", "通識性課程");
            value.Add("0G", "大學預修課程");
            value.Add("0H", "第二外國語文");
            value.Add("0I", "本土語文");
            value.Add("0J", "綜合活動與科技");
            value.Add("0R", "競技體育技能");
            value.Add("0S", "體育專業學科");
            value.Add("0T", "體育專項術科");
            value.Add("11", "特殊需求領域(身心障礙)");
            value.Add("12", "特殊需求領域(資賦優異)");
            value.Add("13", "特殊需求領域(音樂專長)");
            value.Add("14", "特殊需求領域(美術專長)");
            value.Add("15", "特殊需求領域(舞蹈專長)");
            value.Add("16", "特殊需求領域(戲劇專長)");
            value.Add("17", "特殊需求領域(體育專長)");
            value.Add("18", "特殊需求領域(實驗課程)");
            value.Add("91", "藝術才能專長");
            value.Add("92", "生涯發展");
            value.Add("93", "職能探索");
            value.Add("94", "運動防護");
            value.Add("A1", "數值控制技能");
            value.Add("A2", "精密機械製造技能");
            value.Add("A3", "模型設計與鑄造技能");
            value.Add("A4", "電腦輔助機械設計技能");
            value.Add("A5", "自動化整合技能");
            value.Add("A6", "金屬成形與管線技能");
            value.Add("B1", "車輛技能");
            value.Add("B2", "機器腳踏車技能");
            value.Add("B3", "液氣壓技能");
            value.Add("B4", "動力機械技能");
            value.Add("C1", "晶片設計技能");
            value.Add("C2", "微電腦應用技能");
            value.Add("C3", "自動控制技能");
            value.Add("C4", "電機工程技能");
            value.Add("C5", "冷凍空調技能");
            value.Add("D1", "化工及檢驗技能");
            value.Add("D2", "紡染及檢驗技能");
            value.Add("E2", "專業製圖技能");
            value.Add("E3", "土木測量技能");
            value.Add("F1", "商業與財會技能");
            value.Add("F2", "跨境商務技能");
            value.Add("F3", "資訊應用技能");
            value.Add("H1", "平面設計技能");
            value.Add("H2", "立體造形技能");
            value.Add("H3", "數位成型技能");
            value.Add("H4", "數位影音技能");
            value.Add("H5", "互動媒體技能");
            value.Add("H6", "空間設計技能");
            value.Add("I1", "農業生產與休閒生態技能");
            value.Add("I2", "動物飼養及保健技能");
            value.Add("J1", "食品加工技能");
            value.Add("J2", "檢驗分析技能");
            value.Add("K1", "整體造型技能");
            value.Add("K2", "服裝實務技能");
            value.Add("K3", "生活應用技能");
            value.Add("L1", "廚藝技能");
            value.Add("L2", "烘焙技能");
            value.Add("L3", "旅宿技能");
            value.Add("L4", "旅遊技能");
            value.Add("M1", "漁航技能");
            value.Add("M2", "漁業技能");
            value.Add("M3", "水域活動安全技能");
            value.Add("M4", "觀賞水族技能");
            value.Add("M5", "經濟水族技能");
            value.Add("M6", "區域特色水族技能");
            value.Add("M7", "海面養殖技能");
            value.Add("N1", "船舶金工技能");
            value.Add("N2", "船舶機電控制技能");
            value.Add("N3", "船舶動力技能");
            value.Add("N4", "船舶作業技能");
            value.Add("N5", "船舶操縱技能");
            value.Add("N6", "電子導航技能");
            value.Add("N7", "船舶維護與繫固作業技能");
            value.Add("O1", "視覺表現技能");
            value.Add("O2", "展演製作技能");
            value.Add("O3", "數位影音技能");
            value.Add("O4", "音樂藝術技能");
            value.Add("O5", "舞蹈藝術技能");
            value.Add("O6", "表演藝術實務技能");
            value.Add("U1", "車輛整理技能");
            value.Add("U2", "門市技能");
            value.Add("U3", "物品整理技能");
            value.Add("U4", "農園藝技能");
            value.Add("U5", "產品加工技能");
            value.Add("U6", "裝配技能");
            value.Add("U7", "生活照護技能");
            value.Add("U8", "家務處理技能");
            value.Add("U9", "餐飲製作技能");
            value.Add("UA", "旅館房務技能");
            value.Add("UB", "按摩技能");
            value.Add("UC", "紓壓保健技能");
            value.Add("UD", "民俗技能");
            value.Add("UE", "寵物照顧技能");
            value.Add("UF", "美髮技能");
            value.Add("G3", "職場實務技能");
            value.Add("G1", "英語文技能");
            value.Add("G2", "日語文技能");
            foreach(string key in value.Keys)
            {
                DevComponents.Editors.ComboItem item = new DevComponents.Editors.ComboItem();
                item.Text = value[key].ToString();
                item.Tag = key;
                cbox.Items.Add(item);
            }
            cbox.SelectedIndex = 0;
        }



        /// <summary>
        /// 取得科目屬性
        /// </summary>
        public void loadSubjAttrib(ComboBoxEx cbox)
        {
            Dictionary<string, string> value = new Dictionary<string, string>();
            value.Add("0", "不分屬性");
            value.Add("1", "一般科目");
            value.Add("2", "專業科目");
            value.Add("3", "實習科目");
            value.Add("4", "專精科目");
            value.Add("5", "專精科目(核心科目)");
            value.Add("6", "特殊需求領域");
            value.Add("A", "自主學習");
            value.Add("B", "選手培訓");
            value.Add("C", "充實補強(不授予學分)");
            value.Add("D", "充實補強(授予學分)");
            value.Add("E", "學校特色活動");
            value.Add("F", "專精(專業)科目");
            value.Add("G", "專精(實習)科目");
            value.Add("H", "專精(專業)科目(核心)");
            value.Add("I", "專精(實習)科目(核心)");

            foreach (string key in value.Keys)
            {
                ComboItem item = new ComboItem();
                item.Text = value[key].ToString();
                item.Tag = key;
                cbox.Items.Add(item);
            }
            cbox.SelectedIndex = 0;
        }



        public void loadEntryCbo(ComboBoxEx cbox)
        {
            List<string> entryList = new List<string>()
            {"學業","專業科目","實習科目" };
            foreach (string key in entryList)
            {
                ComboItem item = new ComboItem();
                item.Text = key;
                item.Tag = key;
                cbox.Items.Add(item);
            }
            cbox.SelectedIndex = 0;
        }

        public void loadRequiredByCbo(ComboBoxEx cbox)
        {
            List<string> entryList = new List<string>()
            {"校訂","部訂"};
            foreach (string key in entryList)
            {
                ComboItem item = new ComboItem();
                item.Text = key;
                item.Tag = key;
                cbox.Items.Add(item);
            }
            cbox.SelectedIndex = 0;
        }

        public void loadRequiredCbo(ComboBoxEx cbox)
        {
            List<string> entryList = new List<string>()
            {"必修","選修"};
            foreach (string key in entryList)
            {
                DevComponents.Editors.ComboItem item = new DevComponents.Editors.ComboItem();
                item.Text = key;
                item.Tag = key;
                cbox.Items.Add(item);
            }
            cbox.SelectedIndex = 0;
        }


        // 依據課程規劃ID 載入資料項目
        private void loadGPDataGrid()
        {
            dgGraduationPlan.Rows.Clear();
            if(gpTargetCollection.Count > 0 )
            {
                foreach(GraduationPlanSimple dps in gpTargetCollection)
                {
                    int rowIdx = dgGraduationPlan.Rows.Add();
                    dgGraduationPlan.Rows[rowIdx].Tag = dps;
                    dgGraduationPlan.Rows[rowIdx].Cells[0].Value = dps.Domain;
                    dgGraduationPlan.Rows[rowIdx].Cells[1].Value = dps.Attribute;
                    dgGraduationPlan.Rows[rowIdx].Cells[2].Value = dps.Entry;
                    dgGraduationPlan.Rows[rowIdx].Cells[3].Value = dps.SubjectName;
                    dgGraduationPlan.Rows[rowIdx].Cells[4].Value = dps.RequiredBy;
                    dgGraduationPlan.Rows[rowIdx].Cells[5].Value = dps.Required;
                    dgGraduationPlan.Rows[rowIdx].Cells[6].Value = dps.LevelList;

                }
            }
        }
        public void acqGPData(string gpID)
        {
            string query = "SELECT " +
"DISTINCT array_to_string(xpath('//Subject/@SubjectName', subject), '')::text AS SubjectName, " +
"array_to_string(xpath('//Subject/@課程代碼', Subject), '')::text as Code, " +
"array_to_string(xpath('//Subject/@Domain', Subject), '')::text as Domain, " +
"SUBSTRING(array_to_string(xpath('//Subject/@CourseAttr', Subject), '')::text,2,1) as CourseAttrib, " +
"array_to_string(xpath('//Subject/@Entry', Subject), '')::text as Entry, " +
"array_to_string(xpath('//Subject/@RequiredBy', Subject), '')::text as RequiredBy, " +
"array_to_string(xpath('//Subject/@Required', Subject), '')::text as Required, " +
"string_agg(array_to_string(xpath('//Subject/@Level', subject), '')::text, ',') AS LevelList " +
"FROM( " +
"SELECT unnest(xpath('//GraduationPlan/Subject', xmlparse(content content))) AS Subject " +
"FROM graduation_plan " +
"WHERE id = " + gpID + " " +
") AS SubjQry " +
"GROUP BY SubjectName,Code,Domain,CourseAttrib,Entry,RequiredBy,Required " +
"ORDER BY Code ";

            QueryHelper qh = new QueryHelper();
            DataTable dt = qh.Select(query); 
            if(dt.Rows.Count > 0)
            {
                gpSimpleCollection.Clear();
                foreach (DataRow dr in dt.Rows)
                {
                    GraduationPlanSimple gpsd = new GraduationPlanSimple(dr);
                    gpSimpleCollection.Add(gpsd);
                }
            }
        }

        private void cboGPlan_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cboGPlan.SelectedIndex > 0)
            {
                ComboItem item = (ComboItem)cboGPlan.SelectedItem;
                acqGPData(item.Tag.ToString());
                gpTargetCollection.Clear();
                gpTargetCollection = gpSimpleCollection.ToList();
                resetFilter();
                loadGPDataGrid();
            }
            else
            {
                gpTargetCollection.Clear();
                gpSimpleCollection.Clear();
                resetFilter();
                dgGraduationPlan.Rows.Clear();
            }
        }

        private void cbo_SelectedIndexChanged(object sender, EventArgs e)
        {
            processFiltering();
            loadGPDataGrid();

        }
        private void processFiltering()
        {
            gpTargetCollection.Clear();
            gpTargetCollection=gpSimpleCollection.ToList();
            ComboItem itemDomain = (ComboItem)cboDomain.SelectedItem;
            if (cboDomain.SelectedIndex>0)
            {
                gpTargetCollection=gpTargetCollection.Where(x=>x.Domain.Contains(itemDomain.Text)).ToList();
            }
            ComboItem itemAttrib = (ComboItem)cboSubjAttrib.SelectedItem;
            if (cboSubjAttrib.SelectedIndex > 0)
            {
                gpTargetCollection = gpTargetCollection.Where(x => x.Attribute.Contains(itemAttrib.Text)).ToList();
            }
            ComboItem itemEntry = (ComboItem)cboEntry.SelectedItem;
            if (cboEntry.SelectedIndex>0)
            {
                gpTargetCollection = gpTargetCollection.Where(x => x.Entry.Contains(itemEntry.Text)).ToList();
            }
            ComboItem itemRequiredBy = (ComboItem)cboRequiredBy.SelectedItem;
            if (cboRequiredBy.SelectedIndex > 0)
            {
                gpTargetCollection = gpTargetCollection.Where(x => x.RequiredBy.Contains(itemRequiredBy.Text)).ToList();
            }
            ComboItem itemRequired = (ComboItem)cboRequired.SelectedItem;
            if (cboRequired.SelectedIndex>0)
            {
                gpTargetCollection = gpTargetCollection.Where(x => x.Required.Contains(itemRequired.Text)).ToList();
            }
        }
    }
}