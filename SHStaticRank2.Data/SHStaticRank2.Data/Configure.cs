using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;

namespace SHStaticRank2.Data
{
    [FISCA.UDT.TableName("ischool.StudentSHStaticRank2Data.Configure")]
    public class Configure : FISCA.UDT.ActiveRecord
    {
        public Configure()
        {
            PrintSubjectList = new List<string>();
            TagRank1TagList = new List<string>();
            TagRank1SubjectList = new List<string>();
            TagRank2TagList = new List<string>();
            TagRank2SubjectList = new List<string>();
            RankFilterTagList = new List<string>();
            RankFilterUseScoreList = new List<string>();
            RankFilterGradeSemeterList = new List<string>();
        }

        /// <summary>
        /// 設定檔名稱
        /// </summary>
        [FISCA.UDT.Field]
        public string Name { get; set; }
        /// <summary>

        /// <summary>
        /// 列印樣板
        /// </summary>
        [FISCA.UDT.Field]
        private string TemplateStream { get; set; }

        /// <summary>
        /// 列印樣板1
        /// </summary>
        [FISCA.UDT.Field]
        private string TemplateStream1 { get; set; }

        /// <summary>
        /// 列印樣板2
        /// </summary>
        [FISCA.UDT.Field]
        private string TemplateStream2 { get; set; }

        /// <summary>
        /// 列印樣板3
        /// </summary>
        [FISCA.UDT.Field]
        private string TemplateStream3 { get; set; }


        /// <summary>
        /// 樣版P
        /// </summary>
        public Aspose.Words.Document Template { get; set; }


        /// <summary>
        /// 樣版2
        /// </summary>
        public Aspose.Words.Document Template1 { get; set; }


        /// <summary>
        /// 樣版2
        /// </summary>
        public Aspose.Words.Document Template2 { get; set; }


        /// <summary>
        /// 樣版3
        /// </summary>
        public Aspose.Words.Document Template3 { get; set; }

        /// <summary>
        /// 樣板中支援列印科目的最大數
        /// </summary>
        [FISCA.UDT.Field]
        public int SubjectLimit { get; set; }        
      
        /// <summary>
        /// 列印科別
        /// </summary>
        [FISCA.UDT.Field]
        private string PrintSubjectListString { get; set; }
        public List<string> PrintSubjectList { get; private set; }
        /// <summary>
        /// 類別排名1
        /// </summary>
        [FISCA.UDT.Field]
        public string TagRank1TagName { get; set; }
        public List<string> TagRank1TagList { get; private set; }
        /// <summary>
        /// 類別排名1，排名科目
        /// </summary>
        [FISCA.UDT.Field]
        private string TagRank1SubjectListString { get; set; }
        public List<string> TagRank1SubjectList { get; private set; }
        /// <summary>
        /// 類別排名2
        /// </summary>
        [FISCA.UDT.Field]
        public string TagRank2TagName { get; set; }
        public List<string> TagRank2TagList { get; private set; }
        /// <summary>
        /// 類別排名2，排名科目
        /// </summary>
        [FISCA.UDT.Field]
        private string TagRank2SubjectListString { get; set; }
        public List<string> TagRank2SubjectList { get; private set; }
        /// <summary>
        /// 不參與排名學生類別
        /// </summary>
        [FISCA.UDT.Field]
        public string RankFilterTagName { get; set; }
        public List<string> RankFilterTagList { get; private set; }

        /// <summary>
        /// 排名年級對象
        /// </summary>
        [FISCA.UDT.Field]
        public string SortGradeYear { get; set; }

        /// <summary>
        /// 成績年級學期
        /// </summary>
        [FISCA.UDT.Field]
        public string RankFilterGradeSemeter { get; set; }
        public List<string> RankFilterGradeSemeterList { get; private set; }

        /// <summary>
        /// 採計成績
        /// </summary>
        [FISCA.UDT.Field]
        public string RankFilterUseScoreName { get; set; }
        public List<string> RankFilterUseScoreList { get; private set; }

        /// <summary>
        /// 計算學期成績排名
        /// </summary>
        [FISCA.UDT.Field]
        public bool WithCalSemesterScoreRank { get; set; }

        /// <summary>
        /// UI 設定儲存值
        /// </summary>
        [FISCA.UDT.Field]
        public string UIConfigString { get; set; }
        
        /// <summary>
        /// 科目名稱回歸對照使用
        /// </summary>
        [FISCA.UDT.Field]
        public string SubjectMapping { get; set; }
           

        /// <summary>
        /// 在儲存前，把資料填入儲存欄位中
        /// </summary>
        public void Encode()
        {
         
            this.PrintSubjectListString = "";
            foreach (var item in this.PrintSubjectList)
            {
                this.PrintSubjectListString += (this.PrintSubjectListString == "" ? "" : "^^^") + item;
            }         
            this.TagRank1SubjectListString = "";
            foreach (var item in this.TagRank1SubjectList)
            {
                this.TagRank1SubjectListString += (this.TagRank1SubjectListString == "" ? "" : "^^^") + item;
            }
            this.TagRank2SubjectListString = "";
            foreach (var item in this.TagRank2SubjectList)
            {
                this.TagRank2SubjectListString += (this.TagRank2SubjectListString == "" ? "" : "^^^") + item;
            }

            // 畫面上成績年級學期
            this.RankFilterGradeSemeter = "";
            if(this.RankFilterGradeSemeterList==null)
                this.RankFilterGradeSemeterList= new List<string>();
            foreach (var item in this.RankFilterGradeSemeterList)
            {
                this.RankFilterGradeSemeter += (this.RankFilterGradeSemeter == "" ? "" : "^^^") + item;
            }
            // 畫面上採計成績
            this.RankFilterUseScoreName = "";
            if (this.RankFilterUseScoreList == null)
                this.RankFilterUseScoreList = new List<string>();
            foreach (var item in this.RankFilterUseScoreList)
            {
                this.RankFilterUseScoreName += (this.RankFilterUseScoreName == "" ? "" : "^^^") + item;
            }

            if (this.Template != null)
            {
                System.IO.MemoryStream stream = new System.IO.MemoryStream();
                this.Template.Save(stream, Aspose.Words.SaveFormat.Doc);
                this.TemplateStream = Convert.ToBase64String(stream.ToArray());
            }

            if (this.Template1 != null)
            {
                System.IO.MemoryStream stream1 = new System.IO.MemoryStream();
                this.Template1.Save(stream1, Aspose.Words.SaveFormat.Doc);
                this.TemplateStream1 = Convert.ToBase64String(stream1.ToArray());
            }

            if(this.Template2 !=null)
            {
                System.IO.MemoryStream stream2 = new System.IO.MemoryStream();
                this.Template2.Save(stream2, Aspose.Words.SaveFormat.Doc);
                this.TemplateStream2 = Convert.ToBase64String(stream2.ToArray());
            }

            if(this.Template3 !=null)
            {
                System.IO.MemoryStream stream3 = new System.IO.MemoryStream();
                this.Template3.Save(stream3, Aspose.Words.SaveFormat.Doc);
                this.TemplateStream3 = Convert.ToBase64String(stream3.ToArray());
            }

        }
        /// <summary>
        /// 在資料取出後，把資料從儲存欄位轉換至資料欄位
        /// </summary>
        public void Decode()
        {
            
            this.PrintSubjectList = new List<string>(this.PrintSubjectListString.Split(new string[] { "^^^" }, StringSplitOptions.RemoveEmptyEntries));
            
            this.TagRank1SubjectList = new List<string>(this.TagRank1SubjectListString.Split(new string[] { "^^^" }, StringSplitOptions.RemoveEmptyEntries));
            this.TagRank2SubjectList = new List<string>(this.TagRank2SubjectListString.Split(new string[] { "^^^" }, StringSplitOptions.RemoveEmptyEntries));
            // 採計成績
            this.RankFilterUseScoreList = new List<string>(this.RankFilterUseScoreName.Split(new string[] { "^^^" }, StringSplitOptions.RemoveEmptyEntries));
            // 年級學期
            this.RankFilterGradeSemeterList  = new List<string>(this.RankFilterGradeSemeter.Split(new string[] { "^^^" }, StringSplitOptions.RemoveEmptyEntries));

            if(this.TemplateStream !=null)
                this.Template = new Aspose.Words.Document(new MemoryStream(Convert.FromBase64String(this.TemplateStream)));

            if(this.TemplateStream1!=null)
                this.Template1 = new Aspose.Words.Document(new MemoryStream(Convert.FromBase64String(this.TemplateStream1)));

            if(this.TemplateStream2 !=null)
                this.Template2 = new Aspose.Words.Document(new MemoryStream(Convert.FromBase64String(this.TemplateStream2)));

            if(this.TemplateStream3 !=null)
                this.Template3 = new Aspose.Words.Document(new MemoryStream(Convert.FromBase64String(this.TemplateStream3)));
        }


        public bool CalcGradeYear1 { get; set; }
        public bool CalcGradeYear2 { get; set; }
        public bool CalcGradeYear3 { get; set; }
        public bool CalcGradeYear4 { get; set; }
        public string NotRankTag { get; set; }
        public bool use原始成績 { get; set; }
        public bool use補考成績 { get; set; }
        public bool use重修成績 { get; set; }
        public bool use手動調整成績 { get; set; }
        public bool use學年調整成績 { get; set; }
        public string Rank1Tag { get; set; }
        public string Rank2Tag { get; set; }
        public bool DoNotSaveIt { get; set; }
        public bool 計算學業成績排名 { get; set; }
        /// <summary>
        ///勾選年級學期
        /// </summary>
        public List<string> useGradeSemesterList = new List<string>();

        /// <summary>
        /// 勾選列印科目
        /// </summary>
        public List<string> useSubjectPrintList = new List<string>();

        /// <summary>
        /// 勾選類別科目名稱1
        /// </summary>
        public List<string> useSubjecOrder1List = new List<string>();

        /// <summary>
        /// 勾選類別科目名稱2
        /// </summary>
        public List<string> useSubjecOrder2List = new List<string>();

        /// <summary>
        /// 檢查是否產生 PDF 檔
        /// </summary>
        [FISCA.UDT.Field]
        public bool CheckExportPDF { get; set; }

        [FISCA.UDT.Field]
        public bool CheckExportStudent { get; set; }

        /// <summary>
        /// 產生 PDF 檔案名稱是否使用身分證號 Y:身分證 N:學測報名
        /// </summary>
        [FISCA.UDT.Field]
        public bool CheckUseIDNumber { get; set; }


        /// <summary>
        /// 判斷是否只產生回歸科目
        /// </summary>
        [FISCA.UDT.Field]
        public bool CheckExportSubjectMapping { get; set; }
    }
}
