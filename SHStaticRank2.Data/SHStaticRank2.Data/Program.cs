﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SmartSchool.Customization.Data;
using System.IO;
using FISCA.DSAUtil;
using System.Data;
using Aspose.Words;

namespace SHStaticRank2.Data
{
    public class Program
    {
        [FISCA.MainMethod]
        public static void Main()
        {
            /*
             * 1.拉出configure
             * 2.加入OneClassCompleted事件,處理非word merge情況
             * 3.properties下合併欄位總表接口
             * 4.public FolderName
             * 5.加入"類別一分類",類別二分類至 merge table
             */
            FISCA.Permission.Catalog cat = FISCA.Permission.RoleAclSource.Instance["教務作業"]["功能按鈕"];
            cat.Add(new FISCA.Permission.RibbonFeature("SHSchool.SHStaticRank2.Data", "計算固定排名"));

            var button = FISCA.Presentation.MotherForm.RibbonBarItems["教務作業", "批次作業/檢視"]["成績作業"]["計算固定排名"]["計算多學期成績固定排名"];
            button.Enable = FISCA.Permission.UserAcl.Current["SHSchool.SHStaticRank2.Data"].Executable;//MotherForm.RibbonBarItems["教務作業", "批次作業/檢視"]["成績作業"]["學期成績處理"].Enable = CurrentUser.Acl["Button0670"].Executable;
            button.Click += delegate
            {
                System.Windows.Forms.MessageBox.Show(
@"功能說明：
本功能以原有固定功能為基礎，加上類別篩選及類組排名。
排名完將會顯示排名細節，此資料目前暫不儲存，請自行留存，產生過程依個人電腦效能需要5~30分鐘。
1.排名對象為目前為該年級，扣除標註有不排名類別以及非一般狀態學生(非在選取的學期中就讀該年級的學生)。
2.計算過程中所有學生年級、班級、所屬類別等皆以目前的狀態進行計算，不會因為選取以前的學年度而往回推。
3.排名所採用的分數為勾選成績項目中選最高值。
4.若勾選'僅預覽，不儲存結果'項目，排名資料將不會寫入資料庫，僅有排名細節供預覽用。
5.科目排名工作表內最多放入6個科目排名資料。
6.學業分項成績計算：學業、體育、健康與護理、國防通識、實習科目
");
                var conf = new CalcMutilSemeSubjectStatucRank();
                conf.ShowDialog();
                if (conf.DialogResult == System.Windows.Forms.DialogResult.OK)
                {
                    CalcMutilSemeSubjectRank.Setup(conf.Setting);
                    //CalcMutilSemeSubjectRank.OneClassCompleted += delegate
                    //{
                    //    DataTable students = CalcMutilSemeSubjectRank._table.Copy();

                    //    if (students.Rows.Count <= 0)
                    //        return;

                    //    Dictionary<string, string> classData = new Dictionary<string, string>();
                    //    classData.Add("學校名稱", students.Rows[0]["學校名稱"] + "");
                    //    classData.Add("科別", students.Rows[0]["科別"] + "");
                    //    classData.Add("班級", students.Rows[0]["班級"] + "");

                    //    SingleClassMailMerge classm = new SingleClassMailMerge(classData, students);
                    //    MemoryStream ms = new MemoryStream(Properties.Resources.高中部班級歷年成績單);

                    //    Document doc = new Document(ms);
                    //    doc.MailMerge.ExecuteWithRegions(classm);

                    //    string pathW = Path.Combine(System.Windows.Forms.Application.StartupPath, "Reports");
                    //    doc.Save(Path.Combine(pathW, classData["班級"]) + ".docx", SaveFormat.Docx);
                    //};
                }
            };
        }
    }
}
