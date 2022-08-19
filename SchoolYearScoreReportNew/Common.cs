namespace SchoolYearScoreReport
{
    using Aspose.Words;
    using DevComponents.DotNetBar;
    using SmartSchool.Customization.Data;
    using System;
    using System.Collections.Generic;
    using System.Diagnostics;
    using System.IO;
    using System.Windows.Forms;

    public static class Common
    {
        // Methods
        public static string GetNumberString(string p)
        {
            switch (p.Trim())
            {
                case "1":
                    return "Ⅰ";

                case "2":
                    return "Ⅱ";

                case "3":
                    return "Ⅲ";

                case "4":
                    return "Ⅳ";

                case "5":
                    return "Ⅴ";

                case "6":
                    return "Ⅵ";

                case "7":
                    return "Ⅶ";

                case "8":
                    return "Ⅷ";

                case "9":
                    return "Ⅸ";

                case "10":
                    return "Ⅹ";
            }
            return p;
        }

        public static string ParseLevel(decimal score)
        {
            if (!SmartSchool.Customization.Data.SystemInformation.Fields.ContainsKey("Degree"))
            {
                SmartSchool.Customization.Data.SystemInformation.getField("Degree");
            }
            Dictionary<string, decimal> _degreeList = SmartSchool.Customization.Data.SystemInformation.Fields["Degree"] as Dictionary<string, decimal>;
            foreach (string var in _degreeList.Keys)
            {
                if (_degreeList[var] <= score)
                {
                    return (" / " + var);
                }
            }
            return "";
        }

        public static void SaveToFile(string inputReportName, Document inputDoc)
        {
            string reportName = inputReportName;
            Document doc = inputDoc;
            string path = Path.Combine(Application.StartupPath, "Reports");
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
            path = Path.Combine(path, reportName + ".doc");
            if (File.Exists(path))
            {
                int i = 1;
                while (true)
                {
                    string newPath = string.Concat(new object[] { Path.GetDirectoryName(path), @"\", Path.GetFileNameWithoutExtension(path), i++, Path.GetExtension(path) });
                    if (!File.Exists(newPath))
                    {
                        path = newPath;
                        break;
                    }
                }
            }
            try
            {
                doc.Save(path, SaveFormat.Doc);
                Process.Start(path);
            }
            catch
            {
                SaveFileDialog sd = new SaveFileDialog
                {
                    Title = "另存新檔",
                    FileName = reportName + ".doc",
                    Filter = "Word檔案 (*.doc)|*.doc|所有檔案 (*.*)|*.*"
                };
                if (sd.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        doc.Save(sd.FileName, SaveFormat.Doc);
                    }
                    catch
                    {
                        MessageBox.Show("指定路徑無法存取。", "建立檔案失敗", MessageBoxButtons.OK, MessageBoxIcon.Hand);
                    }
                }
            }
        }

        internal static DialogResult Warning(int number)
        {
            return MessageBoxEx.Show("您選取的學生超過 " + number + " 位，可能會造成系統緩慢。\n是否繼續執行？", "", MessageBoxButtons.YesNo);
        }

        // Nested Types
        internal class EntryCompaper : IComparer<ScoreData>
        {
            // Methods
            public int Compare(ScoreData x, ScoreData y)
            {
                return this.SortByEntryName(x.Name, y.Name);
            }

            private int getIntForEntry(string a1)
            {
                List<string> list = new List<string>();
                list.AddRange(new string[] { "學業", "學業成績名次", "實得學分", "累計學分", "學業(原始)", "實習科目(原始)", "實習科目", "專業科目(原始)", "專業科目", "體育", "國防通識", "健康與護理", "德行" });
                int x = list.IndexOf(a1);
                if (x < 0)
                {
                    return list.Count;
                }
                return x;
            }

            private int SortByEntryName(string a, string b)
            {
                int ai = this.getIntForEntry(a);
                int bi = this.getIntForEntry(b);
                if ((ai > 0) || (bi > 0))
                {
                    return ai.CompareTo(bi);
                }
                return a.CompareTo(b);
            }
        }

        internal class SubjectComparer : IComparer<ScoreData>
        {
            // Methods
            public int Compare(ScoreData x, ScoreData y)
            {
                return this.SortBySubjectName(x.Name, y.Name);
            }

            private int getIntForSubject(string a1)
            {
                List<string> list = new List<string>();
                list.AddRange(new string[] { 
                "國", "英", "數", "物", "化", "生", "基", "歷", "地", "公", "文", "礎", "世", "汽", "電", "美", 
                "中", "程"
             });
                int x = list.IndexOf(a1);
                if (x < 0)
                {
                    return list.Count;
                }
                return x;
            }

            private int SortBySubjectName(string a, string b)
            {
                string a1 = (a.Length > 0) ? a.Substring(0, 1) : "";
                string b1 = (b.Length > 0) ? b.Substring(0, 1) : "";
                if (a1 == b1)
                {
                    if ((a.Length > 1) && (b.Length > 1))
                    {
                        return this.SortBySubjectName(a.Substring(1), b.Substring(1));
                    }
                    return a.Length.CompareTo(b.Length);
                }
                int ai = this.getIntForSubject(a1);
                int bi = this.getIntForSubject(b1);
                if ((ai > 0) || (bi > 0))
                {
                    return ai.CompareTo(bi);
                }
                return a1.CompareTo(b1);
            }
        }
    }
}

