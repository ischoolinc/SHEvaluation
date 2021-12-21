using FISCA.Data;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Reflection;


namespace 班級定期評量成績單_固定排名.Model
{
    /// <summary>
    /// 級距資訊
    /// </summary>
    class FixRankIntervalInfo
    {
        /// <summary>
        ///  $"{itemType}^^^{intervalInfo.RankType}^^^grade{intervalInfo.GradeYear}^^^{intervalInfo.RankName}^^^{intervalInfo.ItemName}";
        /// </summary>
        public string Key { get; set; }

        /// <summary>
        /// 班級ID 
        /// </summary>
        public int ClassID { get; set; }

        /// <summary>
        /// 班級名稱(rank_name) 
        /// </summary>
        public string ClassName { get; set; }

        /// <summary>
        /// 科別
        /// </summary>
        public string Dept { get; set; }

        /// <summary>
        /// 年級
        /// </summary>
        public string GradeYear { get; set; }


        /// <summary>
        /// 班排名、科排名、校排名
        /// </summary>
        public string RankType { get; set; }


        /// <summary>
        ///  母群的名稱(ex :普101 、普通科、一年級)
        /// </summary>
        public string RankName { get; set; }


        /// <summary>
        /// 總計成績 OR 科目成績
        /// </summary>
        public string ItemType { get; set; }



        /// <summary>
        /// 排名的標的 (ex :加權平均、平均、加權總分、總分、數學、英文...等等) 
        /// </summary>
        public string ItemName { get; set; }

        /// <summary>
        /// 母群總數
        /// </summary>
        public string Matrix_count { get; set; }

        /// <summary>
        /// 頂標
        /// </summary>
        public string Avg_top_25 { get; set; }

        /// <summary>
        /// 高標
        /// </summary>
        public string Avg_top_50 { get; set; }

        /// <summary>
        /// 均標(平均值)
        /// </summary>
        public string Avg { get; set; }

        /// <summary>
        /// 低標
        /// </summary>
        public string Avg_bottom_50 { get; set; }

        /// <summary>
        /// 底標
        /// </summary>
        public string Avg_bottom_25 { get; set; }

        /// <summary>
        /// 組距人數>=100
        /// </summary>
        public string Level_gte100 { get; set; }

        /// <summary>
        /// 組距人數>=90， 大於100
        /// </summary>
        public string Level_90 { get; set; }

        /// <summary>
        /// 組距人數>=80，大於90
        /// </summary>
        public string Level_80 { get; set; }

        /// <summary>
        /// 組距人數>=70，大於80
        /// </summary>
        public string Level_70 { get; set; }

        /// <summary>
        /// 組距人數>=60，大於70
        /// </summary>
        public string Level_60 { get; set; }


        /// <summary>
        /// 組距人數>=50，大於60
        /// </summary>
        public string Level_50 { get; set; }

        /// <summary>
        /// 組距人數>=40，大於50
        /// </summary>
        public string Level_40 { get; set; }


        /// <summary>
        /// 組距人數>=30，大於40
        /// </summary>
        public string Level_30 { get; set; }

        /// <summary>
        /// 組距人數>=20，大於30
        /// </summary>
        public string Level_20 { get; set; }

        /// <summary>
        /// 組距人數>=10，大於20
        /// </summary>
        public string Level_10 { get; set; }

        /// <summary>
        /// 標準差
        /// </summary>
        public string Std_dev_pop { get; set; }

        /// <summary>
        /// 新頂標
        /// </summary>
        public string Pr_88 { get; set; }

        /// <summary>
        /// 新前標
        /// </summary>
        public string Pr_75 { get; set; }

        /// <summary>
        /// 新均標
        /// </summary>
        public string Pr_50 { get; set; }

        /// <summary>
        /// 新後標
        /// </summary>
        public string Pr_25 { get; set; }

        /// <summary>
        /// 新底標
        /// </summary>
        public string Pr_12 { get; set; }

        public string _Level_lt10 { get; set; }



        /// <summary>
        /// 
        /// </summary>
        /// 


        public string Level_lt10
        {

            get { return _Level_lt10; }
            set
            {
                _Level_lt10 = value;
                //RoundNumber();
                this.PutIvtervalValueToDic();
                this.CaculateAll();
            }
        }



        /// <summary>
        ///  計算UP DOWN用
        /// </summary>
        public Dictionary<int, int> DicCaculateUPDown { get; set; }

        /// <summary>
        /// 儲存以上 的分
        /// </summary>
        public Dictionary<int, int> DicCaculateUpResult { get; set; }

        /// <summary>
        /// 儲存以下
        /// </summary>
        public Dictionary<int, int> DicCaculateDoownResult { get; set; }



        /// <summary>
        /// 取得Key值
        /// </summary>
        /// <returns></returns>
        public string GetKey()
        {
            //Dictionary Key 值
            //string key = $" {intervalInfo.RankType}^^^grade{intervalInfo.GradeYear}^^^{intervalInfo.RankName}^^^{intervalInfo.ItemName}";

            string result = "";

            result = $"{RankType}^^^grade{GradeYear}^^^{RankName}^^^{ItemName}";

            return result;
        }



        /// <summary>
        /// 建構子
        /// </summary>
        /// <param name="rankType"></param>
        /// <param name="rankName"></param>
        public FixRankIntervalInfo(string rankType, string rankName, string gradeYear, string itemName)
        {
            this.ItemName = itemName;
            this.RankType = rankType;
            this.RankName = rankName;
            this.GradeYear = gradeYear;

            //new 統計以上以下 之Dictionary

            this.DicCaculateUPDown = new Dictionary<int, int>();
            this.DicCaculateUpResult = new Dictionary<int, int>();
            this.DicCaculateDoownResult = new Dictionary<int, int>();
        }


        /// <summary>
        /// 計算 幾分以尚
        /// </summary>
        /// <param name="upOrDown"> 以下 或 以上</param>
        /// <param name="scoreStardand"> 標準(幾分)以下或以上</param>
        public int CalculateUpdownCount(string upOrDown, int scoreStardand)
        {
            int totalCount = 0;
            if (upOrDown == "up")
            {
                foreach (int scoreInterval in this.DicCaculateUPDown.Keys)
                {
                    if (scoreInterval >= scoreStardand)
                    {
                        totalCount += this.DicCaculateUPDown[scoreInterval];
                    }
                }
                return totalCount;
            }
            else if (upOrDown == "down")
            {
                foreach (int scoreInterval in this.DicCaculateUPDown.Keys)
                {
                    if (scoreInterval < scoreStardand)
                    {
                        totalCount += this.DicCaculateUPDown[scoreInterval];
                    }
                }

                return totalCount;
            }
            return 0;
        }



        /// <summary>
        /// 放入字典
        /// </summary>
        public void PutIvtervalValueToDic()
        {
            this.DicCaculateUPDown.Add(100, ParseInt(this.Level_gte100));
            this.DicCaculateUPDown.Add(90, ParseInt(this.Level_90));
            this.DicCaculateUPDown.Add(80, ParseInt(this.Level_80));
            this.DicCaculateUPDown.Add(70, ParseInt(this.Level_70));
            this.DicCaculateUPDown.Add(60, ParseInt(this.Level_60));
            this.DicCaculateUPDown.Add(50, ParseInt(this.Level_50));
            this.DicCaculateUPDown.Add(40, ParseInt(this.Level_40));
            this.DicCaculateUPDown.Add(30, ParseInt(this.Level_30));
            this.DicCaculateUPDown.Add(20, ParseInt(this.Level_20));
            this.DicCaculateUPDown.Add(10, ParseInt(this.Level_10));
            this.DicCaculateUPDown.Add(0, ParseInt(this.Level_lt10));

        }

        /// <summary>
        /// 將字串 轉成INT
        /// </summary>
        /// <param name="count"></param>
        /// <returns></returns>
        private int ParseInt(string count)
        {
            int countInt;

            bool success = Int32.TryParse(count, out countInt);

            if (success)
            {
                return countInt;
            }
            else
            {
                return 0;//轉型不成功 就回傳0
            }
        }



        private double ParseDouble(string count)
        {
            double countInt;

            bool success = Double.TryParse(count, out countInt);

            if (success)
            {
                return countInt;
            }
            else
            {
                return 0;//轉型不成功 就回傳0
            }
        }




        /// <summary>
        /// 計算所有的 把資料放進去
        /// </summary>
        private void CaculateAll()
        {

            foreach (int score in DicCaculateUPDown.Keys)
            {

                int upCount = CalculateUpdownCount("up", score);

                this.DicCaculateUpResult.Add(score, upCount);

                int downCount = CalculateUpdownCount("down", score);

                this.DicCaculateDoownResult.Add(score, downCount);

            }
        }

        /// <summary>
        /// 處理四捨五入
        /// </summary>
        private void RoundNumber()
        {
            //this.Avg_top_25 = Math.Round(ParseDouble(Avg_top_25), 2) == 0 ? "" : Math.Round(ParseDouble(Avg_top_25), 2).ToString();
            //this.Avg_top_50 = Math.Round(ParseDouble(Avg_top_50), 2) == 0 ? "" : Math.Round(ParseDouble(Avg_top_50), 2).ToString();
            //this.Avg = Math.Round(ParseDouble(Avg), 2) == 0 ? "" : Math.Round(ParseDouble(Avg), 2).ToString();
            //this.Avg_bottom_25 = Math.Round(ParseDouble(Avg_bottom_25), 2) == 0 ? "" : Math.Round(ParseDouble(Avg_bottom_25), 2).ToString();
            //this.Avg_bottom_50 = Math.Round(ParseDouble(Avg_bottom_50), 2) == 0 ? "" : Math.Round(ParseDouble(Avg_bottom_50), 2).ToString(); 
        }

    }
}
