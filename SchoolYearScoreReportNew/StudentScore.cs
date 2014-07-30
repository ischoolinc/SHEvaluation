namespace SchoolYearScoreReport
{
    using SmartSchool.Common;
    using SmartSchool.Customization.Data;
    using SmartSchool.Customization.Data.StudentExtension;
    using System;
    using System.Collections.Generic;

    internal class StudentScore
    {
        // Fields
        private int _beforeTotalCredit = 0;
        private Config _config;
        private Dictionary<string, ScoreData> _entries = new Dictionary<string, ScoreData>();
        private int _firstTotalCredit = 0;
        private int _grade_year;
        private int _secondTotalCredit = 0;
        private Dictionary<int, decimal> _standard;
        private Dictionary<string, ScoreData> _subjects = new Dictionary<string, ScoreData>();

        // Methods
        public StudentScore(Config config, Dictionary<int, decimal> standard, int currentGradeYear)
        {
            this._config = config;
            this._standard = standard;
            this._grade_year = currentGradeYear;
        }

        public void AddEntry(SchoolYearEntryScoreInfo info)
        {
            if (info.SchoolYear == ((int)this._config.SchoolYear))
            {
                if (!this._entries.ContainsKey(info.Entry))
                {
                    this._entries.Add(info.Entry, new ScoreData());
                }
                ScoreData data = this._entries[info.Entry];
                decimal score = info.Score;
                if (!((info.Score <= 100M) || this._config.AllowOver))
                {
                    score = 100M;
                }
                data.SetInfo(info.Entry, score);
            }
        }

        public void AddEntry(SemesterEntryScoreInfo info)
        {
            if (info.GradeYear == this._grade_year)
            {
                if (!this._entries.ContainsKey(info.Entry))
                {
                    this._entries.Add(info.Entry, new ScoreData());
                }
                ScoreData data = this._entries[info.Entry];
                decimal score = info.Score;
                if (!((info.Score <= 100M) || this._config.AllowOver))
                {
                    score = 100M;
                }
                data.SetInfo(info.Entry, info.Semester, score);
            }
        }

        public void AddSubject(SchoolYearSubjectScoreInfo info)
        {
            if (info.SchoolYear == ((int)this._config.SchoolYear))
            {
                bool found = false;
                foreach (ScoreData var in this._subjects.Values)
                {
                    if (var.Name == info.Subject)
                    {
                        var.SetInfo(info.Subject, info.Score);

                        string sign = "";
                        decimal rS1=-1, rS2=-1;
                        
                        // 判斷是否來自學年補考或重修
                        if (decimal.TryParse(info.Detail.GetAttribute("補考成績"), out rS1))
                        {
                            if (info.Score == rS1)
                                sign = this._config.YearResitSign;
                        }

                        if (decimal.TryParse(info.Detail.GetAttribute("重修成績"), out rS2))
                        {
                            if (info.Score == rS2)
                                sign += this._config.YearRepeatSign;
                        }

                        // 學年成績標示
                        var.SetYearScoreSign(sign);

                        found = true;
                    }
                }
                if (!found)
                {
                    if (!this._subjects.ContainsKey(info.Subject))
                    {
                        this._subjects.Add(info.Subject, new ScoreData());
                    }
                    this._subjects[info.Subject].SetInfo(info.Subject, info.Score);
                }
            }
        }

        public void AddSubject(SemesterSubjectScoreInfo info)
        {
            if (info.Detail.GetAttribute("不計學分") != "是")
            {
                bool noScore = info.Detail.GetAttribute("不需評分") != "是";
                if (info.GradeYear == this._grade_year)
                {
                    ScoreData data = null;
                    bool addNewOne = true;
                    foreach (ScoreData var in this._subjects.Values)
                    {
                        if (var.Name == info.Subject)
                        {
                            addNewOne = false;
                            if (info.Semester == 1)
                            {
                                if (!((var.FirstCredit <= 0) && string.IsNullOrEmpty(var.FirstRequire)))
                                {
                                    addNewOne = true;
                                }
                            }
                            else if ((info.Semester == 2) && !((var.SecondCredit <= 0) && string.IsNullOrEmpty(var.SecondRequire)))
                            {
                                addNewOne = true;
                            }
                            if (!addNewOne)
                            {
                                data = var;
                                data.SetInfo(info.Subject, info.Semester, info.Require, info.Credit, noScore ? info.Score : -1M);
                                data.SetLevel(info.Level, info.Semester);
                                break;
                            }
                        }
                    }
                    if (addNewOne)
                    {
                        if (!this._subjects.ContainsKey(info.Subject + "_" + info.Level))
                        {
                            this._subjects.Add(info.Subject + "_" + info.Level, new ScoreData());
                        }
                        data = this._subjects[info.Subject + "_" + info.Level];
                        data.SetInfo(info.Subject, info.Semester, info.Require, info.Credit, noScore ? info.Score : -1M);
                        data.SetLevel(info.Level, info.Semester);
                    }
                    if (!info.Pass && this._standard.ContainsKey(info.GradeYear))
                    {
                        if (info.Score >= this._standard[info.GradeYear])
                        {
                            data.SetSign(info.Semester, this._config.ResitSign);
                        }
                        else
                        {
                            data.SetSign(info.Semester, this._config.RepeatSign);
                        }
                    }
                }
                XmlHelper helper = new XmlHelper(info.Detail);
                if (info.Pass && (helper.GetText("@不計學分") != "是"))
                {
                    if (info.GradeYear == this._grade_year)
                    {
                        if (info.Semester == 1)
                        {
                            this._firstTotalCredit += info.Credit;
                        }
                        else if (info.Semester == 2)
                        {
                            this._secondTotalCredit += info.Credit;
                        }
                    }
                    if (info.SchoolYear <= this._config.SchoolYear)
                    {
                        this._beforeTotalCredit += info.Credit;
                    }
                }
            }
        }

        public void Clear()
        {
            this._subjects.Clear();
            this._entries.Clear();
            this._standard.Clear();
        }

        public void CreditStatistic()
        {
            ScoreData data = new ScoreData();
            data.SetItem("實得學分", this._firstTotalCredit + "", this._secondTotalCredit + "", (this._firstTotalCredit + this._secondTotalCredit) + "");
            this._entries.Add("實得學分", data);
            data = new ScoreData();
            data.SetItem("累計學分", (this._beforeTotalCredit - this._secondTotalCredit) + "", this._beforeTotalCredit + "", this._beforeTotalCredit + "");
            this._entries.Add("累計學分", data);
        }

        public void Rating(StudentRecord student, int schoolyear)
        {
            SemesterEntryRating semesterRating = new SemesterEntryRating(student);
            SchoolYearEntryRating schoolYearRating = new SchoolYearEntryRating(student);
            ScoreData data = new ScoreData();
            data.SetItem("學業成績名次", semesterRating.GetPlace(schoolyear, 1), semesterRating.GetPlace(schoolyear, 2), schoolYearRating.GetPlace(schoolyear));
            this._entries.Add("學業成績名次", data);
        }

        // Properties
        public Dictionary<string, ScoreData> Entries
        {
            get
            {
                return this._entries;
            }
        }

        public Dictionary<string, ScoreData> Subjects
        {
            get
            {
                return this._subjects;
            }
        }
    }
}

