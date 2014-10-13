namespace SchoolYearScoreReport
{
    using System;
    using System.Collections.Generic;

    internal class ScoreData
    {
        private decimal _firstCredit;
        private string _firstRequire;
        private decimal _firstScore;
        private string _firstSemesterItem;
        private string _firstSign;
        private string _level;
        private string _name;
        private string _schoolYearItem;
        private decimal _schoolYearScore;
        private decimal _secondCredit;
        private string _secondRequire;
        private decimal _secondScore;
        private string _secondSemesterItem;
        private string _secondSign;
        private string _yearScoreSign;


        public void SetInfo(string name, decimal score)
        {
            this._name = name;
            this._schoolYearScore = score;
        }

        public void SetInfo(string name, int semester, decimal score)
        {
            this._name = name;
            if (semester == 1)
            {
                this._firstScore = score;
            }
            else if (semester == 2)
            {
                this._secondScore = score;
            }
        }

        public void SetInfo(string name, int semester, bool require, decimal credit, decimal score)
        {
            this._name = name;
            this._firstSign = "";
            this._secondSign = "";
            this._yearScoreSign = "";
            if (semester == 1)
            {
                this._firstCredit = credit;
                this._firstRequire = require ? "必" : "選";
                this._firstScore = score;
            }
            else if (semester == 2)
            {
                this._secondCredit = credit;
                this._secondRequire = require ? "必" : "選";
                this._secondScore = score;
            }
        }

        public void SetItem(string name, string first, string second, string schoolYear)
        {
            this._name = name;
            this._firstSemesterItem = first;
            this._secondSemesterItem = second;
            this._schoolYearItem = schoolYear;
        }

        public void SetLevel(string level, int semester)
        {
            if (!string.IsNullOrEmpty(level))
            {
                if (string.IsNullOrEmpty(this._level))
                {
                    this._level = level;
                }
                else
                {
                    string[] split = this._level.Split(new string[] { ", " }, StringSplitOptions.RemoveEmptyEntries);
                    List<string> list = new List<string>();
                    foreach (string var in split)
                    {
                        list.Add(var);
                    }
                    if (semester == 1)
                    {
                        list.Insert(0, level);
                    }
                    else if (semester == 2)
                    {
                        list.Add(level);
                    }
                    this._level = "";
                    foreach (string var in list)
                    {
                        if (!string.IsNullOrEmpty(this._level))
                        {
                            this._level = this._level + ", ";
                        }
                        this._level = this._level + var;
                    }
                }
            }
        }

        public void SetSign(int semester, string sign)
        {
            if (semester == 1)
            {
                this._firstSign = sign;
            }
            else if (semester == 2)
            {
                this._secondSign = sign;
            }
        }

        public void SetYearScoreSign(string sign)
        {
            this._yearScoreSign = sign;
        }

        public decimal FirstCredit
        {
            get
            {
                return this._firstCredit;
            }
        }

        public string FirstRequire
        {
            get
            {
                if (string.IsNullOrEmpty(this._firstRequire))
                {
                    return "";
                }
                return this._firstRequire;
            }
        }

        public decimal FirstScore
        {
            get
            {
                return this._firstScore;
            }
        }

        public string FirstSemesterItem
        {
            get
            {
                return this._firstSemesterItem;
            }
        }

        public string FirstSign
        {
            get
            {
                return (" " + this._firstSign);
            }
        }

        public string Level
        {
            get
            {
                if (string.IsNullOrEmpty(this._level))
                {
                    return "";
                }
                string level = "";
                foreach (string var in this._level.Split(new string[] { ", " }, StringSplitOptions.RemoveEmptyEntries))
                {
                    if (!string.IsNullOrEmpty(level))
                    {
                        level = level + ", ";
                    }
                    level = level + Common.GetNumberString(var);
                }
                return level;
            }
        }

        public string Name
        {
            get
            {
                if (string.IsNullOrEmpty(this._name))
                {
                    return "";
                }
                return this._name;
            }
        }

        public string SchoolYearItem
        {
            get
            {
                return this._schoolYearItem;
            }
        }

        public decimal SchoolYearScore
        {
            get
            {
                return this._schoolYearScore;
            }
        }

        public decimal SecondCredit
        {
            get
            {
                return this._secondCredit;
            }
        }

        public string SecondRequire
        {
            get
            {
                if (string.IsNullOrEmpty(this._secondRequire))
                {
                    return "";
                }
                return this._secondRequire;
            }
        }

        public decimal SecondScore
        {
            get
            {
                return this._secondScore;
            }
        }

        public string SecondSemesterItem
        {
            get
            {
                return this._secondSemesterItem;
            }
        }

        public string SecondSign
        {
            get
            {
                return (" " + this._secondSign);
            }
        }

        public string YearScoreSign
        {
            get { return (" " + this._yearScoreSign); }
        }
    }
}

