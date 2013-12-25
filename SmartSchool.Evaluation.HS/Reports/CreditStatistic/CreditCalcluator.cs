using System;
using System.Collections.Generic;
using System.Text;
using SmartSchool.Customization.Data;
using SmartSchool.Customization.Data.StudentExtension;

namespace SmartSchool.Evaluation.Reports.CreditStatistic
{
    internal class CreditCalcluator
    {
        public CreditCalcluator(StudentRecord student)
        {
            foreach (SemesterSubjectScoreInfo eachInfo in student.SemesterSubjectScoreList)
            {
                //不計學分當沒看到
                if ( eachInfo.Detail.GetAttribute("不計學分") == "是" )
                    continue;
                _total_credit += eachInfo.Credit;

                if (eachInfo.Require)
                    _total_require_credit += eachInfo.Credit;

                if (eachInfo.Pass)
                {
                    if (eachInfo.Require)
                        _passed_required_credit += eachInfo.Credit;
                    else
                        _passed_select_credit += eachInfo.Credit;
                }
            }
        }

        private int _total_credit;
        public int TotalCredit
        {
            get { return _total_credit; }
        }

        private int _total_require_credit;
        public int TotalRequireCredit
        {
            get { return _total_require_credit; }
        }

        private int _passed_required_credit;
        public int PassedRequiredCredit
        {
            get { return _passed_required_credit; }
        }

        private int _passed_select_credit;
        public int PassedSelectCredit
        {
            get { return _passed_select_credit; }
        }

        public int TotalPassedCredit
        {
            get { return PassedRequiredCredit + PassedSelectCredit; }
        }

        public int RequiredRestCredit
        {
            get { return TotalRequireCredit - PassedRequiredCredit; }
        }

    }
}
