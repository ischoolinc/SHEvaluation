using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SHStaticRank2.Data
{
    public class StudSemsSubjRating
    {
        public string StudentID { get; set; }
        public string SchoolYear { get; set; }
        public string Semester { get; set; }

        public string GradeYear { get; set; }

        // 班成績人數
        private Dictionary<string, int> _ClassCountDict = new Dictionary<string, int>();

        // 班排名
        private Dictionary<string, int> _ClassRankDict = new Dictionary<string, int>();

        public void AddClassCount(string name,int count)
        {
            if (!_ClassCountDict.ContainsKey(name))
                _ClassCountDict.Add(name, count);
        }

        public void AddClassRank(string name, int rank)
        {
            if (!_ClassRankDict.ContainsKey(name))
                _ClassRankDict.Add(name, rank);
        }

        /// <summary>
        /// 取得班成績人數
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        public int GetClassCount(string name)
        { 
            int value=0;
            if (_ClassCountDict.ContainsKey(name))
                value = _ClassCountDict[name];

            return value;
        }

        /// <summary>
        /// 取得班名次
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        public int GetClassRank(string name)
        {
            int value = 0;
            if (_ClassRankDict.ContainsKey(name))
                value = _ClassRankDict[name];

            return value;
        }
    }
}
