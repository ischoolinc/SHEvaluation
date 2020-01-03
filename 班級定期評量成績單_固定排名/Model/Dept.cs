using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace 班級定期評量成績單_固定排名.Model
{
    class Dept
    {
        /// <summary>
        /// deptID
        /// </summary>
        public string DeptID { get; set; }

        /// <summary>
        ///  Dept 名字 //普通科、商經科等等
        /// </summary>
        public string DeptName { get; set; }

            
        public Dept(string deptID,string deptName)
        {
            this.DeptID = deptID;
            this.DeptName = deptName;
        }

    }
}
