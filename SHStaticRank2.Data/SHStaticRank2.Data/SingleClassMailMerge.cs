using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;

namespace SHStaticRank2.Data
{
    public class SingleClassMailMerge : Aspose.Words.Reporting.IMailMergeDataSource
    {
        private Dictionary<string, string> ClassData { get; set; }
        private DataTable Students { get; set; }
        private int offset = -1;

        public SingleClassMailMerge(Dictionary<string, string> classData, DataTable students)
        {
            ClassData = classData; //班級相關資料。
            Students = students; //學生排名相關資料。
        }

        #region IMailMergeDataSource 成員

        public Aspose.Words.Reporting.IMailMergeDataSource GetChildDataSource(string tableName)
        {
            return new StudentMailMerge(Students);
        }

        public bool GetValue(string fieldName, out object fieldValue)
        {
            if (ClassData.ContainsKey(fieldName))
                fieldValue = ClassData[fieldName];
            else
                fieldValue = string.Empty;

            return true;
        }

        public bool MoveNext()
        {
            offset++;

            return offset == 0; //只有一筆資料。
        }

        public string TableName { get { return "Class"; } }

        #endregion
    }

    public class StudentMailMerge : Aspose.Words.Reporting.IMailMergeDataSource
    {
        private DataTable Students { get; set; }

        private int offset = -1;

        public StudentMailMerge(DataTable students)
        {
            Students = students;
        }

        #region IMailMergeDataSource 成員

        public Aspose.Words.Reporting.IMailMergeDataSource GetChildDataSource(string tableName)
        {
            return null;
        }

        public bool GetValue(string fieldName, out object fieldValue)
        {
            if (Students.Columns.Contains(fieldName))
                fieldValue = Students.Rows[offset][fieldName];
            else
                fieldValue = string.Empty;

            return true;
        }

        public bool MoveNext()
        {
            offset++;
            return offset < Students.Rows.Count;
        }

        public string TableName
        {
            get { return "Student"; }
        }
        #endregion
    }
}
