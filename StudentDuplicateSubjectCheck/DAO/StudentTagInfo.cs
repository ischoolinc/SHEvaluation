using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace StudentDuplicateSubjectCheck.DAO
{
    // 學生類別資訊
    public class StudentTagInfo
    {
        public string StudentID { get; set; }
        // 學生類別名稱  ex: 原住民
        public string Name { get; set; }
        // 學生類別完整 ex: 成績身分:原住民
        public string FullName { get; set; }
        // 學生類別前綴 ex: 成績身分
        public string Prefix { get; set; }
    }
}
