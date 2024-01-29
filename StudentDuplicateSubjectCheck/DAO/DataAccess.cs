using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using FISCA.Data;

namespace StudentDuplicateSubjectCheck.DAO
{
    public class DataAccess
    {
        // 透過學生ID取得學生類別資訊
        public static Dictionary<string, List<StudentTagInfo>> GetStudentTagInfoDictByStduentIDs(List<string> sids)
        {
            Dictionary<string, List<StudentTagInfo>> value = new Dictionary<string, List<StudentTagInfo>>();

            if (sids.Count > 0)
            {
                string sql = string.Format(@"
                SELECT
                    tag_student.ref_student_id AS student_id,
                    tag.prefix AS tag_prefix,
                    tag.name AS tag_name
                FROM
                    tag
                    INNER JOIN tag_student ON tag.id = tag_student.ref_tag_id
                WHERE
                    tag_student.ref_student_id IN ({0})
                ORDER BY
                    tag_student.ref_student_id,
                    tag.prefix,
                    tag.name;	
                ", string.Join(",", sids.ToArray()));

                QueryHelper qh = new QueryHelper();
                DataTable dt = qh.Select(sql);

                foreach (DataRow dr in dt.Rows)
                {
                    StudentTagInfo st = new StudentTagInfo();
                    st.StudentID = dr["student_id"] + "";
                    st.Prefix = dr["tag_prefix"] + "";
                    st.Name = dr["tag_name"] + "";
                    st.FullName = st.Prefix + ":" + st.Name;
                    if (!value.ContainsKey(st.StudentID))
                        value.Add(st.StudentID, new List<StudentTagInfo>());

                    value[st.StudentID].Add(st);
                }
            }

            return value;

        }
    }
}
