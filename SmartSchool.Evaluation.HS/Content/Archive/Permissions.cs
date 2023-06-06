namespace SmartSchool.Evaluation.Content
{
    class Permissions
    {
        public static string 學期成績封存 { get { return "SmartSchool.Evaluation.Content.SemesterScoreDataAchive"; } }
        public static bool 學期成績封存權限
        {
            get
            {
                return FISCA.Permission.UserAcl.Current[學期成績封存].Executable;
            }
        }
    }
}
