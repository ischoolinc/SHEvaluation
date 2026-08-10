namespace SHImportExportScoreRank.DAO
{
    public static class SqlLiteralHelper
    {
        public static string Escape(string value)
        {
            if (value == null)
                return string.Empty;

            return value.Replace("'", "''");
        }
    }
}
