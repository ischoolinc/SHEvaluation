using Campus.DocumentValidator;

namespace SHImportExportScoreRank.ValidationRule
{
    public class SchoolYearEntryRankRowValidatorFactory : IRowValidatorFactory
    {
        IRowVaildator IRowValidatorFactory.CreateRowValidator(
            string typeName,
            System.Xml.XmlElement validatorDescription)
        {
            switch ((typeName ?? string.Empty).ToUpper())
            {
                case "SHIMPORTEXPORTSCORERANKCHECKSTUDENT":
                    return new RowValidator.CheckSchoolYearEntryRankStudentValidator();
                default:
                    return null;
            }
        }
    }
}
