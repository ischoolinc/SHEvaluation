using System.Security;

namespace SmartSchool.Evaluation.Reports.Retake
{
    public class XMLEncoding
    {
        public static string Encoding(string str)
        {
            return SecurityElement.Escape(str).Replace("\r", "&#xD;").Replace("\n", "&#xA;");
        }
    }
}
