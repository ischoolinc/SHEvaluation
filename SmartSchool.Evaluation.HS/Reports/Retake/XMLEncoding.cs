using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SmartSchool.Evaluation.Reports.Retake
{
    public class XMLEncoding
    {
        public static string Encoding(string str)
        {
            string retVal = str.Replace("'", "");

            return retVal;
        }
    }
}
