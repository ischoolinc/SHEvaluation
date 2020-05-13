using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FISCA.Data;
using System.Data;
using System.IO;

namespace ClassSemesterScoreReportFixed_SH
{
    public class Utility
    {
        public static string ParseFileName(string fileName)
        {
            string name = fileName;

            if (fileName == null)
                throw new ArgumentNullException();

            if (name.Length == 0)
                throw new ArgumentException();

            if (name.Length > 245)
                throw new PathTooLongException();

            foreach (char c in System.IO.Path.GetInvalidFileNameChars())
            {
                name = name.Replace(c, '_');
            }

            return name;
        }


    }
}
