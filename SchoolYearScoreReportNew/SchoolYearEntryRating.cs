namespace SchoolYearScoreReport
{
    using SmartSchool.Customization.Data;
    using System;
    using System.Xml;

    internal class SchoolYearEntryRating
    {
        // Fields
        private XmlElement _sy_ratings = null;

        // Methods
        public SchoolYearEntryRating(StudentRecord student)
        {
            if (student.Fields.ContainsKey("SchoolYearEntryClassRating"))
            {
                this._sy_ratings = student.Fields["SchoolYearEntryClassRating"] as XmlElement;
            }
        }

        public string GetPlace(int schoolYear)
        {
            if (this._sy_ratings != null)
            {
                string path = string.Format("SchoolYearEntryScore[SchoolYear='{0}']/ClassRating/Rating/Item[@分項='學業']/@排名", schoolYear);
                XmlNode result = this._sy_ratings.SelectSingleNode(path);
                if (result != null)
                {
                    return result.InnerText;
                }
            }
            return string.Empty;
        }

    }
}

