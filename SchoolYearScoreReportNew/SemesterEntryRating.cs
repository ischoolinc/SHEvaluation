namespace SchoolYearScoreReport
{
    using SmartSchool.Customization.Data;
    using System;
    using System.Xml;

    internal class SemesterEntryRating
    {
        // Fields
        private XmlElement _sems_ratings = null;

        // Methods
        public SemesterEntryRating(StudentRecord student)
        {
            if (student.Fields.ContainsKey("SemesterEntryClassRating"))
            {
                this._sems_ratings = student.Fields["SemesterEntryClassRating"] as XmlElement;
            }
        }

        public string GetPlace(int schoolYear, int semester)
        {
            if (this._sems_ratings != null)
            {
                string path = string.Format("SemesterEntryScore[SchoolYear='{0}' and Semester='{1}']/ClassRating/Rating/Item[@分項='學業']/@排名", schoolYear, semester);
                XmlNode result = this._sems_ratings.SelectSingleNode(path);
                if (result != null)
                {
                    return result.InnerText;
                }
            }
            return string.Empty;
        }
    }


}

