using SmartSchool.Customization.Data;
using System.Collections.Generic;

namespace SmartSchool.Evaluation.Process.Wizards
{
    public interface ISubjectCalcPostProcess
    {
        void ShowConfigForm();
        void PostProcess(int schoolYear, int semester, List<StudentRecord> list);
    }
}
