using DevComponents.DotNetBar;
using System.Windows.Forms;

namespace SmartSchool.Evaluation.Process
{
    public partial class RibbonBarBase : UserControl, IProcess
    {
        public RibbonBarBase()
        {
            InitializeComponent();
        }

        #region IProcess 成員

        public virtual string ProcessTabName
        {
            get { return "未定"; }
        }

        public virtual RibbonBar ProcessRibbon
        {
            get { return MainRibbonBar; }
        }


        private double _Level = 1;
        public virtual double Level
        {
            get { return _Level; }
            set { _Level = value; }
        }
        #endregion
    }
}
