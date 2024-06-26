using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using FISCA.Presentation.Controls;

namespace SHScoreValueManager.UIForm
{
    public partial class frmNotInputChar : BaseForm
    {
        public frmNotInputChar()
        {
            InitializeComponent();
        }

        private void frmNotInputChar_Load(object sender, EventArgs e)
        {

        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
