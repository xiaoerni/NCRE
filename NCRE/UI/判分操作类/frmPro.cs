using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace NCRE学生考试端V1._0
{
    public partial class frmPro : Form
    {
        public static Form frmpro;
        public frmPro()
        {
            InitializeComponent();
            frmpro = this;
            timer1.Enabled = true;
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            if (progressBar1.Value < progressBar1.Maximum)
            {
                progressBar1.Value++;
            }
            else
            {
                timer1.Enabled = false;
            }
        }

        private void frmPro_Load(object sender, EventArgs e)
        {
            timer1.Enabled = true;
            progressBar1.Value = 0;
            progressBar1.Minimum = 0;
            progressBar1.Maximum = 500;
        }
    }
}
