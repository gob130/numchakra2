using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace Omega.Objects
{
    public partial class qnaPickWorkType : Form
    {
        private MainForm mf;
        public qnaPickWorkType(MainForm mainform)
        {
            InitializeComponent();
            mf = mainform;
        }

        private void qnaPickWorkType_Load(object sender, EventArgs e)
        {
            cmbSelectWorkType.Items.Clear();
            foreach (string s in mf.sWorkTypes)
            {
                cmbSelectWorkType.Items.Add(s);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (cmbSelectWorkType.SelectedItem != null)
            {
                mf.sQnAresWorkType = cmbSelectWorkType.SelectedItem.ToString();
            }
            this.Close();
            mf.cmdQnAcalc_Click(null, null);
        }
    }
}
