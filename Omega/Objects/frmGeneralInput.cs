using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace Omega.Objects
{
    public partial class frmGeneralInput : Form
    {
        private string mTitle;
        public frmGeneralInput(string title)
        {
            InitializeComponent();

            mTitle = title;
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            Omega.Objects.InputProvider.Instance.Value = txtInput.Text;
            this.Close();
        }

        private void frmGeneralInput_Load(object sender, EventArgs e)
        {
            txtInput.Text = "";
            this.Icon = Omega.MainForm.ActiveForm.Icon;
            this.Text  = mTitle;
        }
    }
}
