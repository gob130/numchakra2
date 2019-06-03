using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Omega.Objects;
using Omega.Enums;
using Omega.Reports;

namespace Omega.Reports
{
    public partial class frmPrintSpecialReport : Form
    {
        public frmPrintSpecialReport()
        {
            InitializeComponent();
        }

        
        private void frmPrintSpecialReport_Load(object sender, EventArgs e)
        {
            /*
            this.Icon = Omega.MainForm.ActiveForm.Icon;
            SetAllCheckBox(true);


            clbReports.Items.Clear();
            foreach (string str in ReportDataProvider.Instance.ReportTemplets)
            {
                clbReports.Items.Add(System.IO.Path.GetFileNameWithoutExtension(str));
            }

            rbtnFromApp.Checked = true;
            rbtnFromStyle.Checked = false;
            grpPart.Enabled = rbtnFromApp.Checked;
            grpStyle.Enabled = rbtnFromStyle.Checked;
            */
        }

        /*
        private void SetAllCheckBox(bool val)
        {
            cbP1.Checked = val;
            cbP2.Checked = val;
            cbP3.Checked = val;
            cbP4.Checked = val;
            cbP5.Checked = val;
            cbP6.Checked = val;
            cbP71.Checked = val;
            cbP72.Checked = val;
            cbP8.Checked = val;
        }

        private void rbtnFromApp_CheckedChanged(object sender, EventArgs e)
        {
            if (rbtnFromApp.Checked == true)
            {
                grpPart.Enabled = true;
                grpStyle.Enabled = false;
            }
            else
            {
                grpPart.Enabled = false;
                grpStyle.Enabled = true;
            }
            
        }

        private void ChangePrintValues()
        {
            ReportDataProvider.Instance.PrntMain  = cbP1.Checked;
            ReportDataProvider.Instance.PrntLifeCycles = cbP2.Checked;
            ReportDataProvider.Instance.PrntPitagoras = cbP3.Checked;
            ReportDataProvider.Instance.PrntIntansiveMap = cbP4.Checked;
            ReportDataProvider.Instance.PrntChakraOpening = cbP5.Checked;
            ReportDataProvider.Instance.PrntCoupleMatch = cbP6.Checked;
            ReportDataProvider.Instance.PrntBsnnsMulti = cbP72.Checked;
            ReportDataProvider.Instance.PrntBsnnsPersonal = cbP71.Checked;
            ReportDataProvider.Instance.PrntLearnSccss = cbP8.Checked;
        }

        private void btnPrintRegularReport_Click(object sender, EventArgs e)
        {
            ChangePrintValues();

            this.Hide();
            ReportDataProvider.Instance.mMainForm.Show();
            ReportDataProvider.Instance.mMainForm.Focus();

            ReportDataProvider.Instance.PrintReport();

            ReportDataProvider.Instance.Set2PrintAll();
            SetAllCheckBox(true);

            this.Close();
        }

        private void btnPrintFromXML_Click(object sender, EventArgs e)
        {
            this.Hide();
            foreach (string sRepStyleName in clbReports.SelectedItems)
            {
                ReportDataProvider.Instance.mMainForm.Show();
                ReportDataProvider.Instance.mMainForm.Focus();
                

                ReportDataProvider.Instance.PrintReportFromXML(sRepStyleName.ToString());
            }
            
            //ReportDataProvider.Instance.Set2PrintAll();
            //SetAllCheckBoxTrue();

            this.Close();
            MessageBox.Show("הדוחות שנבחרו לייצור נוצרו כולם", "", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            SetAllCheckBox(true);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            SetAllCheckBox(false);
        }
         
         */
    }
    
}
