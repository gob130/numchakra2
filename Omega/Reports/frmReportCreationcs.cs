#region Using
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Xml;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;
//using Microsoft.Office.Core;
using Microsoft.Office.Interop;
using Omega.Objects;
using Omega.Enums;
using System.IO;
#endregion Using

namespace Omega.Reports
{
    public partial class frmReportCreationcs : Form
    {
        private object mMissing = System.Reflection.Missing.Value;
        private Microsoft.Office.Interop.Word.ApplicationClass mWrdApp = new Microsoft.Office.Interop.Word.ApplicationClass();
        
        private Microsoft.Office.Interop.Word.Font curFont = new Microsoft.Office.Interop.Word.FontClass();
        private Microsoft.Office.Interop.Word.Font mDefaltFont = new Microsoft.Office.Interop.Word.FontClass();
        private XmlDocument curRep;
        private XmlElement curPage;
        private XmlElement mRepRoot;
        private string curRepName;
        private string mNewString = "הכנס שם חדש";

        #region Constructor
        public frmReportCreationcs()
        {
            InitializeComponent();
        }

        private void frmReportCreationcs_Load(object sender, EventArgs e)
        {
            curFont = new Microsoft.Office.Interop.Word.FontClass();
            curFont.Bold = 0;
            curFont.Italic = 0;
            curFont.Underline = WdUnderline.wdUnderlineNone;
            curFont.Size = 12;
            curFont.Name = "Times New Roman";
            
            mDefaltFont = curFont;

            #region Loading Report Styles
            if (ReportDataProvider.Instance.ReportTemplets.Count > 0)
            {
                foreach (string strRep in ReportDataProvider.Instance.ReportTemplets)
                {
                    object o = System.IO.Path.GetFileNameWithoutExtension(strRep);
                    cmbxExsistingReport.Items.Add(o);
                }
                grpReportInfo.Enabled = true;

                cmbxExsistingReport.Text = "Select A Report";
                cmbxExsistingReport.Enabled = true;
            }
            else
            {
                grpReportInfo.Enabled = false;
                cmbxExsistingReport.Enabled = true;
                cmbxExsistingReport.Text = mNewString;
            }
            #endregion Loading Report Styles

            ShowCurrentFontOnLabel();

            #region Upadte Reserved Images
            cmbxImageFromControl.Items.Clear();
            List<string> lsTmp =  ReportDataProvider.Instance.GetAllReservedImagesNames();
            foreach (string str in lsTmp)
            {
                cmbxImageFromControl.Items.Add(str.Replace("_", " ").Trim());
            }
            #endregion

            cmbxExsistingReport.Text = mNewString;

            grpPageInfo.Enabled = false;
            grpReportInfo.Enabled = false;
        }
        #endregion Constructor

        private void button1_Click(object sender, EventArgs e)
        {
            mWrdApp.Quit(ref mMissing, ref mMissing, ref mMissing);
            this.Close();
        }

        #region Report stuff
        private void UpdatePages(XmlDocument doc)
        {
            cmbxExsistingPages.Items.Clear();
            foreach (XmlNode page in doc.ChildNodes[1].ChildNodes)
            {
                if (page.Name.Substring(0, 7).ToLower() == "docpage")
                {
                    XmlAttributeCollection atts = page.Attributes;

                    cmbxExsistingPages.Items.Add(atts[0].Value.ToString());
                }
            }
        }

        #endregion Report stuff

        /// <summary>
        /// converts Drawing.Font to FontClass (Office_Interp)
        /// </summary>
        /// <param name="inFont">Drawing.Font object</param>
        /// <returns>FontClass</returns>
        private Microsoft.Office.Interop.Word.Font DrawingFont2FontClass(System.Drawing.Font inFont)
        {
            Microsoft.Office.Interop.Word.Font tmpFont = new Microsoft.Office.Interop.Word.FontClass();
            tmpFont.Bold = Convert.ToInt16(inFont.Bold);
            tmpFont.Italic = Convert.ToInt16(inFont.Italic);
            if (inFont.Underline == false)
            {
                curFont.Underline = WdUnderline.wdUnderlineNone;
            }
            else
            {
                curFont.Underline = WdUnderline.wdUnderlineSingle;
            }
            tmpFont.Name = inFont.Name;
            tmpFont.Size = inFont.Size;

            return tmpFont;
        }

        private void ShowCurrentFontOnLabel()
        {
            lblFontInfo.Text = "Font: " + curFont.Name + " (" + curFont.Size.ToString() + ")" + Environment.NewLine;
            if (curFont.Bold == 1)
            {
                lblFontInfo.Text += "[Bold]";
            }
            if (curFont.Italic == 1)
            {
                lblFontInfo.Text += "[Italic]";
            }
            if (curFont.Underline == WdUnderline.wdUnderlineSingle)
            {
                lblFontInfo.Text += "[Underline]";
            }
        }

        private void rbtnText_CheckedChanged(object sender, EventArgs e)
        {
            if (rbtnText.Checked == false)
            {
                grpTextAdd.Enabled = false;
                grpImageAdd.Enabled = true;
            }
            else
            {
                grpTextAdd.Enabled = true;
                grpImageAdd.Enabled = false;
            }
        }

        private void btnFontSelect_Click(object sender, EventArgs e)
        {
            FontDialog fd = new FontDialog();
            fd.ShowColor = false;
            fd.ShowEffects = true;

            DialogResult dlgRes = fd.ShowDialog();
            if ((dlgRes == DialogResult.Cancel) || (dlgRes == DialogResult.Abort) || (dlgRes == DialogResult.No))
            {
            }
            if (dlgRes == System.Windows.Forms.DialogResult.OK)
            {
                curFont = DrawingFont2FontClass(fd.Font);
            }

            ShowCurrentFontOnLabel();

        }

        private void btnCreateNewReport_Click(object sender, EventArgs e)
        {
            if (cmbxExsistingReport.Text == mNewString)
            {
                MessageBox.Show("חובה להכניס שם לסגנון הדוח", "טעות קלט", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            else
            {
                curRep = ReportDataProvider.Instance.CreateNewReport(cmbxExsistingReport.Text.Trim());

                //mRepRoot = curRep.CreateElement("Report");
                //mRepRoot.SetAttribute("ReportName",cmbxExsistingReport.Text.Trim());

                //curRep.AppendChild(mRepRoot);

                mRepRoot = (XmlElement) curRep.ChildNodes[1];
                curRepName = cmbxExsistingReport.Text.Trim();

                btnCreateNewReport.Enabled = false;
                cmbxExsistingReport.Enabled = false;
                grpReportInfo.Enabled = true;

                UpdatePages(curRep);
                grpPageInfo.Enabled = false;                
            }
        }

        private void btnNewPage_Click(object sender, EventArgs e)
        {
            curPage = curRep.CreateElement("", "DocPage", "");
            curPage.SetAttribute("name", (mRepRoot.ChildNodes.Count + 1).ToString());

            cmbxExsistingReport.Enabled = false;
            grpPageInfo.Enabled = true;

            UpdatePages(curRep);
        }
        
        private void cmbxExsistingReport_SelectedIndexChanged(object sender, EventArgs e)
        {
            curRep = ReportDataProvider.Instance.GetXmlDocByName(cmbxExsistingReport.SelectedItem.ToString());
            UpdatePages(curRep);
            mRepRoot = (XmlElement)curRep.ChildNodes[1];

            grpReportInfo.Enabled = true;
            btnNewPage.Enabled = true;
        }

        private void btnSaveXmlDoc_Click(object sender, EventArgs e)
        {
            if (curRep.ChildNodes.Count > 1)
            {
                curRep.ReplaceChild(mRepRoot, curRep.ChildNodes[1]);
            }
            else
            {
                curRep.AppendChild(mRepRoot);
            }

            bool res = ReportDataProvider.Instance.SaveReport(curRepName, curRep);
            if (res == false)
            {
            }

            rbtnText.Checked = true;

            btnCreateNewReport.Enabled = true;
            cmbxExsistingReport.Enabled = true;

            grpReportInfo.Enabled = false;
            grpPageInfo.Enabled = false;
            
        }

        private void btnSavePage_Click(object sender, EventArgs e)
        {
            bool found = false;
            
            foreach (XmlElement element in mRepRoot.ChildNodes)
            {
                if (element.Attributes.Item(0).Value.ToLower() == curPage.Attributes.Item(0).Value.ToLower())
                {
                    found = true;
                    mRepRoot.ReplaceChild(curPage, element);
                }
            }
            if (found == false)
            {
                mRepRoot.AppendChild(curPage);
            }

            UpdatePages(curRep);
        }

        private void rbtnImage_Prog_CheckedChanged(object sender, EventArgs e)
        {
            if (rbtnImage_Prog.Checked == true)
            {
                cmbxImageFromControl.Enabled = true;
                btnFileSelect.Enabled = false;
                txtFileLocation.Enabled = false;
            }
            else
            {
                cmbxImageFromControl.Enabled = false;
                btnFileSelect.Enabled = true;
                txtFileLocation.Enabled = true;
            }
        }

        private void btnFileSelect_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Title = "בחר תמונה";
            ofd.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            ofd.Multiselect = false;
            ofd.Filter = "All Image Types|*.bmp;*.jpeg;*.jpg;*.ico;*.png;*.tiff";

            DialogResult res = ofd.ShowDialog();
            if (res == System.Windows.Forms.DialogResult.OK)
            {
                txtFileLocation.Text = ofd.FileName;
            }
            else
            {
                return;
            }
        }

        private void btnAddText_Click(object sender, EventArgs e)
        {
            XmlElement txtNode = curRep.CreateElement("", "Text", "");
            txtNode.AppendChild(curRep.CreateTextNode(txtBody.Text.Trim()));
            
            txtNode.SetAttribute("font", curFont.Name.Trim());
            txtNode.SetAttribute("size", curFont.Size.ToString().Trim());

            string val = "";
            if (curFont.Bold == 1)
            {
                val = true.ToString();
            }
            else
            {
                val = false.ToString();
            }
            txtNode.SetAttribute("bold", val.ToLower());

            val = "";
            if (curFont.Italic == 1)
            {
                val = true.ToString();
            }
            else
            {
                val = false.ToString();
            }
            txtNode.SetAttribute("italic", val.ToLower());

            val = "";
            if (curFont.Underline  == WdUnderline.wdUnderlineSingle)
            {
                val = true.ToString();
            }
            else
            {
                val = false.ToString();
            }
            txtNode.SetAttribute("underline", val.ToLower());

            val = "";
            if (rbtnCenter.Checked == true)
            {
                val = "center";
            }
            else
            {
                if (rbtnRight.Checked == true)
                {
                    val = "right";
                }
                else
                {
                    if (rbtnLeft.Checked == true)
                    {
                        val = "left";
                    }
                }
            }
            txtNode.SetAttribute("alignment", val.ToLower());


            curPage.AppendChild(txtNode);
            txtBody.Clear();
        }

        private void btnAddImage_Click(object sender, EventArgs e)
        {
            XmlElement imgNode = curRep.CreateElement("","Image","");
            if (rbtnImage_File.Checked == true)
            {
                imgNode.AppendChild(curRep.CreateTextNode(txtFileLocation.Text.Trim()));
                txtFileLocation.Clear();
            }
            else // from app
            {
                string str = cmbxImageFromControl.SelectedItem.ToString();
                str.Replace(" ", "_");

                imgNode.AppendChild(curRep.CreateTextNode("_" + str + "_"));
            }

            curPage.AppendChild(imgNode);
        }

        private void cmbxExsistingPages_SelectedIndexChanged(object sender, EventArgs e)
        {
            cmbxExsistingReport.Enabled = false;
            grpPageInfo.Enabled = true;

            foreach (XmlElement element in mRepRoot.ChildNodes)
            {
                if (element.Attributes.Item(0).Value.ToLower() == cmbxExsistingPages.SelectedItem.ToString().ToLower())
                {
                    curPage = element;
                }
            }
        }        
    }
}
