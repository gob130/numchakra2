namespace Omega.Reports
{
    partial class frmReportCreationcs
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmReportCreationcs));
            this.cmbxExsistingReports = new System.Windows.Forms.ComboBox();
            this.btnCreateNewReport = new System.Windows.Forms.Button();
            this.grpReportInfo = new System.Windows.Forms.GroupBox();
            this.grpPageInfo = new System.Windows.Forms.GroupBox();
            this.btnSavePage = new System.Windows.Forms.Button();
            this.grpImageAdd = new System.Windows.Forms.GroupBox();
            this.btnAddImage = new System.Windows.Forms.Button();
            this.btnFileSelect = new System.Windows.Forms.Button();
            this.txtFileLocation = new System.Windows.Forms.TextBox();
            this.cmbxImageFromControl = new System.Windows.Forms.ComboBox();
            this.rbtnImage_Prog = new System.Windows.Forms.RadioButton();
            this.rbtnImage_File = new System.Windows.Forms.RadioButton();
            this.grpTextAdd = new System.Windows.Forms.GroupBox();
            this.rbtnLeft = new System.Windows.Forms.RadioButton();
            this.rbtnCenter = new System.Windows.Forms.RadioButton();
            this.rbtnRight = new System.Windows.Forms.RadioButton();
            this.lblFontInfo = new System.Windows.Forms.Label();
            this.btnFontSelect = new System.Windows.Forms.Button();
            this.txtBody = new System.Windows.Forms.TextBox();
            this.btnAddText = new System.Windows.Forms.Button();
            this.rbtnImage = new System.Windows.Forms.RadioButton();
            this.rbtnText = new System.Windows.Forms.RadioButton();
            this.cmbxExsistingPages = new System.Windows.Forms.ComboBox();
            this.btnNewPage = new System.Windows.Forms.Button();
            this.cmbxExsistingReport = new System.Windows.Forms.ComboBox();
            this.btnSaveXmlDoc = new System.Windows.Forms.Button();
            this.button1 = new System.Windows.Forms.Button();
            this.grpReportInfo.SuspendLayout();
            this.grpPageInfo.SuspendLayout();
            this.grpImageAdd.SuspendLayout();
            this.grpTextAdd.SuspendLayout();
            this.SuspendLayout();
            // 
            // cmbxExsistingReports
            // 
            this.cmbxExsistingReports.FormattingEnabled = true;
            this.cmbxExsistingReports.Location = new System.Drawing.Point(-212, 4);
            this.cmbxExsistingReports.Name = "cmbxExsistingReports";
            this.cmbxExsistingReports.Size = new System.Drawing.Size(121, 21);
            this.cmbxExsistingReports.TabIndex = 0;
            // 
            // btnCreateNewReport
            // 
            this.btnCreateNewReport.Location = new System.Drawing.Point(12, 12);
            this.btnCreateNewReport.Name = "btnCreateNewReport";
            this.btnCreateNewReport.Size = new System.Drawing.Size(121, 23);
            this.btnCreateNewReport.TabIndex = 1;
            this.btnCreateNewReport.Text = "צור דוח";
            this.btnCreateNewReport.UseVisualStyleBackColor = true;
            this.btnCreateNewReport.Click += new System.EventHandler(this.btnCreateNewReport_Click);
            // 
            // grpReportInfo
            // 
            this.grpReportInfo.Controls.Add(this.grpPageInfo);
            this.grpReportInfo.Controls.Add(this.cmbxExsistingPages);
            this.grpReportInfo.Controls.Add(this.btnNewPage);
            this.grpReportInfo.Controls.Add(this.cmbxExsistingReports);
            this.grpReportInfo.Location = new System.Drawing.Point(139, 12);
            this.grpReportInfo.Name = "grpReportInfo";
            this.grpReportInfo.Size = new System.Drawing.Size(742, 374);
            this.grpReportInfo.TabIndex = 2;
            this.grpReportInfo.TabStop = false;
            this.grpReportInfo.Text = "פרטי דוח";
            // 
            // grpPageInfo
            // 
            this.grpPageInfo.Controls.Add(this.btnSavePage);
            this.grpPageInfo.Controls.Add(this.grpImageAdd);
            this.grpPageInfo.Controls.Add(this.grpTextAdd);
            this.grpPageInfo.Controls.Add(this.rbtnImage);
            this.grpPageInfo.Controls.Add(this.rbtnText);
            this.grpPageInfo.Location = new System.Drawing.Point(23, 19);
            this.grpPageInfo.Name = "grpPageInfo";
            this.grpPageInfo.Size = new System.Drawing.Size(546, 349);
            this.grpPageInfo.TabIndex = 5;
            this.grpPageInfo.TabStop = false;
            this.grpPageInfo.Text = "פרטי עמוד";
            // 
            // btnSavePage
            // 
            this.btnSavePage.Location = new System.Drawing.Point(15, 320);
            this.btnSavePage.Name = "btnSavePage";
            this.btnSavePage.Size = new System.Drawing.Size(513, 23);
            this.btnSavePage.TabIndex = 2;
            this.btnSavePage.Text = "שמור עמוד";
            this.btnSavePage.UseVisualStyleBackColor = true;
            this.btnSavePage.Click += new System.EventHandler(this.btnSavePage_Click);
            // 
            // grpImageAdd
            // 
            this.grpImageAdd.Controls.Add(this.btnAddImage);
            this.grpImageAdd.Controls.Add(this.btnFileSelect);
            this.grpImageAdd.Controls.Add(this.txtFileLocation);
            this.grpImageAdd.Controls.Add(this.cmbxImageFromControl);
            this.grpImageAdd.Controls.Add(this.rbtnImage_Prog);
            this.grpImageAdd.Controls.Add(this.rbtnImage_File);
            this.grpImageAdd.Enabled = false;
            this.grpImageAdd.Location = new System.Drawing.Point(15, 221);
            this.grpImageAdd.Name = "grpImageAdd";
            this.grpImageAdd.Size = new System.Drawing.Size(406, 89);
            this.grpImageAdd.TabIndex = 1;
            this.grpImageAdd.TabStop = false;
            this.grpImageAdd.Text = "תמונה";
            // 
            // btnAddImage
            // 
            this.btnAddImage.Location = new System.Drawing.Point(6, 19);
            this.btnAddImage.Name = "btnAddImage";
            this.btnAddImage.Size = new System.Drawing.Size(42, 64);
            this.btnAddImage.TabIndex = 8;
            this.btnAddImage.Text = "הוסף";
            this.btnAddImage.UseVisualStyleBackColor = true;
            this.btnAddImage.Click += new System.EventHandler(this.btnAddImage_Click);
            // 
            // btnFileSelect
            // 
            this.btnFileSelect.Enabled = false;
            this.btnFileSelect.Location = new System.Drawing.Point(269, 56);
            this.btnFileSelect.Name = "btnFileSelect";
            this.btnFileSelect.Size = new System.Drawing.Size(47, 19);
            this.btnFileSelect.TabIndex = 7;
            this.btnFileSelect.Text = "בחר";
            this.btnFileSelect.UseVisualStyleBackColor = true;
            this.btnFileSelect.Click += new System.EventHandler(this.btnFileSelect_Click);
            // 
            // txtFileLocation
            // 
            this.txtFileLocation.Enabled = false;
            this.txtFileLocation.Location = new System.Drawing.Point(80, 56);
            this.txtFileLocation.Name = "txtFileLocation";
            this.txtFileLocation.Size = new System.Drawing.Size(183, 20);
            this.txtFileLocation.TabIndex = 6;
            // 
            // cmbxImageFromControl
            // 
            this.cmbxImageFromControl.FormattingEnabled = true;
            this.cmbxImageFromControl.Location = new System.Drawing.Point(171, 19);
            this.cmbxImageFromControl.Name = "cmbxImageFromControl";
            this.cmbxImageFromControl.Size = new System.Drawing.Size(146, 21);
            this.cmbxImageFromControl.TabIndex = 5;
            // 
            // rbtnImage_Prog
            // 
            this.rbtnImage_Prog.AutoSize = true;
            this.rbtnImage_Prog.Checked = true;
            this.rbtnImage_Prog.Location = new System.Drawing.Point(335, 19);
            this.rbtnImage_Prog.Name = "rbtnImage_Prog";
            this.rbtnImage_Prog.Size = new System.Drawing.Size(65, 17);
            this.rbtnImage_Prog.TabIndex = 1;
            this.rbtnImage_Prog.TabStop = true;
            this.rbtnImage_Prog.Text = "מהתכנה";
            this.rbtnImage_Prog.UseVisualStyleBackColor = true;
            this.rbtnImage_Prog.CheckedChanged += new System.EventHandler(this.rbtnImage_Prog_CheckedChanged);
            // 
            // rbtnImage_File
            // 
            this.rbtnImage_File.AutoSize = true;
            this.rbtnImage_File.Location = new System.Drawing.Point(341, 57);
            this.rbtnImage_File.Name = "rbtnImage_File";
            this.rbtnImage_File.Size = new System.Drawing.Size(59, 17);
            this.rbtnImage_File.TabIndex = 1;
            this.rbtnImage_File.Text = "מקובץ";
            this.rbtnImage_File.UseVisualStyleBackColor = true;
            // 
            // grpTextAdd
            // 
            this.grpTextAdd.Controls.Add(this.rbtnLeft);
            this.grpTextAdd.Controls.Add(this.rbtnCenter);
            this.grpTextAdd.Controls.Add(this.rbtnRight);
            this.grpTextAdd.Controls.Add(this.lblFontInfo);
            this.grpTextAdd.Controls.Add(this.btnFontSelect);
            this.grpTextAdd.Controls.Add(this.txtBody);
            this.grpTextAdd.Controls.Add(this.btnAddText);
            this.grpTextAdd.Location = new System.Drawing.Point(15, 21);
            this.grpTextAdd.Name = "grpTextAdd";
            this.grpTextAdd.Size = new System.Drawing.Size(406, 194);
            this.grpTextAdd.TabIndex = 1;
            this.grpTextAdd.TabStop = false;
            this.grpTextAdd.Text = "מלל";
            // 
            // rbtnLeft
            // 
            this.rbtnLeft.AutoSize = true;
            this.rbtnLeft.Location = new System.Drawing.Point(48, 177);
            this.rbtnLeft.Name = "rbtnLeft";
            this.rbtnLeft.Size = new System.Drawing.Size(56, 17);
            this.rbtnLeft.TabIndex = 13;
            this.rbtnLeft.Text = "שמאל";
            this.rbtnLeft.UseVisualStyleBackColor = true;
            // 
            // rbtnCenter
            // 
            this.rbtnCenter.AutoSize = true;
            this.rbtnCenter.Location = new System.Drawing.Point(54, 160);
            this.rbtnCenter.Name = "rbtnCenter";
            this.rbtnCenter.Size = new System.Drawing.Size(50, 17);
            this.rbtnCenter.TabIndex = 13;
            this.rbtnCenter.Text = "מרכז";
            this.rbtnCenter.UseVisualStyleBackColor = true;
            // 
            // rbtnRight
            // 
            this.rbtnRight.AutoSize = true;
            this.rbtnRight.Checked = true;
            this.rbtnRight.Location = new System.Drawing.Point(57, 143);
            this.rbtnRight.Name = "rbtnRight";
            this.rbtnRight.Size = new System.Drawing.Size(47, 17);
            this.rbtnRight.TabIndex = 13;
            this.rbtnRight.TabStop = true;
            this.rbtnRight.Text = "ימין";
            this.rbtnRight.UseVisualStyleBackColor = true;
            // 
            // lblFontInfo
            // 
            this.lblFontInfo.AutoSize = true;
            this.lblFontInfo.Location = new System.Drawing.Point(138, 145);
            this.lblFontInfo.Name = "lblFontInfo";
            this.lblFontInfo.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.lblFontInfo.Size = new System.Drawing.Size(34, 13);
            this.lblFontInfo.TabIndex = 12;
            this.lblFontInfo.Text = "Font: ";
            // 
            // btnFontSelect
            // 
            this.btnFontSelect.Location = new System.Drawing.Point(335, 143);
            this.btnFontSelect.Name = "btnFontSelect";
            this.btnFontSelect.Size = new System.Drawing.Size(55, 40);
            this.btnFontSelect.TabIndex = 11;
            this.btnFontSelect.Text = "בחר פונט";
            this.btnFontSelect.UseVisualStyleBackColor = true;
            this.btnFontSelect.Click += new System.EventHandler(this.btnFontSelect_Click);
            // 
            // txtBody
            // 
            this.txtBody.Location = new System.Drawing.Point(6, 19);
            this.txtBody.Multiline = true;
            this.txtBody.Name = "txtBody";
            this.txtBody.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.txtBody.Size = new System.Drawing.Size(384, 117);
            this.txtBody.TabIndex = 10;
            // 
            // btnAddText
            // 
            this.btnAddText.Location = new System.Drawing.Point(6, 142);
            this.btnAddText.Name = "btnAddText";
            this.btnAddText.Size = new System.Drawing.Size(42, 51);
            this.btnAddText.TabIndex = 9;
            this.btnAddText.Text = "הוסף";
            this.btnAddText.UseVisualStyleBackColor = true;
            this.btnAddText.Click += new System.EventHandler(this.btnAddText_Click);
            // 
            // rbtnImage
            // 
            this.rbtnImage.AutoSize = true;
            this.rbtnImage.Location = new System.Drawing.Point(443, 221);
            this.rbtnImage.Name = "rbtnImage";
            this.rbtnImage.Size = new System.Drawing.Size(85, 17);
            this.rbtnImage.TabIndex = 0;
            this.rbtnImage.Text = "הוסף תמונה";
            this.rbtnImage.UseVisualStyleBackColor = true;
            // 
            // rbtnText
            // 
            this.rbtnText.AutoSize = true;
            this.rbtnText.Checked = true;
            this.rbtnText.Location = new System.Drawing.Point(453, 21);
            this.rbtnText.Name = "rbtnText";
            this.rbtnText.Size = new System.Drawing.Size(75, 17);
            this.rbtnText.TabIndex = 0;
            this.rbtnText.TabStop = true;
            this.rbtnText.Text = "הוסף מלל";
            this.rbtnText.UseVisualStyleBackColor = true;
            this.rbtnText.CheckedChanged += new System.EventHandler(this.rbtnText_CheckedChanged);
            // 
            // cmbxExsistingPages
            // 
            this.cmbxExsistingPages.FormattingEnabled = true;
            this.cmbxExsistingPages.Location = new System.Drawing.Point(609, 48);
            this.cmbxExsistingPages.Name = "cmbxExsistingPages";
            this.cmbxExsistingPages.Size = new System.Drawing.Size(114, 21);
            this.cmbxExsistingPages.TabIndex = 4;
            this.cmbxExsistingPages.SelectedIndexChanged += new System.EventHandler(this.cmbxExsistingPages_SelectedIndexChanged);
            // 
            // btnNewPage
            // 
            this.btnNewPage.Location = new System.Drawing.Point(609, 19);
            this.btnNewPage.Name = "btnNewPage";
            this.btnNewPage.Size = new System.Drawing.Size(114, 23);
            this.btnNewPage.TabIndex = 1;
            this.btnNewPage.Text = "צור עמוד חדש";
            this.btnNewPage.UseVisualStyleBackColor = true;
            this.btnNewPage.Click += new System.EventHandler(this.btnNewPage_Click);
            // 
            // cmbxExsistingReport
            // 
            this.cmbxExsistingReport.FormattingEnabled = true;
            this.cmbxExsistingReport.Location = new System.Drawing.Point(13, 48);
            this.cmbxExsistingReport.Name = "cmbxExsistingReport";
            this.cmbxExsistingReport.Size = new System.Drawing.Size(119, 21);
            this.cmbxExsistingReport.TabIndex = 3;
            this.cmbxExsistingReport.SelectedIndexChanged += new System.EventHandler(this.cmbxExsistingReport_SelectedIndexChanged);
            // 
            // btnSaveXmlDoc
            // 
            this.btnSaveXmlDoc.Location = new System.Drawing.Point(12, 177);
            this.btnSaveXmlDoc.Name = "btnSaveXmlDoc";
            this.btnSaveXmlDoc.Size = new System.Drawing.Size(114, 35);
            this.btnSaveXmlDoc.TabIndex = 7;
            this.btnSaveXmlDoc.Text = "שמור סגנון דוח";
            this.btnSaveXmlDoc.UseVisualStyleBackColor = true;
            this.btnSaveXmlDoc.Click += new System.EventHandler(this.btnSaveXmlDoc_Click);
            // 
            // button1
            // 
            this.button1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button1.Location = new System.Drawing.Point(13, 342);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(120, 44);
            this.button1.TabIndex = 8;
            this.button1.Text = " סגור          X";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // frmReportCreationcs
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(895, 398);
            this.ControlBox = false;
            this.Controls.Add(this.button1);
            this.Controls.Add(this.btnSaveXmlDoc);
            this.Controls.Add(this.cmbxExsistingReport);
            this.Controls.Add(this.grpReportInfo);
            this.Controls.Add(this.btnCreateNewReport);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "frmReportCreationcs";
            this.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.RightToLeftLayout = true;
            this.Text = "הכן סגנונות דוח";
            this.Load += new System.EventHandler(this.frmReportCreationcs_Load);
            this.grpReportInfo.ResumeLayout(false);
            this.grpPageInfo.ResumeLayout(false);
            this.grpPageInfo.PerformLayout();
            this.grpImageAdd.ResumeLayout(false);
            this.grpImageAdd.PerformLayout();
            this.grpTextAdd.ResumeLayout(false);
            this.grpTextAdd.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.ComboBox cmbxExsistingReports;
        private System.Windows.Forms.Button btnCreateNewReport;
        private System.Windows.Forms.GroupBox grpReportInfo;
        private System.Windows.Forms.Button btnNewPage;
        private System.Windows.Forms.ComboBox cmbxExsistingReport;
        private System.Windows.Forms.GroupBox grpPageInfo;
        private System.Windows.Forms.RadioButton rbtnText;
        private System.Windows.Forms.ComboBox cmbxExsistingPages;
        private System.Windows.Forms.GroupBox grpImageAdd;
        private System.Windows.Forms.Button btnFileSelect;
        private System.Windows.Forms.TextBox txtFileLocation;
        private System.Windows.Forms.ComboBox cmbxImageFromControl;
        private System.Windows.Forms.RadioButton rbtnImage_Prog;
        private System.Windows.Forms.RadioButton rbtnImage_File;
        private System.Windows.Forms.GroupBox grpTextAdd;
        private System.Windows.Forms.RadioButton rbtnImage;
        private System.Windows.Forms.Button btnAddImage;
        private System.Windows.Forms.Button btnFontSelect;
        private System.Windows.Forms.TextBox txtBody;
        private System.Windows.Forms.Button btnAddText;
        private System.Windows.Forms.Label lblFontInfo;
        private System.Windows.Forms.Button btnSavePage;
        private System.Windows.Forms.RadioButton rbtnLeft;
        private System.Windows.Forms.RadioButton rbtnCenter;
        private System.Windows.Forms.RadioButton rbtnRight;
        private System.Windows.Forms.Button btnSaveXmlDoc;
        private System.Windows.Forms.Button button1;
    }
}