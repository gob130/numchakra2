using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO;

namespace Omega
{
    public partial class frmReoprtDef : Form
    {
        //public string mChkraPreName = "chkra";
        //public string mNumberPreName = "num";
        //public string mSeporator = "_";

        private MainForm frmParent;

        public frmReoprtDef(MainForm t)
        {
            InitializeComponent();
            frmParent = t;
            cmbSexSelect.SelectedIndex = 0;
            cmbSexSelect.Text  = cmbSexSelect.SelectedItem.ToString();

        }

        private void WriteText2File(string path2file, TextBox tb)
        {
            WriteText2File(path2file, tb.Text);
        }
        private void WriteText2File(string path2file, string info)
        {
            FileInfo file = new FileInfo(path2file);
            if (file.Exists == true)
            {
                file.Delete();
            }


            FileStream fs = new FileStream(path2file, FileMode.CreateNew);
            StreamWriter writer = new StreamWriter(fs);

            writer.WriteLine(info);

            writer.Close();
            fs.Close();

            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        private string ReadTextFromFile(string path2file)
        {
            string outStr = "";

            FileInfo file = new FileInfo(path2file);

            if (file.Exists == true)
            {
                FileStream fs = new FileStream(path2file, FileMode.Open);
                StreamReader fr = new StreamReader(fs);

                outStr = fr.ReadToEnd();

                fr.Close();
                fs.Close();

                GC.Collect();
                GC.WaitForPendingFinalizers();
            }

            return outStr;
        }


        private void button2_Click(object sender, EventArgs e)
        {
            if ((txtSphere.Text == "") || (txtNum.Text == ""))
            {
                return;
            }

            //string path2file = Reports.ReportDataProvider.Instance.mFolderPath2ChakraInfoFiles + "\\" + Reports.ReportDataProvider.Instance.mChkraPreName + txtSphere.Text + Reports.ReportDataProvider.Instance.mSeporator + Reports.ReportDataProvider.Instance.mNumberPreName + txtNum.Text + ".txt";
            string chakraName = "_צאקרה_" + txtSphere.Text + "_";
            //string chakraName = txtSphere.Text;
            string path2file = Reports.ReportDataProvider.Instance.ConstructFilePath2ChkraInfo(chakraName, Convert.ToInt16(txtNum.Text));
            // frmParent.mApplicationMainDir + "\\Templets\\InternalChkraData\\" + mChkraPreName + txtSphere.Text + mSeporator + mNumberPreName + txtNum.Text + ".txt";
            FileInfo file = new FileInfo(path2file);

            if (file.Exists == true)
            {
                FileStream fs = new FileStream(path2file, FileMode.Open);
                StreamReader fr = new StreamReader(fs);

                txtFileView.Text = fr.ReadToEnd();

                fr.Close();
                fs.Close();

                GC.Collect();
                GC.WaitForPendingFinalizers();
            }

            
        }

        private void button3_Click(object sender, EventArgs e)
        {
            txtFileView.Text = "";
            txtNum.Text = "";
            txtSphere.Text = "";
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //string path2file = frmParent.mApplicationMainDir + "\\Templets\\InternalSphereData\\" + mChkraPreName + txtSphere.Text + mSeporator + mNumberPreName + txtNum.Text + ".txt";
            //string path2file = Reports.ReportDataProvider.Instance.mFolderPath2ChakraInfoFiles + "\\" + Reports.ReportDataProvider.Instance.mChkraPreName + txtSphere.Text + Reports.ReportDataProvider.Instance.mSeporator + Reports.ReportDataProvider.Instance.mNumberPreName + txtNum.Text + ".txt";
            string chakraName = "_צאקרה_" + txtSphere.Text + "_";
            string path2file = Reports.ReportDataProvider.Instance.ConstructFilePath2ChkraInfo(chakraName, Convert.ToInt16(txtNum.Text));

            FileInfo file = new FileInfo(path2file);
            if (file.Exists == true)
            {
                file.Delete();
            }


            FileStream fs = new FileStream(path2file,FileMode.CreateNew);
            StreamWriter writer = new StreamWriter(fs);

            writer.Write(txtFileView.Text);

            writer.Close();
            fs.Close();

            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            string path2file = "";
            if (cmbCycleType.SelectedItem.ToString() == "שיא")
            {
                path2file = Reports.ReportDataProvider.Instance.ConstructFilePath2ClimaxInfo(Convert.ToInt16(textBox1.Text));
            }
            if (cmbCycleType.SelectedItem.ToString() == "אתגר")
            {
                path2file = Reports.ReportDataProvider.Instance.ConstructFilePath2ChallengeInfo(Convert.ToInt16(textBox1.Text));
            }
            if (cmbCycleType.SelectedItem.ToString() == "מחזור")
            {
                path2file = Reports.ReportDataProvider.Instance.ConstructFilePath2CycleInfo(Convert.ToInt16(textBox1.Text));
            }

            FileInfo file = new FileInfo(path2file);
            if (file.Exists == true)
            {
                file.Delete();
            }


            FileStream fs = new FileStream(path2file, FileMode.CreateNew);
            StreamWriter writer = new StreamWriter(fs);

            writer.Write(textBox3.Text);

            writer.Close();
            fs.Close();

            GC.Collect();
            GC.WaitForPendingFinalizers();

        }

        private void button7_Click(object sender, EventArgs e)
        {
            string path2file = Reports.ReportDataProvider.Instance.ConstructFilePath2PowerNumberInfo(Convert.ToInt16(textBox2.Text));

            FileInfo file = new FileInfo(path2file);
            if (file.Exists == true)
            {
                file.Delete();
            }


            FileStream fs = new FileStream(path2file, FileMode.CreateNew);
            StreamWriter writer = new StreamWriter(fs);

            writer.WriteLine(textBox5.Text);

            writer.Close();
            fs.Close();

            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        private void button10_Click(object sender, EventArgs e)
        {
            int NameType = Convert.ToInt16(cmbPName.SelectedItem.ToString().Split(".".ToCharArray()[0])[0].Trim());
            string path2file = Reports.ReportDataProvider.Instance.ConstructFilePath2PrivateNameInfo(NameType, Convert.ToInt16(textBox4.Text));

            FileInfo file = new FileInfo(path2file);
            if (file.Exists == true)
            {
                file.Delete();
            }


            FileStream fs = new FileStream(path2file, FileMode.CreateNew);
            StreamWriter writer = new StreamWriter(fs);

            writer.Write(textBox6.Text);

            writer.Close();
            fs.Close();

            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        private void button13_Click(object sender, EventArgs e)
        {
            string path2file = Reports.ReportDataProvider.Instance.ConstructFilePath2PersonalYearInfo(Convert.ToInt16(textBox7.Text));

            FileInfo file = new FileInfo(path2file);
            if (file.Exists == true)
            {
                file.Delete();
            }


            FileStream fs = new FileStream(path2file, FileMode.CreateNew);
            StreamWriter writer = new StreamWriter(fs);

            writer.Write(textBox8.Text);

            writer.Close();
            fs.Close();

            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        private void button16_Click(object sender, EventArgs e)
        {
            string path2file = Reports.ReportDataProvider.Instance.ConstructFilePath2PersonalMonthInfo(Convert.ToInt16(textBox9.Text));

            FileInfo file = new FileInfo(path2file);
            if (file.Exists == true)
            {
                file.Delete();
            }


            FileStream fs = new FileStream(path2file, FileMode.CreateNew);
            StreamWriter writer = new StreamWriter(fs);

            writer.Write(textBox10.Text);

            writer.Close();
            fs.Close();

            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        private void button19_Click(object sender, EventArgs e)
        {
            string path2file = Reports.ReportDataProvider.Instance.ConstructFilePath2PersonalDayInfo(Convert.ToInt16(textBox11.Text));

            FileInfo file = new FileInfo(path2file);
            if (file.Exists == true)
            {
                file.Delete();
            }


            FileStream fs = new FileStream(path2file, FileMode.CreateNew);
            StreamWriter writer = new StreamWriter(fs);

            writer.Write(textBox12.Text);

            writer.Close();
            fs.Close();

            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        private void button22_Click(object sender, EventArgs e)
        {
            int mm = Convert.ToInt16( cmbMonths.SelectedItem.ToString().Split(".".ToCharArray()[0])[0]);
            string path2file = Reports.ReportDataProvider.Instance.ConstructFilePath2AstroLuckInfo(mm);

            FileInfo file = new FileInfo(path2file);
            if (file.Exists == true)
            {
                file.Delete();
            }


            FileStream fs = new FileStream(path2file, FileMode.CreateNew);
            StreamWriter writer = new StreamWriter(fs);

            writer.Write(textBox14.Text);

            writer.Close();
            fs.Close();

            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            string path2file = "";
            if (cmbCycleType.SelectedItem.ToString() == "שיא")
            {
                path2file = Reports.ReportDataProvider.Instance.ConstructFilePath2ClimaxInfo(Convert.ToInt16(textBox1.Text));
                
            }
            if (cmbCycleType.SelectedItem.ToString() == "אתגר")
            {
                path2file = Reports.ReportDataProvider.Instance.ConstructFilePath2ChallengeInfo(Convert.ToInt16(textBox1.Text));

            }
            if (cmbCycleType.SelectedItem.ToString() == "מחזור")
            {
                path2file = Reports.ReportDataProvider.Instance.ConstructFilePath2CycleInfo(Convert.ToInt16(textBox1.Text));

            }

            FileInfo file = new FileInfo(path2file);

            if (file.Exists == true)
            {
                FileStream fs = new FileStream(path2file, FileMode.Open);
                StreamReader fr = new StreamReader(fs);

                textBox3.Text = fr.ReadToEnd();

                fr.Close();
                fs.Close();

                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        private void button9_Click(object sender, EventArgs e)
        {
            string path2file = Reports.ReportDataProvider.Instance.ConstructFilePath2PowerNumberInfo(Convert.ToInt16(textBox2.Text));

            FileInfo file = new FileInfo(path2file);

            if (file.Exists == true)
            {
                FileStream fs = new FileStream(path2file, FileMode.Open);
                StreamReader fr = new StreamReader(fs);

                textBox5.Text = fr.ReadToEnd();

                fr.Close();
                fs.Close();

                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        private void button12_Click(object sender, EventArgs e)
        {
            int NameType = Convert.ToInt16(cmbPName.SelectedItem.ToString().Split(".".ToCharArray()[0])[0].Trim());
            string path2file = Reports.ReportDataProvider.Instance.ConstructFilePath2PrivateNameInfo(NameType,Convert.ToInt16(textBox4.Text));

            FileInfo file = new FileInfo(path2file);

            if (file.Exists == true)
            {
                FileStream fs = new FileStream(path2file, FileMode.Open);
                StreamReader fr = new StreamReader(fs);

                textBox6.Text = fr.ReadToEnd();

                fr.Close();
                fs.Close();

                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        private void button15_Click(object sender, EventArgs e)
        {
            string path2file = Reports.ReportDataProvider.Instance.ConstructFilePath2PersonalYearInfo(Convert.ToInt16(textBox7.Text));

            FileInfo file = new FileInfo(path2file);

            if (file.Exists == true)
            {
                FileStream fs = new FileStream(path2file, FileMode.Open);
                StreamReader fr = new StreamReader(fs);

                textBox8.Text = fr.ReadToEnd();

                fr.Close();
                fs.Close();

                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        private void button18_Click(object sender, EventArgs e)
        {
            string path2file = Reports.ReportDataProvider.Instance.ConstructFilePath2PersonalMonthInfo(Convert.ToInt16(textBox9.Text));

            FileInfo file = new FileInfo(path2file);

            if (file.Exists == true)
            {
                FileStream fs = new FileStream(path2file, FileMode.Open);
                StreamReader fr = new StreamReader(fs);

                textBox10.Text = fr.ReadToEnd();

                fr.Close();
                fs.Close();

                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        private void button21_Click(object sender, EventArgs e)
        {
            string path2file = Reports.ReportDataProvider.Instance.ConstructFilePath2PersonalDayInfo(Convert.ToInt16(textBox11.Text));

            FileInfo file = new FileInfo(path2file);

            if (file.Exists == true)
            {
                FileStream fs = new FileStream(path2file, FileMode.Open);
                StreamReader fr = new StreamReader(fs);

                textBox12.Text = fr.ReadToEnd();

                fr.Close();
                fs.Close();

                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        private void button24_Click(object sender, EventArgs e)
        {
            int val = Convert.ToInt16(cmbMonths.SelectedItem.ToString().Split(".".ToCharArray()[0])[0]);
            string path2file = Reports.ReportDataProvider.Instance.ConstructFilePath2AstroLuckInfo(val);

            FileInfo file = new FileInfo(path2file);

            if (file.Exists == true)
            {
                FileStream fs = new FileStream(path2file, FileMode.Open);
                StreamReader fr = new StreamReader(fs);

                textBox14.Text = fr.ReadToEnd();

                fr.Close();
                fs.Close();

                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        private void button28_Click(object sender, EventArgs e)
        {
            int num = Convert.ToInt16(cmbWorkType.SelectedItem.ToString().Split(".".ToCharArray()[0])[0]) - 1;
            Enums.EnumProvider.ReservedWork enm = (Enums.EnumProvider.ReservedWork)num;
            
            string path2File = Reports.ReportDataProvider.Instance.ConstructFilePath2WorkInfo(enm, Convert.ToInt16(textBox16.Text));

            WriteText2File(path2File, textBox17.Text);
        }

        private void button30_Click(object sender, EventArgs e)
        {
            int num = Convert.ToInt16(cmbWorkType.SelectedItem.ToString().Split(".".ToCharArray()[0])[0]) - 1;
            Enums.EnumProvider.ReservedWork enm = (Enums.EnumProvider.ReservedWork)num;

            string path2File = Reports.ReportDataProvider.Instance.ConstructFilePath2WorkInfo(enm, Convert.ToInt16(textBox16.Text));

            textBox17.Text = ReadTextFromFile(path2File);
        }

        private void button29_Click(object sender, EventArgs e)
        {
            textBox16.Text = "";
            textBox17.Text = "";
            cmbWorkType.Text = "";
        }

        private void button26_Click(object sender, EventArgs e)
        {
            textBox15.Text = "";
            textBox13.Text = "";
            cmbCrysis.Text = "";
        }

        private void button25_Click(object sender, EventArgs e)
        {
            int num = Convert.ToInt16(cmbCrysis.SelectedItem.ToString().Split(".".ToCharArray()[0])[0]) - 1;
            Enums.EnumProvider.ReservedCrysis enm = (Enums.EnumProvider.ReservedCrysis)num;

            string path2File = Reports.ReportDataProvider.Instance.ConstructFilePath2CrysisInfo(enm, Convert.ToInt16(textBox15.Text));

            WriteText2File(path2File, textBox13.Text);
        }

        private void button27_Click(object sender, EventArgs e)
        {
            int num = Convert.ToInt16(cmbCrysis.SelectedItem.ToString().Split(".".ToCharArray()[0])[0]) - 1;
            Enums.EnumProvider.ReservedCrysis enm = (Enums.EnumProvider.ReservedCrysis)num;

            string path2File = Reports.ReportDataProvider.Instance.ConstructFilePath2CrysisInfo(enm, Convert.ToInt16(textBox15.Text));

            textBox13.Text = ReadTextFromFile(path2File);
        }

        private void button31_Click(object sender, EventArgs e)
        {
            int num = Convert.ToInt16(cmbBalance.SelectedItem.ToString().Split(".".ToCharArray()[0])[0]) - 1;
            Enums.EnumProvider.Balance enm = (Enums.EnumProvider.Balance)num;

            string path2File = Reports.ReportDataProvider.Instance.ConstructFilePath2IZUNInfo(enm);

            WriteText2File(path2File, textBox18.Text);
        }

        private void button33_Click(object sender, EventArgs e)
        {
            int num = Convert.ToInt16(cmbBalance.SelectedItem.ToString().Split(".".ToCharArray()[0])[0]) - 1;
            Enums.EnumProvider.Balance enm = (Enums.EnumProvider.Balance)num;

            string path2File = Reports.ReportDataProvider.Instance.ConstructFilePath2IZUNInfo(enm);

            textBox18.Text = ReadTextFromFile(path2File);
        }

        private void button32_Click(object sender, EventArgs e)
        {
            textBox18.Text = "";
            cmbCrysis.Text = "";
        }

        private void button5_Click(object sender, EventArgs e)
        {
            textBox3.Text = "";
            textBox1.Text = "";
            cmbCycleType.Text = "";
        }

        private void button8_Click(object sender, EventArgs e)
        {
            textBox5.Text = "";
            textBox2.Text = "";
        }

        private void button11_Click(object sender, EventArgs e)
        {
            textBox4.Text = "";
            textBox6.Text = "";
            cmbPName.Text = "";
        }

        private void button14_Click(object sender, EventArgs e)
        {
            textBox8.Text = "";
            textBox7.Text = "";
        }

        private void button17_Click(object sender, EventArgs e)
        {
            textBox9.Text = "";
            textBox10.Text = "";
        }

        private void button20_Click(object sender, EventArgs e)
        {
            textBox12.Text = "";
            textBox11.Text = "";
        }

        private void button23_Click(object sender, EventArgs e)
        {
            textBox14.Text = "";
            cmbMonths.Text = "";
        }

        private void button35_Click(object sender, EventArgs e)
        {
            textBox20.Text = "";
            textBox19.Text = "";
        }

        private void button34_Click(object sender, EventArgs e)
        {
            int val = 0;

            if (int.TryParse(textBox19.Text, out val))
            {
                string path2file = Reports.ReportDataProvider.Instance.ConstructFilePath2RectificationInfo(val);

                WriteText2File(path2file, textBox20.Text);

            }

        }

        private void button36_Click(object sender, EventArgs e)
        {
            int val = 0;

            if (int.TryParse(textBox19.Text, out val))
            {
                string path2file = Reports.ReportDataProvider.Instance.ConstructFilePath2RectificationInfo(val);

                textBox20.Text = ReadTextFromFile(path2file);

            }

        }

        private void button38_Click(object sender, EventArgs e)
        {
            textBox21.Text = "";
            textBox22.Text = "";
        }

        private void button37_Click(object sender, EventArgs e)
        {
            int val = Convert.ToInt16(textBox21.Text);
            string path2file = Reports.ReportDataProvider.Instance.ConstructFilePath2DesignationInfo(val);

            WriteText2File(path2file, textBox22.Text);
        }

        private void button39_Click(object sender, EventArgs e)
        {
            int val = Convert.ToInt16(textBox21.Text);
            string path2file = Reports.ReportDataProvider.Instance.ConstructFilePath2DesignationInfo(val);

            textBox22.Text = ReadTextFromFile(path2file);
        }

        private void frmReoprtDef_Load(object sender, EventArgs e)
        {
            if (BLL.AppSettings.Instance.ProgramLanguage == BLL.AppSettings.Language.Hebrew)
            {
            }
            if (BLL.AppSettings.Instance.ProgramLanguage == BLL.AppSettings.Language.English)
            {
            }
            foreach (string sPit in Reports.ReportDataProvider.Instance.mPitPlanes)
            {
                cmbPitPlane.Items.Add(sPit);
            }
        }


        private void button40_Click(object sender, EventArgs e)
        {
            WriteText2File(Reports.ReportDataProvider.Instance.ConstructFilePath2PitagorsInfo(cmbPitPlane.SelectedItem.ToString().Substring(0, 2), cmbPitFull.SelectedItem.ToString().Substring(0, 2)),textBox24.Text);
        }

        private void button42_Click(object sender, EventArgs e)
        {
            textBox24.Text = ReadTextFromFile(Reports.ReportDataProvider.Instance.ConstructFilePath2PitagorsInfo(cmbPitPlane.SelectedItem.ToString().Substring(0, 2), cmbPitFull.SelectedItem.ToString().Substring(0, 2)));
        }

        private void button41_Click(object sender, EventArgs e)
        {
            textBox24.Text = "";
            cmbPitFull.Text = "";
            cmbPitPlane.Text = "";
        }

        private void button47_Click(object sender, EventArgs e)
        {
            textBox27.Text = "";
            cmbBssnssSccss.Text = "סוג";
        }

        private void button48_Click(object sender, EventArgs e)
        {
            string[] split = cmbBssnssSccss.SelectedItem.ToString().Split("-".ToCharArray()[0]);
            double val = ( Convert.ToInt16(split[0]) + Convert.ToInt16(split[1]) ) / 2;

            textBox27.Text = ReadTextFromFile(Reports.ReportDataProvider.Instance.ConstructFilePath2BussinessSuccess(val));
        }

        private void button46_Click(object sender, EventArgs e)
        {
            string[] split = cmbBssnssSccss.SelectedItem.ToString().Split("-".ToCharArray()[0]);
            double val = (Convert.ToInt16(split[0]) + Convert.ToInt16(split[1])) / 2;

            WriteText2File(Reports.ReportDataProvider.Instance.ConstructFilePath2BussinessSuccess(val), textBox27.Text);
        }

        private void button44_Click(object sender, EventArgs e)
        {
            textBox25.Text = "";
            cmbMarrSccss.Text = "סוג";
        }

        private void button45_Click(object sender, EventArgs e)
        {
            string[] split = cmbMarrSccss.SelectedItem.ToString().Split("-".ToCharArray()[0]);
            double val = (Convert.ToInt16(split[0]) + Convert.ToInt16(split[1])) / 2;

            textBox25.Text = ReadTextFromFile(Reports.ReportDataProvider.Instance.ConstructFilePath2MarriageSuccess(val));
        }

        private void button43_Click(object sender, EventArgs e)
        {
            string[] split = cmbMarrSccss.SelectedItem.ToString().Split("-".ToCharArray()[0]);
            double val = (Convert.ToInt16(split[0]) + Convert.ToInt16(split[1])) / 2;

            WriteText2File(Reports.ReportDataProvider.Instance.ConstructFilePath2MarriageSuccess(val), textBox25.Text);
        }

        private void button3_Click_1(object sender, EventArgs e)
        {
            txtFileView.Text = "";
        }

        private void cbPro_CheckedChanged(object sender, EventArgs e)
        {
            Reports.ReportDataProvider.Instance.Pro = cbPro.Checked;
        }

        private void cbLangSelect_SelectedIndexChanged(object sender, EventArgs e)
        {
            string sLang = cbLangSelect.SelectedItem.ToString();

            string[] splt = sLang.Split(" ".ToCharArray());
            switch (splt[1])
            {
                case "HEB":
                    Reports.ReportDataProvider.Instance.Language = BLL.AppSettings.Language.Hebrew;
                    break;
                case "ENG":
                    Reports.ReportDataProvider.Instance.Language = BLL.AppSettings.Language.English;
                    break;
            }
        }

        private void button50_Click(object sender, EventArgs e)
        {
            textBox23.Text = "";
            cbOpenClose.Text = "";
            cbCkra.Text = "";
        }

        private void button51_Click(object sender, EventArgs e)
        {
            int iC = Convert.ToInt16(cbCkra.SelectedItem.ToString().Split(".".ToCharArray()[0])[0]);
            bool oc = true;
            switch (cbOpenClose.SelectedItem.ToString())
            {
                case "פתוח":
                    oc = true;  break;
                case "סגור":
                    oc = false; break;
            }

            textBox23.Text = ReadTextFromFile(Reports.ReportDataProvider.Instance.ConstructFilePath2ChakraOpening(iC,oc));
        }

        private void button49_Click(object sender, EventArgs e)
        {
            int iC = Convert.ToInt16(cbCkra.SelectedItem.ToString().Split(".".ToCharArray()[0])[0]);
            bool oc = true;
            switch (cbOpenClose.SelectedItem.ToString())
            {
                case "פתוח":
                    oc = true; break;
                case "סגור":
                    oc = false; break;
            }

            WriteText2File(Reports.ReportDataProvider.Instance.ConstructFilePath2ChakraOpening(iC, oc), textBox23.Text);
        }

        private void button53_Click(object sender, EventArgs e)
        {
            textBox26.Text = "";
            cbPrsn.Text = "";
        }

        private void button54_Click(object sender, EventArgs e)
        {
            int val = Convert.ToInt16((cbPrsn.SelectedText.Split(".".ToCharArray()[0])[0]));
            
            textBox26.Text = ReadTextFromFile(Reports.ReportDataProvider.Instance.ConstructFilePath2MainPersonality(val-1));
        }

        private void button52_Click(object sender, EventArgs e)
        {
            int val = Convert.ToInt16((cbPrsn.SelectedText.Split(".".ToCharArray()[0])[0]));

            WriteText2File(Reports.ReportDataProvider.Instance.ConstructFilePath2MainPersonality(val - 1), textBox26.Text);
        }

        private void tabPage16_Click(object sender, EventArgs e)
        {

        }

        private void button56_Click(object sender, EventArgs e)
        {
            textBox28.Text = "";
            cbLearnClass.SelectedIndex = 0;
            cbLearnClass.Text = "סוג";

        }

        private void button57_Click(object sender, EventArgs e)
        {
            string[] sVals = cbLearnClass.SelectedItem.ToString().Split("-".ToCharArray());
            double val = (Convert.ToDouble(sVals[0]) + Convert.ToDouble(sVals[1])) / 2.0;

            textBox28.Text = ReadTextFromFile(Reports.ReportDataProvider.Instance.ConstructFilePath2LearnSccss(val));
        }

        private void button55_Click(object sender, EventArgs e)
        {
            string[] sVals = cbLearnClass.SelectedItem.ToString().Split("-".ToCharArray());
            double val = (Convert.ToDouble(sVals[0]) + Convert.ToDouble(sVals[1])) / 2.0;

            WriteText2File(Reports.ReportDataProvider.Instance.ConstructFilePath2LearnSccss(val), textBox28.Text);
        }

        private void button59_Click(object sender, EventArgs e)
        {
            textBox29.Text = "";
            cbAHDH.ResetText();
        }

        private void button60_Click(object sender, EventArgs e)
        {
            string sval = cbAHDH.SelectedItem.ToString().Split(".".ToCharArray()[0])[0];
            textBox29.Text = ReadTextFromFile(Reports.ReportDataProvider.Instance.ConstrcutFilePath2ADHDInfoFiles(Convert.ToInt16(sval)));
        }

        private void button58_Click(object sender, EventArgs e)
        {
            string sval = cbAHDH.SelectedItem.ToString().Split(".".ToCharArray()[0])[0];
            WriteText2File(Reports.ReportDataProvider.Instance.ConstrcutFilePath2ADHDInfoFiles(Convert.ToInt16(sval)), textBox29.Text);
        }

        private void button71_Click(object sender, EventArgs e)
        {
            textBox36.Text = "";
            textBox37.Text = "";

        }

        private void button72_Click(object sender, EventArgs e)
        {
            if (textBox36.Text.Trim().Length > 0)
            {
                textBox37.Text = ReadTextFromFile(Reports.ReportDataProvider.Instance.ConstructFilePath2HealthBody(Convert.ToInt16(textBox36.Text)));
            }
        }

        private void button70_Click(object sender, EventArgs e)
        {
            WriteText2File(Reports.ReportDataProvider.Instance.ConstructFilePath2HealthBody(Convert.ToInt16(textBox36.Text)), textBox37.Text);
        }

        private void button68_Click(object sender, EventArgs e)
        {
            textBox32.Text = "";
            textBox35.Text = "";
        }

        private void button69_Click(object sender, EventArgs e)
        {
            if (textBox35.Text.Trim().Length > 0)
            {
                textBox32.Text = ReadTextFromFile(Reports.ReportDataProvider.Instance.ConstructFilePath2HealthTLC(Convert.ToInt16(textBox35.Text)));
            }
        }

        private void button67_Click(object sender, EventArgs e)
        {
            WriteText2File(Reports.ReportDataProvider.Instance.ConstructFilePath2HealthTLC(Convert.ToInt16(textBox35.Text)), textBox32.Text);
        }

        private void button65_Click(object sender, EventArgs e)
        {
            textBox31.Text = "";
            textBox34.Text = "";
        }

        private void button66_Click(object sender, EventArgs e)
        {
            if (textBox34.Text.Trim().Length > 0)
            {
                textBox31.Text = ReadTextFromFile(Reports.ReportDataProvider.Instance.ConstructFilePath2HealthTakeCare(Convert.ToInt16(textBox34.Text)));
            }
        }

        private void button64_Click(object sender, EventArgs e)
        {
            WriteText2File(Reports.ReportDataProvider.Instance.ConstructFilePath2HealthTakeCare(Convert.ToInt16(textBox34.Text)), textBox31.Text);
        }

        private void button62_Click(object sender, EventArgs e)
        {
            textBox30.Text = "";
            textBox33.Text = "";
        }

        private void button63_Click(object sender, EventArgs e)
        {
            string type = cmbDcsType.SelectedItem.ToString().Split(".".ToCharArray())[0].ToString();
            
            if (textBox33.Text.Trim().Length > 0)
            {
                textBox30.Text = ReadTextFromFile(Reports.ReportDataProvider.Instance.ConstructFilePath2HealthDcs(Convert.ToInt16(type),Convert.ToInt16(textBox33.Text)));
            }
        }

        private void button61_Click(object sender, EventArgs e)
        {
            string type = cmbDcsType.SelectedItem.ToString().Split(".".ToCharArray())[0].ToString();
            WriteText2File(Reports.ReportDataProvider.Instance.ConstructFilePath2HealthDcs(Convert.ToInt16(type),Convert.ToInt16(textBox33.Text)), textBox30.Text);
        }

        private void button73_Click(object sender, EventArgs e)
        {
            WriteText2File(Reports.ReportDataProvider.Instance.ConstructFilePath2HealthPrivateName(Convert.ToInt16(textBox38.Text)), textBox39.Text);
        }

        private void button75_Click(object sender, EventArgs e)
        {
            if (textBox38.Text.Trim().Length > 0)
            {
                textBox39.Text = ReadTextFromFile(Reports.ReportDataProvider.Instance.ConstructFilePath2HealthPrivateName(Convert.ToInt16(textBox38.Text)));
            }
        }

        private void button74_Click(object sender, EventArgs e)
        {
            textBox38.Text = "";
            textBox39.Text = "";
        }

        private void button76_Click(object sender, EventArgs e)
        {
            WriteText2File(Reports.ReportDataProvider.Instance.ConstructFilePath2HealthClimax(Convert.ToInt16(textBox40.Text)), textBox41.Text);
        }

        private void button78_Click(object sender, EventArgs e)
        {
            if (textBox40.Text.Trim().Length > 0)
            {
                textBox41.Text = ReadTextFromFile(Reports.ReportDataProvider.Instance.ConstructFilePath2HealthClimax(Convert.ToInt16(textBox40.Text)));
            }
        }

        private void button77_Click(object sender, EventArgs e)
        {
            textBox41.Text = "";
            textBox40.Text = "";
        }

        private void button79_Click(object sender, EventArgs e)
        {
            WriteText2File(Reports.ReportDataProvider.Instance.ConstructFilePath2HealthPrsnlYear(Convert.ToInt16(textBox42.Text)), textBox43.Text);
        }

        private void button81_Click(object sender, EventArgs e)
        {
            if (textBox42.Text.Trim().Length > 0)
            {
                textBox43.Text = ReadTextFromFile(Reports.ReportDataProvider.Instance.ConstructFilePath2HealthPrsnlYear(Convert.ToInt16(textBox42.Text)));
            }
        }

        private void button80_Click(object sender, EventArgs e)
        {
            textBox42.Text = "";
            textBox43.Text = "";
        }

        private void button82_Click(object sender, EventArgs e)
        {
            WriteText2File(Reports.ReportDataProvider.Instance.ConstructFilePath2HealthAstro(Convert.ToInt16(textBox44.Text)), textBox45.Text);
        }

        private void button84_Click(object sender, EventArgs e)
        {
            if (textBox44.Text.Trim().Length > 0)
            {
                textBox45.Text = ReadTextFromFile(Reports.ReportDataProvider.Instance.ConstructFilePath2HealthAstro(Convert.ToInt16(textBox44.Text)));
            }
        }

        private void button83_Click(object sender, EventArgs e)
        {
            textBox44.Text = "";
            textBox45.Text = "";
        }

        private void button86_Click(object sender, EventArgs e)
        {
            textBox48.Text = "";
        }

        private void button87_Click(object sender, EventArgs e)
        {
            textBox48.Text = ReadTextFromFile(Reports.ReportDataProvider.Instance.ConstructFilePath2FianlySummary());
        }

        private void button85_Click(object sender, EventArgs e)
        {
            WriteText2File(Reports.ReportDataProvider.Instance.ConstructFilePath2FianlySummary(), textBox48.Text);
        }

        private void cmbSexSelect_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cmbSexSelect.SelectedIndex == 0)
            {
                Reports.ReportDataProvider.Instance.Gender = Enums.EnumProvider.Sex.Male;
            }
            if (cmbSexSelect.SelectedIndex == 1)
            {
                Reports.ReportDataProvider.Instance.Gender = Enums.EnumProvider.Sex.Female;
            }

        }

        private void button88_Click(object sender, EventArgs e)
        {
            bool mapstrength = true;
            if (cmbPickStrength.SelectedIndex == 0)
            {
                mapstrength = true;
            }
            if (cmbPickStrength.SelectedIndex == 1)
            {
                mapstrength = false;
            }

            WriteText2File(Reports.ReportDataProvider.Instance.ConstructFilePath2MapStrength(mapstrength),textBox46.Text);
        }

        private void button89_Click(object sender, EventArgs e)
        {
            textBox46.Text = "";
        }

        private void button90_Click(object sender, EventArgs e)
        {
            bool mapstrength = true;
            if (cmbPickStrength.SelectedIndex == 0)
            {
                mapstrength = true;
            }
            if (cmbPickStrength.SelectedIndex == 1)
            {
                mapstrength = false;
            }

            textBox46.Text = ReadTextFromFile(Reports.ReportDataProvider.Instance.ConstructFilePath2MapStrength(mapstrength));
        }

       


    }
}
