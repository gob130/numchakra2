using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
//using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Management;
using BLL;
using System.IO;
using System.Xml;
namespace Proffing
{
    public partial class Form1 : Form
    {
        #region Data Member
        private Calc mCalc = new Calc();
        #endregion

        public Form1()
        {
            InitializeComponent();

            //BLL.Calc Calculator = new BLL.Calc(AppSettings.Instance.ProgramLanguage, true);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //char[] digits = txtIn.Text.ToCharArray();

            //int tmp = (Convert.ToInt16(digits[1])-48) * 1000 + (Convert.ToInt16(digits[3])-48) * 100 + (Convert.ToInt16(digits[4])-48) * 10 + (Convert.ToInt16(digits[6])-48) * 1;
            //tmp = (int) Math.Pow((tmp/200 - 254), 2) +2;

            txtOut.Text = GenerateSecondKEY(txtIn.Text);
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            //string sn = GetWinSN();
            //string[] parts = sn.Split("-".ToCharArray());
            //long tNum;
            //List<long> nums = new List<long>();

            //for (int i = 0; i < parts.Length; i++)
            //{
            //    if (long.TryParse(parts[i],out tNum))
            //    {
            //        nums.Add(tNum);
            //    }
            //}

            //long sum = 0;
            //for (int i = 0; i<nums.Count; i++)
            //{
            //    sum += nums[i];
            //}

            //int size = sum.ToString().Length;
            
            //System.Random rndObj = new Random();
            //long FinalNum = Convert.ToInt64( rndObj.Next(1,9) * Math.Pow(10, (7 - size -1)));

            txtIn.Text = mCalc.Proof_CreateFirstKey() ;//sum.ToString() + FinalNum.ToString();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            //string sn = GetWinSN();
            //string[] parts = sn.Split("-".ToCharArray());
            //long tNum;
            //List<long> nums = new List<long>();

            //for (int i = 0; i < parts.Length; i++)
            //{
            //    if (long.TryParse(parts[i], out tNum))
            //    {
            //        nums.Add(tNum);
            //    }
            //}

            //long sum = 0;
            //for (int i = 0; i < nums.Count; i++)
            //{
            //    sum += nums[i];
            //}

            //int size = sum.ToString().Length;

            //System.Random rndObj = new Random();
            //long FinalNum = rndObj.Next(10 ^ (7 - size), 10 ^ (7 - size + 1));

            txtIn.Text = mCalc.Proof_CreateFirstKey();//sum.ToString() + FinalNum.ToString();
        }

        private string GenerateSecondKEY(string FirstKEY)
        {
            System.Random rndObj = new Random();
            int TwoDigs = rndObj.Next(10, 99);

            int n1 = Convert.ToInt16(FirstKEY.Substring(2, 3));
            int n2 = Convert.ToInt16(FirstKEY.Substring(5, 2));

            int sum1 = n1 + n2;

            n1 = Convert.ToInt16(FirstKEY.Substring(3, 3));
            n2 = Convert.ToInt16(FirstKEY.Substring(2, 1) + FirstKEY.Substring(6, 1));
            int sum2 = n1 * n2;

            return sum1.ToString().Substring(0, 3) + TwoDigs.ToString() + sum2.ToString().Substring(0, 3);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            string path;
            if (txtPath2File.Text == "")
            {
                path = System.Environment.SpecialFolder.Desktop.ToString();
            }
            else
            {
                path = System.IO.Path.GetFullPath(txtPath2File.Text);
            }

            OpenFileDialog ofd = new OpenFileDialog();
            ofd.InitialDirectory = path;
            ofd.Filter = "Text Files(*.txt)|*.txt";
            ofd.Title = "בחר קובץ";
            ofd.Multiselect = false;

            DialogResult res = ofd.ShowDialog();
            if ((res == DialogResult.Cancel) || (res == DialogResult.Abort))
            {
                return;
            }
            if (res == DialogResult.OK)
            {
                path = ofd.FileName;
                txtPath2File.Text = path;
            }
        }

        private void cmdFixNameFile_Click(object sender, EventArgs e)
        {
            string Path = txtPath2File.Text;
            bool ChangeWasMade = false;

            FileInfo mInFile = new FileInfo(Path);

            if (mInFile.Exists == false)
            {
                MessageBox.Show("קובץ לא קיים");
                return;
            }

            List<string> NameList = new List<string>();

            FileStream mFileHandler = new FileStream(Path, FileMode.Open);
            StreamReader mFileReader = new StreamReader(mFileHandler);

            while (mFileReader.EndOfStream == false)
            {
                NameList.Add(mFileReader.ReadLine().Trim());
            }

            mFileReader.Close();

            for (int i = 0; i < NameList.Count-1; i++)
            {
                string CurName = NameList[i];

                for (int j = i + 1; j < NameList.Count; j++)
                {
                    string Name2Test = NameList[j];

                    if (Name2Test == CurName)
                    {
                        ChangeWasMade = true;
                        NameList.RemoveAt(j);
                    }
                }
            }

            if (ChangeWasMade == false )
            {
                MessageBox.Show("הקובץ מוכן לעבודה");
                return;
            }
            
            mFileHandler.Close();

            mInFile.Delete();
            mInFile.Create();
            mInFile.Refresh();
            mInFile = null;

            GC.Collect();
            GC.WaitForPendingFinalizers();

            mFileHandler = new FileStream(Path, FileMode.Open);
            StreamWriter mFileWirter = new StreamWriter(mFileHandler);

            for (int i = 0; i < NameList.Count; i++)
            {
                mFileWirter.WriteLine(NameList[i]);
            }

            mFileWirter.Close();
            mFileHandler.Close();
            

            MessageBox.Show("הקובץ מוכן לעבודה");
        }

        //private void button1_Click_1(object sender, EventArgs e)
        //{
        //    txtOut.Text = GenerateSecondKEY(txtIn.Text);
        //}

        private void button2_Click_1(object sender, EventArgs e)
        {
            txtOut.Text = GenerateSecondKEY(txtIn.Text);
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cmbUserType.SelectedItem.ToString() == "זמני")
            {
                txtTrialDays.Visible = true;
                label5.Visible = true;
            }
            else
            {
                txtTrialDays.Visible = false;
                label5.Visible = false;
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            #region Test
            if ((cmbUserType.Text == "") || (cmbLang.Text == ""))
            {
                MessageBox.Show("יש למלא את כל השדות");
                return;
            }

            if ((cmbUserType.Text == "זמני") && (txtTrialDays.Text == ""))
            {
                MessageBox.Show("יש למלא את כל השדות");
                return;
            }
            #endregion

            FolderBrowserDialog sfd = new FolderBrowserDialog();
            sfd.Description = "בחר מיקום הקובץ";

            string mXMLLocation;

            DialogResult dlgres = sfd.ShowDialog();
            if (dlgres == DialogResult.Abort || dlgres == DialogResult.Cancel)
            {
                return;
            }
            if (dlgres == DialogResult.OK)
            {
                mXMLLocation = sfd.SelectedPath + "\\Settings.xml";
                //FileInfo fiXML = new FileInfo(mXMLLocation);
                //fiXML.Create();
                

                string strOutType, strOutLan;
                strOutLan = "";
                strOutType = "";

                switch (cmbLang.Text)
                {
                    case "עברית":
                        strOutLan = "Hebrew";
                        break;
                    case "אנגלית":
                        strOutLan = "English";
                        break;
                }

                switch (cmbUserType.Text)
                {
                    case "זמני":
                        strOutType = AppSettings.ProgType.Trial.ToString() + "," + txtTrialDays.Text.ToString();
                        break;
                    case "רגיל":
                        strOutType = AppSettings.ProgType.Normal.ToString();
                        break;
                    case "מומחה":
                        strOutType  = AppSettings.ProgType.Expert.ToString();
                        break;
                }

                XmlTextWriter mXmlTextWriter = new XmlTextWriter(mXMLLocation, Encoding.UTF8);
                mXmlTextWriter.Formatting = Formatting.Indented;

                mXmlTextWriter.WriteStartDocument();

                mXmlTextWriter.WriteStartElement("AppSettings");
                mXmlTextWriter.WriteElementString("Language", strOutLan);
                mXmlTextWriter.WriteElementString("Type", strOutType);
                mXmlTextWriter.WriteEndElement();

                mXmlTextWriter.Flush();

                mXmlTextWriter.WriteEndDocument();
                mXmlTextWriter.Close();

                MessageBox.Show("הקובץ נוצר בהצלחה");
            }

            
        }

        private void button5_Click(object sender, EventArgs e)
        {
            DirectoryInfo overwiteDir = new DirectoryInfo(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData) + @"\dc");
            if (overwiteDir.Exists == false)
            {
                overwiteDir.Create();
            }

            FileInfo overwriteFile = new FileInfo(Path.Combine(overwiteDir.FullName, @"rlk767rv.xml"));
            if (overwriteFile.Exists == false )
            {
                overwriteFile.Create();
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            DirectoryInfo overwiteDir = new DirectoryInfo(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData) + @"\dc");
            FileInfo overwriteFile = new FileInfo(Path.Combine(overwiteDir.FullName, @"rlk767rv.xml"));

            if (overwriteFile.Exists == true)
            {
                overwriteFile.Delete();
            }

            if (overwiteDir.Exists == true)
            {
                overwiteDir.Delete(true);
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            /*
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Title = "select csv file....";
            ofd.Multiselect = false;
            ofd.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            
            DialogResult res = ofd.ShowDialog();

            if ((res == DialogResult.Cancel) || (res == DialogResult.Abort) )
            {
                return;
            }

            string sInPath = "";
            if ((res == System.Windows.Forms.DialogResult.OK) && (ofd.CheckFileExists == true))
            {
                sInPath = ofd.FileName;
            }

            FileStream inStream = new FileStream(sInPath, FileMode.Open);
            StreamReader inReader = new StreamReader(inStream, Encoding.UTF8);

            FileStream outStream = new FileStream(System.IO.Path.Combine(System.IO.Path.GetDirectoryName(sInPath), "out_" + System.IO.Path.GetFileNameWithoutExtension(sInPath) + ".txt"), FileMode.CreateNew);
            StreamWriter outWriter = new StreamWriter(outStream, Encoding.UTF8);

            string line = inReader.ReadLine(); // header
            string outStr = "";
            while (inReader.EndOfStream == false)
            {
                outStr = "";
                line = inReader.ReadLine();
                string[] splt = line.Split(",".ToCharArray()[0]);

                if (splt[1] != "NULL") user.mFirstName = splt[1];
                if (splt[2] != "NULL") user.mLastName = splt[2];
                if (splt[3] != "NULL") user.mFatherName = splt[3];
                if (splt[4] != "NULL") user.mMotherName = splt[4];
                    DateTime bd = new DateTime(1900,1,1);
                    if (splt[5] != "NULL")
                    {
                        bool DTres = DateTime.TryParse(splt[5], out bd);
                        if (DTres == true)
                        {
                            user.mB_Date = bd;
                        }
                        else
                        {
                            user.mB_Date = new DateTime(1900, 1, 1);
                        }
                    }
                    else
                    {
                        user.mB_Date = new DateTime(1900, 1, 1);
                    }
                if (splt[6] != "NULL") user.mB_Date = new DateTime(1900, 1, 1); user.mCity = splt[6];
                if (splt[7] != "NULL") user.mStreet = splt[7];
                if (splt[8] != "NULL") user.mBuildingNum = Convert.ToInt32(splt[8]);
                if (splt[9] != "NULL") user.mAppNum = Convert.ToInt32(splt[9]);
                if (splt[10] != "NULL") user.mEMail = splt[10];
                if (splt[11] != "NULL") user.mPhones = splt[11];

            } //while (inReader.EndOfStream == false)

            inReader.Close();
            inStream.Close();
             */
        }

    }
}
