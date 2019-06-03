#region Namespaces

using BLL;
using Omega.Objects;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

#endregion Namespaces

// **********

namespace Omega
{
    public partial class ReName : Form
    {
        #region Data Members
        //private Form frmMain = Omega.MainForm.ActiveForm;
        private MainForm mParentForm;
        private const string ProcessSpinnerName = "ucProcessSpinner";

        private BLL.Calc mCalc;// = new BLL.Calc(AppSettings.Instance.ProgramLanguage);
        bool isMaster;
        private UserInfo OriginalUser;

        private Control.ControlCollection cntrlCls;// = MainForm.ActiveForm.Controls;

        //private string[] NamesList;
        //private string[] mFemaleNamesList = new string[21] { "בתיה", "גליה", "גילה", "דיקלה", "דנה", "אורית", "הדרי", "הודיה", "ורד", "זהר", "חיה", "חניתה", "יהודית, יממה", "יפה", "מור", "מירב", "מלכה", "נעמה", "נתנאלה", "עדנה", "עטרה" };
        //private string[] mMaleNamesList;

        //NamesList = mFemaleNamesList;
        private string[] firstSex = new string[] { "1", "3", "5", "8" };
        private string[] secoundSex = new string[] { "4", "6", "9", "22" };
        private string[] firstThrought = new string[] { "1", "8", "9" };
        private string[] secoundThrought = new string[] { "3", "5" };

        private string[] employeeFirstThrought = new string[] { "9" };
        private string[] employeeSecoundThrought = new string[] { "1", "8" };
        private string[] employeeThirdThrought = new string[] { "3", "5" };
        private string[] employeeQuarterThrought = new string[] { "1", "8", "9" };


        private int round = 0;

    List<ChakraResult> chkrResList = new List<ChakraResult>();
        #endregion

        public ReName(MainForm t)
        {
            InitializeComponent();
            mParentForm = t;
            cntrlCls = mParentForm.Controls;

            isMaster = (cntrlCls.Find("cbPersonMaster", true)[0] as CheckBox).Checked;
            mCalc = new BLL.Calc(AppSettings.Instance.ProgramLanguage, isMaster);

            this.Icon = mParentForm.Icon;

            //NamesList = new string[mFemaleNamesList.Length + mMaleNamesList.Length];
            //for (int i = 0; i < mFemaleNamesList.Length; i++)
            //{
            //    NamesList[i] = mFemaleNamesList[i];
            //}

            //for (int i = 0; i < mMaleNamesList.Length; i++)
            //{
            //    NamesList[mFemaleNamesList.Length + i] = mMaleNamesList[i];
            //}

        }

        private void ReName_Load(object sender, EventArgs e)
        {
            OriginalUser = new UserInfo();
            if (OriginalUser.ProgData2UserInfo(mParentForm) == true)
            {
            }
            else
            {
                throw new Exception("problem in one of the person input parameters.");
            }

            if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
            {
                groupBox4.Visible = false;
            }
        }

        private List<string> ReadNamesFile(string sFilePath)
        {
            List<string> outList = new List<string>();

            FileStream fs = new FileStream(sFilePath, FileMode.Open);
            StreamReader sr = new StreamReader(fs);
            outList = sr.ReadToEnd().Split(Environment.NewLine.ToCharArray(), StringSplitOptions.RemoveEmptyEntries).ToList();
            //while (sr.EndOfStream == false)
            //{
            //    string s = sr.ReadLine().Trim();
            //    if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
            //    {
            //        s = s.ToLower();
            //    }
            //    outList.Add(s);
            //}

            sr.Close();
            fs.Close();
            return outList;
        }

        private async void btnCalc_Click(object sender, EventArgs e)
        {
            if (groupBox7.Visible && !cbIndepended.Checked && !cbIndependedFuture.Checked && !cbEmployee.Checked)
            {
               
                string s1 = "", s2 = "";
                if (BLL.AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                {
                    s1 = "חובה לסמן את 1 או יותר משלושת השדות הראשונים";
                    s2 = "נתונים חסרים : ";
                }
                if (BLL.AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                {
                    s1 = "You must mark 1 or more of the first three fields";
                    s2 = "Missing data...";
                }
                MessageBox.Show(s2 + System.Environment.NewLine + s1, "טעות הכנסת נתונים", MessageBoxButtons.OK, MessageBoxIcon.Warning);

            }
            else
            {
                OriginalUser = new UserInfo();

                bool repres = OriginalUser.ProgData2UserInfo(mParentForm);

                chkrResList.Clear();

                UserInfo CurUser = new UserInfo();

                #region Duplicate

                CurUser = OriginalUser;

                #endregion Duplicate

                string fName = OriginalUser.mFirstName;
                string lName = OriginalUser.mLastName;

                #region GetNames

                string sFilePath = mParentForm.AppMainDir + "\\Names";
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                {
                    if (rbtnFamily.Checked == true)
                    {
                        sFilePath += "\\Heb_JusidhLastNames";
                    }
                    else
                    {
                        if (rbtnJuish.Checked == true)
                        {
                            sFilePath += "\\Juish";
                        }
                        if (rbtnJuishBible.Checked == true)
                        {
                            sFilePath += "\\JuishBible";
                        }
                        if (rbtnArab.Checked == true)
                        {
                            sFilePath += "\\Arab";
                        }
                        if (rbtnMale.Checked == true)
                        {
                            sFilePath += "Male";
                        }
                        else
                        {
                            sFilePath += "Female";
                        }
                    }
                }

                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                {
                    // only male / female / family
                    sFilePath += "\\Eng_";
                    if (rbtnMale.Checked == true)
                    {
                        sFilePath += "MaleNames";
                    }
                    if (rbtnFemale.Checked == true)
                    {
                        sFilePath += "FemaleNames";
                    }
                    if (rbtnFamily.Checked == true)
                    {
                        sFilePath += "LastNames";
                    }
                }

                sFilePath += ".txt";

                #endregion

                List<string> NamesList = ReadNamesFile(sFilePath);

                //MessageBox.Show($"Not Implemented Yet !", string.Empty, MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly);
                //return;

               // await BalanceNamesAsync(NamesList);

                //int iii = 0;
                foreach (string CurAddedName in NamesList)//(int index = 0; index < NamesList.Length; index++)
                {
                    //gcCounter = gcCounter + 1;

                    string finalCurAddedName = CurAddedName.Trim();

                    if (finalCurAddedName.Contains("'"))
                    {
                        finalCurAddedName = finalCurAddedName.Replace("'", "");
                    }

                    if (finalCurAddedName.Contains("-"))
                    {
                        finalCurAddedName = finalCurAddedName.Replace("-", "");
                    }

                    finalCurAddedName = mCalc.ChangeFinalChars(finalCurAddedName);

                    if (rbtnFamily.Checked == false)  // Private Names Testings
                    {
                        if (cbAddOrReplace.Checked == true)// add a first name
                        {
                            CurUser.mFirstName = fName + finalCurAddedName;
                        }
                        else// replace a first name
                        {
                            CurUser.mFirstName = finalCurAddedName;
                        }
                    }
                    else // Last Names Testings
                    {
                        if (cbAddOrReplace.Checked == true) // add a family name
                        {
                            CurUser.mLastName = lName + finalCurAddedName;
                        }
                        else// replace family name
                        {
                            CurUser.mLastName = finalCurAddedName;
                        }
                    }

                    //UpdateMainForm(CurUser);
                    CurUser.ApplyUserInfo2ProgData(mParentForm);

                    RunProgram();

                    ChakraResult curRes = CollectResults(CurAddedName);
                    if (isFilterPassed(curRes))
                    {
                        chkrResList.Add(curRes);
                    }

                    //if (gcCounter > 5)
                    //{
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                    GC.WaitForFullGCComplete();

                    //gcCounter = 0;
                    //}
                }

                dgvRes.Rows.Clear();
                dgvRes.Update();
                UpdateDataGrid();
                
                #region Restore Initial Value
                mParentForm.ClearForm();

                //UpdateMainForm(OriginalUser);
                OriginalUser.mFirstName = fName;
                OriginalUser.mLastName = lName;
                repres = OriginalUser.ApplyUserInfo2ProgData(mParentForm);

                //OriginalUser.mFirstName = fName;
                //OriginalUser.ApplyUserInfo2ProgData(mParentForm);
                RunProgram();
                #endregion

                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.WaitForFullGCComplete();

                if (groupBox7.Visible && rbtnFilterAdvanced.Checked)
                {
                    if (chkrResList.Count == 0 && round <= getNumberOfRound())
                    {
                        round++;
                        btnCalc.PerformClick();
                        //return;

                    }
                    else
                    {
                        MessageBox.Show("אם יש צורך צור קשר עם יעקובי", "הודעה", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
        }

        private int getNumberOfRound()
        {
            //אם שכיר ולא עצמאי
            if (!cbIndepended.Checked && !cbIndependedFuture.Checked && cbEmployee.Checked)
            {
                return 3;
            }
            //כל השאר
            return 2;
        }

        public void SetSpinnerVisiblity(bool isVisible)
        {
            if (groupBox1.Controls.ContainsKey(ProcessSpinnerName))
            {
                groupBox1.Controls[ProcessSpinnerName].Visible = isVisible;

            }

        }

        private async Task BalanceNamesAsync(List<string> namesList)
        {
            await Task.Factory.StartNew(() =>
            {
                foreach (string CurAddedName in namesList)//(int index = 0; index < NamesList.Length; index++)
                {
                    UserInfo CurUser = new UserInfo();

                    //gcCounter = gcCounter + 1;

                    string finalCurAddedName = CurAddedName.Trim();

                    if (finalCurAddedName.Contains("'"))
                    {
                        finalCurAddedName = finalCurAddedName.Replace("'", "");
                    }

                    if (finalCurAddedName.Contains("-"))
                    {
                        finalCurAddedName = finalCurAddedName.Replace("-", "");
                    }

                    finalCurAddedName = mCalc.ChangeFinalChars(finalCurAddedName);

                    if (rbtnFamily.Checked == false)  // Private Names Testings
                    {
                        if (cbAddOrReplace.Checked == true)// add a first name
                        {
                            string fName = OriginalUser.mFirstName;
                            CurUser.mFirstName = fName + finalCurAddedName;
                        }
                        else// replace a first name
                        {
                            CurUser.mFirstName = finalCurAddedName;
                        }
                    }
                    else // Last Names Testings
                    {
                        string fName = OriginalUser.mFirstName;

                        if (cbAddOrReplace.Checked == true) // add a family name
                        {
                            CurUser.mLastName = fName + finalCurAddedName;
                        }
                        else// replace family name
                        {
                            CurUser.mLastName = finalCurAddedName;
                        }
                    }

                    //UpdateMainForm(CurUser);
                    CurUser.ApplyUserInfo2ProgData(mParentForm);

                    //RunProgram();

                    ChakraResult curRes = CollectResults(CurAddedName);
                    if (isFilterPassed(curRes))
                    {
                        chkrResList.Add(curRes);
                    }

                    //if (gcCounter > 5)
                    //{
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                    GC.WaitForFullGCComplete();

                    //gcCounter = 0;
                    //}
                }
            });
        }

        private void UpdateMainForm(UserInfo curUser)
        {
            cntrlCls.Find("txtPrivateName", true)[0].Text = curUser.mFirstName;
            cntrlCls.Find("txtFamilyName", true)[0].Text = curUser.mLastName;
            cntrlCls.Find("txtFatherName", true)[0].Text = curUser.mFatherName;
            cntrlCls.Find("txtMotherName", true)[0].Text = curUser.mMotherName;
            (cntrlCls.Find("DateTimePickerFrom", true)[0] as DateTimePicker).Value = curUser.mB_Date;
            cntrlCls.Find("txtCity", true)[0].Text = curUser.mCity;
            cntrlCls.Find("txtStreet", true)[0].Text = curUser.mStreet;
            cntrlCls.Find("txtBiuldingNum", true)[0].Text = curUser.mBuildingNum.ToString();
            cntrlCls.Find("txtAppNum", true)[0].Text = curUser.mAppNum.ToString();

            (cntrlCls.Find("DateTimePickerTo", true)[0] as DateTimePicker).Value = DateTime.Now;
        }

        private void RunProgram()
        {
           // CheckBox cb = cntrlCls.Find("checkBox1", true)[0] as CheckBox;
           // cb.Checked = false;

            //(cntrlCls.Find("cbPersonMaster", true)[0] as CheckBox).Checked = isMaster;

            mParentForm.runCalc();
        }

        private ChakraResult CollectResults(string sAddedName)
        {
            ChakraResult tmpRes = new ChakraResult("", "", "", "", "", "", "", 0, 0, 0, 0, "", "", 0, 0, 0);

            tmpRes.mAddedName = sAddedName;
            tmpRes.mNameSum = cntrlCls.Find("txtPName_Num", true)[0].Text;
            tmpRes.mChakraSexCreation = cntrlCls.Find("txtNum6", true)[0].Text;
            tmpRes.mCharkraThrought = cntrlCls.Find("txtNum3", true)[0].Text;
            tmpRes.mCharkraBase = cntrlCls.Find("txtNum7", true)[0].Text;
            tmpRes.mCharkraThirdEye = cntrlCls.Find("txtNum2", true)[0].Text;
            tmpRes.mCharkraHeart = cntrlCls.Find("txtNum4", true)[0].Text;
          
            tmpRes.mBussiness = Convert.ToDouble(cntrlCls.Find("txtFinalBusinessMark", true)[0].Text);
            tmpRes.mLrnSccss = Convert.ToDouble(cntrlCls.Find("txtLeanSccss", true)[0].Text);
            tmpRes.mCharkraHealth = Convert.ToDouble(cntrlCls.Find("txtHealthValue", true)[0].Text);

            DataGridView tmpDgv = cntrlCls.Find("dgvCombMapSum", true)[0] as DataGridView;
            DataGridViewRow r = tmpDgv.Rows[0];
            tmpRes.mPhysical = Convert.ToInt16(r.Cells[r.Cells.Count - 1].Value.ToString());
            r = tmpDgv.Rows[1];
            tmpRes.mEmotional = Convert.ToInt16(r.Cells[r.Cells.Count - 1].Value.ToString());
            r = tmpDgv.Rows[2];
            tmpRes.mMental = Convert.ToInt16(r.Cells[r.Cells.Count - 1].Value.ToString());
            r = tmpDgv.Rows[3];
            tmpRes.mEnergetic = Convert.ToInt16(r.Cells[r.Cells.Count - 1].Value.ToString());

            string[] s = cntrlCls.Find("txtInfo1", true)[0].Text.Split(System.Environment.NewLine.ToCharArray()[0]);//" ".ToCharArray()[0]);
            tmpRes.mMapBalance = s[1];
            tmpRes.mMapDepration = cntrlCls.Find("txtDepMap", true)[0].Text;

            return tmpRes;
        }

        private bool isFilterPassed(ChakraResult res)
        {
            if (rbtnNonFilter.Checked) return true;

            bool outRes = true, tmpRes = false;
            #region Filter_Form
            //int num;
            string[] sNums;
            string[] Nums2Test;
            string dlmtr = ", /";

            if (cbCon.Checked == true)
            {
                //num = mCalc.GetCorrectNumberFromSplitedString(cntrlCls.Find("txtNum7",true)[0].Text.Split(mCalc.Delimiter));
                //outRes &= (num.ToString() == txtCon.Text);
                tmpRes = false;
                sNums = cntrlCls.Find("txtNum7", true)[0].Text.Split(mCalc.Delimiter);
                Nums2Test = txtCon.Text.Trim().Split(dlmtr.ToCharArray());
                for (int k = 0; k < Nums2Test.Length; k++)
                {
                    tmpRes |= mCalc.isNumInArray(sNums, Convert.ToInt16(Nums2Test[k]));
                }
                outRes &= tmpRes;
            }

            //if (cbVol.Checked == true)
            //{
            //    num = mCalc.GetCorrectNumberFromSplitedString(cntrlCls.Find("txtNum3", true)[0].Text.Split(mCalc.Delimiter));
            //    outRes &= (num.ToString() == txtVol.Text);
            //}

            if (cbSex.Checked == true || (groupBox7.Visible && rbtnFilterAdvanced.Checked))
            {
                //num = mCalc.GetCorrectNumberFromSplitedString(cntrlCls.Find("txtNum6", true)[0].Text.Split(mCalc.Delimiter));
                //outRes &= (num.ToString() == txtSex.Text);
                tmpRes = false;
                sNums = cntrlCls.Find("txtNum6", true)[0].Text.Split(mCalc.Delimiter);
                Nums2Test = txtSex.Text.Trim().Split(dlmtr.ToCharArray());
                if (groupBox7.Visible && rbtnFilterAdvanced.Checked)
                {
                    //if independed
                    if (cbIndepended.Checked || cbIndependedFuture.Checked) //!cbEmployee
                    {
                        if (round == 0 || round == 1)
                        {
                            Nums2Test = firstSex;// new string[] { "1", "3", "5", "8" };
                        }
                        if (round == 2)
                        {
                            Nums2Test = secoundSex;
                        }
                    }
                    else
                    {
                        //is employee
                        if (round == 0 || round == 1 || round == 2)
                        {
                            Nums2Test = firstSex;// new string[] { "1", "3", "5", "8" };
                        }
                        if (round == 3)
                        {
                            Nums2Test = secoundSex;
                        }
                    }
                }
                for (int k = 0; k < Nums2Test.Length; k++)
                {
                    tmpRes |= mCalc.isNumInArray(sNums, Convert.ToInt16(Nums2Test[k]));
                }
                outRes &= tmpRes;
            }

            if (cbThrought.Checked == true || (groupBox7.Visible && rbtnFilterAdvanced.Checked))
            {
                //num = mCalc.GetCorrectNumberFromSplitedString(cntrlCls.Find("txtNum3", true)[0].Text.Split(mCalc.Delimiter));
                //outRes &= (num.ToString() == txtThr.Text);
                tmpRes = false;
                sNums = cntrlCls.Find("txtNum3", true)[0].Text.Split(mCalc.Delimiter);
                Nums2Test = txtThr.Text.Trim().Split(dlmtr.ToCharArray());
                if (groupBox7.Visible && rbtnFilterAdvanced.Checked)
                {
                    //if independed
                    if (cbIndepended.Checked || cbIndependedFuture.Checked)
                    {
                        if (round == 0 || round == 2)
                        {
                            Nums2Test = firstThrought;// new string[] { "1", "8", "9" };
                        }
                        if (round == 1)
                        {
                            Nums2Test = secoundThrought;
                        }
                    }
                    else
                    {
                        //is employee
                        if (round == 0)
                        {
                            Nums2Test = employeeFirstThrought;// new string[] { "1", "8", "9" };
                        }
                        if(round == 1)
                        {
                            Nums2Test = employeeSecoundThrought;
                        }
                        if (round == 2)
                        {
                            Nums2Test = employeeThirdThrought;
                        }
                        if(round == 3)
                        {
                            Nums2Test = employeeQuarterThrought;
                        }
                    }
                }
                for (int k = 0; k < Nums2Test.Length; k++)
                {
                    tmpRes |= mCalc.isNumInArray(sNums, Convert.ToInt16(Nums2Test[k]));
                }
                outRes &= tmpRes;
            }

            if (cbAddedNameSum.Checked == true)
            {
                //num = mCalc.GetCorrectNumberFromSplitedString(cntrlCls.Find("txtPName_Num", true)[0].Text.Split(mCalc.Delimiter));
                //outRes &= (num.ToString() == txtAddedNameSum.Text);
                tmpRes = false;
                sNums = cntrlCls.Find("txtPName_Num", true)[0].Text.Split(mCalc.Delimiter);
                Nums2Test = txtAddedNameSum.Text.Trim().Split(dlmtr.ToCharArray());
                for (int k = 0; k < Nums2Test.Length; k++)
                {
                    tmpRes |= mCalc.isNumInArray(sNums, Convert.ToInt16(Nums2Test[k]));
                }
                outRes &= tmpRes;
            }
            
            if (cbThirdEyeNoSeven.Checked == true)
            {
                outRes &= (mCalc.GetCorrectNumberFromSplitedString(res.mCharkraThirdEye.Split(mCalc.Delimiter)) != 7);
            }
            if (cbHeartNoZero.Checked == true)
            {
                outRes &= (mCalc.GetCorrectNumberFromSplitedString(res.mCharkraHeart.Split(mCalc.Delimiter)) != 0);
            }
            #endregion

            #region Filter_res - carmatic
            string tmp;
            tmp = res.mNameSum;
            outRes &= !(mCalc.isCarmaticNumber(mCalc.GetCorrectNumberFromSplitedString(tmp.Split(mCalc.Delimiter))));

            tmp = res.mChakraSexCreation;
            if (cbVol2.Checked == false)
            {
                outRes &= !(mCalc.isCarmaticNumber(mCalc.GetCorrectNumberFromSplitedString(tmp.Split(mCalc.Delimiter))));
            }

            tmp = res.mCharkraThrought;
            outRes &= !(mCalc.isCarmaticNumber(mCalc.GetCorrectNumberFromSplitedString(tmp.Split(mCalc.Delimiter))));

            tmp = res.mCharkraBase;
            outRes &= !(mCalc.isCarmaticNumber(mCalc.GetCorrectNumberFromSplitedString(tmp.Split(mCalc.Delimiter))));

            tmp = res.mCharkraThirdEye;
            if (cbVol.Checked == false)
            {
                outRes &= !(mCalc.isCarmaticNumber(mCalc.GetCorrectNumberFromSplitedString(tmp.Split(mCalc.Delimiter))));
            }

            tmp = res.mCharkraHeart;
            outRes &= !(mCalc.isCarmaticNumber(mCalc.GetCorrectNumberFromSplitedString(tmp.Split(mCalc.Delimiter))));
            #endregion



            return outRes;
        }

        private void UpdateDataGrid()
        {
            foreach (ChakraResult curRes in chkrResList)
            {
                DataGridViewRow dR = new DataGridViewRow();
                //dataFuture.Rows.Add(1);

                object[] objects = new string[15];
                objects.Initialize();
                objects.SetValue(curRes.mAddedName, 0);
                objects.SetValue(FlipString(curRes.mNameSum), 1);
                objects.SetValue(FlipString(curRes.mChakraSexCreation), 2);
                objects.SetValue(FlipString(curRes.mCharkraThrought), 3);
                //objects.SetValue(FlipString(curRes.mCharkraBase), 4);
                //objects.SetValue(FlipString(curRes.mCharkraThirdEye), 5);
                objects.SetValue(FlipString(curRes.mBussiness.ToString()), 6);
                objects.SetValue(FlipString(curRes.mLrnSccss.ToString()), 7);
                objects.SetValue(FlipString(curRes.mCharkraHealth.ToString()), 8);
                //objects.SetValue(FlipString(curRes.mCharkraHeart), 8);
                //objects.SetValue(FlipString(curRes.mPhysical.ToString()), 9);
                //objects.SetValue(FlipString(curRes.mEmotional.ToString()), 10);
                //objects.SetValue(FlipString(curRes.mMental.ToString()), 11);
                //objects.SetValue(FlipString(curRes.mEnergetic.ToString()), 12);
                objects.SetValue(FlipString(curRes.mMapBalance), 14);
                //if (curRes.mMapDepration != "")
                //{
                //    objects.SetValue(FlipString(curRes.mMapDepration), 15);
                //}
                //else
                //{
                //    objects.SetValue("לא", 14);
                //}


                dgvRes.Rows.Insert(dgvRes.Rows.Count - 1, objects);
                dgvRes.Update();
            }
        }

        private string FlipString(string s)
        {
            string sout = "";

            string[] ss = s.Split(mCalc.Delimiter);
            if (ss.Length > 1)
            {
                for (int i = 0; i < ss.Length; i++)
                {
                    sout += ss[ss.Length - 1 - i] + mCalc.Delimiter;
                }

                sout = sout.Substring(0, sout.Length - 1);
            }
            else
            {
                sout = s;
            }

            return sout;
        }

        private void rbtnNonFilter_CheckedChanged(object sender, EventArgs e)
        {
            //groupBox3.Enabled = false;
            //groupBox4.Enabled = false;
            groupBox5.Enabled = false;
            groupBox6.Enabled = false;
            groupBox7.Enabled = false;
            groupBox7.Visible = false;
            rbtnJuishBible.Enabled = true;
            rbtnJuishBible.Visible = true;
            rbtnArab.Enabled = true;
            //cbVol.Checked = true;
            //cbVol2.Checked = true;
        }

        private void rbtnFilter_CheckedChanged(object sender, EventArgs e)
        {
            //groupBox3.Enabled = true;
            //groupBox4.Enabled = true;
            groupBox5.Enabled = true;
            groupBox6.Enabled = true;
            groupBox7.Enabled = false;
            groupBox7.Visible = false;
            rbtnJuishBible.Enabled = true;
            rbtnJuishBible.Visible = true;
            rbtnArab.Enabled = true;
            //cbVol.Checked = false;
            //cbVol2.Checked = false;
        }

        private void rbtnFilterAdvanced_CheckedChanged(object sender, EventArgs e)
        {
            //groupBox3.Enabled = true;
            //groupBox4.Enabled = true;
            groupBox5.Enabled = false;
            groupBox6.Enabled = false;
            groupBox7.Enabled = true;
            groupBox7.Visible = true;
            rbtnJuishBible.Enabled = false;
            rbtnJuishBible.Visible = false;
            rbtnArab.Enabled = false;
            rbtnJuish.Checked = true;
            //cbVol.Checked = false;
            //cbVol2.Checked = false;
        }

        private void cmdExport_Click(object sender, EventArgs e)
        {
            if (chkrResList.Count == 0) return;

            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Title = "בחר קובץ לשמירה";
            sfd.InitialDirectory = System.Environment.SpecialFolder.DesktopDirectory.ToString();

            sfd.FileName = OriginalUser.mFirstName + " " + OriginalUser.mLastName + " איזון שם";
            sfd.Filter = "Microsoft Excel (*.csv)|*.csv";
            DialogResult res = sfd.ShowDialog();

            string sOutFile = "";
            if ((res == DialogResult.Cancel) || (res == DialogResult.None)) return;
            if (res == DialogResult.OK)
            {
                sOutFile = sfd.FileName;
            }


            string sEncoding = "";
            if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
            {
                //sEncoding = "iso-8859-8-i"; // Hebrew Logical
                sEncoding = "iso-8859-8"; // Hebrew Logical
                //sEncoding = "windows-1255"; // Windows Hebrew
                //sEncoding = "ibm-862"; // IBM
            }
            if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
            {
                sEncoding = Encoding.Unicode.ToString();
            }

            FileStream fs = new FileStream(sOutFile, FileMode.OpenOrCreate);
            StreamWriter sw = new StreamWriter(fs, Encoding.GetEncoding(sEncoding));

            // Header
            string line = "";
            //string line = "שם נוסף,סכום השם,מין ויצירה,גרון,בסיס,עין שלישית,הצלחה עסקית,הצלחה בלימודים,לב,פיזי,רגשי,מנטאלי,אנרגטי,מפה מאוזנת,מפה דכאונית";
            if (rbtnFamily.Checked == false) // First Names
            {
                //line = "שם נוסף,סכום השם,מין ויצירה,גרון,בסיס,עין שלישית,הצלחה עסקית,הצלחה בלימודים,לב,מפה מאוזנת,מפה דכאונית";
                line = "שם משפחה,סכום השם,מין ויצירה,גרון,הצלחה עסקית,הצלחה בלימודים,בריאות,מפה מאוזנת";
            }
            else // Last Names
            {
                //line = "שם משפחה,סכום השם,מין ויצירה,גרון,בסיס,עין שלישית,הצלחה עסקית,הצלחה בלימודים,לב,מפה מאוזנת,מפה דכאונית";
                line = "שם משפחה,סכום השם,מין ויצירה,גרון,הצלחה עסקית,הצלחה בלימודים,בריאות,מפה מאוזנת";
            }

            sw.WriteLine(line);

            foreach (ChakraResult curRes in chkrResList)
            {
                line = "";
                //line = curRes.mAddedName + "," + curRes.mNameSum + "," + curRes.mChakraSexCreation + "," + curRes.mCharkraThrought + ","
                //        + curRes.mCharkraBase + "," + curRes.mCharkraThirdEye + "," + curRes.mBussiness.ToString() + "," + curRes.mLrnSccss.ToString() + "," + curRes.mCharkraHeart + "," + curRes.mPhysical.ToString() + ","
                //        + curRes.mEmotional.ToString() + "," + curRes.mMental.ToString() + "," + curRes.mEnergetic.ToString() + "," + curRes.mMapBalance.Substring(1,curRes.mMapBalance.Length-1) + ",";

                //line = curRes.mAddedName + "," + curRes.mNameSum + "," + curRes.mChakraSexCreation + "," + curRes.mCharkraThrought + ","
                //        + curRes.mCharkraBase + "," + curRes.mCharkraThirdEye + "," + curRes.mBussiness.ToString() + ","
                //        + curRes.mLrnSccss.ToString() + "," + curRes.mCharkraHeart + ","
                //        + curRes.mMapBalance.Substring(1, curRes.mMapBalance.Length - 1) + ",";

                line = curRes.mAddedName + "," + curRes.mNameSum + "," + curRes.mChakraSexCreation + "," + curRes.mCharkraThrought + ","
                        + curRes.mBussiness.ToString() + ","
                        + curRes.mLrnSccss.ToString() + ","
                         + curRes.mCharkraHealth.ToString() + ","
                        + curRes.mMapBalance.Substring(1, curRes.mMapBalance.Length - 1) + ",";

                //if (curRes.mMapDepration == "")
                //{
                //    line += "לא";
                //}
                //else
                //{
                //    line += curRes.mMapDepration;
                //}

                sw.WriteLine(line);

            }

            sw.Close();
            fs.Close();

            GC.Collect();
            GC.WaitForFullGCComplete();

            MessageBox.Show("תהליך כתיבת קובץ הסתים בהצלחה", "ייצוא", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void rbtnFamily_CheckedChanged(object sender, EventArgs e)
        {
            if (rbtnFamily.Checked == true)
            {
                groupBox4.Enabled = false;
            }
            if (rbtnFamily.Checked == false)
            {
                groupBox4.Enabled = true;
            }
        }

    }

    internal class ChakraResult
    {
        #region Data Members
        public string mAddedName;
        public string mNameSum;
        public string mChakraSexCreation;
        public string mCharkraThrought;
        public string mCharkraBase;
        public string mCharkraThirdEye;
        public string mCharkraHeart;
        public int mPhysical;
        public int mEmotional;
        public int mMental;
        public int mEnergetic;
        public string mMapBalance;
        public string mMapDepration;
        public double mBussiness;
        public double mLrnSccss;
        public double? mCharkraHealth;
        #endregion

        #region Constructor
        public ChakraResult(string sAddedName,
                            string sNameSum,
                            string sChakraSexCreation,
                            string sCharkraThrought,
                            string sCharkraBase,
                            string sCharkraThirdEye,
                            string sCharkraHeart,
                            int sPhysical,
                            int sEmotional,
                            int sMental,
                            int sEnergetic,
                            string sMapBalance,
                            string sMapDepration,
                            double sBussiness,
                            double sLrnSccss,
                            double sCharkraHealth)
        {
            mAddedName = sAddedName;
            mNameSum = sNameSum;
            mChakraSexCreation = sChakraSexCreation;
            mCharkraThrought = sCharkraThrought;
            mCharkraBase = sCharkraBase;
            mCharkraThirdEye = sCharkraThirdEye;
            mCharkraHeart = sCharkraHeart;
            mPhysical = sPhysical;
            mEmotional = sEmotional;
            mMental = sMental;
            mEnergetic = sEnergetic;
            mMapBalance = sMapBalance;
            mMapDepration = sMapDepration;
            mBussiness = sBussiness;
            mLrnSccss = sLrnSccss;
            mCharkraHealth = sCharkraHealth;
        }
        #endregion
    }

    /*
    internal class UserInfo
    {
        #region Data Members
        public string mFirstName;
        public string mLastName;
        public string mMotherName;
        public string mFatherName;
        public DateTime mB_Date;
        public string mCity;
        public string mStreet;
        public double mBuildingNum;
        public double mAppNum;
        #endregion

        #region Constructor
        public UserInfo()
        {
            mFirstName = "";
            mLastName = "";
            mMotherName = "";
            mFatherName = "";
            mB_Date = DateTime.Now;
            mCity = "";
            mStreet = "";
            mBuildingNum = 0 ;
            mAppNum = 0;
        }
        #endregion

        #region Public Methods
        #endregion

        #region Private Methods
        #endregion

    }
    */
}
