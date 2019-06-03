using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;

namespace Omega.Objects
{
    public class UserInfo
    {
        #region Data Members
        public int mId;
        public string mFirstName;
        public string mLastName;
        public string mMotherName;
        public string mFatherName;
        public DateTime mB_Date;
        public string mCity;
        public string mStreet;
        public double mBuildingNum;
        public double mAppNum;
        public string mEMail;
        public string mPhones;
        public string mApplication;
        public Enums.EnumProvider.Sex mSex;
        public Enums.EnumProvider.PassedRectification mPassedRect;
        public Enums.EnumProvider.ReachedMaster mReachMaster;
        #endregion

        #region Constructor
        public UserInfo()
        {
            mId = 0;
            mFirstName = "";
            mLastName = "";
            mMotherName = "";
            mFatherName = "";
            mB_Date = DateTime.Now;
            mCity = "";
            mStreet = "";
            mBuildingNum = 0;
            mAppNum = 0;
            mEMail = "";
            mPhones = "";
            mSex = Enums.EnumProvider.Sex.Male;
            mPassedRect = Enums.EnumProvider.PassedRectification.NotPassed;
            mApplication = "";
            mReachMaster = Enums.EnumProvider.ReachedMaster.No;
        }
        #endregion

        #region Getters/Setters

        public int Age
        {
            get
            {
                return Convert.ToInt16(DateTime.Today.Subtract(mB_Date).TotalDays / 365);

            }

        }

        #endregion Getters/Setters

        #region Public Methods
        public void XmlNode2UserInfo(XmlNode usernode)
        {
            mId = Convert.ToInt16(usernode.Attributes["id"].Value);

            foreach (XmlNode node in usernode.ChildNodes)
            {
                switch (Enums.EnumProvider.Instance.GetInnerXmlDBFieldsEnumFromString(node.Name.ToLower()))
                {
                    case Enums.EnumProvider.InnerXmlDBFields.Appartment:
                        mAppNum = Convert.ToDouble(node.InnerText.ToString());
                        break;
                    case Enums.EnumProvider.InnerXmlDBFields.Application:
                        mApplication = node.InnerText.ToString();
                        break;
                    case Enums.EnumProvider.InnerXmlDBFields.BirthDay:
                        DateTime bd = DateTime.Now;
                        bool res = DateTime.TryParse(node.InnerText.ToString(), out bd);
                        mB_Date = bd;
                        break;
                    case Enums.EnumProvider.InnerXmlDBFields.Building:
                        mBuildingNum = Convert.ToDouble(node.InnerText.ToString());
                        break;
                    case Enums.EnumProvider.InnerXmlDBFields.City:
                        mCity = node.InnerText.ToString();
                        break;
                    case Enums.EnumProvider.InnerXmlDBFields.EMail:
                        mEMail = node.InnerText.ToString();
                        break;
                    case Enums.EnumProvider.InnerXmlDBFields.FatherName:
                        mFatherName = node.InnerText.ToString();
                        break;
                    case Enums.EnumProvider.InnerXmlDBFields.FirstName:
                        mFirstName = node.InnerText.ToString();
                        break;
                    case Enums.EnumProvider.InnerXmlDBFields.LastName:
                        mLastName = node.InnerText.ToString();
                        break;
                    case Enums.EnumProvider.InnerXmlDBFields.MotherName:
                        mMotherName = node.InnerText.ToString();
                        break;
                    case Enums.EnumProvider.InnerXmlDBFields.PassedRect:
                        mPassedRect = Enums.EnumProvider.Instance.GetPassedRectificationEnumFromString(node.InnerText.ToString());
                        break;
                    case Enums.EnumProvider.InnerXmlDBFields.Phones:
                        mPhones = node.InnerText.ToString();
                        break;
                    case Enums.EnumProvider.InnerXmlDBFields.Sex:
                        mSex = Enums.EnumProvider.Instance.GetSexEnumFromString(node.InnerText.ToString());
                        break;
                    case Enums.EnumProvider.InnerXmlDBFields.Street:
                        mStreet = node.InnerText.ToString();
                        break;
                    case Enums.EnumProvider.InnerXmlDBFields.ReachedMaster:
                        mReachMaster = Enums.EnumProvider.Instance.GGetIsMasterEnumFromString(node.InnerText.ToString());
                        break;

                }

            }
        }

        public bool ProgData2UserInfo(MainForm mainForm)
        {
            bool resFInal = true;

            try
            {
                mFirstName = mainForm.Controls.Find("txtPrivateName", true)[0].Text;
                mLastName = mainForm.Controls.Find("txtFamilyName", true)[0].Text;
                mMotherName = mainForm.Controls.Find("txtMotherName", true)[0].Text;
                mFatherName = mainForm.Controls.Find("txtFatherName", true)[0].Text;
                mCity = mainForm.Controls.Find("txtCity", true)[0].Text;
                mStreet = mainForm.Controls.Find("txtStreet", true)[0].Text;
                mBuildingNum = Convert.ToDouble(mainForm.Controls.Find("txtBiuldingNum", true)[0].Text);
                mAppNum = Convert.ToDouble(mainForm.Controls.Find("txtAppNum", true)[0].Text);
                mEMail = mainForm.Controls.Find("txtEMail", true)[0].Text;
                mPhones = mainForm.Controls.Find("txtPhones", true)[0].Text;
                mApplication = mainForm.Controls.Find("txtApplication", true)[0].Text;

                mB_Date = (mainForm.Controls.Find("DateTimePickerFrom", true)[0] as System.Windows.Forms.DateTimePicker).Value;
                string sSex = (mainForm.Controls.Find("cmbSexSelect", true)[0] as System.Windows.Forms.ComboBox).SelectedItem.ToString();
                switch (BLL.AppSettings.Instance.ProgramLanguage)
                {
                    case BLL.AppSettings.Language.Hebrew:
                        mSex = Omega.Enums.EnumProvider.Instance.GetSexEnumFromString(sSex);
                        break;
                    case BLL.AppSettings.Language.English:
                        if (sSex.ToLower() == Omega.Enums.EnumProvider.Sex.Female.ToString().ToLower())
                        {
                            mSex = Omega.Enums.EnumProvider.Sex.Female;
                        }
                        if (sSex.ToLower() == Omega.Enums.EnumProvider.Sex.Male.ToString().ToLower())
                        {
                            mSex = Omega.Enums.EnumProvider.Sex.Male;
                        }
                        break;
                }

                bool cbPassedValue = (mainForm.Controls.Find("cbMainCorrectionDone", true)[0] as System.Windows.Forms.CheckBox).Checked;
                if (cbPassedValue == true)
                {
                    mPassedRect = Enums.EnumProvider.PassedRectification.Passed;
                }
                else
                {
                    mPassedRect = Enums.EnumProvider.PassedRectification.NotPassed;
                }

                bool cbMasterValue = (mainForm.Controls.Find("cbPersonMaster", true)[0] as System.Windows.Forms.CheckBox).Checked;
                if (cbMasterValue == true)
                {
                    mReachMaster = Enums.EnumProvider.ReachedMaster.Yes;
                }
                else
                {
                    mReachMaster = Enums.EnumProvider.ReachedMaster.No;
                }

            }
            catch // (Exception exp)
            {
                resFInal = false;
            }

            return resFInal;
        }

        public bool ApplyUserInfo2ProgData(MainForm mainForm)
        {
            bool resFinal = true;
            try
            {
                mainForm.Controls.Find("txtPrivateName", true)[0].Text = mFirstName;
                mainForm.Controls.Find("txtFamilyName", true)[0].Text = mLastName;
                mainForm.Controls.Find("txtMotherName", true)[0].Text = mMotherName;
                mainForm.Controls.Find("txtFatherName", true)[0].Text = mFatherName;
                mainForm.Controls.Find("txtCity", true)[0].Text = mCity;
                mainForm.Controls.Find("txtStreet", true)[0].Text = mStreet;
                mainForm.Controls.Find("txtBiuldingNum", true)[0].Text = mBuildingNum.ToString();
                mainForm.Controls.Find("txtAppNum", true)[0].Text = mAppNum.ToString();
                mainForm.Controls.Find("txtEMail", true)[0].Text = mEMail;
                mainForm.Controls.Find("txtPhones", true)[0].Text = mPhones;
                mainForm.Controls.Find("txtApplication", true)[0].Text = mApplication;

                (mainForm.Controls.Find("DateTimePickerFrom", true)[0] as System.Windows.Forms.DateTimePicker).Value = mB_Date;

                if (mSex == Enums.EnumProvider.Sex.Male)
                {
                    (mainForm.Controls.Find("cmbSexSelect", true)[0] as System.Windows.Forms.ComboBox).SelectedItem =
                        (mainForm.Controls.Find("cmbSexSelect", true)[0] as System.Windows.Forms.ComboBox).Items[0];
                }
                else
                {
                    (mainForm.Controls.Find("cmbSexSelect", true)[0] as System.Windows.Forms.ComboBox).SelectedItem =
                        (mainForm.Controls.Find("cmbSexSelect", true)[0] as System.Windows.Forms.ComboBox).Items[1];
                }
                (mainForm.Controls.Find("cmbSexSelect", true)[0] as System.Windows.Forms.ComboBox).Text =
                    (mainForm.Controls.Find("cmbSexSelect", true)[0] as System.Windows.Forms.ComboBox).SelectedItem.ToString();

                if (mPassedRect == Enums.EnumProvider.PassedRectification.Passed)
                {
                    (mainForm.Controls.Find("cbMainCorrectionDone", true)[0] as System.Windows.Forms.CheckBox).Checked = true;
                }
                else
                {
                    (mainForm.Controls.Find("cbMainCorrectionDone", true)[0] as System.Windows.Forms.CheckBox).Checked = false;
                }

                if (mReachMaster == Enums.EnumProvider.ReachedMaster.Yes)
                {
                    (mainForm.Controls.Find("cbPersonMaster", true)[0] as System.Windows.Forms.CheckBox).Checked = true;
                }
                else
                {
                    (mainForm.Controls.Find("cbPersonMaster", true)[0] as System.Windows.Forms.CheckBox).Checked = false;
                }

                mainForm.Refresh();
            }
            catch
            {
            }
            return resFinal;
        }
        #endregion

        #region Private Methods
        #endregion

    }
}
