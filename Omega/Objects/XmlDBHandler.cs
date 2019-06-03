using System;
using System.Collections.Generic;
using System.Text;
using BLL;
using System.Xml;
using System.IO;

namespace Omega.Objects
{
    public class XmlDBHandler
    {
        private string mXmlDBPath;
        public static XmlDBHandler Instance = new XmlDBHandler(BLL.AppSettings.Instance.AppmMainDir + "\\db.xml");

        private XmlDBHandler(string path)
        {
            mXmlDBPath = path;
            FileInfo dbFile = new FileInfo(mXmlDBPath);
            dbFile.IsReadOnly = false;
        }

        public void XmlDB2DataGridView(ref System.Windows.Forms.DataGridView dgv)
        {
            dgv.Rows.Clear();

            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.Load(mXmlDBPath);
            XmlNodeList root = xmlDoc.ChildNodes;
            foreach (XmlNode usernode in root[xmlDoc.ChildNodes.Count - 1].ChildNodes)
            {
                UserInfo curUser = new UserInfo();
                curUser.XmlNode2UserInfo(usernode);

                System.Windows.Forms.DataGridViewRow row = new System.Windows.Forms.DataGridViewRow();

                object[] objects = new string[16];
                objects.Initialize();
                //objects.SetValue(usernode.Attributes.GetNamedItem("id").InnerText.ToString(), 0);
                objects.SetValue(curUser.mId.ToString(), 0);
                objects.SetValue(curUser.mFirstName, 1);
                objects.SetValue(curUser.mLastName, 2);
                objects.SetValue(curUser.mFatherName, 3);
                objects.SetValue(curUser.mMotherName, 4);
                objects.SetValue(curUser.mB_Date.ToShortDateString(), 5);
                objects.SetValue(curUser.mCity, 6);
                objects.SetValue(curUser.mStreet, 7);
                objects.SetValue(curUser.mBuildingNum.ToString(), 8);
                objects.SetValue(curUser.mAppNum.ToString(), 9);
                objects.SetValue(curUser.mEMail, 10);
                objects.SetValue(curUser.mPhones, 11);
                objects.SetValue(curUser.mApplication, 12); // application
                objects.SetValue(curUser.mSex.ToString(), 13); // sex
                objects.SetValue(curUser.mPassedRect.ToString(), 14); // passed rectify
                objects.SetValue(curUser.mReachMaster.ToString(), 15); // reach master


                //objects.SetValue(Enums.EnumProvider.Instance.GetSexEnumFromDescription(curUser.mSex.ToString()), 12);
                //objects.SetValue(curUser.mApplication, 13);
                //objects.SetValue(Enums.EnumProvider.Instance.GetPassedRectificationEnumFromDescription(curUser.mPassedRect.ToString()), 14);

                dgv.Rows.Insert(dgv.Rows.Count, objects);
                dgv.Update();
            }
        }

        public bool AddUserToXmlDB(UserInfo user)
        {
            bool resFinal = true;
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.Load(mXmlDBPath);

            try
            {
                XmlNode newNode = xmlDoc.ChildNodes[0].ChildNodes[0].CloneNode(true);

                foreach (XmlNode innernode in newNode.ChildNodes)
                {
                    switch (Enums.EnumProvider.Instance.GetInnerXmlDBFieldsEnumFromString(innernode.Name.ToLower()))
                    {
                        #region convert data
                        case Enums.EnumProvider.InnerXmlDBFields.Appartment:
                            innernode.InnerText = user.mAppNum.ToString();
                            break;
                        case Enums.EnumProvider.InnerXmlDBFields.Application:
                            innernode.InnerText = user.mApplication.ToString();
                            break;
                        case Enums.EnumProvider.InnerXmlDBFields.BirthDay:
                            innernode.InnerText = user.mB_Date.ToString();
                            break;
                        case Enums.EnumProvider.InnerXmlDBFields.Building:
                            innernode.InnerText = user.mBuildingNum.ToString();
                            break;
                        case Enums.EnumProvider.InnerXmlDBFields.City:
                            innernode.InnerText = user.mCity.ToString();
                            break;
                        case Enums.EnumProvider.InnerXmlDBFields.EMail:
                            innernode.InnerText = user.mEMail.ToString();
                            break;
                        case Enums.EnumProvider.InnerXmlDBFields.FatherName:
                            innernode.InnerText = user.mFatherName.ToString();
                            break;
                        case Enums.EnumProvider.InnerXmlDBFields.FirstName:
                            innernode.InnerText = user.mFirstName.ToString();
                            break;
                        case Enums.EnumProvider.InnerXmlDBFields.LastName:
                            innernode.InnerText = user.mLastName.ToString();
                            break;
                        case Enums.EnumProvider.InnerXmlDBFields.MotherName:
                            innernode.InnerText = user.mMotherName.ToString();
                            break;
                        case Enums.EnumProvider.InnerXmlDBFields.PassedRect:
                            innernode.InnerText = user.mPassedRect.ToString();
                            break;
                        case Enums.EnumProvider.InnerXmlDBFields.Phones:
                            innernode.InnerText = user.mPhones.ToString();
                            break;
                        case Enums.EnumProvider.InnerXmlDBFields.Sex:
                            innernode.InnerText = user.mSex.ToString();
                            break;
                        case Enums.EnumProvider.InnerXmlDBFields.Street:
                            innernode.InnerText = user.mStreet.ToString();
                            break;
                        case Enums.EnumProvider.InnerXmlDBFields.ReachedMaster:
                            innernode.InnerText = user.mReachMaster.ToString();
                            break;
                            #endregion convert data
                    }
                }

                user.mId = xmlDoc.ChildNodes[0].ChildNodes.Count + 1;
                newNode.Attributes["id"].InnerText = user.mId.ToString();

                xmlDoc.ChildNodes[0].AppendChild(newNode);

                FileInfo xmlfile = new FileInfo(mXmlDBPath);
                xmlfile.Delete();
                xmlDoc.Save(mXmlDBPath);
            }
            catch
            {
                xmlDoc = new XmlDocument();
                resFinal = false;
            }

            return resFinal;
        }

        public bool AddUserToXmlDB(List<UserInfo> users)
        {
            bool resFinal = true;
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.Load(mXmlDBPath);

            try
            {
                foreach (UserInfo user in users)
                {
                    XmlNode newNode = xmlDoc.ChildNodes[0].ChildNodes[0].CloneNode(true);

                    foreach (XmlNode innernode in newNode.ChildNodes)
                    {
                        switch (Enums.EnumProvider.Instance.GetInnerXmlDBFieldsEnumFromString(innernode.Name.ToLower()))
                        {
                            #region convert data
                            case Enums.EnumProvider.InnerXmlDBFields.Appartment:
                                innernode.InnerText = user.mAppNum.ToString();
                                break;
                            case Enums.EnumProvider.InnerXmlDBFields.Application:
                                innernode.InnerText = user.mApplication.ToString();
                                break;
                            case Enums.EnumProvider.InnerXmlDBFields.BirthDay:
                                innernode.InnerText = user.mB_Date.ToString();
                                break;
                            case Enums.EnumProvider.InnerXmlDBFields.Building:
                                innernode.InnerText = user.mBuildingNum.ToString();
                                break;
                            case Enums.EnumProvider.InnerXmlDBFields.City:
                                innernode.InnerText = user.mCity.ToString();
                                break;
                            case Enums.EnumProvider.InnerXmlDBFields.EMail:
                                innernode.InnerText = user.mEMail.ToString();
                                break;
                            case Enums.EnumProvider.InnerXmlDBFields.FatherName:
                                innernode.InnerText = user.mFatherName.ToString();
                                break;
                            case Enums.EnumProvider.InnerXmlDBFields.FirstName:
                                innernode.InnerText = user.mFirstName.ToString();
                                break;
                            case Enums.EnumProvider.InnerXmlDBFields.LastName:
                                innernode.InnerText = user.mLastName.ToString();
                                break;
                            case Enums.EnumProvider.InnerXmlDBFields.MotherName:
                                innernode.InnerText = user.mMotherName.ToString();
                                break;
                            case Enums.EnumProvider.InnerXmlDBFields.PassedRect:
                                innernode.InnerText = user.mPassedRect.ToString();
                                break;
                            case Enums.EnumProvider.InnerXmlDBFields.Phones:
                                innernode.InnerText = user.mPhones.ToString();
                                break;
                            case Enums.EnumProvider.InnerXmlDBFields.Sex:
                                innernode.InnerText = user.mSex.ToString();
                                break;
                            case Enums.EnumProvider.InnerXmlDBFields.Street:
                                innernode.InnerText = user.mStreet.ToString();
                                break;
                            case Enums.EnumProvider.InnerXmlDBFields.ReachedMaster:
                                innernode.InnerText = user.mReachMaster.ToString();
                                break;
                                #endregion convert data
                        }
                    }

                    newNode.Attributes["id"].InnerText = (xmlDoc.ChildNodes[0].ChildNodes.Count + 1).ToString();

                    xmlDoc.ChildNodes[0].AppendChild(newNode);
                }

                FileInfo xmlfile = new FileInfo(mXmlDBPath);
                xmlfile.Delete();
                xmlDoc.Save(mXmlDBPath);
            }
            catch
            {
                xmlDoc = new XmlDocument();
                resFinal = false;
            }

            return resFinal;
        }

        public bool isUserInXmlDB(UserInfo user)
        {
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.Load(mXmlDBPath);

            bool resFinal = false;

            try
            {
                XmlNodeList root = xmlDoc.ChildNodes;
                foreach (XmlNode curUser in root[0].ChildNodes)
                {
                    //UserInfo thisUser = new UserInfo();
                    //thisUser.XmlNode2UserInfo(curUser);

                    //bool res = true;
                    ////res &= (thisUser.mApplication == user.mApplication);
                    //res &= (thisUser.mAppNum == user.mAppNum);
                    //res &= (thisUser.mB_Date == user.mB_Date);
                    //res &= (thisUser.mBuildingNum == user.mBuildingNum);
                    //res &= (thisUser.mCity  == user.mCity );
                    //res &= (thisUser.mEMail == user.mEMail );
                    //res &= (thisUser.mFatherName == user.mFatherName );
                    //res &= (thisUser.mFirstName == user. mFirstName );
                    //res &= (thisUser.mLastName == user.mLastName );
                    //res &= (thisUser.mMotherName == user.mMotherName );
                    //res &= (thisUser.mPassedRect == user.mPassedRect );
                    //res &= (thisUser.mReachMaster == user.mReachMaster);
                    //res &= (thisUser.mPhones  == user.mPhones );
                    //res &= (thisUser.mSex == user.mSex );
                    //res &= (thisUser.mStreet  == user.mStreet );

                    //if (res == true) return true;
                    //if (curUser.Attributes.GetNamedItem("FirstName").Value.Equals(user.mFirstName, StringComparison.CurrentCultureIgnoreCase) 
                    //    && (curUser.Attributes.GetNamedItem("LastName").Value.Equals(user.mLastName, StringComparison.CurrentCultureIgnoreCase)))
                    if (int.Parse(curUser.Attributes.GetNamedItem("id").Value) == user.mId)
                    {
                        resFinal = true;
                    }

                }

            }
            catch
            {
                resFinal = false;
            }

            return resFinal;

        }

        public bool RemoveUserFromDB(UserInfo user)
        {
            bool res = true;
            try
            {
                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.Load(mXmlDBPath);

                XmlNodeList root = xmlDoc.ChildNodes;
                foreach (XmlNode innernode in root[0].ChildNodes)
                {
                    UserInfo thisUser = new UserInfo();
                    thisUser.XmlNode2UserInfo(innernode);

                    bool resTest = true;
                    resTest &= (thisUser.mApplication == user.mApplication);
                    resTest &= (thisUser.mAppNum == user.mAppNum);
                    resTest &= (thisUser.mB_Date.Date == user.mB_Date.Date);
                    resTest &= (thisUser.mBuildingNum == user.mBuildingNum);
                    resTest &= (thisUser.mCity == user.mCity);
                    resTest &= (thisUser.mEMail == user.mEMail);
                    resTest &= (thisUser.mFatherName == user.mFatherName);
                    resTest &= (thisUser.mFirstName == user.mFirstName);
                    resTest &= (thisUser.mLastName == user.mLastName);
                    resTest &= (thisUser.mMotherName == user.mMotherName);
                    resTest &= (thisUser.mPassedRect == user.mPassedRect);
                    resTest &= (thisUser.mPhones == user.mPhones);
                    resTest &= (thisUser.mSex == user.mSex);
                    resTest &= (thisUser.mStreet == user.mStreet);
                    resTest &= (thisUser.mReachMaster == user.mReachMaster);

                    if (resTest == true)
                    {
                        xmlDoc.ChildNodes[0].RemoveChild(innernode);
                        break;
                    }
                }

                FileInfo xmlfile = new FileInfo(mXmlDBPath);
                xmlfile.Delete();
                xmlDoc.Save(mXmlDBPath);
            }
            catch //(Exception ex)
            {
                res = false;
            }
            return res;
        }
    }
}



//foreach (XmlNode node in user.ChildNodes)
//{
//    UserInfo curUser = new UserInfo();
//    switch (Enums.EnumProvider.Instance.GetInnerXmlDBFieldsEnumFromString(node.Name.ToLower()))
//    {
//        case Enums.EnumProvider.InnerXmlDBFields.Appartment:
//            curUser.mAppNum = Convert.ToDouble(node.Value.ToString());
//            break;
//        case Enums.EnumProvider.InnerXmlDBFields.Application:
//            curUser.mApplication = node.Value.ToString();
//            break;
//        case Enums.EnumProvider.InnerXmlDBFields.BirthDay:
//            DateTime bd = DateTime.Now;
//            bool res = DateTime.TryParse(node.Value.ToString(), out bd);
//            curUser.mB_Date = bd;
//            break;
//        case Enums.EnumProvider.InnerXmlDBFields.Building:
//            curUser.mBuildingNum = Convert.ToDouble(node.Value.ToString());
//            break;
//        case Enums.EnumProvider.InnerXmlDBFields.City:
//            curUser.mCity = node.Value.ToString();
//            break;
//        case Enums.EnumProvider.InnerXmlDBFields.EMail:
//            curUser.mEMail = node.Value.ToString();
//            break;
//        case Enums.EnumProvider.InnerXmlDBFields.FatherName:
//            curUser.mFatherName = node.Value.ToString();
//            break;
//        case Enums.EnumProvider.InnerXmlDBFields.FirstName:
//            curUser.mFirstName = node.Value.ToString();
//            break;
//        case Enums.EnumProvider.InnerXmlDBFields.LastName:
//            curUser.mLastName = node.Value.ToString();
//            break;
//        case Enums.EnumProvider.InnerXmlDBFields.MotherName:
//            curUser.mMotherName = node.Value.ToString();
//            break;
//        case Enums.EnumProvider.InnerXmlDBFields.PassedRect:
//            curUser.mPassedRect = Enums.EnumProvider.Instance.GetPassedRectificationEnumFromString(node.Value.ToString());
//            break;
//        case Enums.EnumProvider.InnerXmlDBFields.Phones:
//            curUser.mPhones = node.Value.ToString();
//            break;
//        case Enums.EnumProvider.InnerXmlDBFields.Sex:
//            curUser.mSex = Enums.EnumProvider.Instance.GetSexEnumFromString(node.Value.ToString());
//            break;
//        case Enums.EnumProvider.InnerXmlDBFields.Street:
//            curUser.mStreet = node.Value.ToString();
//            break;
//    }
//}