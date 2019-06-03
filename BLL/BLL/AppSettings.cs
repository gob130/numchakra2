using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using System.Windows.Forms;
using System.Diagnostics;
using Microsoft.Win32;
using System.Collections;
using System.IO;
using System.Management;
using System.Management.Instrumentation;
using System.Security.AccessControl;
using log4net;

namespace BLL
{
    public class AppSettings
    {
        #region Data Mebers
        private static AppSettings mInstance = new AppSettings(System.IO.Path.GetDirectoryName(System.Windows.Forms.Application.ExecutablePath));
        private ProgType mProgType;
        private Language mLang;
        private int mTrialTimeDuration = 30;

        private string mAppDir;

        private static readonly ILog mlog = LogManager.GetLogger("AppSettings");
        #endregion Data Mebers

        #region Enums
        public enum Language
        {
            Hebrew,
            English
        }

        public enum ProgType
        {
            Normal,
            NormalPlus,
            Expert,
            DAD,
            Trial
        }
        #endregion Enums

        #region Constructor
        public AppSettings(string sPath2AppDir)
        {
            mAppDir = sPath2AppDir;

            string sProgType, sProgLang;
            sProgType = "trial";
            sProgLang = "heb";

            string sPath2SettingsXML = mAppDir + "\\Settings\\Settings.xml";

            XmlDocument XmlDoc = new XmlDocument();
            XmlDoc.Load(sPath2SettingsXML);

            XmlNodeList root = XmlDoc.GetElementsByTagName("AppSettings");

            XmlNodeList AllNodes = root[0].ChildNodes;

            foreach (XmlNode CurrNode in AllNodes)
            {
                if (CurrNode.Name.ToLower() == "language")
                {
                    sProgLang = CurrNode.InnerText.ToLower().Trim();
                }

                if (CurrNode.Name.ToLower() == "type")
                {
                    sProgType = CurrNode.InnerText.ToLower().Trim();
                }
            }

            string[] splt = sProgType.Split(",".ToCharArray()[0]);

            switch (splt[0].ToLower())
            {
                case "trial":
                case "t":
                    mProgType = ProgType.Trial;
                    mTrialTimeDuration = Convert.ToInt16(splt[1]);
                    break;
                case "normal":
                    mProgType = ProgType.Normal;
                    break;
                case "expert":
                    mProgType = ProgType.Expert;
                    break;
                case "dad":
                    mProgType = ProgType.DAD;
                    break;
                default:
                    mProgType = ProgType.Trial;
                    break;
            }

            switch (sProgLang.ToLower())
            {
                case "heb":
                case "hrebrew":
                    mLang = Language.Hebrew;
                    break;
                case "eng":
                case "english":
                    mLang = Language.English;
                    break;
                default:
                    mLang = Language.Hebrew;
                    break;
            }
        }
        #endregion Constructor

        #region Properties
        public static AppSettings Instance
        {
            get
            {
                return mInstance;
            }
        }
        public string AppmMainDir
        {
            get
            {
                return mAppDir;
            }
        }

        public ProgType ProgramType
        {
            get
            {
                return mProgType;
            }
        }

        public Language ProgramLanguage
        {
            get
            {
                return mLang;
            }
        }

        public int TrialTimeDuration
        {
            get
            {
                return mTrialTimeDuration;
            }
        }
        #endregion Properties

        #region Public Methods
        #region Registry
        public bool RegisterDLLs(string dllPath)
        {
            bool res = true;
            try
            {
                //'/s' : indicates regsvr32.exe to run silently.
                string fileinfo = "/s" + " " + "\"" + dllPath + "\"";

                Process reg = new Process();
                reg.StartInfo.FileName = "regsvr32.exe";
                reg.StartInfo.Arguments = fileinfo;
                reg.StartInfo.UseShellExecute = false;
                reg.StartInfo.CreateNoWindow = true;
                reg.StartInfo.RedirectStandardOutput = true;
                reg.Start();
                reg.WaitForExit();
                reg.Close();
            }
            catch //(Exception ex)
            {
                //MessageBox.Show(ex.Message);
                res = false;
            }
            return res;
        }

        public bool isFileInRegistry(string dllName)
        {
            bool FinalRes = false;

            try
            {
                //get into the HKEY_CLASSES_ROOT
                RegistryKey root = Registry.ClassesRoot;

                //generic list to hold all the subkey names
                List<string> subKeys = new List<string>();

                //IEnumerator for enumerating through the subkeys
                IEnumerator RegEnums = root.GetSubKeyNames().GetEnumerator();

                //make sure we still have values
                while (RegEnums.MoveNext())
                {
                    //all registered extensions start with a period (.) so
                    //we need to check for that
                    if (RegEnums.Current.ToString().Contains("office") == true)
                        //valid extension so add it
                        subKeys.Add(RegEnums.Current.ToString());
                }

                foreach (string RegEntry in subKeys)
                {
                    if (RegEntry.Contains(dllName) == true)
                    {
                        FinalRes = true;
                    }
                }

            }
            catch
            {
                FinalRes = false;
            }

            return FinalRes;
        }
        #endregion Registry

        #region Premissions for Owner
        private List<string> GetUsers()
        {
            List<string> lsCurUsers = new List<string>();

            // This query will query for all user account names in our current Domain
            SelectQuery sQuery = new SelectQuery("Win32_UserAccount", "Domain='" + System.Environment.UserDomainName.ToString() + "'");

            try
            {
                // Searching for available Users
                ManagementObjectSearcher mSearcher = new ManagementObjectSearcher(sQuery);

                foreach (ManagementObject mObject in mSearcher.Get())
                {
                    // Adding all user names in our combobox
                    lsCurUsers.Add(mObject["Name"].ToString());
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            return lsCurUsers;
        }

        public void ChangeSecuritySettingsForApplicationDirectory()
        {
            // retrieving the directory information
            DirectoryInfo myDirectoryInfo = new DirectoryInfo(mAppDir);

            // Get a DirectorySecurity object that represents the 
            // current security settings.
            DirectorySecurity myDirectorySecurity = myDirectoryInfo.GetAccessControl();

            foreach (string user in GetUsers())
            {
                string User = System.Environment.UserDomainName + "\\" + user;

                // Add the FileSystemAccessRule to the security settings. 
                // FileSystemRights is a big list we are current using Read property but you 
                // can alter any other or many sme of which are:
                // Create Directories: for sub directories Authority
                // Create Files: for files creation access in a particular folder
                // Delete: for deletion athority on folder
                // Delete Subdirectories and files: for authority of deletion over 
                //subdirectories and files
                // Execute file: For execution accessibility in folder
                // Modify: For folder modification
                // Read: For directory opening
                // Write: to add things in directory
                // Full Control: For administration rights etc etc

                // Also AccessControlType which are of two kinds either “Allow” or “Deny” 
                myDirectorySecurity.AddAccessRule(new FileSystemAccessRule(User, FileSystemRights.FullControl, InheritanceFlags.ContainerInherit, PropagationFlags.InheritOnly, AccessControlType.Allow));

                // Set the new access settings. 
                myDirectoryInfo.SetAccessControl(myDirectorySecurity);
            }

            #region Everyone
            //myDirectorySecurity.AddAccessRule(new FileSystemAccessRule("Everyone", FileSystemRights.FullControl, InheritanceFlags.ContainerInherit, PropagationFlags.InheritOnly, AccessControlType.Allow));
            myDirectorySecurity.AddAccessRule(new FileSystemAccessRule("Everyone", FileSystemRights.Modify, InheritanceFlags.ObjectInherit, PropagationFlags.InheritOnly, AccessControlType.Allow));
            myDirectoryInfo.SetAccessControl(myDirectorySecurity);
            #endregion

            #region adding everyone
            // Also AccessControlType which are of two kinds either “Allow” or “Deny” 
            myDirectorySecurity.AddAccessRule(new FileSystemAccessRule("everyone", FileSystemRights.FullControl, InheritanceFlags.ContainerInherit, PropagationFlags.InheritOnly, AccessControlType.Allow));

            // Set the new access settings. 
            myDirectoryInfo.SetAccessControl(myDirectorySecurity);
            #endregion adding everyone
            // Showing a Succesfully Done Message
            //MessageBox.Show("Permissions Altered Successfully
        }


        public bool AddPremissions2SpecificFolder(string path)
        {
            bool res = true;
            try
            {
                mlog.InfoFormat("Try to set permissions on [{0}] directory", path);
                DirectoryInfo myDirectoryInfo = new DirectoryInfo(path);

                DirectorySecurity myDirectorySecurity = myDirectoryInfo.GetAccessControl();

                foreach (string user in GetUsers())
                {
                    string User = System.Environment.UserDomainName + "\\" + user;

                    // Add the FileSystemAccessRule to the security settings. 
                    // FileSystemRights is a big list we are current using Read property but you 
                    // can alter any other or many sme of which are:
                    // Create Directories: for sub directories Authority
                    // Create Files: for files creation access in a particular folder
                    // Delete: for deletion athority on folder
                    // Delete Subdirectories and files: for authority of deletion over 
                    //subdirectories and files
                    // Execute file: For execution accessibility in folder
                    // Modify: For folder modification
                    // Read: For directory opening
                    // Write: to add things in directory
                    // Full Control: For administration rights etc etc

                    // Also AccessControlType which are of two kinds either “Allow” or “Deny” 
                    myDirectorySecurity.AddAccessRule(new FileSystemAccessRule(User, FileSystemRights.FullControl, InheritanceFlags.ContainerInherit, PropagationFlags.InheritOnly, AccessControlType.Allow));

                    // Set the new access settings. 
                    myDirectoryInfo.SetAccessControl(myDirectorySecurity);
                }

                #region adding everyone
                // Also AccessControlType which are of two kinds either “Allow” or “Deny” 
                myDirectorySecurity.AddAccessRule(new FileSystemAccessRule("everyone", FileSystemRights.FullControl, InheritanceFlags.ContainerInherit, PropagationFlags.InheritOnly, AccessControlType.Allow));

                // Set the new access settings. 
                myDirectoryInfo.SetAccessControl(myDirectorySecurity);
                #endregion adding everyone

                mlog.Info("Permissions successfully set");
            }
            catch(Exception ex)
            {
                mlog.Error(ex);
                res = false;
            }
            return res;
        }
        #endregion Premissions for Owner

        public bool ChangeReportFileNames()
        {
            bool finalres = true;

            try
            {
                string path2TempletsDir = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(Application.ExecutablePath), @"Templets\ReportStyles");
                DirectoryInfo xmlDir = new DirectoryInfo(path2TempletsDir);

                foreach (FileInfo xmlFile in xmlDir.GetFileSystemInfos("*.xml"))
                {
                    XmlDocument xmlDoc = new XmlDocument();
                    xmlDoc.Load(System.IO.Path.Combine(xmlFile.DirectoryName, xmlFile.Name));

                    XmlNodeList root = xmlDoc.ChildNodes;

                    //string[] splt = root[1].OuterXml.Split(" ".ToCharArray()[0]);

                    string XmlName = root[1].Attributes.GetNamedItem("name").InnerText.ToString();

                    xmlFile.CopyTo(System.IO.Path.Combine(xmlFile.DirectoryName, XmlName + ".xml"));
                    xmlFile.Delete();
                }
            }
            catch
            {
                finalres = false;
            }

            return finalres;
        }

        public string GetXmlDbPath()
        {
            return System.IO.Path.Combine(mAppDir, "db.xml");
        }

        public bool CreateNewDB()
        {
            bool finalres = true;

            try
            {
                string xmlpath = System.IO.Path.Combine(mAppDir, "db.xml");

                XmlTextWriter mXmlTextWriter = new XmlTextWriter(xmlpath, Encoding.UTF8);
                mXmlTextWriter.Formatting = Formatting.Indented;

                mXmlTextWriter.WriteStartDocument();

                mXmlTextWriter.WriteStartElement("DB");
                mXmlTextWriter.WriteEndElement();

                mXmlTextWriter.Flush();

                mXmlTextWriter.WriteEndDocument();
                mXmlTextWriter.Close();

                GC.WaitForPendingFinalizers();
                GC.Collect();

                XmlDocument db = new XmlDocument();
                db.Load(xmlpath);

                XmlElement user = db.CreateElement("", "User", "");
                user.SetAttribute("id", "1");

                XmlElement data;
                string val;

                #region First Name
                data = db.CreateElement("", "FirstName", "");
                val = "";
                if (BLL.AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                {
                    val = "יעקבי";
                }
                if (BLL.AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                {
                    val = "yaakobi";
                }
                data.AppendChild(db.CreateTextNode(val));
                user.AppendChild(data);
                #endregion First Name

                #region Last Name
                data = db.CreateElement("", "LastName", "");
                val = "";
                if (BLL.AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                {
                    val = "נשר-סולן";
                }
                if (BLL.AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                {
                    val = "nesher-solan";
                }
                data.AppendChild(db.CreateTextNode(val));
                user.AppendChild(data);
                #endregion Last Name

                #region BirthDay
                data = db.CreateElement("", "BirthDay", "");
                val = (new DateTime(1958, 11, 11)).ToLongDateString();
                data.AppendChild(db.CreateTextNode(val));
                user.AppendChild(data);
                #endregion

                #region FatherName
                data = db.CreateElement("", "FatherName", "");
                val = "";
                if (BLL.AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                {
                    val = "אבא";
                }
                if (BLL.AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                {
                    val = "dad";
                }
                data.AppendChild(db.CreateTextNode(val));
                user.AppendChild(data);
                #endregion

                #region MotherName
                data = db.CreateElement("", "MotherName", "");
                val = "";
                if (BLL.AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                {
                    val = "אמא";
                }
                if (BLL.AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                {
                    val = "mom";
                }
                data.AppendChild(db.CreateTextNode(val));
                user.AppendChild(data);
                #endregion

                #region City
                data = db.CreateElement("", "City", "");
                val = "";
                if (BLL.AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                {
                    val = "חולון";
                }
                if (BLL.AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                {
                    val = "city";
                }
                data.AppendChild(db.CreateTextNode(val));
                user.AppendChild(data);
                #endregion

                #region Street
                data = db.CreateElement("", "Street", "");
                val = "";
                if (BLL.AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                {
                    val = "רחוב";
                }
                if (BLL.AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                {
                    val = "street";
                }
                data.AppendChild(db.CreateTextNode(val));
                user.AppendChild(data);
                #endregion

                #region Building
                data = db.CreateElement("", "Building", "");
                val = "";
                if (BLL.AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                {
                    val = "1";
                }
                if (BLL.AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                {
                    val = "1";
                }
                data.AppendChild(db.CreateTextNode(val));
                user.AppendChild(data);
                #endregion

                #region Appartment
                data = db.CreateElement("", "Appartment", "");
                val = "";
                if (BLL.AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                {
                    val = "0";
                }
                if (BLL.AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                {
                    val = "0";
                }
                data.AppendChild(db.CreateTextNode(val));
                user.AppendChild(data);
                #endregion

                #region EMail
                data = db.CreateElement("", "EMail", "");
                val = "yaakobi999@gmail.com";
                data.AppendChild(db.CreateTextNode(val));
                user.AppendChild(data);
                #endregion

                #region Phones
                data = db.CreateElement("", "Phones", "");
                val = "";
                if (BLL.AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                {
                    val = "052-4694122";
                }
                if (BLL.AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                {
                    val = "+972-52-4694122";
                }
                data.AppendChild(db.CreateTextNode(val));
                user.AppendChild(data);
                #endregion

                #region Sex
                data = db.CreateElement("", "Sex", "");
                val = "";
                if (BLL.AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                {
                    val = "Male";
                }
                if (BLL.AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                {
                    val = "Male";
                }
                data.AppendChild(db.CreateTextNode(val));
                user.AppendChild(data);
                #endregion

                #region Application
                data = db.CreateElement("", "Application", "");
                val = "";
                if (BLL.AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                {
                    val = "נומרולוג";
                }
                if (BLL.AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                {
                    val = "Numerologist";
                }
                data.AppendChild(db.CreateTextNode(val));
                user.AppendChild(data);
                #endregion

                #region PassedRect
                data = db.CreateElement("", "PassedRect", "");
                val = "";
                if (BLL.AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                {
                    val = "Passed";
                }
                if (BLL.AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                {
                    val = "Passed";
                }
                data.AppendChild(db.CreateTextNode(val));
                user.AppendChild(data);
                #endregion

                #region PassedRect
                data = db.CreateElement("", "ReachedMaster", "");
                val = "";
                if (BLL.AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                {
                    val = "Yes";
                }
                if (BLL.AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                {
                    val = "Yes";
                }
                data.AppendChild(db.CreateTextNode(val));
                user.AppendChild(data);
                #endregion

                XmlElement mRepRoot = (XmlElement)db.ChildNodes[1];
                mRepRoot.AppendChild(user);

                FileInfo file = new FileInfo(xmlpath);

                bool found = file.Exists;
                if (found == true)
                {
                    file.Delete();
                }

                db.Save(xmlpath);

                #region rewrite
                FileStream fh = new FileStream(System.IO.Path.Combine(mAppDir, "db.txt"), FileMode.Create);
                StreamWriter wrtr = new StreamWriter(fh);

                FileStream fhDB = new FileStream(xmlpath, FileMode.Open);
                StreamReader fr = new StreamReader(fhDB);

                string line = fr.ReadLine(); // Header!
                string sAll = "";
                while (fr.EndOfStream == false)
                {
                    sAll += fr.ReadLine() + Environment.NewLine;
                }
                sAll = sAll.Trim();
                wrtr.WriteLine(sAll);

                wrtr.Close(); fh.Close();
                fr.Close(); fhDB.Close();

                file.Delete(); // old db.xml handler
                file = new FileInfo(System.IO.Path.Combine(mAppDir, "db.txt"));
                file.CopyTo(xmlpath);
                file.Delete();
                #endregion rewrite
            }
            catch
            {
                finalres = false;
            }

            return finalres;
        }
        #endregion Public Methods

        #region Private Methods
        #endregion Private Methods
    }
}

