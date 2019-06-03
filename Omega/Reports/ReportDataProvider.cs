#region Namespaces

using BLL;
using log4net;
using Omega.Enums;
using Omega.Objects;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;

#endregion Namespaces

// **********

namespace Omega.Reports
{
    public class ReportDataProvider
    {
        public static ReportDataProvider Instance = new ReportDataProvider();
        private List<string> mResevredNames;
        private List<string> mReservedImages;
        private List<string> mReservedChakra;
        private List<string> mReservedCycles;
        private List<string> mReservedOPS;
        private List<string> mBalanced;
        private List<string> mReservedCrysis;
        private List<string> mReservedWork;
        private Omega.Objects.UserInfo mCurrentUser;
        public List<string> ReportTemplets;
        //public Form          mMainForm;
        public MainForm mMainForm;
        private ucSpinner spinner;

        private bool mPro;
        private AppSettings.Language mLang;
        private Enums.EnumProvider.Sex mGender2Print;
        //private string prefix;

        private bool mIsMSOfficeOK = false;

        private string mRepStyle2Create;

        private int mMaxWordsInLine = 10;//max word in row (for Word Wrap)

        private Calc mCalc;// = new BLL.Calc(AppSettings.Instance.ProgramLanguage);
        public List<string> mPitPlanes;
        private static readonly ILog mlog = LogManager.GetLogger("ReportDataProvider");

        #region Print Booleans
        public bool PrntMain = true;
        public bool PrntPitagoras = true;
        public bool PrntLifeCycles = true;
        public bool PrntIntansiveMap = true;
        public bool PrntCombinedMap = true;
        public bool PrntChakraOpening = true;
        public bool PrntBsnnsPersonal = true;
        public bool PrntBsnnsMulti = true;
        public bool PrntCoupleMatch = true;
        public bool PrntLearnSccss = true;
        #endregion 

        #region Chakra_Values_Text
        public string mChkraPreName = "chkra";
        public string mNumberPreName = "num";
        public string mSeporator = "_";
        #endregion Chakra_Values_Text

        #region Reports Data Files Location
        public string mFolderPath2ChakraInfoFiles = "";
        public string mFolderPath2AstroLuckInfoFiles = "";
        public string mFolderPath2PrvtNameInfoFiles = "";
        public string mFolderPath2PwrNumInfoFiles = "";
        public string mFolderPath2ChlngNumInfoFiles = "";
        public string mFolderPath2ClmxNumInfoFiles = "";
        public string mFolderPath2CycleNumInfoFiles = "";
        public string mFolderPath2PrsonalYearInfoFiles = "";
        public string mFolderPath2PrsonalMonthInfoFiles = "";
        public string mFolderPath2PrsonalDayInfoFiles = "";

        public string mFolderPath2WorkInfoFiles = "";
        public string mFolderPath2CrysisInfoFiles = "";
        public string mFolderPath2IZUNInfoFiles = "";
        public string mFolderPath2RectificationInfoFiles = "";
        public string mFolderPath2DesginationInfoFiles = "";
        public string mFolderPath2PitagorasInfoFiles = "";

        public string mFolderPath2MarriageSccssInfoFiles = "";
        public string mFolderPath2BussinessSccssInfoFiles = "";
        public string mFolderPath2LearnSccssInfoFiles = "";
        public string mFolderPath2ADHDInfoFiles = "";

        public string mFolderPath2ChakraOpening = "";
        public string mFolderPath2MainPersonality = "";

        public string mFolderPath2HealthInfoFiles = "";
        public string mFolderPath2FinalSummaryInfoFiles = "";
        public string mFolderPath2FinalMapStrengthInfoFiles = "";

        #endregion Reports Data Files Location

        #region properties
        public int MaxWordsInLine
        {
            get
            {
                return mMaxWordsInLine;
            }
        }

        public bool Pro
        {
            get
            {
                return mPro;
            }
            set
            {
                bool resParse = false;
                resParse = bool.TryParse(value.ToString(), out mPro);

                if (resParse)
                {
                    mPro = value;
                }
            }
        }

        public AppSettings.Language Language
        {
            get
            {
                return mLang;
            }
            set
            {
                mLang = value;
            }
        }

        public Enums.EnumProvider.Sex Gender
        {
            get
            {
                return mGender2Print;
            }
            set
            {
                mGender2Print = value;
            }
        }

        public string Prefix
        {
            get
            {
                string outs = "";
                if (mPro == true)
                {
                    outs += "pro_";
                }
                switch (mLang)
                {
                    case AppSettings.Language.English:
                        outs += "eng_";
                        break;
                    case AppSettings.Language.Hebrew:
                        outs += "heb_";
                        break;
                }
                if (Gender == EnumProvider.Sex.Male)
                {
                    // nothing
                }
                if (Gender == EnumProvider.Sex.Female)
                {
                    outs += "fm_";
                }
                return outs;
            }
        }
        #endregion properties

        private ReportDataProvider()
        {
            #region Reserved Words and Images
            //string path2ReservedNames = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(Application.ExecutablePath), @"Templets\ReportSettings.xml");

            mResevredNames = new List<string>();
            mReservedChakra = new List<string>();
            mReservedImages = new List<string>();
            mReservedCycles = new List<string>();
            mReservedOPS = new List<string>();
            mBalanced = new List<string>();
            mReservedCrysis = new List<string>();
            mReservedWork = new List<string>();

            foreach (Omega.Enums.EnumProvider.ReservedNames item in Enum.GetValues(typeof(Omega.Enums.EnumProvider.ReservedNames)))
            {
                mResevredNames.Add(item.ToString());
            }

            foreach (Omega.Enums.EnumProvider.Balance item in Enum.GetValues(typeof(Omega.Enums.EnumProvider.Balance)))
            {
                mBalanced.Add(item.ToString());
            }

            foreach (Omega.Enums.EnumProvider.ReservedCrysis item in Enum.GetValues(typeof(Omega.Enums.EnumProvider.ReservedCrysis)))
            {
                mReservedCrysis.Add(item.ToString());
            }

            foreach (Omega.Enums.EnumProvider.ReservedWork item in Enum.GetValues(typeof(Omega.Enums.EnumProvider.ReservedWork)))
            {
                mReservedWork.Add(item.ToString());
            }

            foreach (Omega.Enums.EnumProvider.ReservedImages item in Enum.GetValues(typeof(Omega.Enums.EnumProvider.ReservedImages)))
            {
                mReservedImages.Add(item.ToString());
            }

            foreach (Omega.Enums.EnumProvider.ReservedChakra item in Enum.GetValues(typeof(Omega.Enums.EnumProvider.ReservedChakra)))
            {
                mReservedChakra.Add(item.ToString());
            }

            foreach (Omega.Enums.EnumProvider.ReservedOPS item in Enum.GetValues(typeof(Omega.Enums.EnumProvider.ReservedOPS)))
            {
                mReservedOPS.Add(item.ToString());
            }

            foreach (Omega.Enums.EnumProvider.ReservedLifeCycle item in Enum.GetValues(typeof(Omega.Enums.EnumProvider.ReservedLifeCycle)))
            {
                mReservedCycles.Add(item.ToString());
            }

            mPitPlanes = new List<string>() {    "01. (3,6,9)",
                                                 "02. (2,5,8)",
                                                 "03. (1,4,7)",
                                                 "04. (1,5,9)",
                                                 "05. (1,2,3)",
                                                 "06. (4,5,6)",
                                                 "07. (7,8,9)",
                                                 "08. (3,5,7)",
                                                 "09. (1,2,3,4,5,6,7,8,9)"};

            #endregion reserved words

            mMainForm = MainForm.ActiveForm as MainForm;

            //if (mMainForm.Controls["grpStyle"].Controls.ContainsKey(mMainForm.ProcessSpinnerName))
            //{
            //    spinner = (ucSpinner)mMainForm.Controls["grpStyle"].Controls[mMainForm.ProcessSpinnerName];

            //}

            bool isMaster = !((mMainForm.Controls.Find("cbPersonMaster", true)[0] as CheckBox).Checked);
            mCalc = new BLL.Calc(AppSettings.Instance.ProgramLanguage, isMaster);

            #region Report Type List
            if (AppSettings.Instance.ProgramType != AppSettings.ProgType.Trial)
            {
                RefreshReportList();
            }
            #endregion

            #region Reports Info Location
            mFolderPath2ChakraInfoFiles = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(Application.ExecutablePath), @"Templets\InternalChakraValues");
            mFolderPath2AstroLuckInfoFiles = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(Application.ExecutablePath), @"Templets\InternalAstronomicLuckValues");
            mFolderPath2PrvtNameInfoFiles = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(Application.ExecutablePath), @"Templets\InternalPersonalNameValues");
            mFolderPath2PwrNumInfoFiles = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(Application.ExecutablePath), @"Templets\InternalPowerNumberValues");
            mFolderPath2ChlngNumInfoFiles = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(Application.ExecutablePath), @"Templets\InternalChalengeValues");
            mFolderPath2ClmxNumInfoFiles = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(Application.ExecutablePath), @"Templets\InternalClimaxValues");
            mFolderPath2CycleNumInfoFiles = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(Application.ExecutablePath), @"Templets\InternalCycleValues");
            mFolderPath2PrsonalYearInfoFiles = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(Application.ExecutablePath), @"Templets\InternalPersonalYearValues");
            mFolderPath2PrsonalMonthInfoFiles = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(Application.ExecutablePath), @"Templets\InternalPersonalMonthValues");
            mFolderPath2PrsonalDayInfoFiles = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(Application.ExecutablePath), @"Templets\InternalPersonalDayValues");
            mFolderPath2RectificationInfoFiles = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(Application.ExecutablePath), @"Templets\InternalRectificationValues");
            mFolderPath2CrysisInfoFiles = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(Application.ExecutablePath), @"Templets\InternalCrysisValues");
            mFolderPath2IZUNInfoFiles = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(Application.ExecutablePath), @"Templets\InternalIzunValues");
            mFolderPath2DesginationInfoFiles = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(Application.ExecutablePath), @"Templets\InternalDesignationValues");
            mFolderPath2WorkInfoFiles = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(Application.ExecutablePath), @"Templets\InternalWorkValues");
            mFolderPath2PitagorasInfoFiles = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(Application.ExecutablePath), @"Templets\InternalPitagorasValues");
            mFolderPath2MarriageSccssInfoFiles = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(Application.ExecutablePath), @"Templets\InternalMarriageValues");
            mFolderPath2BussinessSccssInfoFiles = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(Application.ExecutablePath), @"Templets\InternalBssnssSccssValues");
            mFolderPath2ChakraOpening = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(Application.ExecutablePath), @"Templets\InternalChakraOpening");
            mFolderPath2MainPersonality = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(Application.ExecutablePath), @"Templets\InternalMainPersonality");
            mFolderPath2LearnSccssInfoFiles = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(Application.ExecutablePath), @"Templets\InternalLearnSccssPersonality");
            mFolderPath2ADHDInfoFiles = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(Application.ExecutablePath), @"Templets\InternalADHDPersonality");
            mFolderPath2HealthInfoFiles = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(Application.ExecutablePath), @"Templets\InternalHealth");
            mFolderPath2FinalSummaryInfoFiles = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(Application.ExecutablePath), @"Templets\InternalFinalSummary");
            mFolderPath2FinalMapStrengthInfoFiles = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(Application.ExecutablePath), @"Templets\InternalMapStrength");
            if (BLL.AppSettings.Instance.ProgramType != AppSettings.ProgType.Trial)
            {
                chkFolder(mFolderPath2ChakraInfoFiles);
                chkFolder(mFolderPath2AstroLuckInfoFiles);
                chkFolder(mFolderPath2PrvtNameInfoFiles);
                chkFolder(mFolderPath2PwrNumInfoFiles);
                chkFolder(mFolderPath2ChlngNumInfoFiles);
                chkFolder(mFolderPath2ClmxNumInfoFiles);
                chkFolder(mFolderPath2CycleNumInfoFiles);
                chkFolder(mFolderPath2PrsonalYearInfoFiles);
                chkFolder(mFolderPath2PrsonalMonthInfoFiles);
                chkFolder(mFolderPath2PrsonalDayInfoFiles);
                chkFolder(mFolderPath2RectificationInfoFiles);
                chkFolder(mFolderPath2CrysisInfoFiles);
                chkFolder(mFolderPath2IZUNInfoFiles);
                chkFolder(mFolderPath2DesginationInfoFiles);
                chkFolder(mFolderPath2WorkInfoFiles);
                chkFolder(mFolderPath2PitagorasInfoFiles);
                chkFolder(mFolderPath2MarriageSccssInfoFiles);
                chkFolder(mFolderPath2BussinessSccssInfoFiles);
                chkFolder(mFolderPath2ChakraOpening);
                chkFolder(mFolderPath2MainPersonality);
                chkFolder(mFolderPath2LearnSccssInfoFiles);
                chkFolder(mFolderPath2ADHDInfoFiles);
                chkFolder(mFolderPath2HealthInfoFiles);
                chkFolder(mFolderPath2FinalSummaryInfoFiles);
                chkFolder(mFolderPath2FinalMapStrengthInfoFiles);
            }
            #endregion Reports Info Location

            mPro = (mMainForm.Controls.Find("cbProRep", true)[0] as CheckBox).Checked;
            mLang = AppSettings.Instance.ProgramLanguage;
            mGender2Print = EnumProvider.Sex.Male; // MALE

        }

        public bool CheckMSOffice()
        {
            bool res = true;
            try
            { // if this one can be done - than MS Office is OK on the PC
                WordWriter ww = new WordWriter();
                Omega.Reports.wordText wt = new wordText();
                ww.CloseWordWriter();
            }
            catch
            {
                res = false;
            }

            mIsMSOfficeOK = res; // setting this value to an inside parameter - needed in the printing stage
            return res;
        }

        public void RefreshReportList()
        {
            ReportTemplets = new List<string>();
            string path2TempletsDir = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(Application.ExecutablePath), @"Templets\ReportStyles");
            DirectoryInfo xmlDir = new DirectoryInfo(path2TempletsDir);

            foreach (FileInfo xmlFile in xmlDir.GetFileSystemInfos("*.xml"))
            {
                //XmlDocument repdoc = new XmlDocument();
                //repdoc.Load(xmlFile.FullName);

                //XmlNodeList root = repdoc.ChildNodes;
                //string repname = root[1].Attributes.GetNamedItem("name").InnerText.ToString();

                //if (repname != System.IO.Path.GetFileNameWithoutExtension(xmlFile.FullName))
                //{
                //    xmlFile.CopyTo(System.IO.Path.Combine(path2TempletsDir, repname + ".xml"));
                //    xmlFile.Delete();
                //}

                //ReportTemplets.Add(repname);
                ReportTemplets.Add(System.IO.Path.GetFileNameWithoutExtension(xmlFile.FullName));
            }
        }

        private void chkFolder(string path)
        {
            DirectoryInfo dir = new DirectoryInfo(path);
            if (dir.Exists == false)
            {
                dir.Create();
            }
        }

        public bool Initialize(Omega.Objects.UserInfo user)
        {
            mCurrentUser = user;

            Gender = mCurrentUser.mSex;
            return true;
        }

        public XmlDocument CreateNewReport(string name)
        {
            string path2TempletsDir = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(Application.ExecutablePath), @"Templets\ReportStyles");
            string mXMLLocation = Path.Combine(path2TempletsDir, name + ".xml");

            XmlTextWriter mXmlTextWriter = new XmlTextWriter(mXMLLocation, Encoding.UTF8);
            mXmlTextWriter.Formatting = Formatting.Indented;

            mXmlTextWriter.WriteStartDocument();

            mXmlTextWriter.WriteStartElement("Report");
            mXmlTextWriter.WriteEndElement();


            mXmlTextWriter.Flush();

            mXmlTextWriter.WriteEndDocument();
            mXmlTextWriter.Close();

            GC.WaitForPendingFinalizers();
            GC.Collect();

            XmlDocument doc = new XmlDocument();
            doc.Load(mXMLLocation);
            return doc;
        }

        public void Set2PrintAll()
        {
            PrntMain = true;
            PrntPitagoras = true;
            PrntLifeCycles = true;
            PrntIntansiveMap = true;
            PrntCombinedMap = true;
            PrntChakraOpening = true;
            PrntBsnnsPersonal = true;
            PrntBsnnsMulti = true;
            PrntCoupleMatch = true;
            PrntLearnSccss = true;
        }

        public bool TextParser(string TextFromNode, out string text2WordDoc)
        {
            text2WordDoc = "";
            TextFromNode = TextFromNode.Replace("\r".ToCharArray()[0], " ".ToCharArray()[0]);
            TextFromNode = TextFromNode.Trim();
            bool res = true;
            try
            {
                string[] splt = TextFromNode.Split("\n".ToCharArray()[0]);
                for (long i = 0; i < splt.Length; i++)
                {
                    string[] splt2 = splt[i].Trim().Split(" ".ToCharArray()[0]);

                    for (int j = 0; j < splt2.Length; j++)
                    {
                        if (isWordReserved(splt2[j]))
                        {
                            if (j > 0)
                                text2WordDoc += " " + Reserved2TextInfo(splt2[j]);
                            else
                                text2WordDoc += Reserved2TextInfo(splt2[j]);
                        }
                        else
                        {
                            if (isChakraReserved(splt2[j]))
                            {
                                if (j > 0)
                                    text2WordDoc += " " + Reserved2ChakraInfo(splt2[j]);
                                else
                                    text2WordDoc += Reserved2ChakraInfo(splt2[j]);
                            }
                            else
                            {
                                if (isBalance(splt2[j]))
                                {
                                    if (j > 0)
                                        text2WordDoc += " " + Reserved2BalancedInfo(splt2[j]);
                                    else
                                        text2WordDoc += Reserved2BalancedInfo(splt2[j]);
                                }
                                else
                                {
                                    if (isReservedCrysis(splt2[j]))
                                    {
                                        if (j > 0)
                                            text2WordDoc += " " + Reserved2CrysisInfo(splt2[j]);
                                        else
                                            text2WordDoc += Reserved2CrysisInfo(splt2[j]);
                                    }
                                    else
                                    {
                                        if (isReservedWork(splt2[j]))
                                        {
                                            if (j > 0)
                                                text2WordDoc += " " + Reserved2WorkInfo(splt2[j]);
                                            else
                                                text2WordDoc += Reserved2WorkInfo(splt2[j]);
                                        }
                                        else
                                        {
                                            if (isOPSReserved(splt2[j]))
                                            {
                                                if (j > 0)
                                                    text2WordDoc += " " + Reserved2OPSInfo(splt2[j]);
                                                else
                                                    text2WordDoc += Reserved2OPSInfo(splt2[j]);
                                            }
                                            else
                                            {
                                                if (j > 0)
                                                    text2WordDoc += " " + splt2[j];
                                                else
                                                    text2WordDoc += splt2[j];
                                            }
                                        }
                                    }
                                }

                            }
                        }
                    }
                    text2WordDoc += Environment.NewLine;
                }
            }
            catch //(Exception ex)
            {
                res = false;
                text2WordDoc = "";
            }

            text2WordDoc = text2WordDoc.Trim();
            return res;
        }

        public bool isWordReserved(string word)
        {
            foreach (string str in mResevredNames)
            {
                if (str == word)
                {
                    return true;
                }
            }
            return false;
        }

        public bool isChakraReserved(string word)
        {
            foreach (string str in mReservedChakra)
            {
                if (str == word)
                {
                    return true;
                }
            }
            return false;
        }

        public bool isImageReserved(string word)
        {
            foreach (string str in mReservedImages)
            {
                if (str == word)
                {
                    return true;
                }
            }
            return false;
        }

        public bool isCycleReserved(string word)
        {
            foreach (string str in mReservedCycles)
            {
                if (str == word)
                {
                    return true;
                }
            }
            return false;
        }

        public bool isOPSReserved(string word)
        {
            foreach (string str in mReservedOPS)
            {
                if (str == word)
                {
                    return true;
                }
            }
            return false;
        }

        public bool isReservedWork(string word)
        {
            foreach (string str in mReservedWork)
            {
                if (str == word)
                {
                    return true;
                }
            }
            return false;
        }

        public bool isReservedCrysis(string word)
        {
            foreach (string str in mReservedCrysis)
            {
                if (str == word)
                {
                    return true;
                }
            }
            return false;
        }

        public bool isBalance(string word)
        {
            foreach (string str in mBalanced)
            {
                if (str == word)
                {
                    return true;
                }
            }
            return false;
        }

        public string Reserved2TextInfo(string word)
        {
            try
            {
                string sCntrlName = EnumProvider.Instance.GetReservedNameEnumFromDescription(word.Trim()).ToString();

                return mMainForm.Controls.Find(sCntrlName, true)[0].Text.Trim();
            }
            catch
            {
                return "";
            }
        }

        public string Reserved2ChakraInfo(string word)
        {
            string outStr = "";
            try
            {
                string path2file = ConstructFilePath2ChkraInfo(word);
                outStr = ReadFromTextFile(path2file);
                /*
                FileInfo file = new FileInfo(path2file );

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
                else
                {
                    outStr = "";
                }
                */
            }
            catch
            {
                return "";
            }

            return outStr;
        }

        private int CalcCurrentCycle()
        {
            int age = Convert.ToInt16(mMainForm.Controls.Find("txtAge", true)[0].Text.Split("(".ToCharArray()[0])[0]);

            int curCycle = 0;
            int[] CycleBoundries = new int[4];
            CycleBoundries[0] = Convert.ToInt16(mMainForm.Controls.Find("txt1_1", true)[0].Text.Split("-".ToCharArray()[0])[1].Trim());
            CycleBoundries[1] = Convert.ToInt16(mMainForm.Controls.Find("txt2_1", true)[0].Text.Split("-".ToCharArray()[0])[1].Trim());
            CycleBoundries[2] = Convert.ToInt16(mMainForm.Controls.Find("txt3_1", true)[0].Text.Split("-".ToCharArray()[0])[1].Trim());
            CycleBoundries[3] = Convert.ToInt16(mMainForm.Controls.Find("txt4_1", true)[0].Text.Split("-".ToCharArray()[0])[1].Trim());

            if (age <= CycleBoundries[0])
            {
                curCycle = 1;
            }
            else if (age <= CycleBoundries[1])
            {
                curCycle = 2;
            }
            else if (age <= CycleBoundries[2])
            {
                curCycle = 3;
            }
            else if (age <= CycleBoundries[3])
            {
                curCycle = 4;
            }

            return curCycle;
        }

        public string Reserved2CycleInfo(string word)
        {
            string outStr = "";
            try
            {
                string path2file = "";// ConstructFilePath2CycleInfo(true, CalcCurrentCycle());
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
                else
                {
                    outStr = "";
                }

            }
            catch
            {
                return "";
            }

            return outStr;
        }

        public string Reserved2OPSInfo(string word)
        {
            string path2file = "";// ConstructFilePath2();
            List<string> path2fileRec = new List<string>();

            if (word == EnumProvider.ReservedOPS._הפוך_מומחה_.ToString())
            {
                Pro = !Pro;
                return "";
            }
            if (word == EnumProvider.ReservedOPS._הפוך_לארוך_.ToString())
            {
                Pro = false;
                return "";
            }
            if (word == EnumProvider.ReservedOPS._הפוך_למקוצר_.ToString())
            {
                Pro = true;
                return "";
            }

            if (word == EnumProvider.ReservedOPS._מפה_חזקה_חלשה_.ToString())
            {
                bool MapStrong;
                string txtbool = mMainForm.Controls.Find(EnumProvider.Instance.GetReservedOPSEnumFromDescription(EnumProvider.ReservedOPS._מפה_חזקה_חלשה_.ToString()), true)[0].Text.Trim();
                bool res = bool.TryParse(txtbool, out MapStrong);

                if (res == false)
                {
                    MessageBox.Show("Map Strength issue....");
                }

                return ReadFromTextFile(ConstructFilePath2MapStrength(MapStrong));
            }

            if (word == EnumProvider.ReservedOPS._סיכום_דוח_והמלצות_.ToString())
            {
                return ReadFromTextFile(ConstructFilePath2FianlySummary());
            }

            #region Health

            if (word == EnumProvider.ReservedOPS._בריאות_.ToString())
            {
                string sCntrlName = EnumProvider.Instance.GetReservedOPSEnumFromDescription(EnumProvider.ReservedOPS._בריאות_.ToString()); ;
                return mMainForm.Controls.Find(sCntrlName, true)[0].Text.Trim();
            }
            if (word == EnumProvider.ReservedOPS._בריאות_סיכום_.ToString())
            {
                string sCntrlName = EnumProvider.Instance.GetReservedOPSEnumFromDescription(EnumProvider.ReservedOPS._בריאות_סיכום_.ToString());
                return mMainForm.Controls.Find(sCntrlName, true)[0].Text.Trim();
            }
            if (word == EnumProvider.ReservedOPS._מחלות_.ToString())
            {
                #region Descises
                string path2file1 = ConstructFilePath2HealthDcs(1);
                string path2file2 = ConstructFilePath2HealthDcs(2);
                string path2file3 = ConstructFilePath2HealthDcs(3);
                string path2file4 = ConstructFilePath2HealthDcs(4);
                string path2file5 = ConstructFilePath2HealthDcs(5);
                string path2file6 = ConstructFilePath2HealthDcs(6);
                string path2file7 = ConstructFilePath2HealthDcs(7);

                string outStrDCS = "";
                outStrDCS = ReadFromTextFile(path2file1);
                outStrDCS += "," + ReadFromTextFile(path2file2);
                outStrDCS += "," + ReadFromTextFile(path2file3);
                outStrDCS += "," + ReadFromTextFile(path2file4);
                outStrDCS += "," + ReadFromTextFile(path2file5);
                outStrDCS += "," + ReadFromTextFile(path2file6);
                outStrDCS += "," + ReadFromTextFile(path2file7);


                List<string> DCStypes = new List<string>();
                string[] types = outStrDCS.Split(",".ToCharArray());

                for (int i = 0; i < types.Length; i++)
                {
                    if (i == 0)
                    {
                        DCStypes.Add(types[0].Trim());
                    }
                    else
                    {
                        if (DCStypes.Contains(types[i].Trim()) == false)
                        {
                            DCStypes.Add(types[i].Trim());
                        }
                    }
                }

                outStrDCS = "";
                for (int i = 0; i < DCStypes.Count; i++)
                {
                    outStrDCS += DCStypes[i].ToString().Trim() + ",";
                }

                if (outStrDCS[0] == ",".ToCharArray()[0])
                {
                    outStrDCS = outStrDCS.Substring(1, outStrDCS.Length - 1);
                }
                if (outStrDCS[outStrDCS.Length - 1] == ",".ToCharArray()[0])
                {
                    outStrDCS = outStrDCS.Substring(0, outStrDCS.Length - 1);
                }
                outStrDCS = outStrDCS.Replace(",", ", ");
                outStrDCS = outStrDCS.Replace(", ,", ", ");
                outStrDCS = outStrDCS.Replace("  ", " ");
                outStrDCS += ".";

                outStrDCS = outStrDCS.Trim();

                #region Word Wrap for HTML
                if (mRepStyle2Create == "html")
                {
                    if (mRepStyle2Create == "html")
                    {
                        string[] temp = outStrDCS.Trim().Split(" ".ToCharArray());

                        int n = Convert.ToInt16(Math.Floor(Convert.ToDecimal(temp.Length / mMaxWordsInLine)));

                        outStrDCS = "";
                        for (int i = 1; i <= n; i++)
                        {
                            for (int j = 1; j <= mMaxWordsInLine; j++)
                            {
                                outStrDCS += temp[j + (i - 1) * mMaxWordsInLine] + " ";
                            }// for (int j = 1; j <= MaxWordsInLine; j++)
                            outStrDCS += Environment.NewLine;
                        } //for (int i = 1; i <= n; i++)

                    }
                }
                #endregion Word Wrap for HTML

                return outStrDCS;//.Substring(0, outStrDCS.Length - 2);

                #endregion Descises
            }
            if (word == EnumProvider.ReservedOPS._יכולת_לטפל_.ToString())
            {
                path2file = ConstructFilePath2HealthTLC();
            }
            if (word == EnumProvider.ReservedOPS._יכולת_לשמור_על_הבריאות_.ToString())
            {
                path2file = ConstructFilePath2HealthTakeCare();
            }
            if (word == EnumProvider.ReservedOPS._תכונות_בסיסיות_לגוף_.ToString())
            {
                path2file = ConstructFilePath2HealthBody();
            }
            if (word == EnumProvider.ReservedOPS._בריאות_שם_פרטי_.ToString())
            {
                path2file = ConstructFilePath2HealthPrivateName();
            }
            if (word == EnumProvider.ReservedOPS._בריאות_שיא_.ToString())
            {
                path2file = ConstructFilePath2HealthClimax();
            }
            if (word == EnumProvider.ReservedOPS._בריאות_שנה_אישית_.ToString())
            {
                path2file = ConstructFilePath2HealthPrsnlYear();
            }
            if (word == EnumProvider.ReservedOPS._בריאות_מזל_אסטרולוגי_.ToString())
            {
                path2file = ConstructFilePath2HealthAstro();
            }
            #endregion Health

            if (word == EnumProvider.ReservedOPS._בדיקת_צאקרות_קרמאתי_.ToString())
            {
                string sout = "";
                if (mCalc.isCarmaticNumber(mCalc.GetCorrectNumberFromSplitedString((mMainForm.Controls.Find("txtNum7", true)[0] as TextBox).Text.Split(mCalc.Delimiter))))
                {
                    sout = ReadFromTextFile(ConstructFilePath2ChkraInfo("_צאקרה_7_"));
                }

                int val = mCalc.GetCorrectNumberFromSplitedString((mMainForm.Controls.Find("txtNum2", true)[0] as TextBox).Text.Split(mCalc.Delimiter));
                if (mCalc.isCarmaticNumber(val) || (val == 7))
                {
                    sout += sout.Trim() + Environment.NewLine + ReadFromTextFile(ConstructFilePath2ChkraInfo("_צאקרה_2_"));
                }

                sout = sout.Trim();
                return sout;
            }

            if (word == EnumProvider.ReservedOPS._מידע_מחזור_נוכחי_שיא_בלבד_.ToString())
            {
                int curCycle = CalcCurrentCycle();
                string sCntrlName = "txt" + curCycle.ToString() + "_1";
                string Story = "בין הגילאים" + " " + mMainForm.Controls.Find(sCntrlName, true)[0].Text.Trim() + ":" + Environment.NewLine;


                sCntrlName = "txt" + curCycle.ToString() + "_3";
                string ClmxNum = mCalc.GetCorrectNumberFromSplitedString(mMainForm.Controls.Find(sCntrlName, true)[0].Text.Trim().Split(mCalc.Delimiter)).ToString();
                //Story += ReadTextFromFile(ConstructFilePath2ClimaxInfo(Convert.ToInt16(ClmxNum)));
                Story += ReadFromTextFile(ConstructFilePath2ClimaxInfo(Convert.ToInt16(ClmxNum)));

                return Story;
            }

            if (word == EnumProvider.ReservedOPS._מידע_מחזור_נוכחי_.ToString())
            {
                return CollectFullInfo4Cycle(CalcCurrentCycle());
            }

            if (word == EnumProvider.ReservedOPS._מידע_שנים_אישיות_עתיד_2_.ToString())
            {
                #region 2 years
                string sCntrlName;

                sCntrlName = EnumProvider.Instance.GetReservedNameEnumFromDescription(EnumProvider.ReservedNames._תאריך_לידה_.ToString());
                DateTime BDate = (mMainForm.Controls.Find(sCntrlName, true)[0] as DateTimePicker).Value;

                sCntrlName = "DateTimePickerTo";// EnumProvider.Instance.GetReservedOPSEnumFromDescription(EnumProvider.ReservedNames._תאריך_לידה_.ToString());
                DateTime toDate = (mMainForm.Controls.Find(sCntrlName, true)[0] as DateTimePicker).Value;

                string outStrPY = "";

                int Times = 2;
                for (int i = 0; i < Times; i++)
                {
                    toDate = toDate.AddYears(1);

                    string pY, pM, pD;
                    mCalc.CalcPersonalInfo(BDate, toDate, out pY, out pM, out pD);

                    path2file = ConstructFilePath2PersonalYearInfo(mCalc.GetCorrectNumberFromSplitedString(pY.Split(mCalc.Delimiter)));

                    if (i == 0)
                    {
                        outStrPY += "בשנה שמתחילה מיום הולדתך הבא תהיה ב";
                    }
                    //outStrPY += "שנה אישית" + " " + pY + " :" + System.Environment.NewLine + ReadTextFromFile(path2file) + System.Environment.NewLine;
                    outStrPY += "שנה אישית" + " " + pY + " :" + System.Environment.NewLine + ReadFromTextFile(path2file) + System.Environment.NewLine;
                }

                return outStrPY;
                #endregion
            }

            if (word == EnumProvider.ReservedOPS._היום_.ToString())
            {
                return (DateTime.Now.Day.ToString() + "." + DateTime.Now.Month.ToString() + "." + DateTime.Now.Year.ToString());
            }

            if (word == EnumProvider.ReservedOPS._מידע_שנים_אישיות_עתיד_.ToString())
            {
                #region Next Years
                string sCntrlName = EnumProvider.Instance.GetReservedOPSEnumFromDescription(EnumProvider.ReservedOPS._מספר_שנה_אישית_.ToString());
                string[] splt = mMainForm.Controls.Find(sCntrlName, true)[0].Text.Trim().Split(mCalc.Delimiter);

                int[] nums = new int[splt.Length];
                for (int i = 0; i < splt.Length; i++)
                {
                    nums[i] = Convert.ToInt16(splt[i]);
                }

                int currentPY = mCalc.MinArr(nums);

                sCntrlName = EnumProvider.Instance.GetReservedNameEnumFromDescription(EnumProvider.ReservedNames._תאריך_לידה_.ToString());
                DateTime BDate = (mMainForm.Controls.Find(sCntrlName, true)[0] as DateTimePicker).Value;

                sCntrlName = "DateTimePickerTo";// EnumProvider.Instance.GetReservedOPSEnumFromDescription(EnumProvider.ReservedNames._תאריך_לידה_.ToString());
                DateTime toDate = (mMainForm.Controls.Find(sCntrlName, true)[0] as DateTimePicker).Value;

                string outStrPY = "";

                int Times = 9 - currentPY + 1;
                for (int i = 0; i < Times; i++)
                {
                    toDate = toDate.AddYears(1);

                    string pY, pM, pD;
                    mCalc.CalcPersonalInfo(BDate, toDate, out pY, out pM, out pD);

                    path2file = ConstructFilePath2PersonalYearInfo(mCalc.GetCorrectNumberFromSplitedString(pY.Split(mCalc.Delimiter)));

                    if (i == 0)
                    {
                        outStrPY += "בשנה שמתחילה מיום הולדתך הבא תהיה ב";
                    }
                    //outStrPY += "שנה אישית" + " " + pY + " :" + System.Environment.NewLine + ReadTextFromFile(path2file) + System.Environment.NewLine + System.Environment.NewLine;
                    outStrPY += "שנה אישית" + " " + pY + " :" + System.Environment.NewLine + ReadFromTextFile(path2file) + System.Environment.NewLine + System.Environment.NewLine;
                }

                return outStrPY;
                #endregion Next Years
            }

            if (word == EnumProvider.ReservedOPS._אישיות_.ToString())
            {
                #region Personality
                int luck1, luck2, luck3;

                string sCntrlName = "txtNum5";//EnumProvider.Instance.GetReservedChakraEnumFromDescription(word.Trim()).ToString();
                luck1 = mCalc.GetCorrectNumberFromSplitedString(mMainForm.Controls.Find(sCntrlName, true)[0].Text.Trim().Split(mCalc.Delimiter));

                sCntrlName = "txtPName_Num";//EnumProvider.Instance.GetReservedChakraEnumFromDescription(word.Trim()).ToString();
                luck2 = mCalc.GetCorrectNumberFromSplitedString(mMainForm.Controls.Find(sCntrlName, true)[0].Text.Trim().Split(mCalc.Delimiter));

                string sVal = mMainForm.Controls.Find("txtAstroName", true)[0].Text.Trim().Split(" ".ToCharArray()[0])[1];
                sVal = sVal.Replace(")", " ");
                sVal = sVal.Replace("(", " ");
                luck3 = Convert.ToInt16(sVal.Trim());


                string sTest1 = mCalc.TestHramonicInfo(luck1, luck2);
                string sTest2 = mCalc.TestHramonicInfo(luck1, luck3);
                string sTest3 = mCalc.TestHramonicInfo(luck3, luck2);

                int IzunValue = -1;
                if ((sTest1 == "לא מאוזן") || (sTest2 == "לא מאוזן") || (sTest3 == "לא מאוזן"))
                {
                    IzunValue = 0;
                }
                else
                {
                    if ((sTest1 == "חצי מאוזן") || (sTest2 == "חצי מאוזן") || (sTest3 == "חצי מאוזן"))
                    {
                        IzunValue = 1;
                    }
                    else
                    {
                        IzunValue = 2;
                    }
                }
                path2file = ConstructFilePath2MainPersonality(IzunValue);

                string outPRSN = "";

                #region print

                string AstroData = mCalc.GetAstroLuckNameByNumber(luck1);
                outPRSN = "א. מבנה אישיותך המרכזי מאופיין ע\"י " + AstroData + " " + luck1.ToString() + " " + "המהווה 50% ממבנה אישיותך." + System.Environment.NewLine;

                AstroData = mCalc.GetAstroLuckNameByNumber(luck2);
                outPRSN += "ב. מבנה אישיותך המשני מאופיין ע\"י " + AstroData + " " + luck2.ToString() + " " + "המהווה 25% ממבנה אישיותך." + System.Environment.NewLine;

                AstroData = mCalc.GetAstroLuckNameByNumber(luck3);
                outPRSN += "ג. מבנה אישיותך המשני הנוסף מאופיין ע\"י " + AstroData + " " + luck3.ToString() + " " + "המהווה 25% ממבנה אישיותך." + System.Environment.NewLine;

                // מסקנה
                //outPRSN += ReadTextFromFile(path2file);
                outPRSN += ReadFromTextFile(path2file);

                #endregion print

                return outPRSN;
                #endregion Personality
            }

            if (word == EnumProvider.ReservedOPS._ערכי_פתיחת_צאקרות_.ToString())
            {
                #region Open Chakra
                string outStrCO = "";
                for (int i = 1; i < 8; i++)
                {
                    path2file = ConstructFilePath2ChakraOpening(i);

                    string ckrNm = "";
                    switch (i)
                    {
                        case 1: ckrNm = "כתר"; break;
                        case 2: ckrNm = "עין שלישית"; break;
                        case 3: ckrNm = "גרון"; break;
                        case 4: ckrNm = "לב"; break;
                        case 5: ckrNm = "מלקעת השמש"; break;
                        case 6: ckrNm = "מין ויצירה"; break;
                        case 7: ckrNm = "בסיס"; break;
                        case 8: ckrNm = "מזל אסטרולוגי"; break;
                        case 9: ckrNm = "שם פרטי"; break;
                    }

                    outStrCO += "צ'אקרת ה" + ckrNm + ":" + System.Environment.NewLine + ReadFromTextFile(path2file) + System.Environment.NewLine;
                }

                return outStrCO;
                #endregion Open Chakra
            }

            if (word == EnumProvider.ReservedOPS._ערך_מזל_אסטרולוגי_.ToString())
            {
                path2file = ConstructFilePath2AstroLuckInfo();
            }

            if (word == EnumProvider.ReservedOPS._ערך_שם_פרטי_ייעוד_.ToString())
            {
                //path2file = ConstructFilePath2PrivateNameInfo((int)EnumProvider.ReservedOPS._ערך_שם_פרטי_ייעוד_);
            }
            if (word == EnumProvider.ReservedOPS._ערך_שם_פרטי_תכונות_אופי_.ToString())
            {
                //path2file = ConstructFilePath2PrivateNameInfo((int)EnumProvider.ReservedOPS._ערך_שם_פרטי_תכונות_אופי_);
            }
            if (word == EnumProvider.ReservedOPS._ערך_שם_פרטי_.ToString())
            {
                path2file = ConstructFilePath2PrivateNameInfo();
            }

            if (word == EnumProvider.ReservedOPS._ערך_מספר_הכוח_.ToString())
            {
                path2file = ConstructFilePath2PowerNumberInfo();
            }
            if (word == EnumProvider.ReservedOPS._ערך_שנה_אישית_.ToString())
            {
                path2file = ConstructFilePath2PersonalYearInfo();
            }
            if (word == EnumProvider.ReservedOPS._ערך_חודש_אישי_.ToString())
            {
                path2file = ConstructFilePath2PersonalMonthInfo();
            }
            if (word == EnumProvider.ReservedOPS._ערך_יום_אישי_.ToString())
            {
                path2file = ConstructFilePath2PersonalDayInfo();
            }

            if (word == EnumProvider.ReservedOPS._ערך_ייעוד_.ToString())
            {
                path2file = ConstructFilePath2DesignationInfo();
            }

            if (word == EnumProvider.ReservedOPS._ערך_חשש_ופחד_.ToString())
            {
                string sout = "";
                #region Fears
                bool res = false;
                List<string> CntrlsNamess = new List<string>() { "txtNum1", "txtNum2", "txtNum3", "txtNum5", "txtNum6", "txtNum7", "txtPName_Num" };
                List<int> nums = new List<int>();

                foreach (string sCntrlName in CntrlsNamess)
                {
                    int Val = mCalc.GetCorrectNumberFromSplitedString(mMainForm.Controls.Find(sCntrlName, true)[0].Text.Trim().Split(mCalc.Delimiter));
                    if (nums.Contains(Val) == false)
                    {
                        nums.Add(Val);
                    }
                }
                #region Addition - Astro
                string sVal = mMainForm.Controls.Find("txtAstroName", true)[0].Text.Trim().Split(" ".ToCharArray()[0])[1];
                sVal = sVal.Replace(")", " ");
                sVal = sVal.Replace("(", " ");
                int iVal = Convert.ToInt16(sVal.Trim());

                if (nums.Contains(iVal) == false)
                {
                    nums.Add(iVal);
                }
                #endregion Primary Addition - Astro

                foreach (int n in nums)
                {
                    res |= (n == 2) || (n == 11);
                }

                if (res == true)
                {
                    if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                    {
                        sout = "יש לך נטייה לחשש ופחד מאירועים לא מוכרים";
                    }
                    if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                    {
                        sout = "You have a tendency to fear the unkown";
                    }

                    iVal = mCalc.GetCorrectNumberFromSplitedString(mMainForm.Controls.Find("txtNum3", true)[0].Text.Split(mCalc.Delimiter));
                    if (iVal == 8)
                    {
                        if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                        {
                            sout += " " + "עם זאת יכולות החשיבה שלך מאפשרים לך להתמודד בקלות עם נטייה זו.";
                        }
                        if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                        {
                            sout += " " + "yet your thinking abilities helps you to cope with that.";
                        }
                    }
                    else
                    {
                        sout += ".";
                    }
                }
                else
                {
                    sout = "";
                }
                return sout;
                #endregion
                //path2file = ConstructFilePath2DesignationInfo();
            }

            if (word == EnumProvider.ReservedOPS._ערך_חרדות_.ToString())
            {
                string sout = "";
                #region Anxity
                bool res = false;
                List<string> CntrlsNamess = new List<string>() { "txtNum1", "txtNum2", "txtNum3", "txtNum5", "txtNum7", "txtPName_Num" };
                #region current cycle
                int c = CalcCurrentCycle();
                CntrlsNamess.Add("txt" + c.ToString() + "_3");
                CntrlsNamess.Add("txt" + c.ToString() + "_2");
                #endregion current cycle
                List<int> nums = new List<int>();

                foreach (string sCntrlName in CntrlsNamess)
                {
                    int Val = mCalc.GetCorrectNumberFromSplitedString(mMainForm.Controls.Find(sCntrlName, true)[0].Text.Trim().Split(mCalc.Delimiter));
                    if (nums.Contains(Val) == false)
                    {
                        nums.Add(Val);
                    }
                }
                #region Addition - Astro
                string sVal = mMainForm.Controls.Find("txtAstroName", true)[0].Text.Trim().Split(" ".ToCharArray()[0])[1];
                sVal = sVal.Replace(")", " ");
                sVal = sVal.Replace("(", " ");
                int iVal = Convert.ToInt16(sVal.Trim());

                if (nums.Contains(iVal) == false)
                {
                    nums.Add(iVal);
                }
                #endregion Primary Addition - Astro

                foreach (int n in nums)
                {
                    res |= (n == 7) || (n == 16);
                }

                if (res == true)
                {
                    if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                    {
                        sout = "יש לך נטייה לחרדות ודיכאון";
                    }
                    if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                    {
                        sout = "You have a tendency for anxity";
                    }

                    iVal = mCalc.GetCorrectNumberFromSplitedString(mMainForm.Controls.Find("txtNum3", true)[0].Text.Split(mCalc.Delimiter));
                    if (iVal == 8)
                    {
                        if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                        {
                            sout += " " + ",עם זאת יכולות החשיבה שלך מאפשרים לך להתמודד בקלות עם נטייה זו.";
                        }
                        if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                        {
                            sout += " " + ",yet your thinking abilities helps you to cope with that.";
                        }
                    }
                    else
                    {
                        sout += ".";
                    }
                }
                else
                {
                    return "";
                }
                return sout;
                #endregion
                //path2file = ConstructFilePath2DesignationInfo();
            }

            if (word == EnumProvider.ReservedOPS._ערך_מפה_דיכאוני_.ToString())
            {
                string sout = "";
                #region Depressed MAP
                bool res2 = false, res11 = false, res7 = false, res16 = false;
                List<string> CntrlsNamess = new List<string>() { "txtNum1", "txtNum2", "txtNum3", "txtNum5", "txtNum6", "txtNum7", "txtPName_Num" };
                List<int> nums = new List<int>();

                foreach (string sCntrlName in CntrlsNamess)
                {
                    int Val = mCalc.GetCorrectNumberFromSplitedString(mMainForm.Controls.Find(sCntrlName, true)[0].Text.Trim().Split(mCalc.Delimiter));
                    if (nums.Contains(Val) == false)
                    {
                        nums.Add(Val);
                    }
                }
                #region Addition - Astro
                string sVal = mMainForm.Controls.Find("txtAstroName", true)[0].Text.Trim().Split(" ".ToCharArray()[0])[1];
                sVal = sVal.Replace(")", " ");
                sVal = sVal.Replace("(", " ");
                int iVal = Convert.ToInt16(sVal.Trim());

                if (nums.Contains(iVal) == false)
                {
                    nums.Add(iVal);
                }
                #endregion Primary Addition - Astro

                foreach (int n in nums)
                {
                    if (n == 2)
                    {
                        res2 = true;
                    }
                    if (n == 11)
                    {
                        res11 = true;
                    }
                    if (n == 7)
                    {
                        res7 = true;
                    }
                    if (n == 16)
                    {
                        res16 = true;
                    }

                }

                if ((res2 || res11) && (res7 || res16))
                {
                    //return "יש לך נטייה למפה דיכאונית.";
                    if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                    {
                        sout = "יש לך נטייה למפה דיכאונית";
                    }
                    if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                    {
                        sout = "You have a tendency for a depressed map";
                    }

                    iVal = mCalc.GetCorrectNumberFromSplitedString(mMainForm.Controls.Find("txtNum3", true)[0].Text.Split(mCalc.Delimiter));
                    if (iVal == 8)
                    {
                        if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                        {
                            sout += " " + ",עם זאת יכולות החשיבה שלך מאפשרים לך להתמודד בקלות עם נטייה זו.";
                        }
                        if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                        {
                            sout += " " + ",yet your thinking abilities helps you to cope with that.";
                        }
                    }
                    else
                    {
                        sout += ".";
                    }
                }
                else
                {
                    sout = "";
                }
                return sout;
                #endregion
                //path2file = ConstructFilePath2DesignationInfo();
            }

            if (word == EnumProvider.ReservedOPS._ערך_מפה_מאוזן_.ToString())
            {
                string sout = "";
                #region IAUN MAP
                path2file = ConstructFilePath2IZUNInfo();

                sout = ReadFromTextFile(path2file);

                EnumProvider.Balance blnc = EnumProvider.Balance._לא_מאוזן_;
                switch (mMainForm.Controls.Find("txtMapBalanceValue", true)[0].Text)
                {
                    case "_מאוזן_":
                        blnc = EnumProvider.Balance._מאוזן_;
                        break;
                    case "_מאוזן_חלקית_":
                        blnc = EnumProvider.Balance._מאוזן_חלקית_;
                        break;
                    case "_לא_מאוזן_":
                        blnc = EnumProvider.Balance._לא_מאוזן_;
                        break;

                }

                if ((mCalc.GetCorrectNumberFromSplitedString(mMainForm.Controls.Find("txtNum3", true)[0].Text.Split(mCalc.Delimiter)) == 8) && (blnc == EnumProvider.Balance._לא_מאוזן_))
                {
                    sout += Environment.NewLine;
                    if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                    {
                        sout += "אם זאת יש לך את היכולת להתגבר על הקשיים ולהצליח בגדול.";
                    }
                    if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                    {
                        sout += "Though, you have the ability to overcome your issues, and achieve a great presonal success.";
                    }
                }

                #endregion
            }

            //_אנרגיות_גרון_ 
            if (word == EnumProvider.ReservedOPS._אנרגיות_גרון_ומיןויצירה_.ToString())
            {
                string str = mMainForm.Controls.Find("txtNum3", true)[0].Text;
                int val = mCalc.GetCorrectNumberFromSplitedString(str.Split(mCalc.Delimiter));

                str = mMainForm.Controls.Find("txtNum6", true)[0].Text;
                int val2 = mCalc.GetCorrectNumberFromSplitedString(str.Split(mCalc.Delimiter));

                str = "";
                if (((val == 8) || (val2 == 8)) && (GetBalanceStatus() == EnumProvider.Balance._לא_מאוזן_))
                {
                    if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                    {
                        str = "היכולות הגלומים באנרגיה זו מאפשרים לך להתגבר על חוסר הזרימה בחיים ולקדם אותך באופן המיטבי עבורך.";
                    }
                    if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                    {
                        str = "The energy conained inside you chakra map enables you to ovecome the lack of energetic flow in your life, and push you in your own right direction.";
                    }
                }
                else
                {
                    str = "";
                }
                return str;
            }

            if (word == EnumProvider.ReservedOPS._חרדות_מולד_.ToString())
            {
                #region חרדות מולד - מספר הכוח
                string str = mMainForm.Controls.Find("txtNum2", true)[0].Text;
                int val = mCalc.GetCorrectNumberFromSplitedString(str.Split(mCalc.Delimiter));

                if ((val == 7) || (val == 16))
                {
                    if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                    {
                        str = "מספר הכוח שלך הינו: ";
                    }
                    if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                    {
                        str = "Your power number is: ";
                    }

                    str += ReadFromTextFile(ConstructFilePath2ChkraInfo(""));
                }
                return str;
                #endregion

            }

            if (word == EnumProvider.ReservedOPS._מתנות_יום_הלידה_.ToString())
            {
                #region מתנות יום הלידה
                string str = "";

                DateTime dt = (mMainForm.Controls.Find("DateTimePickerFrom", true)[0] as DateTimePicker).Value;
                int day = dt.Day;

                List<int> days = new List<int> { 9, 11, 22, 13, 14, 16, 19 };
                if (days.Contains(day) == true)
                {
                    if ((day == 11) || (day == 22) || (day == 9))
                    {
                        if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                        {
                            str = "מתנות צ'אקרת הכתר - יום הלידה: ";
                        }
                        if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                        {
                            str = "Gift of the Crown Chakra - relying on the Day of birth: ";
                        }
                    }
                    else
                    {
                        if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                        {
                            str = "השיעור הקרמאתי המשוייך לצ'אקרת הכתר: ";
                        }
                        if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                        {
                            str = "The carmatic lesson that belongs to the Crown chakra: ";
                        }
                    }
                    str += ReadFromTextFile(ConstructFilePath2ChkraInfo("_צאקרה_1_"));
                }

                return str;
                #endregion
            }

            if (word == EnumProvider.ReservedOPS._מספר_שנה_אישית_.ToString())
            {
                #region Personal Year
                string sCntrlName = EnumProvider.Instance.GetReservedOPSEnumFromDescription(EnumProvider.ReservedOPS._מספר_שנה_אישית_.ToString());
                return mCalc.GetCorrectNumberFromSplitedString(mMainForm.Controls.Find(sCntrlName, true)[0].Text.Trim().Split(mCalc.Delimiter)).ToString();
                #endregion
            }

            if (word == EnumProvider.ReservedOPS._מספר_חודש_אישי_.ToString())
            {
                #region Personal Month
                string sCntrlName = EnumProvider.Instance.GetReservedOPSEnumFromDescription(EnumProvider.ReservedOPS._מספר_חודש_אישי_.ToString());
                return mCalc.GetCorrectNumberFromSplitedString(mMainForm.Controls.Find(sCntrlName, true)[0].Text.Trim().Split(mCalc.Delimiter)).ToString();
                #endregion
            }

            if (word == EnumProvider.ReservedOPS._מספר_יום_אישי_.ToString())
            {
                #region Personal Day
                string sCntrlName = EnumProvider.Instance.GetReservedOPSEnumFromDescription(EnumProvider.ReservedOPS._מספר_יום_אישי_.ToString());
                return mCalc.GetCorrectNumberFromSplitedString(mMainForm.Controls.Find(sCntrlName, true)[0].Text.Trim().Split(mCalc.Delimiter)).ToString();
                #endregion
            }

            if (word == EnumProvider.ReservedOPS._מידע_מחזור_ראשון_.ToString())
            {
                #region 1
                return CollectFullInfo4Cycle(1);
                #endregion
            }

            if (word == EnumProvider.ReservedOPS._מידע_מחזור_שני_.ToString())
            {
                #region 2
                return CollectFullInfo4Cycle(2);
                #endregion
            }

            if (word == EnumProvider.ReservedOPS._מידע_מחזור_שלישי_.ToString())
            {
                #region 3
                return CollectFullInfo4Cycle(3);
                #endregion
            }

            if (word == EnumProvider.ReservedOPS._מידע_מחזור_רביעי_.ToString())
            {
                #region 4
                return CollectFullInfo4Cycle(4);
                #endregion
            }

            if (word == EnumProvider.ReservedOPS._שילובים_.ToString())
            {
                string sCntrlName = EnumProvider.Instance.GetReservedOPSEnumFromDescription(EnumProvider.ReservedOPS._שילובים_.ToString());
                return mMainForm.Controls.Find(sCntrlName, true)[0].Text.Trim();
            }

            #region עסקים
            if (word == EnumProvider.ReservedOPS._ערך_הצלחה_עסקית_.ToString())
            {
                string sCntrlName = EnumProvider.Instance.GetReservedOPSEnumFromDescription(EnumProvider.ReservedOPS._ערך_הצלחה_עסקית_.ToString());
                return (Math.Floor((Convert.ToDouble(mMainForm.Controls.Find(sCntrlName, true)[0].Text.Trim()) * 10)).ToString() + @"%").ToString();
            }

            if (word == EnumProvider.ReservedOPS._מלל_הצלחה_עסקית_.ToString())
            {
                //string outStrBsns = "";
                //string sCntrlName = EnumProvider.Instance.GetReservedOPSEnumFromDescription(EnumProvider.ReservedOPS._ערך_הצלחה_עסקית_.ToString());

                //string str = mMainForm.Controls.Find(sCntrlName, true)[0].Text.Trim();

                //if (str != "")
                //{
                //    double num = Convert.ToDouble(str);
                //    if (num < 8.0)
                //    {
                //        outStrBsns += "לכן אתה מתקשה להצליח בעסקים." + Environment.NewLine + "אם ברצונך לשפר את סיכויי הצלחתך בתחום העסקי, פנה אל יעקבי לקבלת הסבר.";
                //    }
                //}
                //else
                //{
                //    outStrBsns += "אם ברצונך לשפר את סיכויי הצלחתך בתחום העסקי, פנה אל יעקבי לקבלת הסבר.";
                //}

                //return outStrBsns;

                path2file = ConstructFilePath2BussinessSuccess();

                string story = ReadFromTextFile(path2file) + Environment.NewLine + mMainForm.Controls.Find("txtBusinessStory", true)[0].Text.Trim() + Environment.NewLine;

                string sCntrlName = "txtAstroName";
                string sval = mMainForm.Controls.Find(sCntrlName, true)[0].Text.Trim();
                int tmpVal = Convert.ToInt16(sval.Split("(".ToCharArray()[0])[1].Split(")".ToCharArray()[0])[0]);

                if (tmpVal == 22)
                {
                    if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                    {
                        story += "היות ומזלך האסטרולוגי עקרב, קל לך להיות עצמאי מצליח או עובד מוערך.";
                    }
                    if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                    {
                        story += "Being a Zodiac Scorpion, helps becoming a successful independent or a highly appriciated employee.";
                    }
                }

                sCntrlName = "txtNum3";
                sval = mMainForm.Controls.Find(sCntrlName, true)[0].Text.Trim();
                if (new string[] { "2", "4", "6", "7", "13", "14", "16", "19", "33" }.Contains(sval))
                {
                    if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                    {
                        story += "עם זאת יכולות החשיבה שלך מונעים ממך למצות את הפוטנציאל הגלום בך להצלחה גבוהה יותר.מומלץ לך להוסיף שם על מנת לחדד ולשפר את יכולתך העסקית דבר שיבוא לידי ביטויו בפן העסקי והכלכלי.";
                    }
                    if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                    {
                        story += "Being a Zodiac Scorpion, helps becoming a successful independent or a highly appriciated employee.";
                    }
                }

                //sval = mMainForm.Controls.Find("txtNum6", true)[0].Text.Trim();
                //tmpVal = Convert.ToInt16(sval.Split("(".ToCharArray()[0])[1].Split(")".ToCharArray()[0])[0]);

                //sval = mMainForm.Controls.Find("txtNum3", true)[0].Text.Trim();
                //int tmpVal2 = Convert.ToInt16(sval.Split("(".ToCharArray()[0])[1].Split(")".ToCharArray()[0])[0]);

                //if ((tmpVal == 8) || (tmpVal2 == 8))
                //{
                //    story += Environment.NewLine;
                //    if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                //    {
                //        story += "יש לך פוטנציאל עובד מצטיין.";
                //    }
                //    if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                //    {
                //        story += "You have the potential of being an excelent employee.";
                //    }
                //}

                return story;
            }

            //if (word == EnumProvider.ReservedOPS._פרטי_שותפים_.ToString())
            //{
            //    string outStrCur = "";

            //}
            if (word == EnumProvider.ReservedOPS._מלל_שותפות_עסקית_.ToString())
            {
                string sCntrlName = EnumProvider.Instance.GetReservedOPSEnumFromDescription(EnumProvider.ReservedOPS._מלל_שותפות_עסקית_.ToString());
                string sVal = mMainForm.Controls.Find(sCntrlName, true)[0].Text.Trim();
                return sVal;
            }
            if (word == EnumProvider.ReservedOPS._ערך_שותפות_עסקית_.ToString())
            {
                string sCntrlName = EnumProvider.Instance.GetReservedOPSEnumFromDescription(EnumProvider.ReservedOPS._ערך_שותפות_עסקית_.ToString());
                string sVal = mMainForm.Controls.Find(sCntrlName, true)[0].Text.Trim();
                return (Math.Floor((Convert.ToDouble(sVal) * 10)).ToString() + @"%").ToString();
            }

            if (word == EnumProvider.ReservedOPS._ערך_התאמה_למכירות_.ToString())
            {
                string sCntrlName = EnumProvider.Instance.GetReservedOPSEnumFromDescription(EnumProvider.ReservedOPS._ערך_התאמה_למכירות_.ToString());
                return mMainForm.Controls.Find(sCntrlName, true)[0].Text.Trim();

            }
            if (word == EnumProvider.ReservedOPS._מלל_התאמה_למכירות_.ToString())
            {
                string sCntrlName = EnumProvider.Instance.GetReservedOPSEnumFromDescription(EnumProvider.ReservedOPS._מלל_התאמה_למכירות_.ToString());
                return mMainForm.Controls.Find(sCntrlName, true)[0].Text.Trim();

            }

            if (word == EnumProvider.ReservedOPS._מלל_התאמה_קבוצתית_למכירות_.ToString())
            {
                string sCntrlName = EnumProvider.Instance.GetReservedOPSEnumFromDescription(EnumProvider.ReservedOPS._מלל_התאמה_קבוצתית_למכירות_.ToString());
                return mMainForm.Controls.Find(sCntrlName, true)[0].Text.Trim();

            }
            #endregion

            #region זוגיות
            if (word == EnumProvider.ReservedOPS._ערך_הצלחה_זוגית_.ToString())
            {
                string sCntrlName = EnumProvider.Instance.GetReservedOPSEnumFromDescription(EnumProvider.ReservedOPS._ערך_הצלחה_זוגית_.ToString());
                string sVal = mMainForm.Controls.Find(sCntrlName, true)[0].Text.Trim();
                return (Math.Floor((Convert.ToDouble(sVal) * 10)).ToString() + @"%").ToString();
            }

            if (word == EnumProvider.ReservedOPS._מלל_הצלחה_זוגית_.ToString())
            {
                string str = mMainForm.Controls.Find("txtMrrgStory", true)[0].Text.Trim();

                path2file = ConstructFilePath2MarriageSuccess();
                str += Environment.NewLine + ReadFromTextFile(path2file);

                return str;
            }

            if (word == EnumProvider.ReservedOPS._פרטי_בן_זוג_.ToString())
            {
                string outStrCur = "";
                DataGridView SpouceDataTable = mMainForm.Controls.Find("SpouceData", true)[0] as DataGridView;

                DataGridViewRow r = SpouceDataTable.Rows[0];

                UserInfo curPartner = new UserInfo();

                curPartner.mFirstName = CellValue2String(r.Cells[0]);
                curPartner.mLastName = CellValue2String(r.Cells[1]);
                //curPartner.mFatherName = CellValue2String(r.Cells[3]);
                //curPartner.mMotherName = CellValue2String(r.Cells[4]);
                bool resDate = DateTime.TryParse(r.Cells[2].Value.ToString().Trim(), out curPartner.mB_Date);
                //string yesno = CellValue2String(r.Cells[5]);
                //curPartner.mCity = CellValue2String(r.Cells[6]);
                //curPartner.mStreet = CellValue2String(r.Cells[7]);

                //curPartner.mBuildingNum = CellValue2Double(r.Cells[8]);
                //curPartner.mAppNum = CellValue2Double(r.Cells[9]);

                if (BLL.AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                {
                    outStrCur = "Spouce Info. :" + Environment.NewLine + "Name: ";

                }
                if (BLL.AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                {
                    outStrCur = "פרטי בן/בת הזוג הינם:" + Environment.NewLine + "שם: ";
                }

                outStrCur += curPartner.mFirstName + " " + curPartner.mLastName + " (" + curPartner.mB_Date.ToLongDateString() + ").";

                return outStrCur;
            }

            if (word == EnumProvider.ReservedOPS._בדיקת_זוגיות_צאקרות_קארמתי_.ToString())
            {
                string sout = "";
                bool found = false;
                UserResult curPresonRes = new UserResult();
                curPresonRes.GatherDataFromGUI(mMainForm.Controls, CalcCurrentCycle());

                #region curretn preson
                List<int> vals = new List<int> {mCalc.GetCorrectNumberFromSplitedString(curPresonRes.Sun.Split(mCalc.Delimiter)),
                                                mCalc.GetCorrectNumberFromSplitedString(curPresonRes.ThirdEye.Split(mCalc.Delimiter)),
                                                mCalc.GetCorrectNumberFromSplitedString(curPresonRes.PrivateNameNum.Split(mCalc.Delimiter)),
                                                mCalc.GetCorrectNumberFromSplitedString(curPresonRes.Base.Split(mCalc.Delimiter)),
                                                mCalc.GetCorrectNumberFromSplitedString(curPresonRes.Crown.Split(mCalc.Delimiter)),
                                                Convert.ToInt16(curPresonRes.Astro.Split(" ".ToCharArray()[0])[1].Replace("(","").Replace(")","")),
                                                mCalc.GetCorrectNumberFromSplitedString(curPresonRes.Throught.Split(mCalc.Delimiter)),
                                                mCalc.GetCorrectNumberFromSplitedString(curPresonRes.LC_Climax.Split(mCalc.Delimiter)),
                                                mCalc.GetCorrectNumberFromSplitedString(curPresonRes.LC_climx1.Split(mCalc.Delimiter)),
                                                mCalc.GetCorrectNumberFromSplitedString(curPresonRes.LC_cycle.Split(mCalc.Delimiter))};

                for (int i = 0; i < vals.Count; i++)
                {
                    if ((mCalc.isCarmaticNumber(vals[i])) || (vals[i] == 7))
                    {
                        found = true;
                    }
                }
                #endregion
                if (found == true)
                {
                    if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                    {
                        sout = "מפתך מתארת קשיים במערכת זוגית.";
                    }
                    if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                    {
                        sout = "Your Chakra map is describing a difficulty in a relashionship.";
                    }
                }

                #region spuuce 
                found = false;
                UserResult spcRes = mMainForm.getSpouceResults();
                vals = new List<int> {mCalc.GetCorrectNumberFromSplitedString(spcRes.Sun.Split(mCalc.Delimiter)),
                                    mCalc.GetCorrectNumberFromSplitedString(spcRes.ThirdEye.Split(mCalc.Delimiter)),
                                    mCalc.GetCorrectNumberFromSplitedString(spcRes.PrivateNameNum.Split(mCalc.Delimiter)),
                                    mCalc.GetCorrectNumberFromSplitedString(spcRes.Base.Split(mCalc.Delimiter)),
                                    mCalc.GetCorrectNumberFromSplitedString(spcRes.Crown.Split(mCalc.Delimiter)),
                                    Convert.ToInt16(spcRes.Astro.Split(" ".ToCharArray()[0])[1].Replace("(","").Replace(")","")),
                                    mCalc.GetCorrectNumberFromSplitedString(spcRes.Throught.Split(mCalc.Delimiter)),
                                    mCalc.GetCorrectNumberFromSplitedString(spcRes.LC_Climax.Split(mCalc.Delimiter)),
                                    mCalc.GetCorrectNumberFromSplitedString(spcRes.LC_climx1.Split(mCalc.Delimiter)),
                                    mCalc.GetCorrectNumberFromSplitedString(spcRes.LC_cycle.Split(mCalc.Delimiter))};

                for (int i = 0; i < vals.Count; i++)
                {
                    if ((mCalc.isCarmaticNumber(vals[i])) || (vals[i] == 7))
                    {
                        found = true;
                    }
                }
                #endregion
                if (found == true)
                {
                    if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                    {
                        sout += Environment.NewLine + "מפת בן/בת הזוג מתארת קשיים במערכת זוגית.";
                    }
                    if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                    {
                        sout += Environment.NewLine + "Your spouse's Chakra map is describing a difficulty in a relashionship.";
                    }
                }

                return sout;

            }
            #endregion

            #region התאמה מינית

            if (word == EnumProvider.ReservedOPS._ערך_התאמה_מינית_.ToString())
            {
                string sCntrlName = EnumProvider.Instance.GetReservedOPSEnumFromDescription(EnumProvider.ReservedOPS._ערך_התאמה_מינית_.ToString());
                string sVal = mMainForm.Controls.Find(sCntrlName, true)[0].Text.Trim();
                return (Math.Floor((Convert.ToDouble(sVal) * 10)).ToString() + @"%").ToString();
            }

            if (word == EnumProvider.ReservedOPS._מלל_התאמה_מינית_.ToString())
            {
                string sCntrlName = EnumProvider.Instance.GetReservedOPSEnumFromDescription(EnumProvider.ReservedOPS._מלל_התאמה_מינית_.ToString());
                string sVal = mMainForm.Controls.Find(sCntrlName, true)[0].Text.Trim();
                sVal += string.Format(" ({0})", mMainForm.Controls.Find("txtSexMatchMark", true)[0].Text.Trim());

                return sVal;
            }

            if (word == EnumProvider.ReservedOPS._פרטי_בן_זוג_מינית_.ToString())
            {
                string outStrCur = "";
                DataGridView SpouceDataTable = mMainForm.Controls.Find("SpouceSexData", true)[0] as DataGridView;

                DataGridViewRow r = SpouceDataTable.Rows[0];

                UserInfo curPartner = new UserInfo();

                curPartner.mFirstName = CellValue2String(r.Cells[0]);
                curPartner.mLastName = CellValue2String(r.Cells[1]);

                bool resDate = DateTime.TryParse(r.Cells[2].Value.ToString().Trim(), out curPartner.mB_Date);

                if (BLL.AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                {
                    outStrCur = "Spouce Info. :" + Environment.NewLine + "Name: ";

                }
                if (BLL.AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                {
                    outStrCur = "פרטי בן/בת הזוג הינם:" + Environment.NewLine + "שם: ";
                }

                outStrCur += curPartner.mFirstName + " " + curPartner.mLastName + " (" + curPartner.mB_Date.ToLongDateString() + ").";

                return outStrCur;
            }

            #endregion התאמה מינית

            #region LIMUDIM
            if (word == EnumProvider.ReservedOPS._ערך_הצלחה_בלימודים_.ToString())
            {
                string sCntrlName = EnumProvider.Instance.GetReservedOPSEnumFromDescription(EnumProvider.ReservedOPS._ערך_הצלחה_בלימודים_.ToString());
                string sVal = mMainForm.Controls.Find(sCntrlName, true)[0].Text.Trim();
                return sVal;// (Math.Floor((Convert.ToDouble(sVal) * 10)).ToString() + @"%").ToString();
            }

            if (word == EnumProvider.ReservedOPS._מלל_הצלחה_בלימודים_.ToString())
            {
                string sCntrlName = EnumProvider.Instance.GetReservedOPSEnumFromDescription(EnumProvider.ReservedOPS._ערך_הצלחה_בלימודים_.ToString());
                string sVal = mMainForm.Controls.Find(sCntrlName, true)[0].Text.Trim();

                path2file = ConstructFilePath2LearnSccss();

                sCntrlName = EnumProvider.Instance.GetReservedOPSEnumFromDescription(EnumProvider.ReservedOPS._מלל_הצלחה_בלימודים_.ToString());

                return (ReadFromTextFile(path2file) + Environment.NewLine + mMainForm.Controls.Find(sCntrlName, true)[0].Text.Trim());

            }
            #endregion

            #region KESHEV VERIKUZ
            if (word == EnumProvider.ReservedOPS._הפרעות_קשב_וריכוז_.ToString())
            {
                string sOutSext = "";

                //DateTime today = (mMainForm.Controls.Find("DateTimePickerTo", true)[0] as DateTimePicker).Value;

                for (int cycle = 1; cycle < 4; cycle++)
                {
                    //DateTimePicker dtpTo = mMainForm.Controls.Find("DateTimePickerTo", true)[0] as DateTimePicker;

                    //string[] sAges = (mMainForm.Controls.Find(("txt" + cycle.ToString() + "_1").ToString(), true)[0] as TextBox).Text.Split("-".ToCharArray());
                    //double meanAge = Convert.ToDouble((Convert.ToInt16(sAges[0]) + Convert.ToInt16(sAges[1])) / 2.0);

                    //dtpTo.Value = (mMainForm.Controls.Find("DateTimePickerFrom", true)[0] as DateTimePicker).Value.AddYears(Convert.ToInt16 (Math.Floor( meanAge)));

                    // run calc...
                    //(mMainForm as Omega.MainForm).LrnSccssAttPrblm();

                    string[] sLiteProbs = mMainForm.Controls.Find("txtAttMinor", true)[0].Text.Split(",".ToCharArray());
                    string[] sStrongProbs = mMainForm.Controls.Find("txtAttMajor", true)[0].Text.Split(",".ToCharArray());

                    int[] liteProbs = new int[sLiteProbs.Length];
                    int[] strongProbs = new int[sStrongProbs.Length];

                    for (int i = 0; i < sLiteProbs.Length; i++)
                    {
                        liteProbs[i] = Convert.ToInt16(sLiteProbs[i]);
                    }

                    for (int i = 0; i < sStrongProbs.Length; i++)
                    {
                        strongProbs[i] = Convert.ToInt16(sStrongProbs[i]);
                    }

                    (mMainForm as Omega.MainForm).BubbleSort(ref liteProbs);
                    (mMainForm as Omega.MainForm).BubbleSort(ref strongProbs);

                    if (cycle == 1)
                    {
                        sOutSext = "עבור המחזור הראשון:" + Environment.NewLine;
                        sOutSext += "הפרעות קשב וריכוז קלות" + Environment.NewLine;

                        for (int i = 0; i < liteProbs.Length; i++)
                        {
                            sOutSext += ConstrcutFilePath2ADHDInfoFiles(liteProbs[i]) + Environment.NewLine;
                        }

                        sOutSext += "הפרעות קשב וריכוז מורכבות" + Environment.NewLine;
                        for (int i = 0; i < strongProbs.Length; i++)
                        {
                            sOutSext += ConstrcutFilePath2ADHDInfoFiles(strongProbs[i]) + Environment.NewLine;
                        }
                    }
                    else // cycle == 2,3,4
                    {
                        string sCycleName = "";
                        if (cycle == 2) { sCycleName = "שני"; }
                        if (cycle == 3) { sCycleName = "שלישי"; }
                        if (cycle == 4) { sCycleName = "רביעי"; }

                        sOutSext = Environment.NewLine + "עבור המחזור ה:" + sCycleName + Environment.NewLine;

                        int v1 = mCalc.GetCorrectNumberFromSplitedString(mMainForm.Controls.Find("txt" + cycle.ToString() + "_2", true)[0].Text.Split(mCalc.Delimiter));
                        int v2 = mCalc.GetCorrectNumberFromSplitedString(mMainForm.Controls.Find("txt" + cycle.ToString() + "_3", true)[0].Text.Split(mCalc.Delimiter));

                        if ((mCalc.isCarmaticNumber(v1) || mCalc.isCarmaticNumber(v2)) == true)
                        {
                            sOutSext += "הפרעות קשב וריכוז מורכבות" + Environment.NewLine;
                            if (mCalc.isCarmaticNumber(v1) == true)
                            {
                                sOutSext += ConstrcutFilePath2ADHDInfoFiles(v1) + Environment.NewLine;
                            }
                            if (mCalc.isCarmaticNumber(v2) == true)
                            {
                                sOutSext += ConstrcutFilePath2ADHDInfoFiles(v2) + Environment.NewLine;
                            }
                        }
                        else
                        {
                            sOutSext += "אין הפרעות קשב וריכוז חדשות" + Environment.NewLine;
                        }

                    }


                    // TBD:   הוספה של פיחות בהפרעות קשב וריכוז


                }

                //// return to original
                //(mMainForm.Controls.Find("DateTimePickerTo", true)[0] as DateTimePicker).Value = today;
                //(mMainForm as Omega.MainForm).LrnSccssAttPrblm();

                return sOutSext;
            }
            #endregion

            #region המלצות

            if (word == EnumProvider.ReservedOPS._מלל_המלצות_אישיות_.ToString())
            {
                string outStr1 = "";

                #region מפה מאוזנת
                string sCntrlName = "txtInfo1";
                string text = mMainForm.Controls.Find(sCntrlName, true)[0].Text.Trim();
                string[] sTexts = text.Split(Environment.NewLine.ToCharArray()[0]);
                string sEnumVal = sTexts[1].Trim();

                switch (sEnumVal)
                {
                    case "מאוזן":
                        outStr1 += "";
                        break;
                    case "חצי מאוזן":
                        outStr1 += "";
                        break;
                    case "לא מאוזן":
                        outStr1 += "מומלץ לך לאזן מפתך, שכן כיום היא אינה מאפשרת זרימה אנרגטית טובה בחיים ומאפשרת רק 40% של התקדמות בחיים." + Environment.NewLine +
                                "אם ברצונך לשפר את סיכויי הצלחתך בחייך האישיים פנה ליעקבי לקבלת הסבר.";
                        break;
                }

                #endregion

                #region זוגיות
                //sCntrlName = EnumProvider.Instance.GetReservedOPSEnumFromDescription(EnumProvider.ReservedOPS._ערך_הצלחה_זוגית_.ToString());
                //double num = Convert.ToDouble(mMainForm.Controls.Find(sCntrlName, true)[0].Text.Trim());
                #endregion

                return outStr1;
            }

            if (word == EnumProvider.ReservedOPS._פרטים_אישיים_.ToString())
            {
                string sCntrlName = "textBox48";
                string text = mMainForm.Controls.Find(sCntrlName, true)[0].Text.Trim();
                string retVal = string.Empty;
                int fromPos1 = text.IndexOf("[חתימה]");
                int fromPos2 = text.IndexOf("[תיאור מקצועי]");

                if (fromPos1 < fromPos2)
                {
                    retVal = text.Substring(0, fromPos1);
                }
                else
                {
                    retVal = text.Substring(0, fromPos2 - 1);

                }

                return retVal;
            }

            if (word == EnumProvider.ReservedOPS._חתימה_.ToString())
            {
                string sCntrlName = "textBox48";
                string text = mMainForm.Controls.Find(sCntrlName, true)[0].Text.Trim();
                string retVal = string.Empty;
                int fromPos1 = text.IndexOf("[חתימה]") + "[חתימה]".Length;
                int fromPos2 = text.LastIndexOf("[חתימה]");

                if (fromPos2 > fromPos1)
                {
                    retVal = text.Substring(fromPos1, fromPos2 - fromPos1);

                }

                return retVal;
            }

            if (word == EnumProvider.ReservedOPS._תיאור_מקצועי_.ToString())
            {
                string sCntrlName = "textBox48";
                string text = mMainForm.Controls.Find(sCntrlName, true)[0].Text.Trim();
                string retVal = string.Empty;
                int fromPos1 = text.IndexOf("[תיאור מקצועי]") + "[תיאור מקצועי]".Length;
                int fromPos2 = text.LastIndexOf("[תיאור מקצועי]");

                if (fromPos2 > fromPos1)
                {
                    retVal = text.Substring(fromPos1, fromPos2 - fromPos1);

                }

                return retVal;
            }
            #endregion

            if (word == EnumProvider.ReservedOPS._ערכי_פיתגורס_.ToString())
            {
                string outStrCur = "";
                foreach (string sPit in mPitPlanes)
                {
                    string sPlane = sPit.Substring(0, 2);

                    string vals = sPit.Substring(5);
                    vals = vals.Replace(")", "");
                    vals = vals.Replace("(", "");
                    string[] sVals = vals.Split(",".ToCharArray()[0]);

                    List<bool> filled = new List<bool>();
                    for (int i = 0; i < sVals.Length; i++)
                    {
                        string sCntrlName = "txtPit" + sVals[i];
                        string val = mMainForm.Controls.Find(sCntrlName, true)[0].Text.Trim();

                        filled.Add(val != "");
                    }

                    bool isTrue = false, isFalse = false;
                    if (filled.Contains(true)) { isTrue = true; }
                    if (filled.Contains(false)) { isFalse = true; }

                    string fVal = "";
                    if ((isTrue == true) && (isFalse == false))
                    {
                        fVal = "02";
                    }
                    if ((isTrue == false) && (isFalse == true))
                    {
                        fVal = "01";
                    }
                    if ((isTrue == true) && (isFalse == true))
                    {
                        fVal = "03";
                    }
                    path2file = ConstructFilePath2PitagorsInfo(sPlane, fVal);
                    outStrCur += "עבור הערך של מישור מספר " + sPlane + " :" + Environment.NewLine + ReadFromTextFile(path2file) + Environment.NewLine;
                }
                return outStrCur;
            }

            if (word == EnumProvider.ReservedOPS._ערך_תיקון_ראשי_.ToString())
            {
                path2file = "null1";
                path2fileRec = ConstructFilePath2RectificationInfo();

            }

            if (word == EnumProvider.ReservedOPS._ערך_תיקון_מקוצר_.ToString())
            {
                bool temp = Pro;

                Pro = true;
                path2file = "null1";
                path2fileRec = ConstructFilePath2RectificationInfo();
                Pro = temp;
            }

            string outStr = "";
            if (path2fileRec.Count > 0) //_ערך_תיקון_
            {
                string outStrPrime = string.Empty; // "התיקון הראשי הוא: " + Environment.NewLine;
                foreach (string sFilePath in path2fileRec)
                {
                    string[] splt = sFilePath.Split("\\".ToCharArray()[0]);
                    if (splt[splt.Length - 1].Substring(0, 1) == "p")
                    {
                        try
                        {
                            string sNewPath = sFilePath.Replace("prime_", "").Trim();
                            outStrPrime += ReadFromTextFile(sNewPath) + Environment.NewLine + Environment.NewLine + Environment.NewLine;
                            /*
                            FileStream fs = new FileStream(sNewPath, FileMode.Open);
                            StreamReader fr = new StreamReader(fs);

                            outStrPrime += fr.ReadToEnd() + Environment.NewLine;

                            fr.Close();
                            fs.Close();

                            GC.Collect();
                            GC.WaitForPendingFinalizers();
                            */
                        }
                        catch //(Exception ex)
                        {

                            outStrPrime += "";
                        }
                    }
                }

                //string outStrSecond = "התיקון המשני הוא: " + Environment.NewLine;
                string outStrSecond = string.Empty;
                foreach (string sFilePath in path2fileRec)
                {
                    string[] splt = sFilePath.Split("\\".ToCharArray()[0]);
                    if (splt[splt.Length - 1].Substring(0, 1) == "s")
                    {
                        try
                        {
                            outStrSecond += ReadFromTextFile(sFilePath.Replace("second_", "").Trim()) + Environment.NewLine + Environment.NewLine + Environment.NewLine;
                            /*
                            FileStream fs = new FileStream(sFilePath.Replace("second_", "").Trim(), FileMode.Open);
                            StreamReader fr = new StreamReader(fs);

                            outStrSecond += fr.ReadToEnd() + Environment.NewLine;

                            fr.Close();
                            fs.Close();

                            GC.Collect();
                            GC.WaitForPendingFinalizers();
                            */
                        }
                        catch //(Exception ex)
                        {
                            outStrSecond += "";
                        }
                    }
                }

                outStr = outStrPrime + outStrSecond;
                string[] o = outStr.Split(Environment.NewLine.ToCharArray()).Distinct().ToArray();
                outStr = string.Join(Environment.NewLine, o);
            }
            else
            {
                try
                {
                    FileInfo file = new FileInfo(path2file);

                    if (file.Exists == true)
                    {
                        outStr = ReadFromTextFile(path2file);
                        /*
                        FileStream fs = new FileStream(path2file, FileMode.Open);
                        StreamReader fr = new StreamReader(fs);

                        outStr = fr.ReadToEnd();

                        fr.Close();
                        fs.Close();

                        GC.Collect();
                        GC.WaitForPendingFinalizers();
                        */
                    }
                    else
                    {
                        outStr = "";
                    }

                }
                catch
                {
                    outStr = "";
                }
            }

            return outStr;

        }

        private string Reserved2WorkInfo(string word)
        {
            string outStr = "";
            try
            {
                string path2file = ConstructFilePath2WorkInfo(word);
                outStr = ReadFromTextFile(path2file);
                /*
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
                else
                {
                    outStr = "";
                }
                */
            }
            catch
            {
                return "";
            }

            return outStr;
        }
        #region for business & marrigae
        private string CellValue2String(DataGridViewCell cell)
        {
            if (cell.Value == null)
            {
                return "";
            }
            else
            {
                return cell.Value.ToString().Trim();
            }
        }

        private double CellValue2Double(DataGridViewCell cell)
        {
            if (cell.Value == null)
            {
                return 0;
            }
            else
            {
                return Convert.ToDouble(cell.Value.ToString().Trim());
            }
        }
        #endregion 

        private string Reserved2CrysisInfo(string word)
        {
            string outStr = "";
            try
            {
                string path2file = ConstructFilePath2CrysisInfo(word);
                outStr = ReadFromTextFile(path2file);
                /*
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
                else
                {
                    outStr = "";
                }
                */

            }
            catch
            {
                return "";
            }

            return outStr;
        }

        private string Reserved2BalancedInfo(string word)
        {
            string outStr = "";
            try
            {
                string path2file = ConstructFilePath2IZUNInfo();
                outStr = ReadFromTextFile(path2file);
                /* 
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
                else
                {
                    outStr = "";
                }
                */

            }
            catch
            {
                return "";
            }

            return outStr;
        }

        public string ReadFromTextFile(string path)
        {
            string outStr = "";
            try
            {
                FileStream fh = new FileStream(path, FileMode.Open);
                StreamReader sr = new StreamReader(fh);

                outStr += sr.ReadToEnd();

                if (mRepStyle2Create == "html")
                {
                    string[] temp = outStr.Trim().Split(" ".ToCharArray());

                    int n = Convert.ToInt16(Math.Floor(Convert.ToDecimal(temp.Length / mMaxWordsInLine)));

                    outStr = "";
                    for (int i = 1; i <= n; i++)
                    {
                        for (int j = 1; j <= mMaxWordsInLine; j++)
                        {
                            outStr += temp[j + (i - 1) * mMaxWordsInLine] + " ";
                        }// for (int j = 1; j <= MaxWordsInLine; j++)
                        outStr += Environment.NewLine;
                    } //for (int i = 1; i <= n; i++)

                }

                outStr = outStr.Trim() + Environment.NewLine;

                sr.Close();
                fh.Close();

                GC.Collect();
                //GC.WaitForPendingFinalizers();
                //GC.WaitForFullGCComplete();
            }
            catch
            {
                outStr = "";
            }

            return outStr;
        }

        public void WriteText2File(string path2file, string info)
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

        #region Data file Path from Program Value
        /*
        public string ConstructFilePath2ChkraInfo(string word)
        {
            string sCntrlName = EnumProvider.Instance.GetReservedChakraEnumFromDescription(word.Trim()).ToString();

            int Val = mCalc.GetCorrectNumberFromSplitedString(mMainForm.Controls.Find(sCntrlName, true)[0].Text.Trim().Split(mCalc.Delimiter));
            string ChakraNum = word.Split(mSeporator.ToCharArray()[0])[2];

            string path2file = System.IO.Path.GetDirectoryName(Application.ExecutablePath) + "\\Templets\\InternalChakraValues\\";
            path2file += "chkra" + ChakraNum + "_" + "num" + Val.ToString() + ".txt";

            return path2file;
        }
        */

        public string ConstructFilePath2MapStrength(bool val)
        {
            return mFolderPath2FinalMapStrengthInfoFiles + "\\" + Prefix + "strength_" + (Convert.ToInt16(val)).ToString() + ".txt";
        }

        public string ConstructFilePath2FianlySummary()
        {
            return mFolderPath2FinalSummaryInfoFiles + "\\" + Prefix + "fin.txt";
        }

        private int LearnVal2Class(double val)
        {
            int cls = -1;

            if (val <= 5.9)
            {
                cls = 1;
            }

            if ((val >= 6.0) && (val <= 6.9))
            {
                cls = 2;
            }

            if ((val >= 7.0) && (val <= 7.9))
            {
                cls = 3;
            }

            if ((val >= 8.0) && (val <= 8.5))
            {
                cls = 4;
            }

            if ((val >= 8.6) && (val <= 8.9))
            {
                cls = 5;
            }

            if (val >= 9.0)
            {
                cls = 6;
            }
            return cls;
        }

        public string ConstrcutFilePath2ADHDInfoFiles(int val)
        {
            return System.IO.Path.Combine(mFolderPath2ADHDInfoFiles, "adhd_" + val.ToString() + ".txt");
        }

        public string ConstructFilePath2LearnSccss(double val)
        {
            return System.IO.Path.Combine(mFolderPath2MainPersonality, Prefix + "learn_" + LearnVal2Class(val).ToString() + ".txt");
        }

        public string ConstructFilePath2LearnSccss()
        {
            string sCntrlName = EnumProvider.Instance.GetReservedOPSEnumFromDescription(EnumProvider.ReservedOPS._ערך_הצלחה_בלימודים_.ToString());
            //double resVal = Math.Floor((Convert.ToDouble(mMainForm.Controls.Find(sCntrlName, true)[0].Text.Trim()) * 10));
            double resVal = Convert.ToDouble(mMainForm.Controls.Find(sCntrlName, true)[0].Text.Trim());

            return System.IO.Path.Combine(mFolderPath2MainPersonality, Prefix + "learn_" + LearnVal2Class(resVal).ToString() + ".txt");
        }

        #region Health
        public string ConstructFilePath2HealthDcs(int type)
        {
            string sCntrlName = "";
            string sVal = "";

            int currentcycle = CalcCurrentCycle();

            int num1 = mCalc.GetCorrectNumberFromSplitedString(mMainForm.Controls.Find("txtNum5", true)[0].Text.Split(mCalc.Delimiter));
            int num2 = mCalc.GetCorrectNumberFromSplitedString(mMainForm.Controls.Find("txtPName_Num", true)[0].Text.Split(mCalc.Delimiter));
            bool bGetAnyWay = false;

            if ((mCalc.isCarmaticNumber(num1) == true) || (mCalc.isCarmaticNumber(num2) == true))
            {
                bGetAnyWay = true;
            }

            if (type == 2) // ASTRO
            {
                sCntrlName = "txtAstroName";
                sVal = mMainForm.Controls.Find(sCntrlName, true)[0].Text.Split(" ".ToCharArray())[1];
                sVal = sVal.Replace("(", "");
                sVal = sVal.Replace(")", "");
                sVal = sVal.Trim();
            }
            if (type == 1) // miklaat shemesh
            {
                sCntrlName = "txtNum5";
                sVal = mCalc.GetCorrectNumberFromSplitedString(mMainForm.Controls.Find(sCntrlName, true)[0].Text.Split(mCalc.Delimiter)).ToString();
            }
            if (type == 3) // Private Name
            {
                sCntrlName = "txtPName_Num";
                sVal = mCalc.GetCorrectNumberFromSplitedString(mMainForm.Controls.Find(sCntrlName, true)[0].Text.Split(mCalc.Delimiter)).ToString();
            }
            if (type == 4)// Current climax
            {
                if ((bGetAnyWay == true) || (currentcycle >= 3))
                {
                    sCntrlName = "txt" + currentcycle + "_3";
                    sVal = mCalc.GetCorrectNumberFromSplitedString(mMainForm.Controls.Find(sCntrlName, true)[0].Text.Split(mCalc.Delimiter)).ToString();
                }
            }
            if (type == 5)
            {
                if ((bGetAnyWay == true) || (currentcycle >= 3)) // Current cycle
                {
                    sCntrlName = "txt" + currentcycle + "_2";
                    sVal = mCalc.GetCorrectNumberFromSplitedString(mMainForm.Controls.Find(sCntrlName, true)[0].Text.Split(mCalc.Delimiter)).ToString();
                }
            }
            if (type == 6)
            {
                if ((bGetAnyWay == true) || (currentcycle >= 3)) // Current personal year
                {
                    sCntrlName = "txtPYear";
                    sVal = mCalc.GetCorrectNumberFromSplitedString(mMainForm.Controls.Find(sCntrlName, true)[0].Text.Split(mCalc.Delimiter)).ToString();
                }
            }
            if (type == 7)
            {
                if ((bGetAnyWay == true) || (currentcycle >= 3)) // Current personal month
                {
                    sCntrlName = "txtPMonth";
                    sVal = mCalc.GetCorrectNumberFromSplitedString(mMainForm.Controls.Find(sCntrlName, true)[0].Text.Split(mCalc.Delimiter)).ToString();
                }
            }
            if (sVal.Length > 0)
            {
                return ConstructFilePath2HealthDcs(type, Convert.ToInt16(sVal));
            }
            else
            {
                return "";
            }
        }
        public string ConstructFilePath2HealthDcs(int type, int val)
        {
            return System.IO.Path.Combine(mFolderPath2HealthInfoFiles, "health_dcs_" + type.ToString() + "_" + val.ToString() + ".txt");
        }

        public string ConstructFilePath2HealthTakeCare()
        {
            string sCntrlName = EnumProvider.Instance.GetReservedOPSEnumFromDescription(EnumProvider.ReservedOPS._יכולת_לשמור_על_הבריאות_.ToString());

            int v = mCalc.GetCorrectNumberFromSplitedString(mMainForm.Controls.Find(sCntrlName, true)[0].Text.Split(mCalc.Delimiter));
            return ConstructFilePath2HealthTakeCare(v);
        }
        public string ConstructFilePath2HealthTakeCare(int val)
        {
            return System.IO.Path.Combine(mFolderPath2HealthInfoFiles, "health_tc_" + val.ToString() + ".txt");
        }

        public string ConstructFilePath2HealthTLC()
        {
            string sCntrlName = EnumProvider.Instance.GetReservedOPSEnumFromDescription(EnumProvider.ReservedOPS._יכולת_לטפל_.ToString());

            int v = mCalc.GetCorrectNumberFromSplitedString(mMainForm.Controls.Find(sCntrlName, true)[0].Text.Split(mCalc.Delimiter));
            return ConstructFilePath2HealthTLC(v);
        }
        public string ConstructFilePath2HealthTLC(int val)
        {
            return System.IO.Path.Combine(mFolderPath2HealthInfoFiles, "health_tlc_" + val.ToString() + ".txt");
        }

        public string ConstructFilePath2HealthBody()
        {
            string sCntrlName = EnumProvider.Instance.GetReservedOPSEnumFromDescription(EnumProvider.ReservedOPS._תכונות_בסיסיות_לגוף_.ToString());

            int v = mCalc.GetCorrectNumberFromSplitedString(mMainForm.Controls.Find(sCntrlName, true)[0].Text.Split(mCalc.Delimiter));
            return ConstructFilePath2HealthBody(v);
        }
        public string ConstructFilePath2HealthBody(int val)
        {
            return System.IO.Path.Combine(mFolderPath2HealthInfoFiles, "health_body_" + val.ToString() + ".txt");
        }



        public string ConstructFilePath2HealthPrivateName()
        {
            string sCntrlName = EnumProvider.Instance.GetReservedOPSEnumFromDescription(EnumProvider.ReservedOPS._בריאות_שם_פרטי_.ToString());

            int v = mCalc.GetCorrectNumberFromSplitedString(mMainForm.Controls.Find(sCntrlName, true)[0].Text.Split(mCalc.Delimiter));
            return ConstructFilePath2HealthPrivateName(v);
        }
        public string ConstructFilePath2HealthPrivateName(int val)
        {
            return System.IO.Path.Combine(mFolderPath2HealthInfoFiles, "health_pname_" + val.ToString() + ".txt");
        }

        public string ConstructFilePath2HealthClimax()
        {
            int cycle = CalcCurrentCycle();
            string sCntrlName = "txt" + cycle.ToString() + "_3";

            int v = mCalc.GetCorrectNumberFromSplitedString(mMainForm.Controls.Find(sCntrlName, true)[0].Text.Split(mCalc.Delimiter));
            return ConstructFilePath2HealthClimax(v);
        }
        public string ConstructFilePath2HealthClimax(int val)
        {
            return System.IO.Path.Combine(mFolderPath2HealthInfoFiles, "health_clmx_" + val.ToString() + ".txt");
        }

        public string ConstructFilePath2HealthAstro()
        {
            string sCntrlName = EnumProvider.Instance.GetReservedOPSEnumFromDescription(EnumProvider.ReservedOPS._בריאות_מזל_אסטרולוגי_.ToString());

            string v = mMainForm.Controls.Find(sCntrlName, true)[0].Text.Trim().Split("(".ToCharArray())[1].Trim();

            //int v = mCalc.GetCorrectNumberFromSplitedString(mMainForm.Controls.Find(sCntrlName, true)[0].Text.Split(mCalc.Delimiter));
            return ConstructFilePath2HealthAstro(Convert.ToInt16(v));
        }
        public string ConstructFilePath2HealthAstro(int val)
        {
            return System.IO.Path.Combine(mFolderPath2HealthInfoFiles, "health_astro_" + val.ToString() + ".txt");
        }

        public string ConstructFilePath2HealthPrsnlYear()
        {
            string sCntrlName = EnumProvider.Instance.GetReservedOPSEnumFromDescription(EnumProvider.ReservedOPS._בריאות_שנה_אישית_.ToString());

            int v = mCalc.GetCorrectNumberFromSplitedString(mMainForm.Controls.Find(sCntrlName, true)[0].Text.Split(mCalc.Delimiter));
            return ConstructFilePath2HealthPrsnlYear(v);
        }
        public string ConstructFilePath2HealthPrsnlYear(int val)
        {
            return System.IO.Path.Combine(mFolderPath2HealthInfoFiles, "health_prnslyear_" + val.ToString() + ".txt");
        }

        #endregion

        public string ConstructFilePath2MainPersonality(int IZUN_Value)
        {
            return System.IO.Path.Combine(mFolderPath2MainPersonality, Prefix + "prsnlty_" + IZUN_Value.ToString() + ".txt");
        }

        public string ConstructFilePath2BussinessSuccess()
        {
            string sCntrlName = EnumProvider.Instance.GetReservedOPSEnumFromDescription(EnumProvider.ReservedOPS._ערך_הצלחה_עסקית_.ToString());
            double resVal = Math.Floor((Convert.ToDouble(mMainForm.Controls.Find(sCntrlName, true)[0].Text.Trim()) * 10));

            //int cycle = CalcCurrentCycle();
            string sOut = "";

            if (resVal <= 60.0)
            {
                sOut = mFolderPath2BussinessSccssInfoFiles + "\\" + Prefix + "Bssnss1.txt";
            }
            if ((resVal >= 61) && (resVal <= 70.0))
            {
                sOut = mFolderPath2BussinessSccssInfoFiles + "\\" + Prefix + "Bssnss2.txt";
            }
            if ((resVal >= 71) && (resVal <= 80.0))
            {
                sOut = mFolderPath2BussinessSccssInfoFiles + "\\" + Prefix + "Bssnss3.txt";
            }
            if ((resVal >= 81) && (resVal <= 90.0))
            {
                sOut = mFolderPath2BussinessSccssInfoFiles + "\\" + Prefix + "Bssnss4.txt";
            }
            if (resVal > 91.0)
            {
                sOut = mFolderPath2BussinessSccssInfoFiles + "\\" + Prefix + "Bssnss5.txt";
            }

            return sOut;
        }

        public string ConstructFilePath2BussinessSuccess(double val)
        {
            string sOut = "";
            double resVal = val;

            if (resVal <= 60.0)
            {
                sOut = mFolderPath2BussinessSccssInfoFiles + "\\" + Prefix + "Bssnss1.txt";
            }
            if ((resVal >= 61) && (resVal <= 70.0))
            {
                sOut = mFolderPath2BussinessSccssInfoFiles + "\\" + Prefix + "Bssnss2.txt";
            }
            if ((resVal >= 71) && (resVal <= 80.0))
            {
                sOut = mFolderPath2BussinessSccssInfoFiles + "\\" + Prefix + "Bssnss3.txt";
            }
            if ((resVal >= 81) && (resVal <= 90.0))
            {
                sOut = mFolderPath2BussinessSccssInfoFiles + "\\" + Prefix + "Bssnss4.txt";
            }
            if (resVal > 91.0)
            {
                sOut = mFolderPath2BussinessSccssInfoFiles + "\\" + Prefix + "Bssnss5.txt";
            }


            return sOut;

        }

        public string ConstructFilePath2MarriageSuccess()
        {
            string sCntrlName = EnumProvider.Instance.GetReservedOPSEnumFromDescription(EnumProvider.ReservedOPS._ערך_הצלחה_זוגית_.ToString());
            double resVal = Math.Floor((Convert.ToDouble(mMainForm.Controls.Find(sCntrlName, true)[0].Text.Trim()) * 10));

            //int cycle = CalcCurrentCycle();

            string sOut = "";

            if (resVal <= 60.0)
            {
                sOut = mFolderPath2MarriageSccssInfoFiles + "\\" + Prefix + "Marr1.txt";
            }
            if ((resVal >= 61) && (resVal <= 70.0))
            {
                sOut = mFolderPath2MarriageSccssInfoFiles + "\\" + Prefix + "Marr2.txt";
            }
            if ((resVal >= 71) && (resVal <= 80.0))
            {
                sOut = mFolderPath2MarriageSccssInfoFiles + "\\" + Prefix + "Marr3.txt";
            }
            if ((resVal >= 81))
            {
                sOut = mFolderPath2MarriageSccssInfoFiles + "\\" + Prefix + "Marr4.txt";
            }

            return sOut;
        }

        public string ConstructFilePath2MarriageSuccess(double val)
        {
            string sOut = "";
            double resVal = val;

            if (resVal <= 60.0)
            {
                sOut = mFolderPath2MarriageSccssInfoFiles + "\\" + Prefix + "Marr1.txt";
            }
            if ((resVal >= 61) && (resVal <= 65.0))
            {
                sOut = mFolderPath2MarriageSccssInfoFiles + "\\" + Prefix + "Marr2.txt";
            }
            if ((resVal >= 66) && (resVal <= 70.0))
            {
                sOut = mFolderPath2MarriageSccssInfoFiles + "\\" + Prefix + "Marr3.txt";
            }
            if ((resVal >= 71) && (resVal <= 80.0))
            {
                sOut = mFolderPath2MarriageSccssInfoFiles + "\\" + Prefix + "Marr4.txt";
            }
            if ((resVal >= 81))
            {
                sOut = mFolderPath2MarriageSccssInfoFiles + "\\" + Prefix + "Marr5.txt";
            }

            return sOut;
        }

        public string ConstructFilePath2ChkraInfo(string word)
        {
            string sCntrlName = EnumProvider.Instance.GetReservedChakraEnumFromDescription(word.Trim()).ToString();

            int Val = mCalc.GetCorrectNumberFromSplitedString(mMainForm.Controls.Find(sCntrlName, true)[0].Text.Trim().Split(mCalc.Delimiter));
            string ChakraNum = word.Split(mSeporator.ToCharArray()[0])[2];

            return mFolderPath2ChakraInfoFiles + "\\" + Prefix + "chkra" + ChakraNum + "_" + "num" + Val.ToString() + ".txt";
        }
        public string ConstructFilePath2AstroLuckInfo()
        {
            string sCntrlName = "txtAstroName";

            string Val = mMainForm.Controls.Find(sCntrlName, true)[0].Text.Trim().Split("(".ToCharArray())[1].Trim();

            return mFolderPath2AstroLuckInfoFiles + "\\" + Prefix + "astro_" + Val.Substring(0, Val.Length - 1) + ".txt";
        }
        public string ConstructFilePath2PrivateNameInfo()
        {
            int type = -1;
            string sCntrlName = "txtPName_Num";

            int Val = mCalc.GetCorrectNumberFromSplitedString(mMainForm.Controls.Find(sCntrlName, true)[0].Text.Trim().Split(mCalc.Delimiter));

            if (mCalc.isCarmaticNumber(Val) == true)
            {
                type = 3;
            }
            else
            {
                if (mCalc.isMaterNumber(Val) == true)
                {
                    type = 1;
                }
                else //רגיל
                {
                    type = 2;
                }
            }

            return ConstructFilePath2PrivateNameInfo(type, Val);
        }
        public string ConstructFilePath2PowerNumberInfo()
        {
            string sCntrlName = "txtPowerNum";

            int Val = mCalc.GetCorrectNumberFromSplitedString(mMainForm.Controls.Find(sCntrlName, true)[0].Text.Trim().Split(mCalc.Delimiter));

            return mFolderPath2PwrNumInfoFiles + "\\" + Prefix + "power_" + Val.ToString() + ".txt";
        }
        public string ConstructFilePath2ChallengeInfo(bool app, int cycle)
        {
            string sCntrlName = EnumProvider.Instance.GetReservedLifeCycleEnumFromDescription(EnumProvider.ReservedLifeCycle._אתגר_.ToString());
            sCntrlName = sCntrlName.Replace("$", cycle.ToString());

            int Val = mCalc.GetCorrectNumberFromSplitedString(mMainForm.Controls.Find(sCntrlName, true)[0].Text.Trim().Split(mCalc.Delimiter));

            return mFolderPath2ChlngNumInfoFiles + "\\" + Prefix + "chlng_" + Val.ToString() + ".txt";
        }
        public string ConstructFilePath2ClimaxInfo(bool app, int cycle)
        {
            string sCntrlName = EnumProvider.Instance.GetReservedLifeCycleEnumFromDescription(EnumProvider.ReservedLifeCycle._שיא_.ToString());
            sCntrlName = sCntrlName.Replace("$", cycle.ToString());

            int Val = mCalc.GetCorrectNumberFromSplitedString(mMainForm.Controls.Find(sCntrlName, true)[0].Text.Trim().Split(mCalc.Delimiter));

            return mFolderPath2ClmxNumInfoFiles + "\\" + Prefix + "clmx_" + Val.ToString() + ".txt";
        }
        public string ConstructFilePath2CycleInfo(bool app, int cycle)
        {
            string sCntrlName = EnumProvider.Instance.GetReservedLifeCycleEnumFromDescription(EnumProvider.ReservedLifeCycle._מחזור_.ToString());
            sCntrlName = sCntrlName.Replace("$", cycle.ToString());

            int Val = mCalc.GetCorrectNumberFromSplitedString(mMainForm.Controls.Find(sCntrlName, true)[0].Text.Trim().Split(mCalc.Delimiter));

            return mFolderPath2CycleNumInfoFiles + "\\cycle_" + Val.ToString() + ".txt";
        }

        public string ConstructFilePath2PersonalYearInfo()
        {
            string sCntrlName = "txtPYear";

            string[] splitted = mMainForm.Controls.Find(sCntrlName, true)[0].Text.Trim().Split(mCalc.Delimiter);
            int Val = mCalc.GetCorrectNumberFromSplitedString(splitted);

            for (int i = 0; i < splitted.Length; i++)
            {
                if (splitted[i] == "15")
                {
                    Val = 15;
                }
            }

            return mFolderPath2PrsonalYearInfoFiles + "\\" + Prefix + "personalYear" + Val.ToString() + ".txt";
        }
        public string ConstructFilePath2PersonalMonthInfo()
        {
            string sCntrlName = "txtPMonth";

            string[] splitted = mMainForm.Controls.Find(sCntrlName, true)[0].Text.Trim().Split(mCalc.Delimiter);
            int Val = mCalc.GetCorrectNumberFromSplitedString(splitted);

            for (int i = 0; i < splitted.Length; i++)
            {
                if (splitted[i] == "15")
                {
                    Val = 15;
                }
            }

            return mFolderPath2PrsonalMonthInfoFiles + "\\" + Prefix + "personalMonth" + Val.ToString() + ".txt";
        }
        public string ConstructFilePath2PersonalDayInfo()
        {
            string sCntrlName = "txtPDay";

            string[] splitted = mMainForm.Controls.Find(sCntrlName, true)[0].Text.Trim().Split(mCalc.Delimiter);
            int Val = mCalc.GetCorrectNumberFromSplitedString(splitted);

            for (int i = 0; i < splitted.Length; i++)
            {
                if (splitted[i] == "15")
                {
                    Val = 15;
                }
            }

            return mFolderPath2PrsonalDayInfoFiles + "\\" + Prefix + "personalDay" + Val.ToString() + ".txt";
        }

        public string ConstructFilePath2WorkInfo(string enmReservedWork)
        {
            string sCntrlName = EnumProvider.Instance.GetReservedWorkEnumFromDescription(enmReservedWork);
            int Val;
            if (sCntrlName == "") //  בדיקה חריגה
            {
                int v1 = mCalc.GetCorrectNumberFromSplitedString(mMainForm.Controls.Find("txtNum5", true)[0].Text.Split(mCalc.Delimiter));
                int v2 = mCalc.GetCorrectNumberFromSplitedString(mMainForm.Controls.Find("txtNum6", true)[0].Text.Split(mCalc.Delimiter));

                if ((v1 == 8) && (v2 == 8))
                {
                    return AppSettings.Instance.AppmMainDir + "\\Templets\\InternalWorkValues\\heb_work_irregular.txt";
                }
            }
            if (sCntrlName == "txtAstroName".ToLower())
            {
                string sVal = mMainForm.Controls.Find(sCntrlName, true)[0].Text.Trim().Split(" ".ToCharArray()[0])[1];
                sVal = sVal.Replace(")", " ");
                sVal = sVal.Replace("(", " ");
                Val = Convert.ToInt16(sVal.Trim());
            }
            else
            {
                string[] splitted = mMainForm.Controls.Find(sCntrlName, true)[0].Text.Trim().Split(mCalc.Delimiter);
                Val = mCalc.GetCorrectNumberFromSplitedString(splitted);

                for (int i = 0; i < splitted.Length; i++)
                {
                    if (splitted[i] == "33")
                    {
                        Val = 33;
                    }
                }
            }
            return ConstructFilePath2WorkInfo(EnumProvider.Instance.GetReservedWorkEnumFromString(enmReservedWork), Val);
        }

        public string ConstructFilePath2IZUNInfo()
        {
            string sCntrlName = "txtInfo1";
            string text = mMainForm.Controls.Find(sCntrlName, true)[0].Text.Trim();
            string[] sTexts = text.Split(Environment.NewLine.ToCharArray()[0]);
            string sEnumVal = sTexts[1].Trim();

            EnumProvider.Balance enmBalance = EnumProvider.Balance._לא_מאוזן_;
            switch (sEnumVal)
            {
                case "מאוזן":
                    enmBalance = EnumProvider.Balance._מאוזן_;
                    break;
                case "חצי מאוזן":
                    enmBalance = EnumProvider.Balance._מאוזן_חלקית_;
                    break;
                case "לא מאוזן":
                    enmBalance = EnumProvider.Balance._לא_מאוזן_;
                    break;
            }

            //return mFolderPath2IZUNInfoFiles + "\\izun_" + ((int)enmBalance).ToString() + ".txt";
            return ConstructFilePath2IZUNInfo(enmBalance);
        }
        public string ConstructFilePath2CrysisInfo(string enmReservedCrysis)
        {
            string sCntrlName = EnumProvider.Instance.GetReservedCrysisEnumFromDescription(enmReservedCrysis);
            int Val = mCalc.GetCorrectNumberFromSplitedString(mMainForm.Controls.Find(sCntrlName, true)[0].Text.Trim().Split(mCalc.Delimiter)); ;
            return ConstructFilePath2CrysisInfo(EnumProvider.Instance.GetReservedCrysisEnumFromString(enmReservedCrysis), Val);
        }

        public List<string> ConstructFilePath2RectificationInfo()
        {
            List<string> paths = new List<string>();
            // EnumProvider.Instance.GetReservedOPSEnumFromDescription(EnumProvider.ReservedOPS._ערך_תיקון_.ToString());
            List<int> DefaultRecValues = new List<int>() { -1, 0, 9, 11, 22, 13, 14, 16, 19 };

            List<int> nums = new List<int>();

            List<string> CntrlsNames = new List<string>() { "txtNum5", "txtNum1", "txtNum4", "txtPName_Num", "txt1_4", "txt2_4", "txt3_4", "txt4_4", "txt1_3", "txt2_3", "txt3_3", "txt4_3", "txt1_2", "txt2_2", "txt3_2", "txt4_2" };
            #region Primary
            foreach (string sCntrlName in CntrlsNames)
            {
                int Val = mCalc.GetCorrectNumberFromSplitedString(mMainForm.Controls.Find(sCntrlName, true)[0].Text.Trim().Split(mCalc.Delimiter));

                if (DefaultRecValues.Contains(Val) == true)
                {
                    if ((Val == 0) && (sCntrlName.Substring(0, 6) != "txtNum") && (sCntrlName != "txtPName_Num"))
                    {
                        Val = CorrectZeroValue1(sCntrlName);
                    }
                    if (nums.Contains(Val) == false)
                    {
                        nums.Add(Val);
                    }
                }
            }

            #endregion
            #region Primary Addition - Astro
            string sVal = mMainForm.Controls.Find("txtAstroName", true)[0].Text.Trim().Split(" ".ToCharArray()[0])[1];
            sVal = sVal.Replace(")", " ");
            sVal = sVal.Replace("(", " ");
            int iVal = Convert.ToInt16(sVal.Trim());
            if (DefaultRecValues.Contains(iVal) == true)
            {
                if (nums.Contains(iVal) == false)
                {
                    nums.Add(iVal);
                }
            }
            #endregion Primary Addition - Astro
            foreach (int n in nums)
            {
                string finalPath = "";
                string[] str = ConstructFilePath2RectificationInfo(n).Split("\\".ToCharArray()[0]);
                for (int i = 0; i < str.Length - 1; i++)
                {
                    finalPath += str[i] + "\\";
                }
                finalPath += "prime_" + str[str.Length - 1];

                paths.Add(finalPath);
            }

            CntrlsNames = new List<string>() { "txtNum2", "txtNum3", "txtNum6", "txtNum7", "txtFName_Num" };
            #region Secondary
            List<int> scndNums = new List<int>();
            foreach (string sCntrlName in CntrlsNames)
            {
                int Val = mCalc.GetCorrectNumberFromSplitedString(mMainForm.Controls.Find(sCntrlName, true)[0].Text.Trim().Split(mCalc.Delimiter));

                if (DefaultRecValues.Contains(Val) == true)
                {
                    if ((Val == 0) && (sCntrlName == "txtNum3"))
                    {
                        Val = CorrectZeroValue2();
                    }
                    if ((nums.Contains(Val) == false) && (scndNums.Contains(Val) == false))
                    {
                        scndNums.Add(Val);
                    }
                }
            }

            foreach (int n in scndNums)
            {
                string finalPath = "";
                string[] str = ConstructFilePath2RectificationInfo(n).Split("\\".ToCharArray()[0]);
                for (int i = 0; i < str.Length - 1; i++)
                {
                    finalPath += str[i] + "\\";
                }
                finalPath += "second_" + str[str.Length - 1];

                paths.Add(finalPath);
            }
            #endregion

            #region Carmatic Recification 7

            // If 7 in one of those chakras or first climax or in the current age period or in astro number
            CntrlsNames = new List<string>() { "txtNum5", "txtNum3", "txtNum2", "txtNum1", "txt1_3", "txt" + CalcCurrentCycle().ToString() + "_2" };

            // Any of the conditions - add file text
            if ((CntrlsNames.Any(ctrl => (mMainForm.Controls.Find(ctrl, true).Length == 1) &&
                                        (mMainForm.Controls.Find(ctrl, true)[0].GetType().Name == "TextBox") &&
                                        (mCalc.GetCorrectNumberFromSplitedString(mMainForm.Controls.Find(ctrl, true)[0].Text.Split(mCalc.Delimiter)) == 7))) || (iVal == 7))
            {
                string finalPath = string.Empty;
                string[] str = ConstructFilePath2RectificationInfo(7).Split("\\".ToCharArray()[0]);
                for (int i = 0; i < str.Length - 1; i++)
                {
                    finalPath += str[i] + "\\";
                }
                finalPath += "prime_" + str[str.Length - 1];

                paths.Add(finalPath);
            }

            #endregion Carmatic Recification 7

            #region Carmatic Recification 2

            // If 2 in one of those chakras or first climax or in the current age period or in astro number
            CntrlsNames = new List<string>() { "txtNum5", "txtNum3", "txtNum1", "txt1_3", "txt" + CalcCurrentCycle().ToString() + "_2" };

            // Any of the conditions - add file text
            if ((CntrlsNames.Any(ctrl => (mMainForm.Controls.Find(ctrl, true).Length == 1) &&
                                        (mMainForm.Controls.Find(ctrl, true)[0].GetType().Name == "TextBox") &&
                                        (mCalc.GetCorrectNumberFromSplitedString(mMainForm.Controls.Find(ctrl, true)[0].Text.Split(mCalc.Delimiter)) == 2))) || (iVal == 2))
            {
                string finalPath = string.Empty;
                string[] str = ConstructFilePath2RectificationInfo(2).Split("\\".ToCharArray()[0]);
                for (int i = 0; i < str.Length - 1; i++)
                {
                    finalPath += str[i] + "\\";
                }
                finalPath += "prime_" + str[str.Length - 1];

                paths.Add(finalPath);
            }

            #endregion Carmatic Recification 2

            return paths;//;ConstructFilePath2RectificationInfo(Val);
        }
        private int CorrectZeroValue1(string CntrlName)
        {
            int res = 0;

            int Val1 = mCalc.GetCorrectNumberFromSplitedString(mMainForm.Controls.Find(CntrlName.Replace("_4", "_2"), true)[0].Text.Trim().Split(mCalc.Delimiter));
            int Val2 = mCalc.GetCorrectNumberFromSplitedString(mMainForm.Controls.Find(CntrlName.Replace("_4", "_3"), true)[0].Text.Trim().Split(mCalc.Delimiter));

            bool res1 = (mCalc.isMaterNumber(Val1) == true) && (mCalc.isCarmaticNumber(Val2) == false);
            bool res2 = (mCalc.isCarmaticNumber(Val1) == false) && (mCalc.isMaterNumber(Val2) == true);
            bool res3 = (mCalc.isMaterNumber(Val1) == true) && (mCalc.isMaterNumber(Val2) == true);

            if (res1 || res2 || res3)
            {
                res = -1;
            }

            return res;
        }
        private int CorrectZeroValue2()
        {
            int res = 0;
            bool resName, resBalance, mastr, crmt;

            resName = mCalc.isCarmaticNumber(mCalc.GetCorrectNumberFromSplitedString(mMainForm.Controls.Find("txtPName_Num", true)[0].Text.Trim().Split(mCalc.Delimiter)));

            string txt = mMainForm.Controls.Find("txtInfo1", true)[0].Text.Trim().Split(Environment.NewLine.ToCharArray()[0])[1].Trim();
            resBalance = (txt.Trim() == "מאוזן");

            string sVal = mMainForm.Controls.Find("txtAstroName", true)[0].Text.Trim().Split(" ".ToCharArray()[0])[1];
            sVal = sVal.Replace(")", " ");
            sVal = sVal.Replace("(", " ");
            mastr = mCalc.isMaterNumber(Convert.ToInt16(sVal.Trim()));

            mastr &= mCalc.isMaterNumber(mCalc.GetCorrectNumberFromSplitedString(mMainForm.Controls.Find("txtNum1", true)[0].Text.Trim().Split(mCalc.Delimiter)));
            mastr &= mCalc.isMaterNumber(mCalc.GetCorrectNumberFromSplitedString(mMainForm.Controls.Find("txtNum5", true)[0].Text.Trim().Split(mCalc.Delimiter)));

            crmt = mCalc.isCarmaticNumber(mCalc.GetCorrectNumberFromSplitedString(mMainForm.Controls.Find("txtNum1", true)[0].Text.Trim().Split(mCalc.Delimiter)));
            crmt &= mCalc.isCarmaticNumber(mCalc.GetCorrectNumberFromSplitedString(mMainForm.Controls.Find("txtNum5", true)[0].Text.Trim().Split(mCalc.Delimiter)));


            if ((resBalance == true) && (resName == false))
            {
                if ((mastr = true) && (crmt == false))
                {
                    res = -1;
                }
            }

            return res;
        }
        private bool isNumRectification(int tikunNum)
        {
            bool res = false;

            res |= (tikunNum == 0) || (tikunNum == 9) || (tikunNum == 11) || (tikunNum == 22) || (tikunNum == 13) || (tikunNum == 14) || (tikunNum == 16) || (tikunNum == 19);

            return res;
        }

        public string ConstructFilePath2DesignationInfo()
        {
            string sCntrlName = EnumProvider.Instance.GetReservedOPSEnumFromDescription(EnumProvider.ReservedOPS._ערך_ייעוד_.ToString());
            int Val = mCalc.GetCorrectNumberFromSplitedString(mMainForm.Controls.Find(sCntrlName, true)[0].Text.Trim().Split(mCalc.Delimiter)); ;

            return ConstructFilePath2DesignationInfo(Val);
        }

        public string ConstructFilePath2ChakraOpening(int numC, bool OpenOrClose)
        {
            string openclose;
            if (OpenOrClose == true)
            {
                openclose = "open";
            }
            else
            {
                openclose = "close";
            }
            return System.IO.Path.Combine(mFolderPath2ChakraOpening, Prefix + numC.ToString() + "_" + openclose + ".txt");
        }

        #endregion Data file Path from PROGRAM
        public string CollectFullInfo4AllCycles()
        {
            string outStr = "";
            outStr += CollectFullInfo4Cycle(1) + Environment.NewLine;
            outStr += CollectFullInfo4Cycle(2) + Environment.NewLine;
            outStr += CollectFullInfo4Cycle(3) + Environment.NewLine;
            outStr += CollectFullInfo4Cycle(4) + Environment.NewLine;
            return outStr;
        }

        #region Data file Path from Spesific Value
        public string ConstructFilePath2ChkraInfo(string word, int Val)
        {
            //string sCntrlName = EnumProvider.Instance.GetReservedChakraEnumFromDescription(word.Trim()).ToString();

            //int Val = mCalc.GetCorrectNumberFromSplitedString(mMainForm.Controls.Find(sCntrlName, true)[0].Text.Trim().Split(mCalc.Delimiter));
            string ChakraNum = word.Split(mSeporator.ToCharArray()[0])[2];

            return mFolderPath2ChakraInfoFiles + "\\" + Prefix + "chkra" + ChakraNum + "_" + "num" + Val.ToString() + ".txt";
        }
        public string ConstructFilePath2AstroLuckInfo(int Val)
        {
            //string sCntrlName = "txtAstroName";

            //string Val = mMainForm.Controls.Find(sCntrlName, true)[0].Text.Trim().Split("(".ToCharArray())[1].Trim();

            return mFolderPath2AstroLuckInfoFiles + "\\" + Prefix + "astro_" + Val.ToString() + ".txt";
        }
        public string ConstructFilePath2PrivateNameInfo(int Type, int Val)
        {
            //string sCntrlName = "txtPName_Num";

            //int Val = mCalc.GetCorrectNumberFromSplitedString(mMainForm.Controls.Find(sCntrlName, true)[0].Text.Trim().Split(mCalc.Delimiter));

            return mFolderPath2PrvtNameInfoFiles + "\\" + Prefix + "name_" + Type.ToString() + "_" + Val.ToString() + ".txt";
        }
        public string ConstructFilePath2PowerNumberInfo(int Val)
        {
            //string sCntrlName = "txtPowerNum";

            //int Val = mCalc.GetCorrectNumberFromSplitedString(mMainForm.Controls.Find(sCntrlName, true)[0].Text.Trim().Split(mCalc.Delimiter));

            return mFolderPath2PwrNumInfoFiles + "\\" + Prefix + "power_" + Val.ToString() + ".txt";
        }
        public string ConstructFilePath2ChallengeInfo(int Val)
        {
            //string sCntrlName = EnumProvider.Instance.GetReservedLifeCycleEnumFromDescription(EnumProvider.ReservedLifeCycle._אתגר_.ToString());

            //int Val = mCalc.GetCorrectNumberFromSplitedString(mMainForm.Controls.Find(sCntrlName, true)[0].Text.Trim().Split(mCalc.Delimiter));

            return mFolderPath2ChlngNumInfoFiles + "\\" + Prefix + "chlng_" + Val.ToString() + ".txt";
        }
        public string ConstructFilePath2ClimaxInfo(int Val)
        {
            //string sCntrlName = EnumProvider.Instance.GetReservedLifeCycleEnumFromDescription(EnumProvider.ReservedLifeCycle._שיא_.ToString());

            //int Val = mCalc.GetCorrectNumberFromSplitedString(mMainForm.Controls.Find(sCntrlName, true)[0].Text.Trim().Split(mCalc.Delimiter));

            return mFolderPath2ClmxNumInfoFiles + "\\" + Prefix + "clmx_" + Val.ToString() + ".txt";
        }
        public string ConstructFilePath2CycleInfo(int Val)
        {
            //string sCntrlName = EnumProvider.Instance.GetReservedLifeCycleEnumFromDescription(EnumProvider.ReservedLifeCycle._מחזור_.ToString());

            //int Val = mCalc.GetCorrectNumberFromSplitedString(mMainForm.Controls.Find(sCntrlName, true)[0].Text.Trim().Split(mCalc.Delimiter));

            return mFolderPath2CycleNumInfoFiles + "\\" + Prefix + "cycle_" + Val.ToString() + ".txt";
        }
        public string ConstructFilePath2PersonalYearInfo(int Val)
        {
            //string sCntrlName = "txtPYear";

            //int Val = mCalc.GetCorrectNumberFromSplitedString(mMainForm.Controls.Find(sCntrlName, true)[0].Text.Trim().Split(mCalc.Delimiter));

            return mFolderPath2PrsonalYearInfoFiles + "\\" + Prefix + "personalYear" + Val.ToString() + ".txt";
        }
        public string ConstructFilePath2PersonalMonthInfo(int Val)
        {
            //string sCntrlName = "txtPMonth";

            //int Val = mCalc.GetCorrectNumberFromSplitedString(mMainForm.Controls.Find(sCntrlName, true)[0].Text.Trim().Split(mCalc.Delimiter));

            return mFolderPath2PrsonalMonthInfoFiles + "\\" + Prefix + "personalMonth" + Val.ToString() + ".txt";
        }
        public string ConstructFilePath2PersonalDayInfo(int Val)
        {
            //string sCntrlName = "txtPDay";

            //int Val = mCalc.GetCorrectNumberFromSplitedString(mMainForm.Controls.Find(sCntrlName, true)[0].Text.Trim().Split(mCalc.Delimiter));

            return mFolderPath2PrsonalDayInfoFiles + "\\" + Prefix + "personalDay" + Val.ToString() + ".txt";
        }

        public string ConstructFilePath2IZUNInfo(EnumProvider.Balance enmBalance)
        {
            return mFolderPath2IZUNInfoFiles + "\\" + Prefix + "izun_" + ((int)enmBalance).ToString() + ".txt";
        }

        public string ConstructFilePath2RectificationInfo(int Val)
        {
            return mFolderPath2RectificationInfoFiles + "\\" + Prefix + "rectify_" + Val.ToString() + ".txt";
        }
        public string ConstructFilePath2DesignationInfo(int Val)
        {
            return mFolderPath2DesginationInfoFiles + "\\" + Prefix + "designation_" + Val.ToString() + ".txt";
        }

        public string ConstructFilePath2WorkInfo(EnumProvider.ReservedWork enmWork, int Val)
        {
            return mFolderPath2WorkInfoFiles + "\\" + Prefix + "work_" + ((int)enmWork).ToString() + "_" + Val.ToString() + ".txt";
        }

        /*
        public string ConstructFilePath2IZUNInfo(EnumProvider.Balance enmBalance, int Val)
        {
            return mFolderPath2IZUNInfoFiles + "\\izun_" + ((int)enmBalance).ToString() + "_" + Val.ToString() + ".txt";
        }
        */
        public string ConstructFilePath2PitagorsInfo(string pitNum, string valueNum)
        {
            return mFolderPath2PitagorasInfoFiles + "\\" + Prefix + "pit" + pitNum + "_" + "num" + valueNum + ".txt"; ;
        }

        public string ConstructFilePath2CrysisInfo(EnumProvider.ReservedCrysis enmCrysis, int Val)
        {
            return mFolderPath2CrysisInfoFiles + "\\" + Prefix + "crisys_" + ((int)enmCrysis).ToString() + "_" + Val.ToString() + ".txt";
        }

        public string CollectFullInfo4Cycle(int cycle)
        {
            bool isPrintChllng = false;
            isPrintChllng = !(mMainForm.Controls.Find("cbRemoveChlng", true)[0] as CheckBox).Checked;

            string outStr = "";
            string Ages, ChlngNum, ChlngStory, ClmxNum, ClmxStory, CycleNum, CycleStory;

            string sCntrlName = "txt" + cycle.ToString() + "_1";
            Ages = mMainForm.Controls.Find(sCntrlName, true)[0].Text.Trim();

            sCntrlName = "txt" + cycle.ToString() + "_2";
            CycleNum = mCalc.GetCorrectNumberFromSplitedString(mMainForm.Controls.Find(sCntrlName, true)[0].Text.Trim().Split(mCalc.Delimiter)).ToString();
            //CycleStory = ReadTextFromFile(ConstructFilePath2CycleInfo(Convert.ToInt16(CycleNum)));
            CycleStory = ReadFromTextFile(ConstructFilePath2CycleInfo(Convert.ToInt16(CycleNum)));

            sCntrlName = "txt" + cycle.ToString() + "_3";
            ClmxNum = mCalc.GetCorrectNumberFromSplitedString(mMainForm.Controls.Find(sCntrlName, true)[0].Text.Trim().Split(mCalc.Delimiter)).ToString();
            //ClmxStory = ReadTextFromFile(ConstructFilePath2ClimaxInfo(Convert.ToInt16(ClmxNum)));
            ClmxStory = ReadFromTextFile(ConstructFilePath2ClimaxInfo(Convert.ToInt16(ClmxNum)));//

            ChlngStory = "";
            sCntrlName = "txt" + cycle.ToString() + "_4";
            ChlngNum = mCalc.GetCorrectNumberFromSplitedString(mMainForm.Controls.Find(sCntrlName, true)[0].Text.Trim().Split(mCalc.Delimiter)).ToString();
            if (Convert.ToInt16(ChlngNum) == 0)
            {
                ChlngNum = CorrectZeroValue1(sCntrlName).ToString();
                //ChlngStory = ReadTextFromFile(ConstructFilePath2ChallengeInfo(Convert.ToInt16(ChlngNum)));
                ChlngStory = ReadFromTextFile(ConstructFilePath2ChallengeInfo(Convert.ToInt16(ChlngNum)));

                isPrintChllng = true;
            }
            else
            {
                if (isPrintChllng == true)
                {
                    ChlngStory = ReadFromTextFile(ConstructFilePath2ChallengeInfo(Convert.ToInt16(ChlngNum)));
                }
            }

            if (Convert.ToInt16(ChlngNum) < 0)
            {
                ChlngNum = "0";
            }

            outStr = "בתקופת החיים ה";
            switch (cycle)
            {
                case 1:
                    outStr += "ראשונה";
                    break;
                case 2:
                    outStr += "שנייה";
                    break;
                case 3:
                    outStr += "שלישית";
                    break;
                case 4:
                    outStr += "רביעית";
                    break;
            }

            outStr += " " + " בין הגילאים" + " " + Ages + Environment.NewLine + Environment.NewLine; ;
            outStr += "האתגר שלך (" + ChlngNum + ") :" + ChlngStory + Environment.NewLine + Environment.NewLine; ;
            //if (isPrintChllng == true)
            //{
            outStr += "השיא שלך (" + ClmxNum + ") :" + ClmxStory + Environment.NewLine + Environment.NewLine; ;
            //}
            outStr += "המחזור/המתנה שלך (" + CycleNum + ") :" + CycleStory + Environment.NewLine + Environment.NewLine;
            return outStr;
        }

        public string ConstructFilePath2ChakraOpening(int numC)
        {
            string tbName = "";
            switch (numC)
            {
                case 1:
                    tbName = "txtDC1";
                    break;
                case 2:
                    tbName = "txtDC2";
                    break;
                case 3:
                    tbName = "txtDC3";
                    break;
                case 4:
                    tbName = "txtDC4";
                    break;
                case 5:
                    tbName = "txtDC5";
                    break;
                case 6:
                    tbName = "txtDC6";
                    break;
                case 7:
                    tbName = "txtDC7";
                    break;
                case 9: // מזל אסטרולוגי
                    tbName = "txtDC9";
                    break;
                case 8: // שם פרטי
                    tbName = "txtDC8";
                    break;

            }

            string Val = mMainForm.Controls.Find(tbName, true)[0].Text;
            bool bOpenOrClose;
            if (Val.Trim() == "")
            {
                bOpenOrClose = false;
            }
            else
            {
                bOpenOrClose = true;
            }

            return ConstructFilePath2ChakraOpening(numC, bOpenOrClose);
        }

        private Omega.Enums.EnumProvider.Balance GetBalanceStatus()
        {
            string sCntrlName = "txtInfo1";
            string text = mMainForm.Controls.Find(sCntrlName, true)[0].Text.Trim();
            string[] sTexts = text.Split(Environment.NewLine.ToCharArray()[0]);
            string sEnumVal = sTexts[1].Trim();

            EnumProvider.Balance enmBalance = EnumProvider.Balance._לא_מאוזן_;
            switch (sEnumVal)
            {
                case "מאוזן":
                    enmBalance = EnumProvider.Balance._מאוזן_;
                    break;
                case "חצי מאוזן":
                    enmBalance = EnumProvider.Balance._מאוזן_חלקית_;
                    break;
                case "לא מאוזן":
                    enmBalance = EnumProvider.Balance._לא_מאוזן_;
                    break;
            }

            return enmBalance;
        }
        #endregion Spesific Value

        private string ReadTextFromFile(string path2file)
        {
            string outStr = "";

            FileInfo file = new FileInfo(path2file);

            if (file.Exists == true)
            {
                FileStream fs = new FileStream(path2file, FileMode.Open, FileAccess.Read);
                StreamReader fr = new StreamReader(fs, Encoding.UTF8);

                outStr = fr.ReadToEnd();

                fr.Close();
                fs.Close();

                GC.Collect();
                GC.WaitForPendingFinalizers();
            }

            return outStr;
        }

        public string Reserved2ControlName(string word)
        {
            return EnumProvider.Instance.GetReservedImageEnumFromDescription(word.Trim()).ToString();
        }

        // **********

        public XmlDocument GetXmlDocByName(string name)
        {
            string path2TempletsDir = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(Application.ExecutablePath), @"Templets\ReportStyles");

            XmlDocument tmpDoc = new XmlDocument();
            foreach (string str in ReportTemplets)
            {
                if (name == Path.GetFileNameWithoutExtension(str))
                {
                    tmpDoc.Load(System.IO.Path.Combine(path2TempletsDir, str + ".xml"));
                    return tmpDoc;
                }
            }
            return null;
        }

        // **********

        public bool SaveReport(string name, XmlDocument doc)
        {
            bool res = true;

            try
            {
                string path2TempletsDir = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(Application.ExecutablePath), @"Templets\ReportStyles");
                string path2report = Path.Combine(path2TempletsDir, name + ".xml");
                FileInfo file = new FileInfo(path2report);

                bool found = file.Exists;
                if (found == true)
                {
                    file.Delete();
                }

                doc.Save(path2report);

                if (found == false)
                {
                    ReportTemplets.Add(path2report);
                }
            }
            catch
            {
                res = false;
            }
            return res;
        }

        public List<string> GetAllReservedImagesNames()
        {
            List<string> outList = new List<string>();

            foreach (Omega.Enums.EnumProvider.ReservedImages eRI in Enum.GetValues(typeof(Omega.Enums.EnumProvider.ReservedImages)))
            {
                string str = eRI.ToString();
                str.Replace("_", " ");
                outList.Add(str.Trim());
            }

            return outList;
        }

        private void KillWord()
        {
            string sProcessName = "winword";
            //here we're going to get a list of all running processes on
            //the computer
            try
            {
                mlog.Info("Try to kill Office Word application...");
                foreach (Process clsProcess in Process.GetProcesses())
                {
                    //now we're going to see if any of the running processes
                    //match the currently running processes by using the StartsWith Method,
                    //this prevents us from incluing the .EXE for the process we're looking for.
                    //. Be sure to not
                    //add the .exe to the name you provide, i.e: NOTEPAD,
                    //not NOTEPAD.EXE or false is always returned even if
                    //notepad is running
                    if (clsProcess.ProcessName.ToLower().StartsWith(sProcessName.ToLower()))
                    {
                        //since we found the proccess we now need to use the
                        //Kill Method to kill the process. Remember, if you have
                        //the process running more than once, say IE open 4
                        //times the loop thr way it is now will close all 4,
                        //if you want it to just close the first one it finds
                        //then add a return; after the Kill
                        clsProcess.Kill();
                        mlog.Info("Office Word application closed");

                        //process killed, return true
                        return;

                    }

                }
                mlog.Info("Office Word application closed");

            }
            catch (Exception ex)
            {
                mlog.Error(ex);

            }

            //process not found, return 
        }

        private bool isWordOpen()
        {
            string sProcessName = "winword";
            //here we're going to get a list of all running processes on
            //the computer
            foreach (Process clsProcess in Process.GetProcesses())
            {
                //now we're going to see if any of the running processes
                //match the currently running processes by using the StartsWith Method,
                //this prevents us from incluing the .EXE for the process we're looking for.
                //. Be sure to not
                //add the .exe to the name you provide, i.e: NOTEPAD,
                //not NOTEPAD.EXE or false is always returned even if
                //notepad is running
                if (clsProcess.ProcessName.ToLower().StartsWith(sProcessName.ToLower()))
                {
                    //since we found the proccess we now need to use the
                    //Kill Method to kill the process. Remember, if you have
                    //the process running more than once, say IE open 4
                    //times the loop thr way it is now will close all 4,
                    //if you want it to just close the first one it finds
                    //then add a return; after the Kill

                    //clsProcess.Kill();

                    //process killed, return true
                    return true;
                }
            }
            //process not found, return 
            return false;
        }

        public void PrintReport()
        {
            string sMsgText = "";
            string sMsgCaption = "";
            if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
            {
                sMsgCaption = "הדפסת דוח נומרולוגי";
                sMsgText = "שים לב, לצורך כתיבת הדוח יש לסגור את כל מסמכי האופיס (וורד WORD) הפתוחים !" + Environment.NewLine +
                            "לכן מומלץ לשמור ולסגור את כל מסמכי האופיס הפתוחים כעת" + Environment.NewLine
                            + Environment.NewLine + "האם ניתן להדפיס כרגע את הדוח?";
            }
            if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
            {
                sMsgCaption = "Printing a Numerologic Report";

                sMsgText = "In order to generate a report all Microsoft Office Word documents must be closed !" + Environment.NewLine +
                    "It is recommanded to save and close all open documents..." + Environment.NewLine
                            + Environment.NewLine + "Are all Word Odcuments are closed? Can the report generation run?";
            }

            if (isWordOpen() == true)
            {
                DialogResult dlgRes = MessageBox.Show(sMsgText, sMsgCaption, MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                if (dlgRes == DialogResult.No)
                {
                    return;
                }
            }

            KillWord();

            Control.ControlCollection Cntrls = mMainForm.Controls;

            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Title = "בחר מיקום הדוח";
            sfd.InitialDirectory = System.Environment.SpecialFolder.Desktop.ToString();
            sfd.Filter = "Microsoft Word (2003 Format)(*.doc)|*.doc";
            //sfd.FileName = "" + Omega.MainForm.ActiveForm.Controls.Find("txtPrivateName", true)[0].Text + " " + Omega.MainForm.ActiveForm.Controls.Find("txtFamilyName", true)[0].Text + ".doc";
            sfd.FileName = "" + mCurrentUser.mFirstName.Trim() + " " + mCurrentUser.mLastName.Trim() + ".doc";

            mPro = (mMainForm.Controls.Find("cbProRep", true)[0] as CheckBox).Checked;
            mLang = AppSettings.Instance.ProgramLanguage;

            DialogResult res = sfd.ShowDialog();
            switch (res)
            {
                case DialogResult.Cancel:
                    return;
                case DialogResult.OK:
                    #region Blind Fold
                    Cntrls.Find("label29", true)[0].Visible = false;
                    Cntrls.Find("label98", true)[0].Visible = false;
                    Cntrls.Find("txtUnique", true)[0].Visible = false;
                    Cntrls.Find("txtParentsPresent", true)[0].Visible = false;

                    bool vis1 = Cntrls.Find("picUnique", true)[0].Visible;
                    Cntrls.Find("picUnique", true)[0].Visible = false;

                    bool vis2 = Cntrls.Find("picParentsPresent", true)[0].Visible;
                    Cntrls.Find("picParentsPresent", true)[0].Visible = false;

                    //grpMoreInfo.Visible = false;
                    Cntrls.Find("tabPage1", true)[0].BackgroundImage = Omega.Properties.Resources.backgrond_to_print;
                    #endregion Blind Fold

                    WordWriter wFileWriter = new WordWriter(Path.GetDirectoryName(Application.ExecutablePath), sfd.FileName, mCurrentUser.mFirstName, mCurrentUser.mLastName);
                    #region Build Report
                    wFileWriter.CreateDocument();
                    wFileWriter.TypeHeader();

                    TabControl Tabs = Cntrls.Find("mTabs", true)[0] as TabControl;

                    if (PrntMain == true)
                    {
                        Tabs.TabPages[0].Show();
                        wFileWriter.TypeBodyWithPicture(Tabs.TabPages[0], "מפה כללית:");
                        //wFileWriter.InsertPageBreak();
                    }

                    if (PrntLifeCycles == true)
                    {
                        Tabs.TabPages[1].Show();
                        wFileWriter.TypeBodyWithPicture(Tabs.TabPages[1], "מחזורי החיים:");
                        //wFileWriter.InsertPageBreak();
                    }

                    if (PrntIntansiveMap == true)
                    {
                        Tabs.TabPages[2].Show();
                        wFileWriter.TypeBodyWithPicture(Tabs.TabPages[2], "מפה אינטנסיבית:");
                        //wFileWriter.InsertPageBreak();
                    }

                    if (PrntPitagoras == true)
                    {
                        Tabs.TabPages[3].Show();
                        wFileWriter.TypeBodyWithPicture(Tabs.TabPages[3], "ריבועי פיתגורס:");
                        //wFileWriter.InsertPageBreak();
                    }

                    if (PrntCombinedMap == true)
                    {
                        Tabs.TabPages[4].Show();
                        wFileWriter.TypeBodyWithPicture(Tabs.TabPages[4], "מפה משולבת:");
                        //wFileWriter.InsertPageBreak();
                    }

                    if (PrntChakraOpening == true)
                    {
                        Tabs.TabPages[5].Show();
                        wFileWriter.TypeBodyWithPicture(Tabs.TabPages[5], "פתיחת צ'אקרות:");
                        //wFileWriter.InsertPageBreak();
                    }


                    Tabs.TabPages[6].Show();
                    TabControl BsnsTab = Cntrls.Find("tabControl1", true)[0] as TabControl;

                    if (PrntBsnnsPersonal == true)
                    {
                        BsnsTab.TabPages[0].Show();
                        BsnsTab.Refresh();
                        wFileWriter.TypeBodyWithPicture(BsnsTab.TabPages[0], "התאמה עסקית אישית:");
                    }

                    if (PrntBsnnsMulti == true)
                    {
                        BsnsTab.TabPages[BsnsTab.TabPages.Count - 1].Show();
                        BsnsTab.Refresh();
                        Tabs.TabPages[6].Show();
                        wFileWriter.TypeBodyWithPicture(BsnsTab.TabPages[1], "התאמה עסקית משותפת:");
                    }

                    if (PrntCoupleMatch == true)
                    {
                        Tabs.TabPages[7].Show();
                        wFileWriter.TypeBodyWithPicture(Tabs.TabPages[7], "התאמה זוגית:");
                    }

                    if (PrntLearnSccss == true)
                    {
                        Tabs.TabPages[8].Show();
                        wFileWriter.TypeBodyWithPicture(Tabs.TabPages[8], "הצלחה בלימודים / קשב וריכוז");
                    }

                    BsnsTab.TabPages[0].Show();
                    Tabs.TabPages[0].Show();
                    wFileWriter.FinishDoc();

                    //DialogResult dlgRes = MessageBox.Show("האם לפתוח את הקובץ?", "הייצוא הסתיים", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    //if (dlgRes == DialogResult.OK)
                    //{
                    //    wFileWriter.OpenWordFile();
                    //}
                    #endregion

                    #region Remove Blind Fold
                    Cntrls.Find("label29", true)[0].Visible = true;
                    Cntrls.Find("label98", true)[0].Visible = true;
                    Cntrls.Find("txtUnique", true)[0].Visible = true;
                    Cntrls.Find("txtParentsPresent", true)[0].Visible = true;
                    Cntrls.Find("picUnique", true)[0].Visible = vis1;
                    Cntrls.Find("picParentsPresent", true)[0].Visible = vis2;
                    //grpMoreInfo.Visible = true;

                    Cntrls.Find("tabPage1", true)[0].BackgroundImage = Omega.Properties.Resources.backgrond;
                    #endregion
                    break;
            }

            KillWord();
        }

        public void PrintReportFromXML(string XmlReportName)
        {
            string sMsgText = "";
            string sMsgCaption = "";

            //mIsMSOfficeOK = CheckMSOffice();

            //mIsMSOfficeOK = false;

            mlog.Info("Selecting file location");
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Title = "בחר מיקום הדוח";
            sfd.InitialDirectory = System.Environment.SpecialFolder.Desktop.ToString();

            if (mIsMSOfficeOK == true)
            { // filter with MSO word option
                sfd.Filter = "Microsoft Word (2003 Format)(*.doc)|*.doc|HTML Report (HTM Format)(*.html)|*.html";
                sfd.FileName = "" + mCurrentUser.mFirstName.Trim() + " " + mCurrentUser.mLastName.Trim() + " - " + XmlReportName + ".doc";
                mlog.Info("Selected filename is: " + sfd.FileName);
            }
            else // non MSOffice option      //(mIsMSOfficeOK == false)
            {
                #region Office no good
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                {
                    sMsgCaption = "הדפסת דוח נומרולוגי";
                    sMsgText = "לא ניתן לייצר מסמך וורד, ככל הנראה אחת מהבעיות הבאות קיימות" + Environment.NewLine +
                                "1. לא קיים אופיס חוקי על גבי המחשב" + Environment.NewLine +
                                "2. התקנת האופיס במחשב זה איננה תקינה" + Environment.NewLine +
                                "לאחר התקנה של אופיס תקין במחשב זה ניתן יהיה להדפיס מסמך וורד";
                }
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                {
                    sMsgCaption = "Printing a Numerologic Report";

                    sMsgText = "Creating a Microsoft Office Word report is unavaliable, due to one or more of the following reasons:" + Environment.NewLine +
                                    "1. Original Microsoft Office is not installed on this PC" + Environment.NewLine +
                                    "2. The installation of the current Microsoft Office is corrupted" + Environment.NewLine +
                                    "After a propper installation of Microsoft Office on this PC a Word document will be possible.";
                }

                MessageBox.Show(sMsgText, sMsgCaption, MessageBoxButtons.OK, MessageBoxIcon.Error);
                #endregion Office no good
                sfd.Filter = "HTML Report (HTM Format)(*.html)|*.html";
                sfd.FileName = "" + mCurrentUser.mFirstName.Trim() + " " + mCurrentUser.mLastName.Trim() + " - " + XmlReportName + ".html";
                mlog.Info("Selected filename is: " + sfd.FileName);

            }
            //sfd.FileName = "" + Omega.MainForm.ActiveForm.Controls.Find("txtPrivateName", true)[0].Text + " " + Omega.MainForm.ActiveForm.Controls.Find("txtFamilyName", true)[0].Text + ".doc";


            mPro = (mMainForm.Controls.Find("cbProRep", true)[0] as CheckBox).Checked;
            mLang = AppSettings.Instance.ProgramLanguage;

            DialogResult res = sfd.ShowDialog();
            switch (res)
            {
                case DialogResult.Cancel:
                case DialogResult.No:
                case DialogResult.Abort:
                    return;
                case DialogResult.OK:
                    string ext = System.IO.Path.GetExtension(sfd.FileName);

                    mRepStyle2Create = "";

                    XmlDocument xmlReport;
                    switch (ext)
                    {
                        case ".doc":

                            mRepStyle2Create = "doc";
                            if (mIsMSOfficeOK == true)
                            {
                                #region Microsoft Office Word

                                #region Test Open Word
                                //if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                                //{
                                //    sMsgCaption = "הדפסת דוח נומרולוגי";
                                //    //sMsgText = "שים לב, לצורך כתיבת הדוח יש לסגור את כל מסמכי האופיס (וורד WORD) הפתוחים !" + Environment.NewLine +
                                //    //            "לכן מומלץ לשמור ולסגור את כל מסמכי האופיס הפתוחים כעת" + Environment.NewLine
                                //    //            + Environment.NewLine + "האם ניתן להדפיס כרגע את הדוח?";
                                //    sMsgText = "האם להפיק את הדו\"חות שנבחרו ?";

                                //}
                                //if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                                //{
                                //    sMsgCaption = "Printing a Numerologic Report";

                                //    sMsgText = "In order to generate a report all Microsoft Office Word documents must be closed !" + Environment.NewLine +
                                //        "It is recommanded to save and close all open documents..." + Environment.NewLine
                                //                + Environment.NewLine + "Are all Word Odcuments are closed? Can the report generation run?";
                                //}

                                //if (isWordOpen() == true)
                                //{
                                //    DialogResult dlgRes = MessageBox.Show(sMsgText, sMsgCaption, MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                                //    if (dlgRes == DialogResult.No)
                                //    {
                                //        return;
                                //    }

                                //    KillWord();
                                //}
                                #endregion Test Open Word

                                mMainForm.Cursor = Cursors.WaitCursor;
                                xmlReport = GetXmlDocByName(XmlReportName);

                                mlog.DebugFormat("Creating WordWriter object: [ApplicationMainDir: {0}], [Document Path: {1}], [FirstName: {2}], [LastName: {3}]",
                                    Path.GetDirectoryName(Application.ExecutablePath), sfd.FileName, mCurrentUser.mFirstName, mCurrentUser.mLastName);

                                WordWriter wFileWriter = new WordWriter(Path.GetDirectoryName(Application.ExecutablePath), sfd.FileName, mCurrentUser.mFirstName, mCurrentUser.mLastName);

                                wFileWriter.CreateDocument();

                                wFileWriter.TypeXml2Doc(xmlReport);

                                wFileWriter.FinishDoc();

                                mMainForm.Cursor = Cursors.Default;

                                KillWord();

                                #endregion Microsoft Office Word

                            }
                            else // MSOffice is not properlly installed on PC
                            {
                                #region Office no good
                                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                                {
                                    sMsgCaption = "הדפסת דוח נומרולוגי";
                                    sMsgText = "לא ניתן לייצר מסמך וורד, ככל הנראה אחת מהבעיות הבאות קיימות" + Environment.NewLine +
                                                "1. לא קיים אופיס חוקי על גבי המחשב" + Environment.NewLine +
                                                "2. התקנת האופיס במחשב זה איננה תקינה" + Environment.NewLine +
                                                "לאחר התקנה של אופיס תקין במחשב זה ניתן יהיה להדפיס מסמך וורד";
                                }
                                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                                {
                                    sMsgCaption = "Printing a Numerologic Report";

                                    sMsgText = "Creating a Microsoft Office Word report is unavaliable, due to one or more of the following reasons:" + Environment.NewLine +
                                                    "1. Original Microsoft Office is not installed on this PC" + Environment.NewLine +
                                                    "2. The installation of the current Microsoft Office is corrupted" + Environment.NewLine +
                                                    "After a propper installation of Microsoft Office on this PC a Word document will be possible.";
                                }

                                MessageBox.Show(sMsgText, sMsgCaption, MessageBoxButtons.OK, MessageBoxIcon.Error);

                                return;
                                #endregion Office no good

                            }
                            break;

                        case ".html":
                            mRepStyle2Create = "html";
                            #region HTML
                            xmlReport = GetXmlDocByName(XmlReportName);

                            htmlWriter hw = new htmlWriter(sfd.FileName);
                            hw.TypeXml2HTML(xmlReport);

                            hw.OpenReoprtOnDefaultBrowser();
                            #endregion HTML
                            break;
                    }
                    mRepStyle2Create = "";
                    break;
            }

        }

        public async Task PrintReportFromXMLAsync(string XmlReportName)
        {
            string sMsgText = string.Empty;
            string sMsgCaption = string.Empty;

            mlog.Info("Selecting file location");
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Title = "בחר מיקום הדוח";
            sfd.InitialDirectory = System.Environment.SpecialFolder.Desktop.ToString();

            if (mIsMSOfficeOK == true)
            { // filter with MSO word option
                sfd.Filter = "Microsoft Word (2003 Format)(*.doc)|*.doc|HTML Report (HTM Format)(*.html)|*.html";
                sfd.FileName = "" + mCurrentUser.mFirstName.Trim() + " " + mCurrentUser.mLastName.Trim() + " - " + XmlReportName + ".doc";
                mlog.Info("Selected filename is: " + sfd.FileName);
            }
            else // non MSOffice option      //(mIsMSOfficeOK == false)
            {
                #region Office no good
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                {
                    sMsgCaption = "הדפסת דוח נומרולוגי";
                    sMsgText = "לא ניתן לייצר מסמך וורד, ככל הנראה אחת מהבעיות הבאות קיימות" + Environment.NewLine +
                                "1. לא קיים אופיס חוקי על גבי המחשב" + Environment.NewLine +
                                "2. התקנת האופיס במחשב זה איננה תקינה" + Environment.NewLine +
                                "לאחר התקנה של אופיס תקין במחשב זה ניתן יהיה להדפיס מסמך וורד";
                }
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                {
                    sMsgCaption = "Printing a Numerologic Report";

                    sMsgText = "Creating a Microsoft Office Word report is unavaliable, due to one or more of the following reasons:" + Environment.NewLine +
                                    "1. Original Microsoft Office is not installed on this PC" + Environment.NewLine +
                                    "2. The installation of the current Microsoft Office is corrupted" + Environment.NewLine +
                                    "After a propper installation of Microsoft Office on this PC a Word document will be possible.";
                }

                MessageBox.Show(sMsgText, sMsgCaption, MessageBoxButtons.OK, MessageBoxIcon.Error);
                #endregion Office no good
                sfd.Filter = "HTML Report (HTM Format)(*.html)|*.html";
                sfd.FileName = "" + mCurrentUser.mFirstName.Trim() + " " + mCurrentUser.mLastName.Trim() + " - " + XmlReportName + ".html";
                mlog.Info("Selected filename is: " + sfd.FileName);

            }

            mPro = (mMainForm.Controls.Find("cbProRep", true)[0] as CheckBox).Checked;
            mLang = AppSettings.Instance.ProgramLanguage;

            DialogResult res = sfd.ShowDialog();
            switch (res)
            {
                case DialogResult.Cancel:
                case DialogResult.No:
                case DialogResult.Abort:
                    {
                        return;
                    }
                case DialogResult.OK:
                    {
                        string ext = System.IO.Path.GetExtension(sfd.FileName);

                        mRepStyle2Create = "";

                        XmlDocument xmlReport;
                        switch (ext)
                        {
                            case ".doc":
                                {
                                    mRepStyle2Create = "doc";
                                    if (mIsMSOfficeOK == true)
                                    {
                                        #region Microsoft Office Word

                                        mMainForm.Cursor = Cursors.WaitCursor;
                                        xmlReport = GetXmlDocByName(XmlReportName);

                                        mlog.DebugFormat("Creating WordWriter object: [ApplicationMainDir: {0}], [Document Path: {1}], [FirstName: {2}], [LastName: {3}]",
                                            Path.GetDirectoryName(Application.ExecutablePath), sfd.FileName, mCurrentUser.mFirstName, mCurrentUser.mLastName);

                                        WordWriter wFileWriter = new WordWriter(Path.GetDirectoryName(Application.ExecutablePath), sfd.FileName, mCurrentUser.mFirstName, mCurrentUser.mLastName);

                                        await Task.Run(() => 
                                        {
                                            //spinner.SetSpinnerVisiblity(true);
                                            
                                            wFileWriter.CreateDocument();

                                            wFileWriter.TypeXml2Doc(xmlReport);

                                            wFileWriter.FinishDoc();

                                            //spinner.SetSpinnerVisiblity(false);

                                        });

                                        mMainForm.Cursor = Cursors.Default;

                                        KillWord();

                                        #endregion Microsoft Office Word

                                    }
                                    else // MSOffice is not properlly installed on PC
                                    {
                                        #region Office no good
                                        if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                                        {
                                            sMsgCaption = "הדפסת דוח נומרולוגי";
                                            sMsgText = "לא ניתן לייצר מסמך וורד, ככל הנראה אחת מהבעיות הבאות קיימות" + Environment.NewLine +
                                                        "1. לא קיים אופיס חוקי על גבי המחשב" + Environment.NewLine +
                                                        "2. התקנת האופיס במחשב זה איננה תקינה" + Environment.NewLine +
                                                        "לאחר התקנה של אופיס תקין במחשב זה ניתן יהיה להדפיס מסמך וורד";
                                        }
                                        if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                                        {
                                            sMsgCaption = "Printing a Numerologic Report";

                                            sMsgText = "Creating a Microsoft Office Word report is unavaliable, due to one or more of the following reasons:" + Environment.NewLine +
                                                            "1. Original Microsoft Office is not installed on this PC" + Environment.NewLine +
                                                            "2. The installation of the current Microsoft Office is corrupted" + Environment.NewLine +
                                                            "After a propper installation of Microsoft Office on this PC a Word document will be possible.";
                                        }

                                        MessageBox.Show(sMsgText, sMsgCaption, MessageBoxButtons.OK, MessageBoxIcon.Error);

                                        return;
                                        #endregion Office no good

                                    }
                                    break;

                                }
                            case ".html":
                                {
                                    mRepStyle2Create = "html";
                                    #region HTML
                                    xmlReport = GetXmlDocByName(XmlReportName);

                                    htmlWriter hw = new htmlWriter(sfd.FileName);
                                    hw.TypeXml2HTML(xmlReport);

                                    hw.OpenReoprtOnDefaultBrowser();
                                    #endregion HTML
                                    break;
                                }
                        }
                        break;

                    }

            }

        }

    }

    public class StringComparer : IEqualityComparer<string>
    {
        public bool Equals(string x, string y)
        {
            return x != y;
        }

        public int GetHashCode(string obj)
        {
            throw new NotImplementedException();
        }
    }

}
