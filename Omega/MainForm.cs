#region Namespaces

using BLL;
using log4net;
using log4net.Config;
using Omega.Enums;
using Omega.Objects;
using Omega.Reports;
using Shell32;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;

#endregion Namespaces

// **********

namespace Omega
{

    public partial class MainForm : Form
    {
        #region Data Members

        public BLL.Calc Calculator;
        public string mApplicationMainDir = AppSettings.Instance.AppmMainDir;
        private int mLuckNumber;
        private string path2ProffongFile = AppSettings.Instance.AppmMainDir + "\\Settings\\ProofConfig.txt";
        private TabPage tcLifeCycles = new TabPage();
        private readonly DateTime dateVersion = new DateTime(2018, 8, 16);
        private static readonly ILog mlog = LogManager.GetLogger("NumChakra");

        #region Personal Info.

        private string mFirstName;
        private string mLastName;
        private string mMotherName;
        private string mFatherName;
        private DateTime mB_Date;
        private string mCity;
        private string mStreet;
        private double mBuildingNum;
        private double mAppNum;

        private string mEMail;
        private string mPhones;

        private UserInfo CurrentUserData;

        private string mApplication;
        private Enums.EnumProvider.Sex mSex;
        private Enums.EnumProvider.PassedRectification mPassedRect;
        private Enums.EnumProvider.ReachedMaster mReachMaster;

        #endregion  Personal Info.

        #region Color Management

        System.Drawing.Color mHalfCarmaticColor = System.Drawing.Color.FromArgb(191, 191, 191); // GREY
        System.Drawing.Color mMasterColor = System.Drawing.Color.FromArgb(248, 40, 90);// RED //(254, 56, 56);
        System.Drawing.Color mEnhancedColor = System.Drawing.Color.FromArgb(255, 0, 0); //RED
        System.Drawing.Color mRegularColor = System.Drawing.Color.FromArgb(255, 255, 255); // WHITE
        System.Drawing.Color mBlack = System.Drawing.Color.FromArgb(0, 0, 0); // BLACK
        System.Drawing.Color mBlue = System.Drawing.Color.FromArgb(56, 198, 255); // AZURE
        System.Drawing.Color mGreen = System.Drawing.Color.FromArgb(0, 255, 0); // GREEN

        #endregion Color Managment

        #region Text Arrays

        TextBox[] mIntsMapTextArr; //= new TextBox[9];
        TextBox[] mPitgMapTextArr; //= new TextBox[9];
        TextBox[] mCombMapTextArr;// = new TextBox[9];

        #endregion Text Arrays

        private string mTempDirPath;

        private bool isMultiBusineesCalc = false;

        private bool isMultiSecereteryCalc = false;

        private bool isNeedToCalculateWithMsala = false;

        //private bool isMultiAccountingCalc = false;

        #endregion Data Members
        #region QnA
        public bool isRunningNameBalance = false;
        public bool isRunningQnA = false, isRunningQnAspouce = false;
        public int currentQ = -1;
        public string sQnAresWorkType;
        public List<string> sWorkTypes;
        List<string> txtOut = new List<string>();
        #endregion QnA

        public readonly string ProcessSpinnerName = "ucProcessSpinner";
        private const string TabPageLifeCycle2 = "tcLifeCycles2";

        private UserResult SpouceResult;
        private Omega.Objects.UserInfo curPartner;

        public string AppMainDir
        {
            get
            {
                return mApplicationMainDir;
            }
        }

        public delegate void cmdRunProg();
        public cmdRunProg dlgRunProg;

        private bool mPro;
        private AppSettings.Language mLang;
        private ucSpinner spinner;

        public MainForm()
        {
            XmlConfigurator.Configure();
            //MessageBox.Show($"גירסה: {dateVersion.ToLongDateString()}", string.Empty, MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);

            mlog.Info("\n\n\nNumChakra Session Started");

            InitializeComponent();
            //EnumProvider.Instance.Init();

            SetAppOptionsByAppSettings();

            Calculator = new BLL.Calc(AppSettings.Instance.ProgramLanguage, true);

            mCombMapTextArr = new TextBox[9] { txtMComb1, txtMComb2, txtMComb3, txtMComb4, txtMComb5, txtMComb6, txtMComb7, txtMComb8, txtMComb9 };
            mPitgMapTextArr = new TextBox[9] { txtPit1, txtPit2, txtPit3, txtPit4, txtPit5, txtPit6, txtPit7, txtPit8, txtPit9 };
            mIntsMapTextArr = new TextBox[9] { txtIMap1, txtIMap2, txtIMap3, txtIMap4, txtIMap5, txtIMap6, txtIMap7, txtIMap8, txtIMap9 };

            mTempDirPath = System.IO.Path.Combine(System.Environment.CurrentDirectory, "TmpDir");

            dlgRunProg = new cmdRunProg(runCalc);
            grpPrint.SendToBack();

            if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
            {
                sWorkTypes = new List<string> { "ניהול", "שירות", "פקידות", "ספורט" };
            }
            if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
            {
                sWorkTypes = new List<string> { "ניהול", "שירות", "פקידות", "ספורט" };
            }

        }

        // **********

        private void MainForm_Load(object sender, EventArgs e)
        {
            #region Q_n_A
            /*
01. עתידות קלאסי
02. מתי כדאי לי להתחתן?
03. מתי כדאי לי לרכשו דירה?
04. מתי כדאי לי לפתוח עסק?
05. מתי כדאי לי לבצע פעולה עסקית?
06. מתי כדאי לי לגשת למבחן תיאוריה או טסט לרכב?
07. מהו הזמן המתאים למכור דירה?
08. מתי כדאי לי לשנות שם?
09. שותפיות עסקיות
10. האם הוא אוהב אותי?
11. מתי אמצא זוגיות?
12. האם בן הזוג שלי מתאים לי?
13. האם אנחנו נתחתן?
14. האם כדאי להתחתן עם בן הזוג?
15. האם הוא יחזור אלי?
16. האם כדאי שיחזור אלי?
17. האם הוא בוגד בי?
18. למה הזוגיות מתעכבת?
19. מתי אסתדר כלכלית?
20. מתי אמצא עבודה?
21. האם לעבור לעבודה חדשה?
22. האם לעשות שינוי בקריירה?
23. במה מתאים לי לעסוק?
24. מתי אקבל קידום בעבודה?
25. האם כדאי לפתוח עסק?
26. באיזה תחום כדאי לפתוח עסק?
27. מה היא  הדרך הנכונה להצלחה כלכלית?
28. מתי דברים יתחילו להסתדר עבורי?
29. האם השם שלי מתאים לי?
30. מה צופן לי העתיד?
31. למה כלכך קשה לי בחיים?
32. מתי יפתח לי המזל?
33. מתי בן\בת ימצאו זוגיות?
34. האם הבן\בת יתחתנו בקרוב?
35. האם אצליח להכנס להריון?
36. האם כדאי לעבור דירה?
37. האם עכשיו זה זמן טוב לקנות דירה?
            */
            #endregion
            #region Shurtcut
            string linkName = "";
            switch (BLL.AppSettings.Instance.ProgramLanguage)
            {
                case AppSettings.Language.Hebrew:
                    linkName = "נומרולוגיית הצ'אקרות הדינאמיות";
                    break;
                case AppSettings.Language.English:
                    linkName = "Dynamic Chakra Numerology";
                    break;
            }
            appShortcutToDesktop(linkName);
            #endregion Shurtcut

            //איזוןשםToolStripMenuItem.Visible = false;

            if (InitTestPassword() == true)
            {
                tcInput.BringToFront();
                grpPrint.SendToBack();
                tcInput.Show();

                #region Ms MDB
                // TODO: This line of code loads data into the 'omegaDataSet.Clients' table. You can move, or remove it, as needed.
                //this.clientsTableAdapter.Fill(this.omegaDataSet.Clients);
                #endregion Ms MDB

                #region XML_DB
                XmlDBHandler.Instance.XmlDB2DataGridView(ref gvDBview);
                #endregion XML_DB

                #region BirthDays Comming
                string BDmssg = "לאנשים הבאים יש יום-הולדת בקרוב:" + System.Environment.NewLine + "שם - ימים" + System.Environment.NewLine;
                List<string> names = new List<string>();
                for (int i = 0; i < gvDBview.Rows.Count; i++)
                {
                    DataGridViewRow curRow = gvDBview.Rows[i];
                    //curRow.Cells[i].Value.ToString()
                    string dbDate = curRow.Cells[5].Value.ToString();
                    DateTime thisDate = new DateTime();
                    if (DateTime.TryParse(dbDate, out thisDate))
                    {
                        int y = 0, m = 0, d = 0;
                        //string[] splt = dbDate.Split(" ".ToCharArray())[0].Split("-".ToCharArray());

                        //y = Convert.ToInt16(splt[2]);
                        y = thisDate.Year;
                        //if (y.ToString().Length < 4)
                        //{
                        //    if (y >= 50)
                        //    {
                        //        y = 1900 + y;
                        //    }
                        //    else
                        //    {
                        //        y = 2000 + y;
                        //    }
                        //}

                        m = thisDate.Month;
                        d = thisDate.Day;
                        try
                        {
                            DateTime bd = new DateTime(DateTime.Now.Year, m, d);

                            if (Math.Abs(DateTime.Now.Subtract(bd).Days) < 7)
                            {

                                names.Add(curRow.Cells[0].Value.ToString() + " " + curRow.Cells[1].Value.ToString() + "," + DateTime.Now.Subtract(bd).Days.ToString());
                                BDmssg += names[names.Count - 1] + System.Environment.NewLine;


                            }
                        }
                        catch //(Exception expBDs)
                        {
                            //MessageBox.Show(expBDs.Message);
                        }
                    }

                }

                if (names.Count > 0)
                {
                    MessageBox.Show(BDmssg, "יומולדת שמח ל", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                #endregion

                ClearForm();
            }
            else
            {

            }

            dlgRunProg = new cmdRunProg(runCalc);
            txtHebDate.Text = Calculator.GeorgianDate2HebrewJewishDateString(DateTime.Now);
            dgvQnA.Columns[1].HeaderText = "קוסמו-\nנומרולוגי";

            // Save current extra life cycle tag instance to variable
            tcLifeCycles2 = tcInput.TabPages[TabPageLifeCycle2];
            tcInput.TabPages.RemoveByKey(TabPageLifeCycle2);

        }

        private string ChangeMonthName2MonthNum(string str)
        {
            int outVal = 0;
            switch (str.ToLower())
            {
                case "jan":
                case "ינו":
                    outVal = 1;
                    break;
                case "fe":
                case "פבר":
                    outVal = 2;
                    break;
                case "mar":
                case "מרץ":
                    outVal = 3;
                    break;
                case "apr":
                case "אפר":
                    outVal = 4;
                    break;
                case "may":
                case "מאי":
                    outVal = 5;
                    break;
                case "jun":
                case "יונ":
                    outVal = 6;
                    break;
                case "jul":
                case "יול":
                    outVal = 7;
                    break;
                case "aug":
                case "אוג":
                    outVal = 8;
                    break;
                case "sep":
                case "ספט":
                    outVal = 9;
                    break;
                case "oct":
                case "אוק":
                    outVal = 10;
                    break;
                case "nov":
                case "נוב":
                    outVal = 11;
                    break;
                case "dec":
                case "דצמ":
                    outVal = 12;
                    break;
            }



            return outVal.ToString();
        }

        private void SetAppOptionsByAppSettings()
        {
            #region Type
            if (AppSettings.Instance.ProgramType == AppSettings.ProgType.Trial)
            {
                #region Trial
                menuStrip1.Items.Find("איזוןשםToolStripMenuItem", true)[0].Visible = false;
                //menuStrip1.Items.Find("צורדוחToolStripMenuItem", true)[0].Visible = false;
                //menuStrip1.Items.Find("החלףמסמךבסיסToolStripMenuItem", true)[0].Visible = false;
                menuStrip1.Items.Find("הכןערכיצאקרותומספריםToolStripMenuItem", true)[0].Visible = false;
                menuStrip1.Items.Find("דוחותToolStripMenuItem", true)[0].Enabled = false;

                mTabs.TabPages.Remove(mTabs.TabPages[6]);
                mTabs.Refresh();
                #endregion Trial
            }

            if (AppSettings.Instance.ProgramType == AppSettings.ProgType.Normal)
            {
                #region Normal
                menuStrip1.Items.Find("איזוןשםToolStripMenuItem", true)[0].Visible = false;
                //menuStrip1.Items.Find("צורדוחToolStripMenuItem", true)[0].Visible = true;
                menuStrip1.Items.Find("החלףמסמךבסיסToolStripMenuItem", true)[0].Visible = true;

                //mTabs.TabPages.RemoveByKey("tabPage2");
                #endregion Normal
                menuStrip1.Items.Find("הכןערכיצאקרותומספריםToolStripMenuItem", true)[0].Visible = false;
            }
            if (AppSettings.Instance.ProgramType == AppSettings.ProgType.Expert)
            {
                #region Expert
                menuStrip1.Items.Find("איזוןשםToolStripMenuItem", true)[0].Visible = true;
                menuStrip1.Items.Find("צורדוחToolStripMenuItem", true)[0].Enabled = true;
                menuStrip1.Items.Find("החלףמסמךבסיסToolStripMenuItem", true)[0].Visible = true;
                #endregion Expert
                menuStrip1.Items.Find("הכןערכיצאקרותומספריםToolStripMenuItem", true)[0].Visible = false;
            }
            if (AppSettings.Instance.ProgramType == AppSettings.ProgType.DAD)
            {
                #region DAD
                menuStrip1.Items.Find("איזוןשםToolStripMenuItem", true)[0].Visible = true;
                //menuStrip1.Items.Find("צורדוחToolStripMenuItem", true)[0].Visible = true;
                menuStrip1.Items.Find("החלףמסמךבסיסToolStripMenuItem", true)[0].Visible = true;
                #endregion DAD
                menuStrip1.Items.Find("הכןערכיצאקרותומספריםToolStripMenuItem", true)[0].Visible = true;
            }
            #endregion Type

            #region Langugae
            switch (AppSettings.Instance.ProgramLanguage)
            {
                case AppSettings.Language.Hebrew:
                    break;
                case AppSettings.Language.English:
                    //MessageBox.Show("TBD: Need to fix UI");
                    #region US-En
                    this.RightToLeft = System.Windows.Forms.RightToLeft.No;
                    this.Text = "Classic Chakra Numerology - by Yaakobi Nesher-Solan" + ((char)169).ToString();

                    #region Strip Menu
                    אודותToolStripMenuItem1.Text = "About";
                    איזוןשםToolStripMenuItem.Text = "First name balancing";
                    בסיסנתוניםToolStripMenuItem.Text = "DataBase";
                    החלףמסמךבסיסToolStripMenuItem.Text = "Replace word templet";
                    הכןסגנונותדוחToolStripMenuItem.Text = "Preper personal report styles";
                    הכןערכיצאקרותומספריםToolStripMenuItem.Text = "Preper personal chakras values";
                    ייבואToolStripMenuItem.Text = "Import to DataBase";
                    ייצואגיבויToolStripMenuItem.Text = "Export \\ Backup DataBase";
                    יציאהToolStripMenuItem.Text = "Exit";
                    בצעחישובToolStripMenuItem.Text = "Run calculation";
                    לקוחחדשToolStripMenuItem.Text = "New client";
                    צורדוחToolStripMenuItem.Text = "Create plain report";
                    צורדוחמסגנוןToolStripMenuItem.Text = "Create report from style";
                    קובץToolStripMenuItem.Text = "File";
                    דוחותToolStripMenuItem.Text = "Reports";
                    #endregion Strip Menu


                    txtHebDate.Visible = false;
                    btnConvertHebDate2GregorianDate.Visible = false;

                    tcInput.TabPages[0].Text = "Clients Info";
                    tcInput.TabPages[1].Text = "All Clients";
                    tcInput.TabPages[2].Text = "Search";

                    mTabs.TabPages[0].Text = "Main output";
                    mTabs.TabPages[1].Text = "Life Cycles";
                    mTabs.TabPages[2].Text = "Energetic Map";
                    mTabs.TabPages[3].Text = "Pitagors Squares";
                    mTabs.TabPages[4].Text = "Combained Map";
                    mTabs.TabPages[5].Text = "Chakras Status";
                    mTabs.TabPages[6].Text = "Dynamic Chakras Status";
                    mTabs.TabPages[7].Text = "Business Potential";
                    mTabs.TabPages[8].Text = "Coupling Potential";
                    mTabs.TabPages[9].Text = "Learning " + ((char)38).ToString() + " ADD";
                    mTabs.TabPages[10].Text = "Health";



                    checkBox1.Text = "Auto Save" + Environment.NewLine + "to DataBase";
                    cmdSave.Text = "Save To DataBase";
                    cmdRun.Text = "Run and Save";
                    cbPersonMaster.Text = "Did not reached the Master potential";
                    cbPersonMaster.RightToLeft = System.Windows.Forms.RightToLeft.No;
                    cbMainCorrectionDone.Text = "Passed the Main rectification";
                    cbMainCorrectionDone.RightToLeft = System.Windows.Forms.RightToLeft.No;
                    cbMainCorrectionDone.Left = cbPersonMaster.Left;

                    groupBox1.Text = "Personal Info";
                    groupBox1.RightToLeft = System.Windows.Forms.RightToLeft.No;
                    groupBox2.Text = "Location";
                    grpMoreInfo.Text = "More Info";

                    groupBox7.Text = "Birthday Info";
                    groupBox7.Height = 45;
                    groupBox7.Top = txtEMail.Top;

                    label113.Text = "All rights reserved to Yaakobi Nesher-Solan" + ((char)174).ToString();
                    label10.Text = "Crown";
                    label11.Text = "Third Eye";
                    label12.Text = "Throught";
                    label13.Text = "Heart";
                    label14.Text = "Solar Plexus";
                    label15.Text = "Sex " + ((char)38).ToString() + " Creativity";
                    label16.Text = "Base";
                    label17.Text = "Solar Plexus";
                    label114.Text = "Name Chakra";
                    label115.Text = "Universal Chakra";
                    label28.Text = "Last Name";
                    label27.Text = "Zodiac";
                    label26.Text = "Age";
                    label25.Text = "First Name";
                    label23.Text = "Strength No.";
                    label62.Text = "City";
                    label63.Text = "Appartment";
                    label1.Text = "First Name";
                    label2.Text = "Last Name";
                    label3.Text = "Father's Name";
                    label4.Text = "Mother's Name";
                    label5.Text = "E-Mail";
                    label110.Text = "Phones";
                    label6.Text = "City";
                    label7.Text = "Street";
                    label8.Text = "Building No.";
                    label9.Text = "Appartment No.";
                    label22.Text = "Calculation For Date:";

                    cmdLoad.Text = "Load to project";
                    cmdNewClient.Text = "New Client";
                    button1.Text = "Today";

                    #region DB Tab
                    btnLoad2Project2.Text = "Load to project";
                    gvDBview.Columns[0].HeaderText = "#";
                    gvDBview.Columns[1].HeaderText = "Name";
                    gvDBview.Columns[2].HeaderText = "Last Name";
                    gvDBview.Columns[3].HeaderText = "Father's Name";
                    gvDBview.Columns[4].HeaderText = "Mother's Name";
                    gvDBview.Columns[5].HeaderText = "BirthDay";
                    gvDBview.Columns[6].HeaderText = "City";
                    gvDBview.Columns[7].HeaderText = "Street";
                    gvDBview.Columns[8].HeaderText = "Building";
                    gvDBview.Columns[9].HeaderText = "Appartment";
                    gvDBview.Columns[10].HeaderText = "E-Mail";
                    gvDBview.Columns[11].HeaderText = "Phones";
                    #endregion DB Tab

                    #region Search Tab
                    label112.Text = "words to search";
                    btnSearchDB.Text = "Search";
                    btnResetSearch.Text = "Clear Search";
                    dgvSearch.Columns[0].HeaderText = "#";
                    dgvSearch.Columns[1].HeaderText = "Name";
                    dgvSearch.Columns[2].HeaderText = "Last Name";
                    dgvSearch.Columns[3].HeaderText = "Father's Name";
                    dgvSearch.Columns[4].HeaderText = "Mother's Name";
                    dgvSearch.Columns[5].HeaderText = "BirthDay";
                    dgvSearch.Columns[6].HeaderText = "City";
                    dgvSearch.Columns[7].HeaderText = "Street";
                    dgvSearch.Columns[8].HeaderText = "Building";
                    dgvSearch.Columns[9].HeaderText = "Appartment";
                    dgvSearch.Columns[10].HeaderText = "E-Mail";
                    dgvSearch.Columns[11].HeaderText = "Phones";

                    txtSearchText.Text = "Insert words to search";
                    #endregion Search Tab

                    #region Life Cycles Tab
                    label52.Text = "Univeral / Life Cycles map and Future Prediction";

                    label53.Text = "Cycle No.";
                    label54.Text = "Cycle Value";
                    label56.Text = "Ages";
                    label57.Text = "Climax Value";
                    label58.Text = "Challange Value";

                    groupBox3.Text = "Presonal Information";
                    label18.Text = "Zodiac";
                    label97.Text = "Age";
                    label19.Text = "Personal Year";
                    label20.Text = "Personal Month";
                    label21.Text = "Personal Day";

                    groupFuture.Text = "Future Values";
                    label75.Text = "Time Interval";
                    label77.Text = "Qty.";
                    label76.Text = "untill";
                    btnFutureCalc.Text = "Go";
                    listIntervalType.Text = "Day";
                    listIntervalType.Items.Clear();
                    listIntervalType.Items.Add("Day");
                    listIntervalType.Items.Add("Week");
                    listIntervalType.Items.Add("Month");
                    listIntervalType.Items.Add("Year");

                    dataFuture.Columns[0].HeaderText = "Date (Fate)";
                    dataFuture.Columns[1].HeaderText = "Persoanl Year";
                    dataFuture.Columns[2].HeaderText = "Personal Month";
                    dataFuture.Columns[3].HeaderText = "Personal Day";
                    #endregion 

                    #region Personal Name Tab
                    label30.Text = "Presonal Name Map :";
                    label31.Text = "Data inside the table is orgenised by:" + Environment.NewLine + "Physical = 4 , 5" + Environment.NewLine + "Emotional = 2 , 3 , 6" + Environment.NewLine + "Mental = 1 , 8" + Environment.NewLine + "Energetic = 7 , 9";

                    dgvIntsMapSum.Columns[0].HeaderText = "Type";
                    dgvIntsMapSum.Columns[1].HeaderText = "Qty.";
                    #endregion 

                    cmbSexSelect.Text = "Gender";
                    cmbSexSelect.Items.Clear();
                    cmbSexSelect.Items.Add("Male" as object);
                    cmbSexSelect.Items.Add("Female" as object);

                    #endregion US-En
                    break;
            }
            #endregion Langugae


        }

        private void appShortcutToDesktop(string linkName)
        {
            string deskDir = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
            string urlPath = deskDir + "\\" + linkName + ".lnk";

            string lnkDir = System.IO.Path.GetDirectoryName(urlPath);
            string lnkName = System.IO.Path.GetFileNameWithoutExtension(urlPath);

            string appPath = System.Reflection.Assembly.GetExecutingAssembly().Location;

            FileInfo urlFile = new FileInfo(urlPath);
            if (urlFile.Exists == true) return;

            //using (StreamWriter writer = new StreamWriter(urlPath))
            //{
            //    string app = System.Reflection.Assembly.GetExecutingAssembly().Location;
            //    writer.WriteLine("[InternetShortcut]");
            //    writer.WriteLine("URL=file:///" + app);
            //    writer.WriteLine("IconIndex=0");
            //    string icon = app.Replace('\\', '/');
            //    writer.WriteLine("IconFile=" + icon);
            //    writer.Flush();
            //}

            System.IO.File.WriteAllBytes(urlPath, new byte[] { });

            // Initialize a ShellLinkObject for that .lnk file
            Shell shl = new Shell();
            Shell32.Folder dir = shl.NameSpace(lnkDir);
            Shell32.FolderItem itm = dir.Items().Item(lnkName + ".lnk");
            Shell32.ShellLinkObject lnk = (Shell32.ShellLinkObject)itm.GetLink;

            // We'll just dummy a link to notepad
            lnk.Path = appPath;
            lnk.Description = "Dynamic Chakra Numerology";
            lnk.Arguments = @"";
            lnk.WorkingDirectory = @"c:\";

            // And dummy an icon (it will use notepad's)
            lnk.SetIconLocation(appPath, 0);

            // Done, save it
            lnk.Save(urlPath);
        }

        #region Menu Strip
        private void בדיקותToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            DateTime time = new DateTime(2000, 2, 30);

            DateTime res;
            DateTime.TryParse(time.ToLongDateString(), out res);
            MessageBox.Show(res.ToLongDateString());
        }

        private void יציאהToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Calculator.Terminate();
            MainForm.ActiveForm.Hide();
            Close();
        }

        private void אודותToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            AboutBox dlgAbout = new AboutBox();

            dlgAbout.Show();
        }

        private void איזוןשםToolStripMenuItem_Click(object sender, EventArgs e)
        {
            isRunningNameBalance = true;
            if ((txtPrivateName.Text == "") || (txtFamilyName.Text == "")) return;

            ReName frmReName = new ReName(this);

            frmReName.Show();
        }

        private void לקוחחדשToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ClearForm();
        }

        private void בצעחישובToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (GatherData() == true)
            {
                NumericalCalc1();

                SetStylesToObjects();
            }
            else
            {
                return; // nothing was calclated
            }
        }

        #region reports
        private void צורדוחToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if ((txtPrivateName.Text == "") && (txtFamilyName.Text == "") || (txtNum1.Text == ""))
            {
                MessageBox.Show("אין חישוב עבור אדם - לכן לא יווצר דוח", "חוסר מידע בתכנה", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Title = "בחר מיקום הדוח";
            sfd.InitialDirectory = System.Environment.SpecialFolder.Desktop.ToString();
            sfd.Filter = "Microsoft Word (2003 Format)(*.doc)|*.doc";
            sfd.FileName = "" + txtPrivateName.Text + " " + txtFamilyName.Text + ".doc";

            DialogResult res = sfd.ShowDialog();
            switch (res)
            {
                case DialogResult.Cancel:
                    return;
                case DialogResult.OK:
                    #region Blind Fold
                    label29.Visible = false;
                    label98.Visible = false;
                    txtUnique.Visible = false;
                    txtParentsPresent.Visible = false;

                    bool vis1 = picUnique.Visible;
                    picUnique.Visible = false;

                    bool vis2 = picParentsPresent.Visible;
                    picParentsPresent.Visible = false;

                    //grpMoreInfo.Visible = false;
                    tabPage1.BackgroundImage = Omega.Properties.Resources.backgrond_to_print;
                    #endregion Blind Fold

                    WordWriter wFileWriter = new WordWriter(mApplicationMainDir, sfd.FileName, txtPrivateName.Text, txtFamilyName.Text);
                    #region Build Report
                    wFileWriter.CreateDocument();
                    wFileWriter.TypeHeader();

                    mTabs.TabPages[0].Show();
                    wFileWriter.TypeBodyWithPicture(mTabs.TabPages[0], "מפה כללית:");
                    //wFileWriter.InsertPageBreak();

                    mTabs.TabPages[1].Show();
                    wFileWriter.TypeBodyWithPicture(mTabs.TabPages[1], "מחזורי החיים:");
                    //wFileWriter.InsertPageBreak();

                    mTabs.TabPages[2].Show();
                    wFileWriter.TypeBodyWithPicture(mTabs.TabPages[2], "מפה אינטנסיבית:");
                    //wFileWriter.InsertPageBreak();

                    mTabs.TabPages[3].Show();
                    wFileWriter.TypeBodyWithPicture(mTabs.TabPages[3], "ריבועי פיתגורס:");
                    //wFileWriter.InsertPageBreak();

                    mTabs.TabPages[4].Show();
                    wFileWriter.TypeBodyWithPicture(mTabs.TabPages[4], "מפה משולבת:");
                    //wFileWriter.InsertPageBreak();

                    mTabs.TabPages[5].Show();
                    wFileWriter.TypeBodyWithPicture(mTabs.TabPages[5], "פתיחת צ'אקרות:");
                    //wFileWriter.InsertPageBreak();

                    mTabs.TabPages[6].Show();
                    tabControl1.TabPages[0].Show();
                    tabControl1.Refresh();
                    wFileWriter.TypeBodyWithPicture(tabControl1.TabPages[0], "התאמה עסקית אישית:");

                    tabControl1.TabPages[tabControl1.TabPages.Count - 1].Show();
                    tabControl1.Refresh();
                    mTabs.TabPages[6].Show();
                    wFileWriter.TypeBodyWithPicture(tabControl1.TabPages[1], "התאמה עסקית משותפת:");

                    mTabs.TabPages[7].Show();
                    wFileWriter.TypeBodyWithPicture(mTabs.TabPages[7], "התאמה זוגית:");

                    mTabs.TabPages[8].Show();
                    wFileWriter.TypeBodyWithPicture(mTabs.TabPages[8], "הצלחה בלימודים / קשב וריכוז");

                    tabControl1.TabPages[0].Show();
                    mTabs.TabPages[0].Show();
                    wFileWriter.FinishDoc();

                    //DialogResult dlgRes = MessageBox.Show("האם לפתוח את הקובץ?", "הייצוא הסתיים", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    //if (dlgRes == DialogResult.OK)
                    //{
                    //    wFileWriter.OpenWordFile();
                    //}
                    #endregion

                    #region Remove Blind Fold
                    label29.Visible = true;
                    label98.Visible = true;
                    txtUnique.Visible = true;
                    txtParentsPresent.Visible = true;
                    picUnique.Visible = vis1;
                    picParentsPresent.Visible = vis2;
                    //grpMoreInfo.Visible = true;

                    tabPage1.BackgroundImage = Omega.Properties.Resources.backgrond;
                    #endregion
                    break;
            }


        }

        private void הכןסגנונותדוחToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Omega.Reports.frmReportCreationcs frmRepCrt = new frmReportCreationcs();
            frmRepCrt.Show();
        }

        private void צורדוחמסגנוןToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if ((txtPrivateName.Text == "") && (txtFamilyName.Text == "") || (txtNum1.Text == ""))
            {
                MessageBox.Show("אין חישוב עבור אדם - לכן לא יווצר דוח", "חוסר מידע בתכנה", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            ReportDataProvider.Instance.RefreshReportList();
            Omega.Objects.UserInfo currentuser = new Omega.Objects.UserInfo();
            bool collectres = currentuser.ProgData2UserInfo(this);
            ReportDataProvider.Instance.Initialize(currentuser);//CollectCurrentUserInfo()

            //frmPrintSpecialReport prntSpecial = new frmPrintSpecialReport();
            //prntSpecial.Show();

            SetAllCheckBox(false);


            clbReports.Items.Clear();

            foreach (string str in ReportDataProvider.Instance.ReportTemplets)
            {
                // Convert a string to utf-8 bytes.
                byte[] utf8Bytes = System.Text.Encoding.UTF8.GetBytes(str);

                // Convert utf-8 bytes to a string.
                string strUft8 = System.Text.Encoding.UTF8.GetString(utf8Bytes);
                clbReports.Items.Add(System.IO.Path.GetFileNameWithoutExtension(strUft8));
            }

            //rbtnFromApp.Checked = true;
            //rbtnFromStyle.Checked = false;
            //grpPart.Enabled = rbtnFromApp.Checked;
            //grpStyle.Enabled = rbtnFromStyle.Checked;

            grpPrint.Visible = true;
            grpPrint.BringToFront();
        }

        private void הכןערכיצאקרותומספריםToolStripMenuItem_Click(object sender, EventArgs e)
        {
            frmReoprtDef frmDefs = new frmReoprtDef(this);

            frmDefs.Show();
        }

        #region for reports

        private async void btnPrintFromXML_Click(object sender, EventArgs e)
        {
            mlog.Info("Try to check registered Office Word object");
            bool isMSOfficeOK = Reports.ReportDataProvider.Instance.CheckMSOffice();

            if (isMSOfficeOK)
            {
                mlog.Info("Office word object successfully registered");

            }

            mlog.Info("Prepare reports");
            InitSpinnerUC(true);

            foreach (object cic in clbReports.CheckedItems)
            {
                string sRepStyleName = cic.ToString();

                mlog.InfoFormat("Prepare [{0}] report", sRepStyleName);

                string sMsgText = string.Empty;
                string sMsgCaption = string.Empty;

                mlog.Info("Selecting file location");
                SaveFileDialog sfd = new SaveFileDialog();
                sfd.Title = "בחר מיקום הדוח";
                sfd.InitialDirectory = System.Environment.SpecialFolder.Desktop.ToString();

                if (isMSOfficeOK == true)
                {
                    const string shorterDocName = "-מקוצר";
                    string shorter = string.Empty;

                    // filter with MSO word option
                    sfd.Filter = "Microsoft Word (2003 Format)(*.doc)|*.doc|HTML Report (HTM Format)(*.html)|*.html";

                    if (cbProRep.Checked) shorter = shorterDocName;

                    sfd.FileName = "" + txtPrivateName.Text.Trim() + " " + txtFamilyName.Text.Trim() + " - " + sRepStyleName + shorter + ".doc";
                    mlog.Info("Selected filename is: " + sfd.FileName);

                }
                else
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
                    sfd.FileName = "" + txtPrivateName.Text.Trim() + " " + txtFamilyName.Text.Trim() + " - " + sRepStyleName + ".html";
                    mlog.Info("Selected filename is: " + sfd.FileName);

                    MessageBox.Show(sMsgText, sMsgCaption, MessageBoxButtons.OK, MessageBoxIcon.Error);

                }                

                mPro = ReportDataProvider.Instance.Pro = cbProRep.Checked;
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
                            string mRepStyle2Create = string.Empty;

                            XmlDocument xmlReport;

                            switch (ext)
                            {
                                case ".doc":
                                    {
                                        mRepStyle2Create = "doc";
                                        if (isMSOfficeOK == true)
                                        {
                                            #region Microsoft Office Word

                                            this.Cursor = Cursors.WaitCursor;
                                            xmlReport = ReportDataProvider.Instance.GetXmlDocByName(sRepStyleName);
                                            ReportDataProvider.Instance.mMainForm = this;

                                            mlog.DebugFormat("Creating WordWriter object: [ApplicationMainDir: {0}], [Document Path: {1}], [FirstName: {2}], [LastName: {3}]",
                                                Path.GetDirectoryName(System.Windows.Forms.Application.ExecutablePath), sfd.FileName, txtPrivateName.Text, txtFamilyName.Text);


                                            WordWriter wFileWriter = new WordWriter(Path.GetDirectoryName(System.Windows.Forms.Application.ExecutablePath), sfd.FileName, txtPrivateName.Text, txtFamilyName.Text);
                                            spinner.Visible = true;
                                            
                                            await Task.Run(() =>
                                            {
                                                wFileWriter.CreateDocument();

                                                wFileWriter.TypeXml2Doc(xmlReport);

                                                wFileWriter.FinishDoc();

                                            });
                                            spinner.Visible = false;
                                            RemoveSpinnerUC();

                                            this.Cursor = Cursors.Default;

                                            #endregion Microsoft Office Word

                                        }
                                        break;

                                    }

                            }
                            break;

                        }


                        //ReportDataProvider.Instance.PrintReportFromXMLAsync(sRepStyleName.ToString());


                        //ReportDataProvider.Instance.mMainForm.Show();
                        //ReportDataProvider.Instance.mMainForm.Focus();

                        //await ReportDataProvider.Instance.PrintReportFromXMLAsync(sRepStyleName.ToString());

                }

                MessageBox.Show("הדוחו\"ת שבחרת מהרשימה, נוצרו בהצלחה", "", MessageBoxButtons.OK, MessageBoxIcon.Information);

                grpPrint.Visible = false;

            }

        }

        // **********

        private void RemoveSpinnerUC()
        {
            mlog.Info("Try to remove spinner from form controls");

            if (grpStyle.Controls.ContainsKey(ProcessSpinnerName))
            {
                grpStyle.Controls.RemoveByKey(ProcessSpinnerName);
                mlog.Info("Spinner removed");
            }
            else
            {
                mlog.Info("Spinner not found");

            }

        }

        // **********

        private void InitSpinnerUC(bool initState = false)
        {
            mlog.Info("Initializing Spinner controls");
            spinner = new ucSpinner(false, "יוצר דו\"ח...")
            {
                Top = 185,
                Left = 3,
                Name = ProcessSpinnerName,
                Visible = initState,

            };

            mlog.Info("Add spinner to group");
            grpStyle.Controls.Add(spinner);

        }

        // **********

        private void btnPrintRegularReport_Click(object sender, EventArgs e)
        {
            ChangePrintValues();

            ReportDataProvider.Instance.mMainForm.Show();
            ReportDataProvider.Instance.mMainForm.Focus();

            ReportDataProvider.Instance.PrintReport();

            ReportDataProvider.Instance.Set2PrintAll();
            SetAllCheckBox(true);

            grpPrint.Visible = false;
        }

        private void ChangePrintValues()
        {
            //ReportDataProvider.Instance.PrntMain = cbP1.Checked;
            //ReportDataProvider.Instance.PrntLifeCycles = cbP2.Checked;
            //ReportDataProvider.Instance.PrntPitagoras = cbP3.Checked;
            //ReportDataProvider.Instance.PrntIntansiveMap = cbP4.Checked;
            //ReportDataProvider.Instance.PrntCombinedMap = cbComb.Checked;
            //ReportDataProvider.Instance.PrntChakraOpening = cbP5.Checked;
            //ReportDataProvider.Instance.PrntCoupleMatch = cbP6.Checked;
            //ReportDataProvider.Instance.PrntBsnnsMulti = cbP72.Checked;
            //ReportDataProvider.Instance.PrntBsnnsPersonal = cbP71.Checked;
            //ReportDataProvider.Instance.PrntLearnSccss = cbP8.Checked;
        }
        private void SetAllCheckBox(bool val)
        {
            //cbP1.Checked = val;
            //cbP2.Checked = val;
            //cbP3.Checked = val;
            //cbP4.Checked = val;
            //cbP5.Checked = val;
            //cbP6.Checked = val;
            //cbP71.Checked = val;
            //cbP72.Checked = val;
            //cbP8.Checked = val;
            //cbComb.Checked = val;
            cbProRep.Checked = val;
            cbRemoveChlng.Checked = val;

        }

        private void rbtnFromApp_CheckedChanged(object sender, EventArgs e)
        {
            //if (rbtnFromApp.Checked == true)
            //{
            //    grpPart.Enabled = true;
            //    grpStyle.Enabled = false;
            //}
            //else
            //{
            //    grpPart.Enabled = false;
            //    grpStyle.Enabled = true;
            //}
        }

        private void btnBack_Click(object sender, EventArgs e)
        {
            grpPrint.Visible = false;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            SetAllCheckBox(true);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            SetAllCheckBox(false);
        }
        #endregion
        #endregion reports

        private void החלףמסמךבסיסToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Title = "בחר קובץ מקור חדש";
            ofd.Filter = "2003 Word File (*.doc)|*.doc|2007 Word File(*.docx)|*.docx|Word 2003 Templet(*.dot)|*.dot|Word 2007 Templet(*.dotx)|*.dotx";
            ofd.Multiselect = false;
            ofd.InitialDirectory = @"C:\";

            DialogResult ofdRes = ofd.ShowDialog();
            if ((ofdRes == DialogResult.Cancel) || (ofdRes == DialogResult.Abort))
            {
                return;
            }
            if (ofdRes == DialogResult.OK)
            {
                string NewFile = ofd.FileName;
                string NewFileExt = Path.GetExtension(NewFile);
                string NewFileDir = Path.GetDirectoryName(NewFile);

                if (NewFileExt.ToLower() == "dot")
                {
                    FileInfo fOldTemplet = new FileInfo(mApplicationMainDir + "\\Templets\\ThisTemplet.dot");
                    fOldTemplet.Delete();

                    FileInfo fNewTemplet = new FileInfo(NewFile);
                    fNewTemplet.CopyTo(mApplicationMainDir + "\\Templets\\ThisTemplet.dot");
                }
                else
                {
                    WordWriter wwTmpletHandler = new WordWriter();

                    wwTmpletHandler.ReplcaeTempletFile(NewFile, mApplicationMainDir + "\\Templets\\ThisTemplet.dot");



                }
            }

        }

        private void ייצואגיבויToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Title = "ייצוא בסיס נתונים";
            sfd.Filter = "CSV (*.csv)|*.csv";
            sfd.FileName = "ChakraNumerology_ExportedDB_" + DateTime.Now.Year.ToString() + "_" + DateTime.Now.Year.ToString() + "_" + DateTime.Now.Month.ToString() + "_" + DateTime.Now.Day.ToString() + ".csv";
            sfd.InitialDirectory = System.Environment.SpecialFolder.DesktopDirectory.ToString();

            DialogResult dlgres = sfd.ShowDialog();
            if (dlgres == DialogResult.Abort || dlgres == DialogResult.Cancel)
            {
                return;
            }
            if (dlgres == DialogResult.OK)
            {
                //ExportData2XML(sfd.FileName);
                MessageBox.Show("כרגע לא ניתן לבצע את הפעולה");
            }
        }

        private void ייבואToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenFileDialog sfd = new OpenFileDialog();
            sfd.Title = "ייבוא אל בסיס נתונים";
            sfd.Filter = "CSV (*.csv)|*.csv";
            sfd.InitialDirectory = System.Environment.GetFolderPath(System.Environment.SpecialFolder.DesktopDirectory);
            sfd.Multiselect = false;

            DialogResult dlgres = sfd.ShowDialog();
            if (dlgres == DialogResult.Abort || dlgres == DialogResult.Cancel)
            {
                return;
            }

            string sInPath = "";
            if (dlgres == DialogResult.OK)
            {
                sInPath = sfd.FileName;
            }

            FileStream inStream = new FileStream(sInPath, FileMode.Open);
            StreamReader inReader = new StreamReader(inStream, Encoding.UTF8);

            string line = inReader.ReadLine(); // header
            UserInfo user = new UserInfo();
            List<UserInfo> AllNewUsers = new List<UserInfo>();

            while (inReader.EndOfStream == false)
            {
                line = inReader.ReadLine();
                string[] splt = line.Split(",".ToCharArray()[0]);

                user = new UserInfo();

                if (splt[1] != "NULL") user.mFirstName = splt[1];
                if (splt[2] != "NULL") user.mLastName = splt[2];
                if (splt[3] != "NULL") user.mFatherName = splt[3];
                if (splt[4] != "NULL") user.mMotherName = splt[4];
                DateTime bd = new DateTime(1900, 1, 1);
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
                if (splt[6] != "NULL") user.mCity = splt[6];
                if (splt[7] != "NULL") user.mStreet = splt[7];
                if (splt[8] != "NULL") user.mBuildingNum = Convert.ToInt32(splt[8]);
                if (splt[9] != "NULL") user.mAppNum = Convert.ToInt32(splt[9]);
                if (splt[10] != "NULL") user.mEMail = splt[10];
                if (splt[11] != "NULL") user.mPhones = splt[11];

                if (XmlDBHandler.Instance.isUserInXmlDB(user) == false)
                {
                    AllNewUsers.Add(user);
                }
            } //while (inReader.EndOfStream == false)

            bool res = XmlDBHandler.Instance.AddUserToXmlDB(AllNewUsers);

            inReader.Close();
            inStream.Close();

            gvDBview.Rows.Clear();
            XmlDBHandler.Instance.XmlDB2DataGridView(ref gvDBview);
        }
        #endregion //Menu Strip

        #region Proofing
        private bool InitTestPassword()
        {
            bool res = true;

            string FirstKey = "";
            string SecondKey = "";
            FileInfo ProofInfo = new FileInfo(path2ProffongFile);

            //DBDataSetTableAdapters.PSSTableAdapter AdapPSS = new Omega.DBDataSetTableAdapters.PSSTableAdapter();

            //if (omegaDataSet.PSS.Rows.Count == 0) // First Time Program was run
            if (ProofInfo.Exists == false) // First Time Program was run
            {
                #region Set View For Password Mode Only
                groupPSSWRD.Visible = true;
                tcInput.Visible = false;
                mTabs.Visible = false;
                #endregion //Set View For Password Mode Only

                FirstKey = Calculator.Proof_CreateFirstKey();
                txtFirstKey.Text = FirstKey;
                txtSecondKey.Text = "";
                // wait for user.... then OK will be presse

                res = false;
            }
            else // Not The First Use - might be moved to another PC
            {
                //DBDataSet.PSSRow pRow = omegaDataSet.PSS.Rows[omegaDataSet.PSS.Rows.Count - 1] as DBDataSet.PSSRow;
                //FirstKey = pRow.ItemArray[pRow.ItemArray.Length - 2].ToString();
                //SecondKey = pRow.ItemArray[pRow.ItemArray.Length - 1].ToString();

                FileStream fsProof = new FileStream(path2ProffongFile, FileMode.OpenOrCreate);
                StreamReader ProofReader = new StreamReader(fsProof);

                FirstKey = ProofReader.ReadLine();
                SecondKey = ProofReader.ReadLine();

                bool resProof = Calculator.Proof_Test_First_Key_With_PC(FirstKey);
                resProof &= Calculator.Proof_TestKeys(FirstKey, SecondKey);

                if (resProof)
                {
                    #region Set View For Password Mode Only
                    groupPSSWRD.Visible = false;
                    tcInput.Visible = true;
                    mTabs.Visible = true;
                    #endregion //Set View For Password Mode Only
                    res = true;
                }
                else
                {
                    #region Set View For Password Mode Only
                    groupPSSWRD.Visible = true;
                    tcInput.Visible = false;
                    mTabs.Visible = false;
                    #endregion //Set View For Password Mode Only
                    MessageBox.Show("הסיסמאות אינן תואמות את המחשב - או אחת את השניה", "שגיאת סיסמא", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    res = false;
                }

                fsProof.Close();
                ProofReader.Close();

                if (res == false)
                {
                    ProofInfo.Delete();
                }
            }

            return res;
        }

        private void cmdCreate_Click(object sender, EventArgs e)
        {
            txtFirstKey.Text = Calculator.Proof_CreateFirstKey();
        }

        private void cmdProofOK_Click(object sender, EventArgs e)
        {
            //DBDataSet.PSSRow pRow = omegaDataSet.PSS.NewPSSRow() as DBDataSet.PSSRow;
            //DBDataSetTableAdapters.PSSTableAdapter AdapPSS = new Omega.DBDataSetTableAdapters.PSSTableAdapter();

            if (Calculator.Proof_TestKeys(txtFirstKey.Text, txtSecondKey.Text))
            {
                //pRow.Time = "OK";
                //pRow.num = txtFirstKey.Text;
                //pRow.psw = txtSecondKey.Text;

                //omegaDataSet.PSS.AddPSSRow(pRow);
                //AdapPSS.Adapter.Update(omegaDataSet.PSS.DataSet);
                //omegaDataSet.PSS.AcceptChanges();

                #region Set View For Password Mode Only
                groupPSSWRD.Visible = false;
                tcInput.Visible = true;
                mTabs.Visible = true;
                #endregion //Set View For Password Mode Only

                //this.clientsTableAdapter.Fill(this.omegaDataSet.Clients);

                FileInfo ProofInfo = new FileInfo(path2ProffongFile);
                FileStream fsProof = new FileStream(path2ProffongFile, FileMode.OpenOrCreate);
                StreamWriter swProof = new StreamWriter(fsProof);

                swProof.WriteLine(txtFirstKey.Text);
                swProof.WriteLine(txtSecondKey.Text);

                swProof.Close();
                fsProof.Close();

                XmlDBHandler.Instance.XmlDB2DataGridView(ref gvDBview);
            }
            else
            {
                MessageBox.Show("הסיסמא שגויה", "טעות!!!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        #endregion // proofing

        #region DB_Handle
        /// <summary> Loading user data to the aplication main form
        /// Loading user data to the aplication main form
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void cmdLoad_Click(object sender, EventArgs e)
        {
            //if (gvDBview.SelectedRows.Count == 0)
            //{
            //    MessageBox.Show("צריך לבחור שורה שלמה", "שגיאה בסימון שורה", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //    return;
            //}

            if (tcInput.TabPages[TabPageLifeCycle2] != null)
            {
                tcInput.TabPages.RemoveByKey(TabPageLifeCycle2);

            }

            if ((gvDBview.SelectedRows.Count == 0) && (gvDBview.SelectedCells.Count == 0))
            {
                return;
            }

            ClearForm();

            checkBox1.Checked = false;
            checkBox1.CheckState = CheckState.Unchecked;
            checkBox1.Update();
            checkBox1_CheckedChanged(null, null);


            //int RowNum = gvDBview.SelectedRows[0].Index;
            int RowNum = gvDBview.SelectedCells[0].RowIndex;
            //DataGridViewRow CurClient = gvDBview.SelectedRows[0];
            DataGridViewRow CurClient = gvDBview.Rows[RowNum];
            /*
            DataGridViewRow gvClientRow = gvDBview.Rows[RowNum];

            DBDataSet.ClientsRow CurClient = omegaDataSet.Clients.NewRow() as DBDataSet.ClientsRow;
            CurClient = omegaDataSet.Clients.Rows[RowNum] as DBDataSet.ClientsRow;

            txtPrivateName.Text = CurClient.PrivateName;
            txtFamilyName.Text = CurClient.LastName;
            txtMotherName.Text = CurClient.MotherName;
            txtFatherName.Text = CurClient.FatherName;
            DateTimePickerFrom.Value = CurClient.B_Date;
            txtCity.Text = CurClient.City;
            txtStreet.Text = CurClient.Street;
            txtBiuldingNum.Text = CurClient.BuildingNum.ToString();
            txtAppNum.Text = CurClient.AppNum.ToString();
            txtEMail.Text = CurClient.EMail.ToString();
            txtPhones.Text = CurClient.Phones.ToString();
            */
            txtClientId.Text = CurClient.Cells[0].Value.ToString();
            txtPrivateName.Text = CurClient.Cells[1].Value.ToString();
            txtFamilyName.Text = CurClient.Cells[2].Value.ToString();
            txtMotherName.Text = CurClient.Cells[4].Value.ToString();
            txtFatherName.Text = CurClient.Cells[3].Value.ToString();
            DateTimePickerFrom.Value = DateTime.Parse(CurClient.Cells[5].Value.ToString());
            txtCity.Text = CurClient.Cells[6].Value.ToString();
            txtStreet.Text = CurClient.Cells[7].Value.ToString();
            txtBiuldingNum.Text = CurClient.Cells[8].Value.ToString();
            txtAppNum.Text = CurClient.Cells[9].Value.ToString();
            txtEMail.Text = CurClient.Cells[10].Value.ToString();
            txtPhones.Text = CurClient.Cells[11].Value.ToString();
            txtApplication.Text = CurClient.Cells[12].Value.ToString();

            EnumProvider.Sex sex = EnumProvider.Sex.Male;
            EnumProvider.PassedRectification passedrect = EnumProvider.PassedRectification.NotPassed;
            EnumProvider.ReachedMaster reachmaster = EnumProvider.ReachedMaster.No;

            if (BLL.AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
            {
                sex = EnumProvider.Instance.GetSexEnumFromString(CurClient.Cells[13].Value.ToString());
                passedrect = EnumProvider.Instance.GetPassedRectificationEnumFromString(CurClient.Cells[14].Value.ToString());
                reachmaster = EnumProvider.Instance.GGetIsMasterEnumFromString(CurClient.Cells[15].Value.ToString());
            }
            if (BLL.AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
            {
                string s = CurClient.Cells[13].Value.ToString(); // SEX
                if (s.ToLower() == EnumProvider.Sex.Female.ToString().ToLower())
                {
                    sex = EnumProvider.Sex.Female;
                }
                if (s.ToLower() == EnumProvider.Sex.Male.ToString().ToLower())
                {
                    sex = EnumProvider.Sex.Male;
                }

                s = CurClient.Cells[14].Value.ToString(); // RECTIFICATION
                if (s.ToLower() == EnumProvider.PassedRectification.NotPassed.ToString().ToLower())
                {
                    passedrect = EnumProvider.PassedRectification.NotPassed;
                }
                if (s.ToLower() == EnumProvider.PassedRectification.Passed.ToString().ToLower())
                {
                    passedrect = EnumProvider.PassedRectification.Passed;
                }

                s = CurClient.Cells[15].Value.ToString(); // MASTER
                if (s.ToLower() == EnumProvider.ReachedMaster.No.ToString().ToLower())
                {
                    reachmaster = EnumProvider.ReachedMaster.No;
                }
                if (s.ToLower() == EnumProvider.ReachedMaster.Yes.ToString().ToLower())
                {
                    reachmaster = EnumProvider.ReachedMaster.Yes;
                }
            }

            if (sex == EnumProvider.Sex.Male)
            {
                cmbSexSelect.Text = cmbSexSelect.Items[0].ToString();
                //cmbSexSelect.Select(0, 1);
                cmbSexSelect.SelectedIndex = cmbSexSelect.FindStringExact(cmbSexSelect.Text);
            }
            else
            {
                cmbSexSelect.Text = cmbSexSelect.Items[1].ToString();
                //cmbSexSelect.Select(0, 1);
                cmbSexSelect.SelectedIndex = cmbSexSelect.FindStringExact(cmbSexSelect.Text);
            }

            if (passedrect == EnumProvider.PassedRectification.NotPassed)
            {
                cbMainCorrectionDone.Checked = false;
                cmbSelfFix.SelectedIndex = cmbSelfFix.FindStringExact(cmbSelfFix.Items[0].ToString());
            }
            else
            {
                cbMainCorrectionDone.Checked = true;
                cmbSelfFix.SelectedIndex = cmbSelfFix.FindStringExact(cmbSelfFix.Items[1].ToString());
            }

            //if (reachmaster == EnumProvider.ReachedMaster.No)
            //{
            //    cbPersonMaster.Checked = true;
            //}
            //else
            //{
            //    cbPersonMaster.Checked = false;
            //}

            tcInput.SelectTab("ClientDataPage");

        }

        /// <summary> Saving user info inthe DB
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void cmdSave_Click(object sender, EventArgs e)
        {
            #region MS DB
            //if (isPersonInDB() == false)
            //{
            //    SaveData2DB();
            //}
            #endregion MS DB

            #region XmlDB
            UserInfo curUser = DGVrow2UserInfo();

            if (XmlDBHandler.Instance.isUserInXmlDB(curUser) == false)
            {
                XmlDBHandler.Instance.AddUserToXmlDB(curUser);
                gvDBview.Rows.Clear();
                XmlDBHandler.Instance.XmlDB2DataGridView(ref gvDBview);
            }
            #endregion XmlDB
        }

        private void btnDelUser_Click(object sender, EventArgs e)
        {
            DeleteRecordFromXMLDB(ref gvDBview);
        }

        private void DeleteRecordFromXMLDB(ref DataGridView grd)
        {
            //if ((gvDBview.SelectedRows.Count == 0) && (gvDBview.SelectedCells.Count == 0))
            if ((grd.SelectedRows.Count == 0) && (grd.SelectedCells.Count == 0))
            {
                return;
            }

            string mssg = "", cptn = "";
            if (BLL.AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
            {
                mssg = "האם את/ה משוכנע/ת לגבי מחיקת רשומה זו?";
                cptn = "תכנת נומרולוגיית הצ'אקרות הדינאמיות - מחיקת רשומה";
            }
            if (BLL.AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
            {
                mssg = "Are you sure about deleting this record?";
                cptn = "Dynamic Chakra Numerology App - Record Delete";
            }
            DialogResult dlgres = MessageBox.Show(mssg, cptn, MessageBoxButtons.YesNo, MessageBoxIcon.Stop);

            if (dlgres == DialogResult.No) return;

            foreach (DataGridViewCell curCell in grd.SelectedCells)
            {
                #region 
                int RowNum = curCell.RowIndex;
                DataGridViewRow CurClient = grd.Rows[RowNum];

                UserInfo user = new UserInfo();

                user.mId = Convert.ToInt16(CurClient.Cells[0].Value);
                user.mFirstName = CurClient.Cells[1].Value.ToString();
                user.mLastName = CurClient.Cells[2].Value.ToString();
                user.mMotherName = CurClient.Cells[4].Value.ToString();
                user.mFatherName = CurClient.Cells[3].Value.ToString();
                user.mB_Date = DateTime.Parse(CurClient.Cells[5].Value.ToString());
                user.mCity = CurClient.Cells[6].Value.ToString();
                user.mStreet = CurClient.Cells[7].Value.ToString();
                user.mBuildingNum = Convert.ToInt16(CurClient.Cells[8].Value.ToString());
                user.mAppNum = Convert.ToInt16(CurClient.Cells[9].Value.ToString());
                user.mEMail = CurClient.Cells[10].Value.ToString();
                user.mPhones = CurClient.Cells[11].Value.ToString();
                user.mApplication = CurClient.Cells[12].Value.ToString();
                if (BLL.AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                {
                    user.mSex = EnumProvider.Instance.GetSexEnumFromString(CurClient.Cells[13].Value.ToString());
                    user.mPassedRect = EnumProvider.Instance.GetPassedRectificationEnumFromString(CurClient.Cells[14].Value.ToString());
                    user.mReachMaster = EnumProvider.Instance.GGetIsMasterEnumFromString(CurClient.Cells[15].Value.ToString());
                }
                if (BLL.AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                {
                    string s = CurClient.Cells[13].Value.ToString(); // SEX
                    if (s.ToLower() == EnumProvider.Sex.Female.ToString().ToLower())
                    {
                        user.mSex = EnumProvider.Sex.Female;
                    }
                    if (s.ToLower() == EnumProvider.Sex.Male.ToString().ToLower())
                    {
                        user.mSex = EnumProvider.Sex.Male;
                    }

                    s = CurClient.Cells[14].Value.ToString(); // RECTIFICATION
                    if (s.ToLower() == EnumProvider.PassedRectification.NotPassed.ToString().ToLower())
                    {
                        user.mPassedRect = EnumProvider.PassedRectification.NotPassed;
                    }
                    if (s.ToLower() == EnumProvider.PassedRectification.Passed.ToString().ToLower())
                    {
                        user.mPassedRect = EnumProvider.PassedRectification.Passed;
                    }

                    s = CurClient.Cells[15].Value.ToString(); // MASTER
                    if (s.ToLower() == EnumProvider.ReachedMaster.No.ToString().ToLower())
                    {
                        user.mReachMaster = EnumProvider.ReachedMaster.No;
                    }
                    if (s.ToLower() == EnumProvider.ReachedMaster.Yes.ToString().ToLower())
                    {
                        user.mReachMaster = EnumProvider.ReachedMaster.Yes;
                    }
                }


                XmlDBHandler.Instance.RemoveUserFromDB(user);
                #endregion
            }
            grd.Rows.Clear();
            XmlDBHandler.Instance.XmlDB2DataGridView(ref grd);
            // refresh main list anyway
            XmlDBHandler.Instance.XmlDB2DataGridView(ref gvDBview);
        }

        #region Ms DB (DB.mdb)

        //private bool isPersonInDB()
        //{
        //    bool res = false;
        //    DBDataSet.ClientsRow curDataRow;

        //    for (int i = 0; i < omegaDataSet.Clients.Rows.Count; i++)
        //    {
        //        curDataRow = omegaDataSet.Clients.Rows[i] as DBDataSet.ClientsRow;
        //        res = (curDataRow.PrivateName == txtPrivateName.Text);
        //        res &= (curDataRow.LastName == txtFamilyName.Text);
        //        res &= (curDataRow.FatherName == txtFatherName.Text);
        //        res &= (curDataRow.MotherName == txtMotherName.Text);
        //        res &= (curDataRow.City == txtCity.Text);
        //        res &= (curDataRow.Street == txtStreet.Text);
        //        res &= (curDataRow.AppNum == Convert.ToDouble(txtAppNum.Text));
        //        res &= (curDataRow.BuildingNum == Convert.ToDouble(txtBiuldingNum.Text));
        //        res &= (curDataRow.B_Date.Subtract(DateTimePickerFrom.Value).Days == 0);
        //        res &= (curDataRow.EMail == txtEMail.Text);
        //        res &= (curDataRow.Phones == txtPhones.Text);
        //        if (res == true)
        //        {
        //            return res;
        //        }
        //    }
        //    return res;
        //}

        //private void SaveData2DB()
        //{
        //    DBDataSet.ClientsRow CurClient = omegaDataSet.Clients.NewRow() as DBDataSet.ClientsRow;

        //    DBDataSetTableAdapters.ClientsTableAdapter AdapClients = new Omega.DBDataSetTableAdapters.ClientsTableAdapter();

        //    BindingSource bSource = new BindingSource();
        //    bSource.DataSource = omegaDataSet.Clients as DataTable;
        //    gvDBview.DataSource = bSource.DataSource;

        //    #region Data Collection
        //    CurClient.PrivateName = txtPrivateName.Text.Trim();
        //    CurClient.LastName = txtFamilyName.Text.Trim();
        //    CurClient.MotherName = txtMotherName.Text.Trim();
        //    CurClient.FatherName = txtFatherName.Text.Trim();
        //    CurClient.ClientID = omegaDataSet.Clients.Rows.Count + 1;
        //    CurClient.B_Date = DateTimePickerFrom.Value;
        //    CurClient.City = txtCity.Text.Trim();
        //    CurClient.Street = txtStreet.Text.Trim();
        //    CurClient.BuildingNum = Convert.ToDouble(txtBiuldingNum.Text.Trim());
        //    CurClient.AppNum = Convert.ToDouble(txtAppNum.Text.Trim());

        //    CurClient.EMail = txtEMail.Text.Trim();
        //    CurClient.Phones = txtPhones.Text.Trim();
        //    #endregion

        //    //omegaDataSet.Clients.AddClientsRow(CurClient);
        //    omegaDataSet.Clients.Rows.Add(CurClient);
        //    AdapClients.Adapter.Update(omegaDataSet.Clients.DataSet);
        //    omegaDataSet.Clients.AcceptChanges();

        //    gvDBview.Refresh();
        //    gvDBview.Update();
        //}

        //private void ExportData2XML(string Path2ExportedXML)
        //{
        //    XmlTextWriter mXmlTextWriter = new XmlTextWriter(Path2ExportedXML, Encoding.UTF8);
        //    mXmlTextWriter.Formatting = Formatting.Indented;

        //    mXmlTextWriter.WriteStartDocument();

        //    mXmlTextWriter.WriteStartElement("DB_Records");

        //    for (int i = 0; i < omegaDataSet.Clients.Rows.Count; i++)
        //    {
        //        DBDataSet.ClientsRow CurClient = omegaDataSet.Clients.NewRow() as DBDataSet.ClientsRow;
        //        CurClient = omegaDataSet.Clients.Rows[i] as DBDataSet.ClientsRow;

        //        mXmlTextWriter.WriteStartElement("Client_" + i.ToString());
        //            mXmlTextWriter.WriteElementString("PrivateName", CurClient.PrivateName.Trim());
        //            mXmlTextWriter.WriteElementString("LastName", CurClient.LastName.Trim());
        //            mXmlTextWriter.WriteElementString("B_Date", CurClient.B_Date.Year.ToString() + "-" + CurClient.B_Date.Month.ToString() + "-" + CurClient.B_Date.Day.ToString());
        //            mXmlTextWriter.WriteElementString("FatherName", CurClient.FatherName.Trim());
        //            mXmlTextWriter.WriteElementString("MotherName", CurClient.MotherName.Trim());
        //            mXmlTextWriter.WriteElementString("City", CurClient.City.Trim());
        //            mXmlTextWriter.WriteElementString("Street", CurClient.Street.Trim());
        //            mXmlTextWriter.WriteElementString("BuildingNum", CurClient.BuildingNum.ToString().Trim());
        //            mXmlTextWriter.WriteElementString("AppNum", CurClient.AppNum.ToString().Trim());
        //        mXmlTextWriter.WriteEndElement();
        //    }

        //    mXmlTextWriter.Flush();

        //    mXmlTextWriter.WriteEndDocument();
        //    mXmlTextWriter.Close();

        //}

        //private void ImportDataFromXML(string Path2ExportedXML)
        //{
        //    XmlDocument xmlFile = new XmlDocument();

        //    xmlFile.Load(Path2ExportedXML);

        //    XmlNodeList root = xmlFile.GetElementsByTagName("DB_Records");

        //    foreach (XmlNode curClientNode in root[0].ChildNodes)
        //    {
        //        foreach (XmlNode data in curClientNode.ChildNodes)
        //        {
        //            string txt = data.InnerText.Trim();
        //            if (data.Name == "PrivateName")
        //            {
        //                txtPrivateName.Text = txt;
        //            }
        //            if (data.Name == "LastName")
        //            {
        //                txtFamilyName.Text = txt;
        //            }
        //            if (data.Name == "B_Date")
        //            {
        //                string[] spl = txt.Split("-".ToCharArray()[0]);
        //                DateTimePickerFrom.Value = new DateTime(Convert.ToInt32(spl[0]), Convert.ToInt32(spl[1]), Convert.ToInt32(spl[2]), 0, 0, 0);
        //            }
        //            if (data.Name == "FatherName")
        //            {
        //                txtFatherName.Text = txt;
        //            }
        //            if (data.Name == "MotherName")
        //            {
        //                txtMotherName.Text = txt;
        //            }
        //            if (data.Name == "City")
        //            {
        //                txtCity.Text = txt;
        //            }
        //            if (data.Name == "Street")
        //            {
        //                txtStreet.Text = txt;
        //            }
        //            if (data.Name == "BuildingNum")
        //            {
        //                txtBiuldingNum.Text = txt;
        //            }
        //            if (data.Name == "AppNum")
        //            {
        //                txtAppNum.Text = txt;
        //            }
        //        }

        //        runCalc();
        //    }

        //    ClearForm();
        //}

        #endregion Ms DB (DB.mdb)

        #region XML DB

        // handeled by the XmlDBHAndler
        private UserInfo DGVrow2UserInfo()
        {
            DataGridViewRow row = gvDBview.Rows[gvDBview.SelectedCells[0].RowIndex];
            UserInfo ui = new UserInfo();

            try
            {
                ui.mFirstName = row.Cells[1].Value.ToString();
                ui.mLastName = row.Cells[2].Value.ToString();
                ui.mFatherName = row.Cells[2].Value.ToString();
                ui.mMotherName = row.Cells[2].Value.ToString();
                ui.mB_Date = DateTime.Parse(row.Cells[2].Value.ToString());
                ui.mCity = row.Cells[2].Value.ToString();
                ui.mStreet = row.Cells[2].Value.ToString();
                ui.mBuildingNum = Convert.ToDouble(row.Cells[2].Value.ToString());
                ui.mAppNum = Convert.ToDouble(row.Cells[2].Value.ToString());
                ui.mEMail = row.Cells[2].Value.ToString();
                ui.mPhones = row.Cells[2].Value.ToString();
                //ui.mSex = row.Cells[2].Value.ToString();
                //ui.mPassedRect = row.Cells[2].Value.ToString();
                //ui.mApplication = row.Cells[2].Value.ToString();
            }
            catch
            {
                ui = new UserInfo();
            }

            return ui;
        }

        #endregion XML DB
        #endregion //DB_Handle

        public void runCalc()
        {
            if (GatherData() == true)
            {
                #region DB Save if needed
                if ((checkBox1.Checked == true) && (isRunningNameBalance == false))
                { // Saving will be done by this opeartion just as well
                    #region MS MDB
                    //if (isPersonInDB() == false)
                    //    SaveData2DB();
                    #endregion MS MDB
                    //TODO: Auto Save spouce data
                    #region XML DB
                    // user Info gathered @ GatherData()
                    if (XmlDBHandler.Instance.isUserInXmlDB(CurrentUserData) == false)
                    {
                        XmlDBHandler.Instance.AddUserToXmlDB(CurrentUserData);
                        txtClientId.Text = CurrentUserData.mId.ToString();
                        gvDBview.Rows.Clear();
                        XmlDBHandler.Instance.XmlDB2DataGridView(ref gvDBview);
                    }
                    #endregion XML DB
                }
                else
                {
                    // noting will be saved
                }
                #endregion

                if (cbMainCorrectionDone.Checked == true)
                {
                    cmbIsFix.SelectedIndex = 0;
                    cmbIsFix.Text = cmbIsFix.SelectedItem.ToString();
                }

                ClearOutputForm(); // Clears the Screen from other Data

                #region Life Cycles
                int[] LC = Calculator.CareteLifeCycles(mB_Date);

                SetLifeCycleValues(LC);

                #region Personal Info.
                string pY, pM, pD;
                Calculator.CalcPersonalInfo(mB_Date, DateTimePickerTo.Value, out pY, out pM, out pD);

                txtPYear.Text = txtPYearX.Text = pY.ToString();
                txtPMonth.Text = txtPMonthX.Text = pM.ToString();
                txtPDay.Text = txtPDayX.Text = pD.ToString();
                #endregion

                // Expose the extra life cycle tab
                if (!tcInput.TabPages.ContainsKey(TabPageLifeCycle2))
                {
                    tcInput.TabPages.Add(tcLifeCycles2);
                    tcInput.SelectedTab = tcLifeCycles2;

                }

                #endregion

                #region General Calculations
                NumericalCalc1();
                #endregion

                txtLuck.Text = txtXLuck.Text = txtAstroName.Text;
                txtPrivateNameCycle.Text = txtPrivateName.Text;
                txtFamilyNameCycle.Text = txtFamilyName.Text;
                DateTimePickerFromCycle.Text = DateTimePickerFrom.Text;

                txtNum8.Text = txtPName_Num.Text;
                txtNum9.Text = txtAstroName.Text;

                #region Intensive Map
                string[] IM = Calculator.CreateIntensiveMap(mFirstName, mLastName);

                if (string.Join(string.Empty, IM).Trim().Length == 0)
                {
                    ClearForm();
                    MessageBox.Show("לא ניתן ליצור מפה.", "שגיאת נתונים", MessageBoxButtons.OK, MessageBoxIcon.Error);

                    return;
                }

                SetINtensiveMapValues(IM);

                SetSumTableFromTextBoxes(mIntsMapTextArr, dgvIntsMapSum);
                #endregion

                #region Pitagoras Squares
                string[] PS = Calculator.CreatePitagorasSquares(mB_Date);

                SetPitagorasSquaresValues(PS);

                SetSumTableFromTextBoxes(mPitgMapTextArr, dgvPitMapSum);
                #endregion

                #region Chakras
                //txtCMaster.Text = Calculator.CalcMasterChkra(mFirstName + mLastName);//על
                //txtCCrown.Text = Calculator.CalcCrownChakra(mFirstName + mLastName, txtNum1.Text);//כתר
                //txtCThirdEye.Text = Calculator.CalcThirdEyeChakra(mFirstName + mLastName, txtNum2.Text);//עין שלישית 
                //txtCThrought.Text = Calculator.CalcThroughtChakra(mFirstName + mLastName, txtNum3.Text);//גרון
                //txtCHeart.Text = Calculator.CalcHeartChakra(mFirstName + mLastName, txtNum4.Text);//לב
                //txtCMikSun.Text = Calculator.CalcSunChakra(mFirstName + mLastName, txtNum5.Text);//מקלעת השמש
                //txtCSexCreat.Text = Calculator.CalcSex_CreationChakra(mFirstName + mLastName, txtNum6.Text);//מין ויצירה
                //txtCRoot.Text = Calculator.CalcRootChakra(mFirstName + mLastName, txtNum7.Text);//בסיס
                #endregion

                #region Dynaic Chkras

                DynamicChakraOpening();

                #endregion

                #region Combined Map
                string[] inIntensiveMapData = new string[9] { txtIMap1.Text, txtIMap2.Text, txtIMap3.Text, txtIMap4.Text, txtIMap5.Text, txtIMap6.Text, txtIMap7.Text, txtIMap8.Text, txtIMap9.Text };
                string[] inPiatgorasMapData = new string[9] { txtPit1.Text, txtPit2.Text, txtPit3.Text, txtPit4.Text, txtPit5.Text, txtPit6.Text, txtPit7.Text, txtPit8.Text, txtPit9.Text };
                string[] outCombinedMapData = Calculator.CalcCombinedMap(inIntensiveMapData, inPiatgorasMapData);

                txtMComb1.Text = outCombinedMapData[0].ToString();
                txtMComb2.Text = outCombinedMapData[1].ToString();
                txtMComb3.Text = outCombinedMapData[2].ToString();
                txtMComb4.Text = outCombinedMapData[3].ToString();
                txtMComb5.Text = outCombinedMapData[4].ToString();
                txtMComb6.Text = outCombinedMapData[5].ToString();
                txtMComb7.Text = outCombinedMapData[6].ToString();
                txtMComb8.Text = outCombinedMapData[7].ToString();
                txtMComb9.Text = outCombinedMapData[8].ToString();


                SetSumTableFromTextBoxes(mCombMapTextArr, dgvCombMapSum);

                dgvCombMapSum.Visible = true;

                #region Info
                lblCombInfo.Text = "";
                int[] indexSum = new int[9];
                for (int i = 0; i < 9; i++)
                {
                    string[] sTmp = outCombinedMapData[i].Split(",".ToCharArray()[0]);
                    if (sTmp[0] == "")
                    {
                        indexSum[i] = 0;

                    }
                    else
                    {
                        indexSum[i] = sTmp.Length;
                    }
                }

                List<int> max, nun;
                GetMinMax4CombMap(indexSum, out max, out nun);

                lblCombInfo.Text = "מספרי העודף הינם:  ";// +System.Environment.NewLine;
                string os = "";
                foreach (int i in max)
                {
                    os += "," + (i + 1).ToString();
                }
                lblCombInfo.Text += os.Substring(1, os.Length - 1) + System.Environment.NewLine;


                lblCombInfo.Text += "מספרי החוסר הינם:  ";// +System.Environment.NewLine;
                os = "";
                if (nun.Count > 0)
                {
                    foreach (int i in nun)
                    {
                        os += "," + (i + 1).ToString();
                    }

                    os = os.Substring(1, os.Length - 1);
                }
                lblCombInfo.Text += os;

                #endregion

                #endregion

                #region BUSNSS
                if (isRunningNameBalance == true) // איזון שם רץ  אין צורך ביותר מדי הרצות פנימיות
                {
                    SingleBusinessMatchCalc(cmbBusinessBonus.SelectedItem.ToString().Trim(), cmbSelfFix.SelectedItem.ToString().Trim());
                }
                if (isRunningNameBalance == false) // חישוב מלא  ניתן להציג את כל הפרמטרים
                {
                    RunBssnssCalc4cycles();
                }

                txtSalesStory.Text = string.Empty;

                CalcSalesMatch();

                #endregion

                #region LearnSccss_AttPrbl
                LrnSccssAttPrblm();
                #endregion

                #region Health
                int c = CalcCurrentCycle();
                DateTime dt = DateTimePickerTo.Value;
                double v = 0;

                int fromC = 0, toC = 0;
                if (isRunningNameBalance == true)// ריצה לטובת איזון שם  אין צורך ביותר מדי עבודה
                {
                    fromC = c;
                    toC = c + 1;
                }
                if (isRunningNameBalance == false) // ריצה גרילה ניתן להראות את כלל הפרמטרים
                {
                    fromC = 1;
                    toC = 5;
                }
                for (int i = fromC; i < toC; i++)
                {
                    #region MoveDate 4 diff. life-cycle
                    string[] temp = (this.Controls.Find("txt" + i.ToString() + "_1", true)[0] as TextBox).Text.Split("-".ToCharArray());
                    DateTimePickerTo.Value = mB_Date.AddYears((Convert.ToInt16(temp[0]) + Convert.ToInt16(temp[1])) / 2);
                    #endregion MoveDate 4 diff. life-cycle

                    int CurrentCycle = CalcCurrentCycle();

                    List<int> HealthPersonalValues = new List<int>();

                    #region Gather Data


                    int tmpnum = -1;
                    //כתר
                    int crown = Calculator.GetCorrectNumberFromSplitedString(txtNum1.Text.Split(Calculator.Delimiter));

                    //עין שלישית
                    int thirdeye = Calculator.GetCorrectNumberFromSplitedString(txtNum2.Text.Split(Calculator.Delimiter));

                    //מקלעת השמש
                    tmpnum = Calculator.GetCorrectNumberFromSplitedString(txtNum5.Text.Split(Calculator.Delimiter));
                    HealthPersonalValues.Add(tmpnum);

                    //שם פרטי
                    tmpnum = Calculator.GetCorrectNumberFromSplitedString(txtPName_Num.Text.Split(Calculator.Delimiter));
                    HealthPersonalValues.Add(tmpnum);

                    //מזל אסטרולוגי
                    string[] strAstr = txtAstroName.Text.Split(" ".ToCharArray()[0]);
                    string strAstr2 = strAstr[1];
                    strAstr2 = strAstr2.Replace("(", "");
                    strAstr2 = strAstr2.Replace(")", "");
                    tmpnum = Convert.ToInt16(strAstr2);
                    HealthPersonalValues.Add(tmpnum);

                    //מפה מאוזנת
                    tmpnum = -1;
                    string[] info = txtInfo1.Text.Split("\n".ToCharArray());
                    if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                    {
                        if (info[1].Trim() == "מאוזן")
                        {
                            tmpnum = 10;
                        }
                        if (info[1].Trim() == "מאוזן חלקית")
                        {
                            tmpnum = 8;
                        }
                        if (info[1].Trim() == "לא מאוזן")
                        {
                            tmpnum = 4;
                        }
                    }
                    if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                    {
                        if (info[1].Trim() == "Balanced")
                        {
                            tmpnum = 10;
                        }
                        if (info[1].Trim() == "Half Balanced")
                        {
                            tmpnum = 8;
                        }
                        if (info[1].Trim() == "Not Balanced")
                        {
                            tmpnum = 4;
                        }
                    }
                    HealthPersonalValues.Add(tmpnum);

                    //שיא
                    string nums = this.Controls.Find("txt" + CurrentCycle.ToString() + "_3", true)[0].Text;
                    tmpnum = Calculator.GetCorrectNumberFromSplitedString(nums.Split(Calculator.Delimiter));
                    HealthPersonalValues.Add(tmpnum);

                    //מחזור
                    nums = this.Controls.Find("txt" + CurrentCycle.ToString() + "_2", true)[0].Text;
                    tmpnum = Calculator.GetCorrectNumberFromSplitedString(nums.Split(Calculator.Delimiter));
                    HealthPersonalValues.Add(tmpnum);

                    //שנה אישית
                    tmpnum = Calculator.GetCorrectNumberFromSplitedString(txtPYear.Text.Split(Calculator.Delimiter));
                    HealthPersonalValues.Add(tmpnum);

                    //חודש אישי
                    tmpnum = Calculator.GetCorrectNumberFromSplitedString(txtPMonth.Text.Split(Calculator.Delimiter));
                    HealthPersonalValues.Add(tmpnum);

                    //מין ויצירה
                    tmpnum = Calculator.GetCorrectNumberFromSplitedString(txtNum6.Text.Split(Calculator.Delimiter));
                    HealthPersonalValues.Add(tmpnum);
                    #endregion

                    double HealthVal = Calculator.CalcHealth(HealthPersonalValues, CurrentCycle, crown, thirdeye);
                    HealthVal = Math.Round(HealthVal, 2);
                    txtHealthValue.Text += HealthVal.ToString();
                    if (c == i)
                    {
                        v = HealthVal;
                    }

                    if (BLL.AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                    {
                        switch (i)
                        {
                            case 1:
                                txtHealthStory.Text += "בתקופת החיים הראשונה: (";
                                break;
                            case 2:
                                txtHealthStory.Text += "בתקופת החיים השנייה: (";
                                break;
                            case 3:
                                txtHealthStory.Text += "בתקופת החיים השלישית: (";
                                break;
                            case 4:
                                txtHealthStory.Text += "בתקופת החיים הרביעית: (";
                                break;
                        }
                    }
                    if (BLL.AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                    {
                        switch (i)
                        {
                            case 1:
                                txtHealthStory.Text += "First cycle: (";
                                break;
                            case 2:
                                txtHealthStory.Text += "Second cycle: (";
                                break;
                            case 3:
                                txtHealthStory.Text += "Third cycle: (";
                                break;
                            case 4:
                                txtHealthStory.Text += "Fourth cycle: (";
                                break;
                        }
                    }
                    txtHealthStory.Text += HealthVal.ToString() + ")" + Environment.NewLine;

                    if ((HealthVal > 0) && (HealthVal <= 6))
                    {
                        txtHealthStory.Text += "בריאות חלשה מאוד.";
                    }
                    if ((HealthVal > 6.01) && (HealthVal <= 7))
                    {
                        txtHealthStory.Text += "בריאות חלשה.";
                    }
                    if ((HealthVal > 7.01) && (HealthVal <= 8))
                    {
                        txtHealthStory.Text += "בריאות ממוצעת.";
                    }
                    if ((HealthVal > 8.01) && (HealthVal <= 9))
                    {
                        txtHealthStory.Text += "בריאות טובה.";
                    }
                    if (HealthVal > 9)
                    {
                        txtHealthStory.Text += "בריאות טובה מאוד.";
                    }

                    if (HealthVal < 7)
                    {
                        txtHealthStory.Text += Environment.NewLine + "מומלץ לך לשפר בריאותך.";
                    }

                    txtHealthStory.Text += Environment.NewLine + Environment.NewLine;
                }
                txtHealthStory.Text = txtHealthStory.Text.Trim();
                txtHealthStory.Text += Environment.NewLine + Environment.NewLine + "סוגי המחלות המאפיינות את גופך הן:" + Environment.NewLine + ReportDataProvider.Instance.Reserved2OPSInfo("_מחלות_");
                txtHealthValue.Text = v.ToString();
                DateTimePickerTo.Value = dt;
                #endregion Health

                SetStylesToObjects();

                #region Heb 2 Other Lang
                switch (AppSettings.Instance.ProgramLanguage)
                {
                    case AppSettings.Language.Hebrew:
                        break;
                    case AppSettings.Language.English:
                        #region US-EN

                        #endregion
                        break;
                }
                #endregion

                textBox48.Text = ReportDataProvider.Instance.ReadFromTextFile(Reports.ReportDataProvider.Instance.ConstructFilePath2FianlySummary());

                #region Calculate Combinations

                txtCombinations.Lines = CalcCombinations();

                #endregion Calculate Combinations

            }
            else
            {
                return; // nothing was calclated
            }
        }

        #region Buttons
        private void cmdRun_Click_1(object sender, EventArgs e)
        {
            isRunningNameBalance = false;// לא מריץ איזון שם  אלא חישוב מלא של 4 מחזורים איפה שיש

            ReportDataProvider.Instance.Set2PrintAll();
            ReportDataProvider.Instance.mMainForm = this;

            EnumProvider.Instance.Init();

            Calculator.SetMainMasterValue(!(cbPersonMaster.Checked));

            // Remove extra life cycle tab if exists before run calculation again
            if (tcInput.TabPages[TabPageLifeCycle2] != null)
            {
                tcInput.TabPages.RemoveByKey(TabPageLifeCycle2);

            }

            runCalc();

            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.WaitForFullGCComplete();
            #region remark
            /*
            if (GatherData() == true)
            {
                #region DB Save if needed
                if (checkBox1.Checked == true)
                { // Saving will be done by this opeartion just as well
                    if (isPersonInDB() == false)
                    SaveData2DB();
                }
                else
                {
                    // noting will be saved
                }
                #endregion

                ClearOutputForm(); // Clears the Screen from other Data

                #region General Calculations
                NumericalCalc1();
                #endregion

                #region Intensive Map
                string[] IM = Calculator.CreateIntensiveMap(mFirstName, mLastName);

                SetINtensiveMapValues(IM);

                SetSumTableFromTextBoxes( mIntsMapTextArr, dgvIntsMapSum);
                #endregion

                #region Pitagoras Squares
                string[] PS = Calculator.CreatePitagorasSquares(mB_Date);

                SetPitagorasSquaresValues(PS);

                SetSumTableFromTextBoxes(mPitgMapTextArr , dgvPitMapSum);
                #endregion

                #region Life Cycles
                int[] LC = Calculator.CareteLifeCycles(mB_Date);

                SetLifeCycleValues(LC);

                #region Personal Info.
                txtLuck.Text = txtAstroName.Text;

                string pY, pM, pD;
                Calculator.CalcPersonalInfo(mB_Date, DateTimePickerTo.Value, out pY, out pM, out pD);

                txtPYear.Text = pY.ToString();
                txtPMonth.Text = pM.ToString();
                txtPDay.Text = pD.ToString();
                #endregion
                #endregion

                #region Chakras
                //txtCMaster.Text = Calculator.CalcMasterChkra(mFirstName + mLastName);//על
                txtCCrown.Text = Calculator.CalcCrownChakra(mFirstName + mLastName);//כתר
                txtCThirdEye.Text = Calculator.CalcThirdEyeChakra(mFirstName + mLastName);//עין שלישית 
                txtCThrought.Text = Calculator.CalcThroughtChakra(mFirstName + mLastName);//גרון
                txtCHeart.Text = Calculator.CalcHeartChakra(mFirstName + mLastName);//לב
                txtCMikSun.Text = Calculator.CalcSunChakra(mFirstName + mLastName);//מקלעת השמש
                txtCSexCreat.Text = Calculator.CalcSex_CreationChakra(mFirstName + mLastName);//מין ויצירה
                txtCRoot.Text = Calculator.CalcRootChakra(mFirstName + mLastName);//בסיס
                #endregion 

                #region Combined Map
                string[] inIntensiveMapData = new string[9] { txtIMap1.Text, txtIMap2.Text, txtIMap3.Text, txtIMap4.Text, txtIMap5.Text, txtIMap6.Text, txtIMap7.Text, txtIMap8.Text, txtIMap9.Text };
                string[] inPiatgorasMapData = new string[9] { txtPit1.Text, txtPit2.Text, txtPit3.Text, txtPit4.Text, txtPit5.Text, txtPit6.Text, txtPit7.Text, txtPit8.Text, txtPit9.Text };
                string[] outCombinedMapData = Calculator.CalcCombinedMap(inIntensiveMapData, inPiatgorasMapData);

                txtMComb1.Text = outCombinedMapData[0].ToString();
                txtMComb2.Text = outCombinedMapData[1].ToString();
                txtMComb3.Text = outCombinedMapData[2].ToString();
                txtMComb4.Text = outCombinedMapData[3].ToString();
                txtMComb5.Text = outCombinedMapData[4].ToString();
                txtMComb6.Text = outCombinedMapData[5].ToString();
                txtMComb7.Text = outCombinedMapData[6].ToString();
                txtMComb8.Text = outCombinedMapData[7].ToString();
                txtMComb9.Text = outCombinedMapData[8].ToString();

                SetSumTableFromTextBoxes( mCombMapTextArr, dgvCombMapSum);
                #endregion

                SetStylesToObjects();
            }
            else
            {
                return; // nothing was calclated
            }
            */
            #endregion remark
        }

        private void cmdNewClient_Click(object sender, EventArgs e)
        {
            if (tcInput.TabPages[TabPageLifeCycle2] != null)
            {
                tcInput.TabPages.RemoveByKey(TabPageLifeCycle2);

            }

            ClearForm();
        }

        private void gvDBview_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            int iRowNum = gvDBview.SelectedCells[0].RowIndex;
            gvDBview.Rows[gvDBview.SelectedCells[0].RowIndex].Selected = true;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            DateTimePickerTo.Value = DateTime.Now;
        }

        #region Future
        private int CalcDaysSpan(int amount, string type, int times)
        {
            int res = 0;
            switch (type)
            {
                case "יום":
                case "Day":
                    {
                        res = amount;
                        break;
                    }
                case "שבוע":
                case "Week":
                    {
                        res = amount * 7;
                        break;
                    }
                case "חודש":
                case "Month":
                    {
                        res = amount * 28;
                        break;
                    }
                case "שנה":
                case "Year":
                    {
                        res = amount * 365;
                        break;
                    }
            }

            return res * times;
        }

        private string FlipString(string s)
        {
            string sout = "";

            string[] ss = s.Split(Calculator.Delimiter);
            if (ss.Length > 1)
            {
                for (int i = 0; i < ss.Length; i++)
                {
                    sout += ss[ss.Length - 1 - i] + Calculator.Delimiter;
                }

                sout = sout.Substring(0, sout.Length - 1);
            }
            else
            {
                sout = s;
            }

            return sout;
        }

        private void listInterval_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            int days = CalcDaysSpan(Convert.ToInt16(listInterval.Text), listIntervalType.Text, Convert.ToInt16(listTimes.Text));
            FutureTime.Value = DateTime.Now.AddDays(days);
        }

        private void listIntervalType_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            int days = CalcDaysSpan(Convert.ToInt16(listInterval.Text), listIntervalType.Text, Convert.ToInt16(listTimes.Text));
            FutureTime.Value = DateTime.Now.AddDays(days);
        }

        private void btnFutureCalc_Click(object sender, EventArgs e)
        {
            #region Q and A - Filters
            // to show in output table
            List<string> sQAfilter = new List<string>() {   "-1",
                                                                ",1,2,3,4,5,6,8,9,11,22,33,",
                                                                ",1,2,4,6,8,9,11,22,33,",
                                                                ",1,4,8,9,11,22,",
                                                                ",1,3,4,5,8,9,11,22,",
                                                                ",1,2,3,4,5,6,8,9,11,14,19,22,33,",
                                                                ",1,2,4,6,8,9,11,22,33,",
                                                                "",
                                                                "",
                                                                "",
                                                                "",
                                                                "",
                                                                "",
                                                                "",
                                                                "",
                                                                "",
                                                                "",
                                                                "",};
            #endregion
            dataFuture.Rows.Clear();

            List<DateTime> DateList = CalcDates2FutureCalc(Convert.ToInt16(listInterval.Text), listIntervalType.Text, Convert.ToInt16(listTimes.Text));

            //int days = CalcDaysSpan(Convert.ToInt16(listInterval.Text), listIntervalType.Text,1);
            //for (int i = 1; i <= Convert.ToInt16(listTimes.Text); i++)
            for (int i = 0; i < DateList.Count; i++)
            {
                DateTime futuredate = DateList[i];//DateTime.Now.AddDays(days * i);

                string pY, pM, pD, strDate;
                Calculator.CalcPersonalInfo(mB_Date, futuredate, out pY, out pM, out pD);

                DataGridViewRow dR = new DataGridViewRow();
                //dataFuture.Rows.Add(1);

                strDate = futuredate.Day.ToString() + "." + futuredate.Month.ToString() + "." + futuredate.Year.ToString() + "  (" + FlipString(Calculator.FateCalc(futuredate)) + ")";

                #region Climax - Future
                DateTime newBD = futuredate;
                int[] fLC = Calculator.CareteLifeCycles(newBD);

                // first climax
                string ftrClmx = Calculator.CalcSum(fLC[2]);
                #endregion

                #region Question Filter
                string sQ = cmbQnA.SelectedItem.ToString();
                string[] spltQ = sQ.Split(".".ToCharArray()[0]);

                bool passedfilter = true;
                int tmpPy = Calculator.GetCorrectNumberFromSplitedString(pY.Split(Calculator.Delimiter));
                int tmpPm = Calculator.GetCorrectNumberFromSplitedString(pM.Split(Calculator.Delimiter));
                int tmpPd = Calculator.GetCorrectNumberFromSplitedString(pD.Split(Calculator.Delimiter));
                int tmpFate = Calculator.GetCorrectNumberFromSplitedString(Calculator.FateCalc(futuredate).Split(Calculator.Delimiter));
                int tmpFtrClmx = Calculator.GetCorrectNumberFromSplitedString(ftrClmx.Split(Calculator.Delimiter));

                if (Convert.ToInt16(spltQ[0]) != 1) // with filter
                {
                    if (Calculator.isCarmaticNumber(futuredate.Day) == true)
                    {
                        passedfilter = false;
                    }
                    else
                    {
                        string sCurFilter = sQAfilter[Convert.ToInt16(spltQ[0]) - 1];

                        if (cbPYtake.Checked == true)
                        {
                            passedfilter &= sCurFilter.Contains("," + tmpPy.ToString() + ",");
                        }

                        passedfilter &= sCurFilter.Contains("," + tmpPm.ToString() + ",");
                        passedfilter &= sCurFilter.Contains("," + tmpPd.ToString() + ",");
                        passedfilter &= sCurFilter.Contains("," + tmpFate.ToString() + ",");
                        passedfilter &= sCurFilter.Contains("," + tmpFtrClmx.ToString() + ",");
                    }
                }


                #endregion Question Filter

                if (passedfilter == true)
                {
                    object[] objects = new string[5];
                    objects.Initialize();
                    objects.SetValue(strDate, 0);
                    objects.SetValue(FlipString(pY), 1);
                    objects.SetValue(FlipString(pM), 2);
                    objects.SetValue(FlipString(pD), 3);
                    objects.SetValue(FlipString(ftrClmx), 4);

                    dataFuture.Rows.Insert(dataFuture.Rows.Count, objects);
                    dataFuture.Update();
                }
                //txtFuture.Text += futuredate.Date.ToShortDateString() + mSpaceTab + pY + mSpaceTab + pM + mSpaceTab + pD + System.Environment.NewLine;
            }
            dataFuture.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataFuture.Columns[1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataFuture.Columns[2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataFuture.Columns[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataFuture.Columns[4].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

            dataFuture.Columns[0].Width = 115;
            dataFuture.Columns[1].Width = 50;
            dataFuture.Columns[2].Width = 50;
            dataFuture.Columns[3].Width = 50;
            dataFuture.Columns[4].Width = 50;

            dataFuture.Update();
        }

        private List<DateTime> CalcDates2FutureCalc(int amount, string type, int times)
        {
            List<DateTime> FinalDates = new List<DateTime>();
            DateTime curDate = DateTimePickerTo.Value;

            FinalDates.Add(curDate);
            DateTime mLastDate = FinalDates[0];
            //DateTimeComparer dtc = new DateTimeComparer();

            for (int i = 1; i <= times; i++)
            {
                switch (type)
                {
                    case "יום":
                    case "Day":
                        {
                            mLastDate = mLastDate.AddDays(amount);
                            break;
                        }
                    case "שבוע":
                    case "Week":
                        {
                            mLastDate = mLastDate.AddDays(amount * 7);
                            break;
                        }
                    case "חודש":
                    case "Month":
                        {
                            mLastDate = mLastDate.AddMonths(amount);
                            break;
                        }
                    case "שנה":
                    case "Year":
                        {
                            mLastDate = mLastDate.AddYears(amount);
                            break;
                        }
                }

                FinalDates.Add(mLastDate);
            }

            // Sort list ordered by descending date
            return FinalDates.OrderByDescending(d => d.Date).ToList();
        }

        #endregion Future

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            string txt2lbl = "";
            if (checkBox1.Checked == true)
            {
                cmdSave.Enabled = false;
                switch (AppSettings.Instance.ProgramLanguage)
                {
                    case AppSettings.Language.Hebrew:
                        txt2lbl = "שמור וחשב";
                        break;
                    case AppSettings.Language.English:
                        txt2lbl = "Run and Save";
                        break;
                }
            }
            else
            {
                cmdSave.Enabled = true;
                switch (AppSettings.Instance.ProgramLanguage)
                {
                    case AppSettings.Language.Hebrew:
                        txt2lbl = "חשב";
                        break;
                    case AppSettings.Language.English:
                        txt2lbl = "Run";
                        break;
                }
            }

            cmdRun.Text = txt2lbl;
        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            txtBusinessStory.Text = "";
            RunBssnssCalc4cycles();
        }
        private void RunBssnssCalc4cycles()
        {
            int c = CalcCurrentCycle();
            DateTime dt = DateTimePickerTo.Value;
            string t = "";
            for (int i = 1; i < 5; i++)
            {
                txtBusinessStory.Text += Environment.NewLine;

                string[] s = (this.Controls.Find("txt" + i.ToString() + "_1", true)[0] as TextBox).Text.Split("-".ToCharArray());
                DateTimePickerTo.Value = mB_Date.AddYears((Convert.ToInt16(s[0]) + Convert.ToInt16(s[1])) / 2);

                #region intro 4 each cycle
                if (i == 1)
                {
                    if (BLL.AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                    {
                        txtBusinessStory.Text += "במחזור ראשון: ";
                    }
                    if (BLL.AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                    {
                        txtBusinessStory.Text += "First Cycle: ";
                    }
                }
                if (i == 2)
                {
                    if (BLL.AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                    {
                        txtBusinessStory.Text += "במחזור שני: ";
                    }
                    if (BLL.AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                    {
                        txtBusinessStory.Text += "Second Cycle: ";
                    }
                }
                if (i == 3)
                {
                    if (BLL.AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                    {
                        txtBusinessStory.Text += "במחזור שלישי: ";
                    }
                    if (BLL.AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                    {
                        txtBusinessStory.Text += "Third Cycle: ";
                    }
                }
                if (i == 4)
                {
                    if (BLL.AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                    {
                        txtBusinessStory.Text += "במחזור רביעי: ";
                    }
                    if (BLL.AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                    {
                        txtBusinessStory.Text += "Fourth Cycle: ";
                    }
                }
                #endregion intro 4 each cycle
                SingleBusinessMatchCalc(cmbBusinessBonus.SelectedItem.ToString().Trim(), cmbSelfFix.SelectedItem.ToString().Trim());
                txtBusinessStory.Text += Environment.NewLine;

                if (c == i)
                {
                    t = txtFinalBusinessMark.Text;
                }
            }

            #region "employee of the month"
            if ((Convert.ToInt16(Calculator.GetCorrectNumberFromSplitedString(txtNum6.Text.Split(Calculator.Delimiter))) == 8) || (Convert.ToInt16(Calculator.GetCorrectNumberFromSplitedString(txtNum3.Text.Split(Calculator.Delimiter))) == 8))
            {
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                {
                    txtBusinessStory.Text += Environment.NewLine + "יש לך פוטנציאל עובד מצטיין";
                }
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                {
                    txtBusinessStory.Text += Environment.NewLine + "You have the potential of being an excelent employee";
                }
            }
            #endregion

            #region Comment - 7
            int av = Convert.ToInt16(txtAstroName.Text.Split(" ".ToCharArray()[0])[1].Replace("(", "").Replace(")", ""));

            //if ((Calculator.GetCorrectNumberFromSplitedString(txtNum3.Text.Split(Calculator.Delimiter)) == 7) ||
            //     (Calculator.GetCorrectNumberFromSplitedString(txtNum5.Text.Split(Calculator.Delimiter)) == 7) ||
            //     (av == 7))
            if (new int[] { 2, 7, 13, 14, 16, 19 }.Contains(Calculator.GetCorrectNumberFromSplitedString(txtNum3.Text.Split(Calculator.Delimiter))))
            {
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                {
                    txtBusinessStory.Text += "יש לך פוטנציאל עובד מצטיין";
                    //"יכולות החשיבה שלך מונעות ממך למצות את הפוטנציאל הגלום בך להצלחה גבוהה יותר. מומלץ לך להוסיף שם על מנת לחדד ולשפר את יכולתך העסקית דבר שיבוא לידי ביטויו בפן העסקי והכלכלי. ";

                }
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                {
                    txtBusinessStory.Text += "Though your thinking abilities prevent the accomplishment of your potential to greater success, it is recommended for you to add a name in order to refine your financial strength.";
                }
            }
            #endregion Comment - 7


            DateTimePickerTo.Value = dt;
            txtFinalBusinessMark.Text = t;
            txtBusinessStory.Text = txtBusinessStory.Text.Trim();
        }


        private void btnMultiPartnersInBusiness_Click(object sender, EventArgs e)
        {
            isMultiBusineesCalc = true;

            List<UserInfo> allPertners = new List<UserInfo>();
            List<string> yesno = new List<string>();
            List<string> selffix = new List<string>();
            List<DataGridViewRow> RowList = new List<DataGridViewRow>();
            DateTime toCalcDay = DateTimePickerTo.Value;

            #region Gather Data From DataGridView
            UserInfo curPartner;
            for (int i = 0; i < mPartners.Rows.Count - 1; i++)
            {
                DataGridViewRow r = mPartners.Rows[i];
                RowList.Add(r);

                curPartner = new UserInfo();

                curPartner.mFirstName = CellValue2String(r.Cells[0]);
                curPartner.mLastName = CellValue2String(r.Cells[1]);
                curPartner.mFatherName = CellValue2String(r.Cells[3]);
                curPartner.mMotherName = CellValue2String(r.Cells[4]);
                bool resDate = DateTime.TryParse(r.Cells[2].Value.ToString().Trim(), out curPartner.mB_Date);
                if (resDate == false)
                {
                    MessageBox.Show("שגיאה בכתיבת תאריך", "שגיאת נתונים", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                yesno.Add(CellValue2String(r.Cells[5]));
                selffix.Add(CellValue2String(r.Cells[6]));
                curPartner.mCity = CellValue2String(r.Cells[7]);
                curPartner.mStreet = CellValue2String(r.Cells[8]);

                curPartner.mBuildingNum = CellValue2Double(r.Cells[9]);
                curPartner.mAppNum = CellValue2Double(r.Cells[10]);

                curPartner.mPhones = "";
                curPartner.mEMail = "";

                allPertners.Add(curPartner);
            }

            curPartner = new UserInfo();
            curPartner.mFirstName = txtPrivateName.Text.Trim();
            curPartner.mLastName = txtFamilyName.Text.Trim();
            curPartner.mFatherName = txtFatherName.Text.Trim();
            curPartner.mMotherName = txtMotherName.Text.Trim();
            curPartner.mB_Date = DateTimePickerFrom.Value;
            curPartner.mPhones = txtPhones.Text;
            curPartner.mEMail = txtEMail.Text;

            yesno.Add(cmbBusinessBonus.SelectedItem.ToString());
            selffix.Add(cmbSelfFix.SelectedItem.ToString());

            curPartner.mCity = txtCity.Text.Trim();
            curPartner.mStreet = txtStreet.Text.Trim();
            curPartner.mBuildingNum = Convert.ToDouble(txtBiuldingNum.Text);
            curPartner.mAppNum = Convert.ToDouble(txtAppNum.Text);

            allPertners.Add(curPartner);
            #endregion

            double finalMultiResult = 0;
            List<string> personalResults = new List<string>();
            List<double> persoanlBssnsRes = new List<double>();
            #region Calc Business Mark For Each Oertner

            for (int i = 0; i < yesno.Count; i++)
            {
                curPartner = allPertners[i];
                string bonusyesno = yesno[i];
                string bunosselffix = selffix[i];

                txtPrivateName.Text = curPartner.mFirstName;
                txtFamilyName.Text = curPartner.mLastName;
                txtFatherName.Text = curPartner.mFatherName;
                txtMotherName.Text = curPartner.mMotherName;
                DateTimePickerFrom.Value = curPartner.mB_Date;
                txtCity.Text = curPartner.mCity;
                txtStreet.Text = curPartner.mStreet;
                txtBiuldingNum.Text = curPartner.mBuildingNum.ToString();
                txtAppNum.Text = curPartner.mAppNum.ToString();

                DateTimePickerTo.Value = toCalcDay;

                runCalc();

                SingleBusinessMatchCalc(bonusyesno, bunosselffix);

                finalMultiResult += Convert.ToDouble(txtFinalBusinessMark.Text.Trim());
                persoanlBssnsRes.Add(Convert.ToDouble(txtFinalBusinessMark.Text.Trim()));

                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                {
                    personalResults.Add(curPartner.mFirstName + " " + curPartner.mLastName + ", תוצאה עסקית - " + txtFinalBusinessMark.Text.Trim() + " , " + txtBusinessStory.Text);
                }
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                {
                    personalResults.Add(curPartner.mFirstName + " " + curPartner.mLastName + ", Bussiness Index - " + txtFinalBusinessMark.Text.Trim() + " , " + txtBusinessStory.Text);
                }
            }

            finalMultiResult = finalMultiResult / allPertners.Count;
            #endregion

            for (int i = 0; i < RowList.Count; i++)
            {
                DataGridViewRow r = RowList[i];
                mPartners.Rows.Add(r);
            }
            mPartners.Update();
            mPartners.Refresh();

            #region Show results
            txtFinalMultipleBusineesMartk.Text = finalMultiResult.ToString();

            string story = "";
            if ((finalMultiResult >= 9) & (finalMultiResult < 10))
            {
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                {
                    story = "סיכוי מצויין להצלחה עסקית משותפת ויכולת לעסק לשרוד מעל חמש שנים";
                }
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                {
                    story = "Excellent chance of joint business success and business ability to survive over five years";
                }
            }

            if ((finalMultiResult >= 8) & (finalMultiResult < 9))
            {
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                {
                    story = "סיכוי טוב מאוד להצלחה עיסקית משותפת ויכולת לעסק לשרוד מעל חמש שנים";
                }
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                {
                    story = "Very good chance of success shared business and business ability to survive over five years";
                }
            }

            if ((finalMultiResult >= 7) & (finalMultiResult < 8))
            {
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                {
                    story = "סיכוי בינוני להצלחה עסקית משותפת ויכולת לעסק לשרוד מעל חמש שנים";
                }
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                {
                    story = "Medium chance of joint business success and ability to survive in business over five years";
                }
            }

            if ((finalMultiResult >= 6) & (finalMultiResult < 7))
            {
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                {
                    story = "סיכוי בינוני להצלחה עסקית משותפת ויכולת לעסק לשרוד עד שנתיים";
                }
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                {
                    story = "Moderate chance a joint business success and business ability to survive up to two years";
                }
            }

            if ((finalMultiResult >= 2) & (finalMultiResult < 6))
            {
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                {
                    story = "סיכוי קטן מאוד להצלחה עסקית משותפת ויכולת לעסק לשרוד עד מספר חודשים";
                }
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                {
                    story = "Very little chance of a joint business success and business ability to survive up to several months";
                }
            }

            if ((finalMultiResult >= 0) & (finalMultiResult < 2))
            {
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                {
                    story = "אין סיכוי להצלחה עסקית משותפת";
                }
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                {
                    story = "No chance of a joint business success";
                }
            }

            txtMultiPartnerStory.Text = story + System.Environment.NewLine + System.Environment.NewLine + "פרוט עבור השותפים:";

            //txtMultiPartnerStory.Text += System.Environment.NewLine +personalResults[personalResults.Count - 1];

            ReorderBusinessParteners(ref persoanlBssnsRes, ref personalResults);

            for (int i = 0; i < personalResults.Count; i++)
            {
                txtMultiPartnerStory.Text += System.Environment.NewLine + personalResults[i];

                //DataGridViewRow r = RowList[i];
                //mPartners.Rows.Add(r);
            }
            //mPartners.Update();
            //mPartners.Refresh();
            #endregion

            isMultiBusineesCalc = false;
        }

        private void ReorderBusinessParteners(ref List<double> pRes, ref List<string> pStr)
        {
            for (int pass = 0; pass <= (pRes.Count - 1); pass++)
            {
                for (int i = 0; i <= (pRes.Count - 2); i++)
                {
                    if (pRes[i] < pRes[i + 1])
                    {
                        bsSwapInt(ref pRes, i);
                        bsSwapStr(ref pStr, i);
                    }
                }
            }
        }
        private void bsSwapInt(ref List<double> arr, int pos)
        {
            arr.Insert(pos, arr[pos + 1]);
            arr.RemoveAt(pos + 2);
        }
        private void bsSwapStr(ref List<string> arr, int pos)
        {
            arr.Insert(pos, arr[pos + 1]);
            arr.RemoveAt(pos + 2);
        }

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

        private void btnConvertHebDate2GregorianDate_Click(object sender, EventArgs e)
        {
            try
            {
                DateTimePickerFrom.Value = Calculator.HebrewJewishDateString2GeorgianDate(txtHebDate.Text);
            }
            catch
            {
                MessageBox.Show("שגיאה בהכנסת תאריך עברי", "טעות קלט", MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtHebDate.Text = Calculator.GeorgianDate2HebrewJewishDateString(DateTimePickerFrom.Value);
            }

        }

        private void DateTimePickerFrom_ValueChanged(object sender, EventArgs e)
        {
            txtHebDate.Text = Calculator.GeorgianDate2HebrewJewishDateString(DateTimePickerFrom.Value);
        }
        #endregion

        #region Private Methods

        public void ClearForm()
        {
            cbPersonMaster.Checked = false;
            cbMainCorrectionDone.Checked = false;

            txtClientId.Text = string.Empty;
            txtAppNum.Text = "0";
            txtBiuldingNum.Text = "0";
            txtCity.Text = "";
            txtStreet.Text = "";
            txtUnique.Text = "";
            txtFamilyName.Text = "";
            txtPrivateName.Text = "";
            txtFatherName.Text = "";
            txtMotherName.Text = "";
            DateTimePickerFrom.Value = DateTime.Now;
            DateTimePickerTo.Value = DateTime.Now;
            txtEMail.Text = "";
            txtPhones.Text = "";

            txtNum1.Text = "";
            txtNum2.Text = "";
            txtNum3.Text = "";
            txtNum4.Text = "";
            txtNum5.Text = "";
            txtNum6.Text = "";
            txtNum7.Text = "";

            ClearOutputForm();

            tcInput.SelectTab("ClientDataPage");
            mTabs.SelectTab("tabPage1");

            cmbSexSelect.SelectedItem = cmbSexSelect.Items[0];
            if (BLL.AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
            {
                cmbSexSelect.Text = "Gender";
            }
            if (BLL.AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
            {
                cmbSexSelect.Text = "בחר מין";
            }
            txtApplication.Text = "";

        }

        private void ClearOutputForm()
        {
            checkBox1.Checked = true;
            checkBox1.CheckState = CheckState.Checked;
            checkBox1_CheckedChanged(null, null);

            #region General Calculations
            txtNum1.BackColor = mRegularColor;
            txtNum2.BackColor = mRegularColor;
            txtNum3.BackColor = mRegularColor;
            txtNum4.BackColor = mRegularColor;
            txtNum5.BackColor = mRegularColor;
            txtNum6.BackColor = mRegularColor;
            txtNum7.BackColor = mRegularColor;
            txtNum9.BackColor = mRegularColor;
            txtNum8.BackColor = mRegularColor;
            txtSpiralCode.BackColor = mRegularColor;
            txtParentsPresent.Text = "";
            txtParentsPresent.BackColor = mRegularColor;
            txtParentsPresent.ForeColor = mBlack;

            txtNum1.ForeColor = mBlack;
            txtNum2.ForeColor = mBlack;
            txtNum3.ForeColor = mBlack;
            txtNum4.ForeColor = mBlack;
            txtNum5.ForeColor = mBlack;
            txtNum6.ForeColor = mBlack;
            txtNum7.ForeColor = mBlack;
            txtNum9.ForeColor = mBlack;
            txtNum8.ForeColor = mBlack;
            txtSpiralCode.ForeColor = mBlack;

            picFName.Visible = false;
            picLName.Visible = false;
            picSml1.Visible = false;
            picSml2.Visible = false;
            picSml3.Visible = false;
            picSml4.Visible = false;
            picSml5.Visible = false;
            picSml6.Visible = false;
            picSml7.Visible = false;
            picSmlNINE.Visible = false;
            picSmlMik.Visible = false;
            picParentsPresent.Visible = false;
            picUnique.Visible = false;
            picPowerNum.Visible = false;

            txtNum9.Text = "";
            txtNum8.Text = "";

            picSml8.Visible = false;
            picSml9.Visible = false;

            txtAge.Text = "";
            txtPName_Num.Text = "";
            txtFName_Num.Text = "";
            txtPName_Num.BackColor = mRegularColor;
            txtPName_Num.ForeColor = mBlack;
            txtFName_Num.BackColor = mRegularColor;
            txtFName_Num.ForeColor = mBlack;
            txtSpiralCode.Text = "";

            txtAstroName.Text = "";
            txtAstroName.ForeColor = mBlack;
            txtAstroName.BackColor = mRegularColor;

            txtPowerNum.Text = "";
            txtPowerNum.ForeColor = mBlack;
            txtPowerNum.BackColor = mRegularColor;

            txtInfo1.Text = "";
            txtInfo1.ForeColor = mBlack;
            txtInfo1.BackColor = mRegularColor;

            txtAppComp.Text = "";
            //txtAppComp.ForeColor = mBlack;
            //txtAppComp.BackColor = mRegularColor;
            picApp.Visible = false;

            txtCityComp.Text = "";
            //txtCityComp.ForeColor = mBlack;
            //txtCityComp.BackColor = mRegularColor;
            picCity.Visible = false;
            #endregion

            #region IntensiveMap
            txtIMap1.Text = "";
            txtIMap2.Text = "";
            txtIMap3.Text = "";
            txtIMap4.Text = "";
            txtIMap5.Text = "";
            txtIMap6.Text = "";
            txtIMap7.Text = "";
            txtIMap8.Text = "";
            txtIMap9.Text = "";

            dgvIntsMapSum.Rows.Clear();
            #endregion

            #region Pitagors Squares
            txtPit1.Text = "";
            txtPit2.Text = "";
            txtPit3.Text = "";
            txtPit4.Text = "";
            txtPit5.Text = "";
            txtPit6.Text = "";
            txtPit7.Text = "";
            txtPit8.Text = "";
            txtPit9.Text = "";

            dgvPitMapSum.Rows.Clear();
            #endregion

            #region Life Cycles
            txt1_1.Text = "";
            txt1_2.Text = "";
            txt1_3.Text = "";
            txt1_4.Text = "";
            txt2_1.Text = "";
            txt2_2.Text = "";
            txt2_3.Text = "";
            txt2_4.Text = "";
            txt3_1.Text = "";
            txt3_2.Text = "";
            txt3_3.Text = "";
            txt3_4.Text = "";
            txt4_1.Text = "";
            txt4_2.Text = "";
            txt4_3.Text = "";
            txt4_4.Text = "";
            txtAge2.Text = "";

            txt1_1.ForeColor = mBlack;
            txt1_2.ForeColor = mBlack;
            txt1_3.ForeColor = mBlack;
            txt1_4.ForeColor = mBlack;
            txt1_1.BackColor = mRegularColor;
            txt1_2.BackColor = mRegularColor;
            txt1_3.BackColor = mRegularColor;
            txt1_4.BackColor = mRegularColor;

            txt2_1.ForeColor = mBlack;
            txt2_2.ForeColor = mBlack;
            txt2_3.ForeColor = mBlack;
            txt2_4.ForeColor = mBlack;
            txt2_1.BackColor = mRegularColor;
            txt2_2.BackColor = mRegularColor;
            txt2_3.BackColor = mRegularColor;
            txt2_4.BackColor = mRegularColor;

            txt3_1.ForeColor = mBlack;
            txt3_2.ForeColor = mBlack;
            txt3_3.ForeColor = mBlack;
            txt3_4.ForeColor = mBlack;
            txt3_1.BackColor = mRegularColor;
            txt3_2.BackColor = mRegularColor;
            txt3_3.BackColor = mRegularColor;
            txt3_4.BackColor = mRegularColor;

            txt4_1.ForeColor = mBlack;
            txt4_2.ForeColor = mBlack;
            txt4_3.ForeColor = mBlack;
            txt4_4.ForeColor = mBlack;
            txt4_1.BackColor = mRegularColor;
            txt4_2.BackColor = mRegularColor;
            txt4_3.BackColor = mRegularColor;
            txt4_4.BackColor = mRegularColor;

            LCpic1_2.Visible = false;
            LCpic2_2.Visible = false;
            LCpic3_2.Visible = false;
            LCpic4_2.Visible = false;

            LCpic1_3.Visible = false;
            LCpic2_3.Visible = false;
            LCpic3_3.Visible = false;
            LCpic4_3.Visible = false;

            LCpic1_4.Visible = false;
            LCpic2_4.Visible = false;
            LCpic3_4.Visible = false;
            LCpic4_4.Visible = false;

            LCpicY.Visible = false;
            LCpicM.Visible = false;
            LCpicD.Visible = false;

            IMpic1.Visible = false;
            IMpic2.Visible = false;
            IMpic3.Visible = false;
            IMpic4.Visible = false;
            IMpic5.Visible = false;
            IMpic6.Visible = false;
            IMpic7.Visible = false;
            IMpic8.Visible = false;
            IMpic9.Visible = false;

            PSpic1.Visible = false;
            PSpic2.Visible = false;
            PSpic3.Visible = false;
            PSpic4.Visible = false;
            PSpic5.Visible = false;
            PSpic6.Visible = false;
            PSpic7.Visible = false;
            PSpic8.Visible = false;
            PSpic9.Visible = false;

            txtLuck.Text = "";
            txtLuck.ForeColor = mBlack;
            txtLuck.BackColor = mRegularColor;

            txtXLuck.Text = "";
            txtXLuck.ForeColor = mBlack;
            txtXLuck.BackColor = mRegularColor;

            txtPYear.Text = "";
            txtPYear.ForeColor = mBlack;
            txtPYear.BackColor = mRegularColor;
            txtPMonth.Text = "";
            txtPMonth.ForeColor = mBlack;
            txtPMonth.BackColor = mRegularColor;
            txtPDay.Text = "";
            txtPDay.ForeColor = mBlack;
            txtPDay.BackColor = mRegularColor;

            //User details in מחזורי חיים
            txtPrivateNameCycle.Text = "";
            txtPrivateNameCycle.ForeColor = mBlack;
            txtPrivateNameCycle.BackColor = mRegularColor;

            txtFamilyNameCycle.Text = "";
            txtFamilyNameCycle.ForeColor = mBlack;
            txtFamilyNameCycle.BackColor = mRegularColor;

            DateTimePickerFromCycle.Text = "";

            #endregion

            #region Life Cycles Extra
            txtX1_1.Text = "";
            txtX1_2.Text = "";
            txtX1_3.Text = "";
            txtX1_4.Text = "";
            txtX2_1.Text = "";
            txtX2_2.Text = "";
            txtX2_3.Text = "";
            txtX2_4.Text = "";
            txtX3_1.Text = "";
            txtX3_2.Text = "";
            txtX3_3.Text = "";
            txtX3_4.Text = "";
            txtX4_1.Text = "";
            txtX4_2.Text = "";
            txtX4_3.Text = "";
            txtX4_4.Text = "";
            txtAgeX2.Text = "";

            txtX1_1.ForeColor = mBlack;
            txtX1_2.ForeColor = mBlack;
            txtX1_3.ForeColor = mBlack;
            txtX1_4.ForeColor = mBlack;
            txtX1_1.BackColor = mRegularColor;
            txtX1_2.BackColor = mRegularColor;
            txtX1_3.BackColor = mRegularColor;
            txtX1_4.BackColor = mRegularColor;

            txtX2_1.ForeColor = mBlack;
            txtX2_2.ForeColor = mBlack;
            txtX2_3.ForeColor = mBlack;
            txtX2_4.ForeColor = mBlack;
            txtX2_1.BackColor = mRegularColor;
            txtX2_2.BackColor = mRegularColor;
            txtX2_3.BackColor = mRegularColor;
            txtX2_4.BackColor = mRegularColor;

            txtX3_1.ForeColor = mBlack;
            txtX3_2.ForeColor = mBlack;
            txtX3_3.ForeColor = mBlack;
            txtX3_4.ForeColor = mBlack;
            txtX3_1.BackColor = mRegularColor;
            txtX3_2.BackColor = mRegularColor;
            txtX3_3.BackColor = mRegularColor;
            txtX3_4.BackColor = mRegularColor;

            txtX4_1.ForeColor = mBlack;
            txtX4_2.ForeColor = mBlack;
            txtX4_3.ForeColor = mBlack;
            txtX4_4.ForeColor = mBlack;
            txtX4_1.BackColor = mRegularColor;
            txtX4_2.BackColor = mRegularColor;
            txtX4_3.BackColor = mRegularColor;
            txtX4_4.BackColor = mRegularColor;

            LCpicX1_2.Visible = false;
            LCpicX2_2.Visible = false;
            LCpicX3_2.Visible = false;
            LCpicX4_2.Visible = false;

            LCpicX1_3.Visible = false;
            LCpicX2_3.Visible = false;
            LCpicX3_3.Visible = false;
            LCpicX4_3.Visible = false;

            LCpicX1_4.Visible = false;
            LCpicX2_4.Visible = false;
            LCpicX3_4.Visible = false;
            LCpicX4_4.Visible = false;

            LCpicYX.Visible = false;
            LCpicMX.Visible = false;
            LCpicDX.Visible = false;

            txtXLuck.Text = "";
            txtXLuck.ForeColor = mBlack;
            txtXLuck.BackColor = mRegularColor;

            txtPYearX.Text = "";
            txtPYearX.ForeColor = mBlack;
            txtPYearX.BackColor = mRegularColor;
            txtPMonthX.Text = "";
            txtPMonthX.ForeColor = mBlack;
            txtPMonthX.BackColor = mRegularColor;
            txtPDayX.Text = "";
            txtPDayX.ForeColor = mBlack;
            txtPDayX.BackColor = mRegularColor;
            #endregion

            #region Future Calc
            listInterval.Text = listInterval.Items[0].ToString();
            listIntervalType.Text = listIntervalType.Items[0].ToString();
            listTimes.Text = listTimes.Items[0].ToString();

            int days = CalcDaysSpan(Convert.ToInt16(listInterval.Text), listIntervalType.Text, Convert.ToInt16(listTimes.Text));
            FutureTime.Value = DateTime.Now.AddDays(days);
            dataFuture.Rows.Clear();
            //txtFuture.Text = "";

            cmbQnA.Text = cmbQnA.Items[0].ToString();
            #endregion

            #region Chakras
            //txtCMaster.Text = "";//על
            //txtCCrown.Text = "";//כתר
            //txtCThirdEye.Text = "";//עין שלישית 
            //txtCThrought.Text = ""; //גרון
            //txtCHeart.Text = "";//לב
            //txtCMikSun.Text = "";//מקלעת השמש
            //txtCSexCreat.Text = "";//מין ויצירה
            //txtCRoot.Text = "";//בסיס

            //picMaster.Visible = false;
            //picCrown.Visible = true;
            //picThirdEye.Visible = true;
            //picMikSun.Visible = true;
            //picSexCreat.Visible = true;
            //picThrought.Visible = true;
            //picHeart.Visible = true;
            //picRoot.Visible = true;

            //picCrown.Image = Omega.Properties.Resources.Crown1;
            //picThirdEye.Image = Omega.Properties.Resources.ThirdEye;
            //picMikSun.Image = Omega.Properties.Resources.MikSun;
            //picSexCreat.Image = Omega.Properties.Resources.SexCreation;
            //picThrought.Image = Omega.Properties.Resources.Throught;
            //picHeart.Image = Omega.Properties.Resources.Heart;
            //picRoot.Image = Omega.Properties.Resources.Root;
            #endregion

            #region Dynamic Chakras

            for (int i = 1; i < 10; i++)
            {
                TextBox t = this.Controls.Find("txtDC" + i.ToString(), true)[0] as TextBox;
                t.Text = string.Empty;

                PictureBox p = this.Controls.Find("picDC" + i.ToString(), true)[0] as PictureBox;
                p.Image = null;

            }

            #endregion Dynamic Chakras

            #region Combined Map
            txtMComb1.Text = "";
            txtMComb2.Text = "";
            txtMComb3.Text = "";
            txtMComb4.Text = "";
            txtMComb5.Text = "";
            txtMComb6.Text = "";
            txtMComb7.Text = "";
            txtMComb8.Text = "";
            txtMComb9.Text = "";

            picMComb1.Visible = false;
            picMComb2.Visible = false;
            picMComb3.Visible = false;
            picMComb4.Visible = false;
            picMComb5.Visible = false;
            picMComb6.Visible = false;
            picMComb7.Visible = false;
            picMComb8.Visible = false;
            picMComb9.Visible = false;

            dgvCombMapSum.Rows.Clear();
            lblCombInfo.Text = "";
            #endregion //Combined Map

            txtBusinessStory.Text = "";
            txtFinalBusinessMark.Text = "";
            cmbBusinessBonus.Text = "לא";
            if (cbMainCorrectionDone.Checked == true)
            {
                cmbSelfFix.SelectedItem = cmbSelfFix.Items[0];
            }
            else
            {
                cmbSelfFix.SelectedItem = cmbSelfFix.Items[1];
            }
            cmbSelfFix.Text = cmbSelfFix.SelectedItem.ToString();

            txtFinalMultipleBusineesMartk.Text = "";
            txtFinalMultipleSeniorMngsMartk.Text = "";
            txtFinalMultipleJuniorMngsMartk.Text = "";
            txtMultiPartnerStory.Text = "";
            txtMultiSecreteryStory.Text = "";
            txtMultiAccountingStory.Text = "";
            txtMultiSeniorMngsStory.Text = "";
            txtMultiJuniorrMngsStory.Text = "";
            mPartners.Rows.Clear();
            mSeniorMngPartners.Rows.Clear();
            mJuniorMngPartners.Rows.Clear();
            mPartnersSecretery.Rows.Clear();
            mPartnersAccounting.Rows.Clear();

            txtMrrgMark.Text = "";
            txtMrrgStory.Text = "";
            rbtnYesMarriage.Checked = true;
            rbtnNoMarriage.Checked = false;

            txtFinalSalesMark.Text = txtSalesStory.Text = string.Empty;

            #region Sexual Match

            SpouceData.Rows.Clear();
            SpouceSexData.Rows.Clear();
            txtSexMatchMark.Text = txtSexMatchStory.Text = string.Empty;

            #endregion Sexual Match

            #region Sccss_AttPrb
            txtLeanSccss.Text = "";
            txtLearnStory.Text = "";

            txtAtt.Text = "";
            txtAttStory.Text = "";

            cmbtnSccssSelect.Text = cmbtnSccssSelect.Items[0].ToString();
            txtAttMajor.Text = "";
            txtAttMinor.Text = "";
            #endregion

            #region Health
            txtHealthStory.Text = "";
            txtHealthValue.Text = "";
            #endregion

            #region Questions

            txtQnAres.Clear();
            txtQfilter.Clear();
            cbQPYmode.Checked = true;
            cbQPYmode.CheckState = CheckState.Checked;
            cmbQ.SelectedIndex = -1;
            cmbQF.SelectedIndex = -1;
            cmbQ.Text = "כל השאלות";
            cmbQF.Text = "כל השאלות";
            cmbInterval.SelectedIndex = -1;
            cmbInterval.Text = "1";
            cmbType.SelectedIndex = -1;
            cmbType.Text = "יום";
            cmbTimes.SelectedIndex = -1;
            cmbTimes.Text = "500";
            dgvQnA.Rows.Clear();

            #endregion Questions

        }

        /// <summary> Tests the form to its input data for sufficient data inserted
        /// 
        /// </summary>
        /// <returns>True \ False</returns>
        private bool isSufficientValuesIn()
        {
            return ((txtPrivateName.Text != "") && (txtFamilyName.Text != ""));
        }

        /// <summary> Gathering Data from the MainForm Personal Info. Groupbox
        /// 
        /// </summary>
        /// <returns>true OR flase - rather there is not a sufficient data in</returns>
        private bool GatherData()
        {
            bool resGatherData = true;

            #region Validate Data
            if (isSufficientValuesIn() == false)
            {
                string s1 = "", s2 = "";
                if (BLL.AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                {
                    s1 = "חובה להכניס שם פרטי ושם משפחה בצורה מלאה";
                    s2 = "נתונים חסרים : ";
                }
                if (BLL.AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                {
                    s1 = "First name and Last name are obligatory in this process";
                    s2 = "Missing data...";
                }
                MessageBox.Show(s2 + System.Environment.NewLine + s1, "טעות הכנסת נתונים", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                resGatherData = false;
                return false;
            }
            #endregion

            try
            {
                #region Change_Comma_To_ELSE
                txtPrivateName.Text = txtPrivateName.Text.Replace(",", ";");
                txtFamilyName.Text = txtFamilyName.Text.Replace(",", ";");
                txtMotherName.Text = txtMotherName.Text.Replace(",", ";");
                txtFatherName.Text = txtFatherName.Text.Replace(",", ";");

                txtCity.Text = txtCity.Text.Replace(",", ";");
                txtStreet.Text = txtStreet.Text.Replace(",", ";");
                txtEMail.Text = txtEMail.Text.Replace(",", ";");
                txtPhones.Text = txtPhones.Text.Replace(",", ";");
                #endregion Change_Comma_To_ELSE

                mFirstName = Calculator.ChangeFinalChars(txtPrivateName.Text);
                mLastName = Calculator.ChangeFinalChars(txtFamilyName.Text);
                mMotherName = Calculator.ChangeFinalChars(txtMotherName.Text);
                mFatherName = Calculator.ChangeFinalChars(txtFatherName.Text);

                mB_Date = DateTimePickerFrom.Value;
                mCity = Calculator.ChangeFinalChars(txtCity.Text);
                mStreet = Calculator.ChangeFinalChars(txtStreet.Text);
                mBuildingNum = Convert.ToDouble(txtBiuldingNum.Text);
                mAppNum = Convert.ToDouble(txtAppNum.Text);

                mEMail = txtEMail.Text;
                mPhones = txtPhones.Text;
                if (cmbSexSelect.SelectedItem == null)
                {
                    cmbSexSelect.SelectedItem = cmbSexSelect.Items[0];
                }
                string s = cmbSexSelect.SelectedItem.ToString();
                if (BLL.AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                {
                    mSex = EnumProvider.Instance.GetSexEnumFromString(s);
                }
                if (BLL.AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                {
                    if (s.ToLower() == EnumProvider.Sex.Female.ToString().ToLower())
                    {
                        mSex = EnumProvider.Sex.Female;
                    }
                    if (s.ToLower() == EnumProvider.Sex.Male.ToString().ToLower())
                    {
                        mSex = EnumProvider.Sex.Male;
                    }
                }
                //cmbSexSelect.SelectedIndex = cmbSexSelect.FindStringExact(s);
                //cmbSexSelect.Text = cmbSexSelect.SelectedIndex.ToString();
                mApplication = txtApplication.Text;
                if (cbMainCorrectionDone.Checked == true)
                {
                    mPassedRect = EnumProvider.PassedRectification.Passed;
                    cmbSelfFix.SelectedIndex = 0;
                    cmbSelfFix.Text = cmbSelfFix.Items[0].ToString();
                }
                else
                {
                    mPassedRect = EnumProvider.PassedRectification.NotPassed;
                    cmbSelfFix.SelectedIndex = 1;
                    cmbSelfFix.Text = cmbSelfFix.Items[1].ToString();
                }
                if (cbPersonMaster.Checked == true)
                {
                    mReachMaster = EnumProvider.ReachedMaster.No;
                }
                else
                {
                    mReachMaster = EnumProvider.ReachedMaster.Yes;
                }

                //mSex = EnumProvider.Sex.Male;
                //mApplication = "";
                //mPassedRect = EnumProvider.PassedRectification.NotPassed;

                CurrentUserData = new UserInfo();

                int.TryParse(txtClientId.Text, out CurrentUserData.mId);
                CurrentUserData.mFirstName = txtPrivateName.Text;
                CurrentUserData.mLastName = txtFamilyName.Text;
                CurrentUserData.mFatherName = txtFatherName.Text;
                CurrentUserData.mMotherName = txtMotherName.Text;
                CurrentUserData.mB_Date = mB_Date;
                CurrentUserData.mCity = txtCity.Text;
                CurrentUserData.mStreet = txtStreet.Text;
                CurrentUserData.mBuildingNum = Convert.ToDouble(txtBiuldingNum.Text);
                CurrentUserData.mAppNum = Convert.ToDouble(txtAppNum.Text);
                CurrentUserData.mEMail = txtEMail.Text;
                CurrentUserData.mPhones = txtPhones.Text;
                CurrentUserData.mApplication = mApplication;
                CurrentUserData.mSex = mSex;
                CurrentUserData.mPassedRect = mPassedRect;
                CurrentUserData.mReachMaster = mReachMaster;

                resGatherData = true;
            }
            catch (Exception exp)
            {
                resGatherData = false;
                throw new Exception("Problem in one of the parameters of input:" + Environment.NewLine, exp);
            }
            return resGatherData;
        }

        /// <summary> Calculates the Numerical Data on a persons information.
        /// and inserts it to the correct textboxes.
        /// </summary>
        private void NumericalCalc1()
        {
            txtNum1.Text = Calculator.DayOfTheMonth(mB_Date); // יום בחודש - כתר
            txtNum3.Text = Calculator.VouleCalc(mFirstName, mLastName); // גרון - ראש
            txtNum5.Text = Calculator.FateCalc(mB_Date); // מקלעת השמש - גורל
            txtNum6.Text = Calculator.FullNameCalc(mFirstName, mLastName); // מין ויצירה - שם מלא (כל האותיות
            txtNum7.Text = Calculator.ConsonantsCalc(mFirstName, mLastName); //רגליים - בסיס
            txtNum2.Text = Calculator.MaturityCalc(txtNum6.Text, txtNum5.Text);// בגרות - עין שלישית
            txtNum4.Text = Calculator.WoundCalc(txtNum6.Text, txtNum5.Text);  // לב - פצע

            //////////
            txtSpiralCode.Text = txtNum5.Text;
            txtPName_Num.Text = Calculator.PrivateNameCalc(mFirstName);
            txtFName_Num.Text = Calculator.PrivateNameCalc(mLastName);

            string luckName = "";
            mLuckNumber = Calculator.GetAstroData(mB_Date, out luckName);
            txtAstroName.Text = luckName + " (" + mLuckNumber.ToString() + ")";

            DateTime today = DateTime.Now;
            //TimeSpan ts = today.Subtract(mB_Date);
            txtAge.Text = sCalcAge_wMonths(mB_Date); // Math.Floor(Convert.ToDouble(ts.Days / 360)).ToString();
            txtAge2.Text = txtAgeX2.Text = txtAge.Text;

            txtUnique.Text = Calculator.CalcUnique(mFirstName, mLastName, mB_Date).ToString();

            if ((mMotherName != "") && (mFatherName != ""))
            {
                txtParentsPresent.Text = Calculator.CalcParentsPresent(mFirstName, mLastName, mB_Date, mMotherName, mFatherName).ToString();
            }

            if (txtCity.Text != "")
            {
                txtCityComp.Text = Calculator.CityCompetability(txtNum5.Text, mCity);
            }
            else
            {
                txtCityComp.Text = "-";
            }
            if (txtAppNum.Text != "0")
            {
                txtAppComp.Text = Calculator.AppCompetability(txtNum5.Text, Convert.ToInt16(mAppNum));
            }
            else
            {
                txtAppComp.Text = "-";
            }

            if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
            {
                txtInfo1.Text = "המפה : ";
            }
            if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
            {
                txtInfo1.Text = "The Map : ";
            }

            #region User_info

            string[] sFN = txtNum6.Text.Split(Calculator.Delimiter);
            string[] sF = txtNum5.Text.Split(Calculator.Delimiter);

            int num1 = -9;
            int num2 = -9;

            if (cbPersonMaster.Checked == false)
            {
                num1 = Convert.ToInt16(Calculator.GetCorrectNumberFromSplitedString(txtNum6.Text.Split(Calculator.Delimiter)));
                num2 = Convert.ToInt16(Calculator.GetCorrectNumberFromSplitedString(txtNum5.Text.Split(Calculator.Delimiter)));
            }
            else //(cbPersonMaster.Checked == true)
            {
                num1 = Convert.ToInt16(sFN[sFN.Length - 1]);
                num2 = Convert.ToInt16(sF[sF.Length - 1]);
            }


            if (cbMainCorrectionDone.Checked == false)
            {
                num2 = Convert.ToInt16(Calculator.GetCorrectNumberFromSplitedString(txtNum5.Text.Split(Calculator.Delimiter)));
            }
            else //(cbPersonMaster.Checked == true)
            {
                num2 = Convert.ToInt16(sF[sF.Length - 1]);
            }


            bool num1test = true, num2test = true;
            bool res1 = true, res2 = true, res3 = true, res4 = true, res5 = true, res6 = true, res7 = true, res8 = true;
            for (int i = 0; i < sFN.Length; i++)
            {
                if (Calculator.isCarmaticNumber(Convert.ToInt16(sFN[i])))
                    num1test = false;
            }

            //for (int i = 0; i < sF.Length; i++)
            //{
            //    if (Calculator.isCarmaticNumber(Convert.ToInt16(sF[i])))
            //        num2test = false;
            //}
            num2test = !Calculator.isCarmaticNumber(Calculator.GetCorrectNumberFromSplitedString(num2.ToString().Split(Calculator.Delimiter)));

            string sVal = txtAstroName.Text.Trim().Split(" ".ToCharArray()[0])[1];
            sVal = sVal.Replace(")", " ");
            sVal = sVal.Replace("(", " ");
            res8 = !Calculator.isCarmaticNumber(Convert.ToInt16(sVal.Trim()));

            string outStr = "";

            res1 = !Calculator.isCarmaticNumber(Calculator.GetCorrectNumberFromSplitedString(txtPName_Num.Text.Split(Calculator.Delimiter)));
            res2 = !Calculator.isCarmaticNumber(Calculator.GetCorrectNumberFromSplitedString(txtNum3.Text.Split(Calculator.Delimiter)));
            res3 = !Calculator.isCarmaticNumber(Calculator.GetCorrectNumberFromSplitedString(txtNum1.Text.Split(Calculator.Delimiter)));
            res4 = !(Calculator.GetCorrectNumberFromSplitedString(txtNum4.Text.Split(Calculator.Delimiter)) == 0);
            res5 = !Calculator.isCarmaticNumber(Calculator.GetCorrectNumberFromSplitedString(txtNum2.Text.Split(Calculator.Delimiter)));
            res6 = !Calculator.isCarmaticNumber(Calculator.GetCorrectNumberFromSplitedString(txtNum7.Text.Split(Calculator.Delimiter)));
            res7 = !Calculator.isCarmaticNumber(Calculator.GetCorrectNumberFromSplitedString(txtNum4.Text.Split(Calculator.Delimiter)));

            if ((num1test & num2test & res1 & res2 & res3 & res5 & res6 & res7 & res8) == false)
            {
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                {
                    outStr = "לא מאוזן";
                }
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                {
                    outStr = "Not Balanced";
                }
            }
            else
            {
                string tableTestValue = Calculator.TestHramonicInfo(num1, num2);
                outStr = tableTestValue;
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                {
                }
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                {
                    if (outStr == "מאוזן")
                    {
                        outStr = "Balanced";
                    }
                    if (outStr == "לא מאוזן")
                    {
                        outStr = "Not Balanced";
                    }
                    if (outStr == "חצי מאוזן")
                    {
                        outStr = "Half Balanced";
                    }

                }


                string strCompare = "";
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                {
                    strCompare = "לא מאוזן";
                }
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                {
                    strCompare = "Not Balanced";
                }

                if (strCompare == tableTestValue)
                {
                    if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                    {
                        outStr = "חצי מאוזן";
                    }
                    if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                    {
                        outStr = "Half Balanced";
                    }
                }



                if (Calculator.GetCorrectNumberFromSplitedString(txtNum4.Text.Split(Calculator.Delimiter)) == 0)
                {
                    bool test0 = false;

                    int[] vals = new int[] { 1, 2, 4, 6, 7, 8, 33, 13, 14, 16, 19 };
                    int vv1 = Calculator.GetCorrectNumberFromSplitedString(txtNum5.Text.Split(Calculator.Delimiter));
                    int vv2 = Calculator.GetCorrectNumberFromSplitedString(txtNum6.Text.Split(Calculator.Delimiter));
                    for (int i = 0; i < vals.Length; i++)
                    {
                        if ((vv1 == vals[i]) && (vv2 == vals[i]))
                        {
                            test0 = true;
                        }
                    }

                    if (test0 == true)
                    {
                        if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                        {
                            outStr = "לא מאוזן";
                        }
                        if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                        {
                            outStr = "Not Balanced";
                        }
                    }
                    else
                    {
                        #region "Joker Card"
                        res1 = Calculator.isMaterNumber(Calculator.GetCorrectNumberFromSplitedString(txtPName_Num.Text.Split(Calculator.Delimiter)));
                        //res2 = Calculator.isMaterNumber(Calculator.GetCorrectNumberFromSplitedString(txtNum5.Text.Split(Calculator.Delimiter)));
                        res2 = Calculator.isMaterNumber(Calculator.GetCorrectNumberFromSplitedString(num2.ToString().Split(Calculator.Delimiter)));

                        sVal = txtAstroName.Text.Trim().Split(" ".ToCharArray()[0])[1];
                        sVal = sVal.Replace(")", " ");
                        sVal = sVal.Replace("(", " ");
                        res3 = Calculator.isMaterNumber(Convert.ToInt16(sVal.Trim()));

                        List<int> nums2check = new List<int>();
                        bool resCC = false;
                        #region Check for carmatic numbers in Chakra Values
                        nums2check.Clear();
                        nums2check.Add(Calculator.GetCorrectNumberFromSplitedString(txtPName_Num.Text.Split(Calculator.Delimiter)));
                        nums2check.Add(Convert.ToInt16(sVal.Trim()));
                        nums2check.Add(Calculator.GetCorrectNumberFromSplitedString(txtNum1.Text.Split(Calculator.Delimiter)));
                        nums2check.Add(Calculator.GetCorrectNumberFromSplitedString(txtNum2.Text.Split(Calculator.Delimiter)));
                        nums2check.Add(Calculator.GetCorrectNumberFromSplitedString(txtNum3.Text.Split(Calculator.Delimiter)));
                        nums2check.Add(Calculator.GetCorrectNumberFromSplitedString(txtNum4.Text.Split(Calculator.Delimiter)));
                        //nums2check.Add(Calculator.GetCorrectNumberFromSplitedString(txtNum5.Text.Split(Calculator.Delimiter)));
                        nums2check.Add(Calculator.GetCorrectNumberFromSplitedString(num2.ToString().Split(Calculator.Delimiter)));
                        nums2check.Add(Calculator.GetCorrectNumberFromSplitedString(txtNum6.Text.Split(Calculator.Delimiter)));
                        nums2check.Add(Calculator.GetCorrectNumberFromSplitedString(txtNum7.Text.Split(Calculator.Delimiter)));

                        foreach (int n in nums2check)
                        {
                            resCC = resCC || Calculator.isCarmaticNumber(n);
                        }
                        #endregion

                        bool resCL = false;
                        #region Check for carmatic numbers in LifeCycle Values
                        nums2check.Clear();
                        #region cycle number....
                        int curCycle = 0;
                        int age = Convert.ToInt16(txtAge.Text.Split("(".ToCharArray()[0])[0]);
                        int[] CycleBoundries = new int[4];
                        CycleBoundries[0] = Convert.ToInt16(txt1_1.Text.Split("-".ToCharArray()[0])[1].Trim());
                        CycleBoundries[1] = Convert.ToInt16(txt2_1.Text.Split("-".ToCharArray()[0])[1].Trim());
                        CycleBoundries[2] = Convert.ToInt16(txt3_1.Text.Split("-".ToCharArray()[0])[1].Trim());
                        CycleBoundries[3] = Convert.ToInt16(txt4_1.Text.Split("-".ToCharArray()[0])[1].Trim());

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
                        #endregion cycle number....
                        nums2check.Add(Calculator.GetCorrectNumberFromSplitedString(this.Controls.Find("txt" + curCycle.ToString() + "_4", true)[0].Text.Split(Calculator.Delimiter)));
                        nums2check.Add(Calculator.GetCorrectNumberFromSplitedString(this.Controls.Find("txt" + curCycle.ToString() + "_3", true)[0].Text.Split(Calculator.Delimiter)));
                        nums2check.Add(Calculator.GetCorrectNumberFromSplitedString(this.Controls.Find("txt" + curCycle.ToString() + "_2", true)[0].Text.Split(Calculator.Delimiter)));


                        foreach (int n in nums2check)
                        {
                            resCL = resCL || Calculator.isCarmaticNumber(n);
                        }
                        #endregion

                        if ((res1 | res2 | res3) && ((resCL == false) && (resCC == false)))
                        {
                            if (tableTestValue == "חצי מאוזן")
                            {
                                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                                {
                                    outStr = "מאוזן";
                                }
                                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                                {
                                    outStr = "Balanced";
                                }
                            }
                        }
                        #endregion
                    }
                }
            }

            #region Keep a-side // txtMapBalanceValue
            if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
            {
                if (outStr == "מאוזן")
                {
                    txtMapBalanceValue.Text = EnumProvider.Balance._מאוזן_.ToString();
                }
                if (outStr == "לא מאוזן")
                {
                    txtMapBalanceValue.Text = EnumProvider.Balance._לא_מאוזן_.ToString();
                }
                if (outStr == "חצי מאוזן")
                {
                    txtMapBalanceValue.Text = EnumProvider.Balance._מאוזן_חלקית_.ToString();
                }
            }
            if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
            {
                if (outStr == "Balanced")
                {
                    txtMapBalanceValue.Text = EnumProvider.Balance._מאוזן_.ToString();
                }
                if (outStr == "Not Balanced")
                {
                    txtMapBalanceValue.Text = EnumProvider.Balance._לא_מאוזן_.ToString();
                }
                if (outStr == "Half Balanced")
                {
                    txtMapBalanceValue.Text = EnumProvider.Balance._מאוזן_חלקית_.ToString();
                }
            }
            #endregion
            txtInfo1.Text += System.Environment.NewLine + outStr;
            string sMapCond = outStr;

            setSytle4UserInfo(txtInfo1, outStr);
            #endregion User_info

            List<int> ckra = new List<int>();
            ckra.Add(Calculator.GetCorrectNumberFromSplitedString(txtNum1.Text.Split(Calculator.Delimiter)));
            ckra.Add(Calculator.GetCorrectNumberFromSplitedString(txtNum2.Text.Split(Calculator.Delimiter)));
            ckra.Add(Calculator.GetCorrectNumberFromSplitedString(txtNum5.Text.Split(Calculator.Delimiter)));
            ckra.Add(Calculator.GetCorrectNumberFromSplitedString(txtNum7.Text.Split(Calculator.Delimiter)));
            ckra.Add(Calculator.GetCorrectNumberFromSplitedString(txtPName_Num.Text.Split(Calculator.Delimiter)));
            ckra.Add(Convert.ToInt16(txtAstroName.Text.Split(" ".ToCharArray())[1].Replace("(", "").Replace(")", "")));

            bool MapStrong = Calculator.isMapStrong(ckra, // כל הצאקרות
                                        Calculator.GetCorrectNumberFromSplitedString(txtNum3.Text.Split(Calculator.Delimiter)), // גרון
                                        Calculator.GetCorrectNumberFromSplitedString(txtNum6.Text.Split(Calculator.Delimiter))); // מין ויצירה

            txtMapStrong.Text = MapStrong.ToString();

            txtInfo1.Text += Environment.NewLine;
            if (MapStrong == true)
            {
                if (BLL.AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                {
                    txtInfo1.Text += "המפה חזקה";
                }
                if (BLL.AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                {
                    txtInfo1.Text += "Map is strong";
                }
            }
            else //(MapStrong == false)
            {
                if (BLL.AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                {
                    txtInfo1.Text += "המפה חלשה";
                }
                if (BLL.AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                {
                    txtInfo1.Text += "Map is weak";
                }
            }
            //txtInfo1.Text += Environment.NewLine;

            if (Calculator.isCarmaticNumber(Calculator.GetCorrectNumberFromSplitedString(txtPName_Num.Text.Split(Calculator.Delimiter))))
            {
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                {
                    txtInfo1.Text += System.Environment.NewLine + "שם פרטי לא מאוזן";
                }
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                {
                    txtInfo1.Text += System.Environment.NewLine + "Personal name is not balanced";
                }
            }

            int cc = CalcCurrentCycle();
            string[] schk = new string[6] { txtNum1.Text, txtNum2.Text, txtNum3.Text, txtNum5.Text,
                                            (this.Controls.Find("txt"+cc.ToString()+"_2",true)[0] as TextBox).Text,
                                            (this.Controls.Find("txt"+cc.ToString()+"_3",true)[0] as TextBox).Text};

            //string[] schk = new string[] { txtNum1.Text, txtNum2.Text, txtNum3.Text, txtNum5.Text };

            string[] s = txtAstroName.Text.Split("(".ToCharArray()[0]);
            string[] s2 = s[1].Split(")".ToCharArray()[0]);
            int iAstroNum = Convert.ToInt16(s2[0]);
            txtDepMap.Text = Calculator.CheckDepresedMap(schk, txtPName_Num.Text, iAstroNum);
            if (txtDepMap.Text != "")
            {
                txtInfo1.Text += System.Environment.NewLine + txtDepMap.Text;

            }

            schk = new string[6] {  txtNum1.Text, txtNum2.Text, txtNum3.Text, txtNum5.Text,
                                    (this.Controls.Find("txt"+cc.ToString()+"_2",true)[0] as TextBox).Text,
                                    (this.Controls.Find("txt"+cc.ToString()+"_3",true)[0] as TextBox).Text};

            //schk = new string[] { txtNum1.Text, txtNum2.Text, txtNum3.Text, txtNum5.Text };
            int firstClimax = Calculator.GetCorrectNumberFromSplitedString(txt1_3.Text.Split(Calculator.Delimiter)); // שיא ראשון
            outStr = Calculator.CheckStress(schk, txtPName_Num.Text, iAstroNum, firstClimax);
            if (outStr != "")
            {
                txtInfo1.Text += System.Environment.NewLine + outStr;
            }

            schk = new string[6] { txtNum1.Text, txtNum2.Text, txtNum3.Text, txtNum5.Text, txtNum6.Text, txtNum7.Text };
            outStr = Calculator.CheckAttention(schk, txtPName_Num.Text, iAstroNum);
            if (outStr != "")
            {
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                {
                    if ((sMapCond == "מאוזן"))
                    {
                        outStr += " " + "קלות";
                    }
                    else
                    {
                        outStr += " " + "מורכבות";
                    }
                }
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                {
                    if ((sMapCond == "Balanced"))
                    {
                        outStr += " " + "lite-ly";
                    }
                    else
                    {
                        outStr += " " + "in a complex manner";
                    }
                }

                txtInfo1.Text += System.Environment.NewLine + outStr;
            }

            //schk = new string[6] { txtNum1.Text, txtNum2.Text, txtNum3.Text, txtNum5.Text, txtNum6.Text, txtNum7.Text };
            schk = new string[] { txtNum1.Text, txtNum3.Text, txtNum5.Text };
            outStr = Calculator.CheckFearFromUnkown(schk, txtPName_Num.Text, iAstroNum);
            if (outStr != "")
            {
                txtInfo1.Text += System.Environment.NewLine + outStr;
            }

            // If there is 2,7,11 in the throat, set as weak map instead
            string num = Calculator.GetCorrectNumberFromSplitedString(txtNum3.Text.Split(Calculator.Delimiter)).ToString();
            if (new string[] { "2", "7", "11" }.Any(n => n.Equals(num)))
            {
                if (MapStrong == true)
                {
                    if (BLL.AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                    {
                        txtInfo1.Text = txtInfo1.Text.Replace("המפה חזקה", "המפה חלשה");

                    }
                    if (BLL.AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                    {
                        txtInfo1.Text = txtInfo1.Text.Replace("Map is strong", "Map is weak");

                    }

                }
                else
                {
                    if (BLL.AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                    {
                        txtInfo1.Text += "\nהמפה חלשה";

                    }
                    if (BLL.AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                    {
                        txtInfo1.Text += "\nMap is weak";

                    }

                }

            }

        }

        /// <summary>
        /// Dynamic Chkra Openning
        /// </summary>
        private void DynamicChakraOpening()
        {
            /*for (int i = 1; i < 10; i++)
            {
                string sCntrlName;
                string[] sVals = new string[] { };
                #region  collect numbers
                if (i != 9)
                {
                    sCntrlName = "txtNum" + i.ToString();
                    sVals = this.Controls.Find(sCntrlName, true)[0].Text.Trim().Split(Calculator.Delimiter);
                }
                else // astro data
                {
                    sVals = new string[1];
                    string stmp = txtAstroName.Text.Split(" ".ToCharArray())[1].Substring(1, 2);
                    sVals[0] = stmp.Split(")".ToCharArray())[0];
                }

                int[] nums = new int[sVals.Length];
                for (int j = 0; j < nums.Length; j++)
                {
                    nums[j] = Convert.ToInt16(sVals[j]);
                }
                #endregion

                string cFound = "";
                if (Calculator.isDynamicChakraOpen(i, nums, out cFound) == true)
                {
                    sCntrlName = "txtDC" + i.ToString();
                    this.Controls.Find(sCntrlName, true)[0].Text = cFound;
                }
                else
                {
                    PictureBox p = this.Controls.Find("picDC" + i.ToString(), true)[0] as PictureBox;
                    p.Image = Omega.Properties.Resources.NotOpen1;
                }
            } //for (int i = 1; i < 10; i++)*/

            for (int i = 1; i < 10; i++)
            {
                string sCntrlName;
                string chakraCntrlName;
                int[] sVals = new int[] { };
                chakraCntrlName = "txtNum" + i.ToString();

                //בשביל צאקרת כתר
                if (i == 1) {
                    sVals = new int[] { 1, 3, 4, 5, 6, 8, 9, 11, 22, 33 };
                }
                //בשביל צאקרת עין שלישית
                if (i == 2)
                {
                    sVals = new int[] { 1, 2, 3, 4, 5, 6, 8, 9, 11, 22, 33 };
                }
                //בשביל צאקרת הגרון
                if (i == 3)
                {
                    sVals = new int[] { 1, 3, 4, 5, 6, 8, 9, 22, 33 };
                }
                //בשביל צאקרת הלב
                if (i == 4)
                {
                    sVals = new int[] { 1, 2, 3, 4, 5, 6, 7, 8 };
                }
                //בשביל צאקרת מקלעת השמש
                if (i == 5)
                {
                    sVals = new int[] { 1, 3, 4, 5, 6, 8, 9, 22, 33};
                }
                //בשביל צאקרת מין ויצירה
                if (i == 6)
                {
                    sVals = new int[] { 1, 3, 5, 8};
                }
                //בשביל צאקרת בסיס
                if (i == 7)
                {
                    sVals = new int[] { 1, 2, 3, 4, 5, 6, 7, 8, 9, 11, 22, 33};
                }
                //בשביל צאקרת העל
                if (i == 8)
                {
                    sVals = new int[] { 1, 2, 3, 4, 5, 6, 7, 8, 9, 11, 22, 33};
                }
                //בשביל צאקרת היקום
                if (i == 9)
                {
                    sVals = new int[] { 1, 3, 4, 5, 6, 8, 9, 22, 33}; 
                }

                int cfoundNumber = 0;
                if ( i == 9)
                {
                    // astro data
                    cfoundNumber = Calculator.GetUniverseNumberFromChakra(this.Controls.Find(chakraCntrlName, true)[0].Text.Trim());
                }
                else
                {
                    cfoundNumber = Calculator.GetCorrectNumberFromSplitedString(this.Controls.Find(chakraCntrlName, true)[0].Text.Trim().Split(Calculator.Delimiter));
                }
                if (sVals.Contains(cfoundNumber))
                {
                    sCntrlName = "txtDC" + i.ToString();
                    this.Controls.Find(sCntrlName, true)[0].Text = cfoundNumber.ToString();

                   //this.Controls.Find(chakraCntrlName, true)[0].Text;
                }
                else
                {
                    PictureBox p = this.Controls.Find("picDC" + i.ToString(), true)[0] as PictureBox;
                    p.Image = Omega.Properties.Resources.NotOpen1;
                }
            }
        }

        private string sCalcAge_wMonths(DateTime bDay)
        {
            DateTime today = DateTimePickerTo.Value;
            TimeSpan ts = today.Subtract(bDay);
            int years = today.Year - bDay.Year; //Math.Floor(Convert.ToDouble(ts.Days / 366)).ToString(); 

            DateTime moovedBDay = bDay.AddYears(years);
            ts = today.Subtract(moovedBDay);
            int months = (int)Math.Floor(ts.Days / 30.00);

            if (months < 0)
            {
                years -= 1;
                months += 12;
            }

            return "" + years.ToString() + " (" + months.ToString() + ")";
        }

        /// <summary> Changes the style of the textboxes and pictures according to thier contents
        /// 
        /// </summary>
        private void SetStylesToObjects()
        {
            #region General Calc
            if (cbMainCorrectionDone.Checked == true)
            {
                #region txtNum1
                string oriVal = txtNum1.Text;
                string[] curSplt = txtNum1.Text.Split("=".ToCharArray()[0]);

                txtNum1.Text = curSplt[curSplt.Length - 1];
                ChangeStyleTextboxAndPictureBox(txtNum1, picSml1);

                txtNum1.Text = oriVal;
                #endregion

                #region txtNum5
                oriVal = txtNum5.Text;
                curSplt = txtNum5.Text.Split("=".ToCharArray()[0]);

                txtNum5.Text = curSplt[curSplt.Length - 1];
                ChangeStyleTextboxAndPictureBox(txtNum5, picSml5);

                txtNum5.Text = oriVal;
                #endregion

                //ChangeStyleTextboxAndPictureBox(txtSpiralCode, picSmlMik); // miklaat
                #region txtSpiralCode
                oriVal = txtSpiralCode.Text;
                curSplt = txtSpiralCode.Text.Split("=".ToCharArray()[0]);

                txtSpiralCode.Text = curSplt[curSplt.Length - 1];
                ChangeStyleTextboxAndPictureBox(txtSpiralCode, picSmlMik);

                txtSpiralCode.Text = oriVal;
                #endregion
            }
            else
            {
                ChangeStyleTextboxAndPictureBox(txtNum1, picSml1);
                ChangeStyleTextboxAndPictureBox(txtNum5, picSml5);
                ChangeStyleTextboxAndPictureBox(txtSpiralCode, picSmlMik); // miklaat
            }

            ChangeStyleTextboxAndPictureBox(txtNum2, picSml2);
            ChangeStyleTextboxAndPictureBox(txtNum3, picSml3);

            ChangeStyleTextboxAndPictureBox(txtNum4, picSml4);

            ChangeStyleTextboxAndPictureBox(txtNum6, picSml6);
            ChangeStyleTextboxAndPictureBox(txtNum7, picSml7);
            ChangeStyleTextboxAndPictureBox(txtPName_Num, picFName);
            ChangeStyleTextboxAndPictureBox(txtFName_Num, picLName);
            ChangeStyleTextboxAndPictureBox(txtParentsPresent, picParentsPresent);
            ChangeStyleTextboxAndPictureBox(txtUnique, picUnique);

            ChangeStyleForTextbox_Astro(txtAstroName);

            ChangeStyleForTextbox_Astro(txtNum9);
            ChangeStyleTextboxAndPictureBox(txtNum8, picSml9);

            txtPowerNum.Text = txtNum2.Text;
            //ChangeStyleTextboxAndPictureBox(txtPowerNum, picPowerNum);

            ChangeStyleByHousing();
            #endregion

            #region Life Cycles
            ChangeStyleTextboxAndPictureBox(txt1_2, LCpic1_2);
            ChangeStyleTextboxAndPictureBox(txt2_2, LCpic2_2);
            ChangeStyleTextboxAndPictureBox(txt3_2, LCpic3_2);
            ChangeStyleTextboxAndPictureBox(txt4_2, LCpic4_2);

            ChangeStyleTextboxAndPictureBox(txt1_3, LCpic1_3);
            ChangeStyleTextboxAndPictureBox(txt2_3, LCpic2_3);
            ChangeStyleTextboxAndPictureBox(txt3_3, LCpic3_3);
            ChangeStyleTextboxAndPictureBox(txt4_3, LCpic4_3);

            ChangeStyleTextboxAndPictureBox(txt1_4, LCpic1_4);
            ChangeStyleTextboxAndPictureBox(txt2_4, LCpic2_4);
            ChangeStyleTextboxAndPictureBox(txt3_4, LCpic3_4);
            ChangeStyleTextboxAndPictureBox(txt4_4, LCpic4_4);

            ChangeStyleTextboxAndPictureBox(txtPDay, LCpicD);
            ChangeStyleTextboxAndPictureBox(txtPMonth, LCpicM);
            ChangeStyleTextboxAndPictureBox(txtPYear, LCpicY);

            ChangeStyleForTextbox_Astro(txtLuck);


            #endregion

            #region Life Cycles Extra
            ChangeStyleTextboxAndPictureBox(txtX1_2, LCpicX1_2);
            ChangeStyleTextboxAndPictureBox(txtX2_2, LCpicX2_2);
            ChangeStyleTextboxAndPictureBox(txtX3_2, LCpicX3_2);
            ChangeStyleTextboxAndPictureBox(txtX4_2, LCpicX4_2);

            ChangeStyleTextboxAndPictureBox(txtX1_3, LCpicX1_3);
            ChangeStyleTextboxAndPictureBox(txtX2_3, LCpicX2_3);
            ChangeStyleTextboxAndPictureBox(txtX3_3, LCpicX3_3);
            ChangeStyleTextboxAndPictureBox(txtX4_3, LCpicX4_3);

            ChangeStyleTextboxAndPictureBox(txtX1_4, LCpicX1_4);
            ChangeStyleTextboxAndPictureBox(txtX2_4, LCpicX2_4);
            ChangeStyleTextboxAndPictureBox(txtX3_4, LCpicX3_4);
            ChangeStyleTextboxAndPictureBox(txtX4_4, LCpicX4_4);

            ChangeStyleTextboxAndPictureBox(txtPDayX, LCpicDX);
            ChangeStyleTextboxAndPictureBox(txtPMonthX, LCpicMX);
            ChangeStyleTextboxAndPictureBox(txtPYearX, LCpicYX);

            ChangeStyleForTextbox_Astro(txtXLuck);

            #endregion

            #region Intensive Map
            ChangeStyleForPicsOnly(txtIMap1, IMpic1);
            ChangeStyleForPicsOnly(txtIMap2, IMpic2);
            ChangeStyleForPicsOnly(txtIMap3, IMpic3);
            ChangeStyleForPicsOnly(txtIMap4, IMpic4);
            ChangeStyleForPicsOnly(txtIMap5, IMpic5);
            ChangeStyleForPicsOnly(txtIMap6, IMpic6);
            ChangeStyleForPicsOnly(txtIMap7, IMpic7);
            ChangeStyleForPicsOnly(txtIMap8, IMpic8);
            ChangeStyleForPicsOnly(txtIMap9, IMpic9);
            #endregion

            #region Pitagors Squares
            ChangeStyleForPicsOnly(txtPit1, PSpic1);
            ChangeStyleForPicsOnly(txtPit2, PSpic2);
            ChangeStyleForPicsOnly(txtPit3, PSpic3);
            ChangeStyleForPicsOnly(txtPit4, PSpic4);
            ChangeStyleForPicsOnly(txtPit5, PSpic5);
            ChangeStyleForPicsOnly(txtPit6, PSpic6);
            ChangeStyleForPicsOnly(txtPit7, PSpic7);
            ChangeStyleForPicsOnly(txtPit8, PSpic8);
            ChangeStyleForPicsOnly(txtPit9, PSpic9);
            #endregion

            #region Chakras
            //if (txtCMaster.Text != "") picMaster.Visible = true;
            //ChangeChakraCircle(txtCCrown, picCrown, Omega.Properties.Resources.Crown1);
            //ChangeChakraCircle(txtCThirdEye, picThirdEye, Omega.Properties.Resources.ThirdEye);
            //ChangeChakraCircle(txtCMikSun, picMikSun, Omega.Properties.Resources.MikSun);
            //ChangeChakraCircle(txtCSexCreat, picSexCreat, Omega.Properties.Resources.SexCreation);
            //ChangeChakraCircle(txtCThrought, picThrought, Omega.Properties.Resources.Throught);
            //ChangeChakraCircle(txtCHeart, picHeart, Omega.Properties.Resources.Heart);
            //ChangeChakraCircle(txtCRoot, picRoot, Omega.Properties.Resources.Root);
            #endregion

            #region Combined Map
            ChangeStyleForPicsOnly(txtMComb1, picMComb1);
            ChangeStyleForPicsOnly(txtMComb2, picMComb2);
            ChangeStyleForPicsOnly(txtMComb3, picMComb3);
            ChangeStyleForPicsOnly(txtMComb4, picMComb4);
            ChangeStyleForPicsOnly(txtMComb5, picMComb5);
            ChangeStyleForPicsOnly(txtMComb6, picMComb6);
            ChangeStyleForPicsOnly(txtMComb7, picMComb7);
            ChangeStyleForPicsOnly(txtMComb8, picMComb8);
            ChangeStyleForPicsOnly(txtMComb9, picMComb9);
            #endregion
        }

        #region Styles

        private void ChangeStyleByHousing()
        {
            string[] s = txtAstroName.Text.Split("(".ToCharArray()[0]);
            string[] s2 = s[1].Split(")".ToCharArray()[0]);
            int iAstroNum = Convert.ToInt16(s2[0]);

            int iApp = Calculator.GetCorrectNumberFromSplitedString(Calculator.CalcSum(Convert.ToInt16(mAppNum)).Split(Calculator.Delimiter));

            if (iAstroNum == iApp)
            {
                //picApp.Visible = true;
            }

            string sCityNum = Calculator.PrivateNameCalc(mCity);
            int iCityNum = Calculator.GetCorrectNumberFromSplitedString(sCityNum.Split(Calculator.Delimiter));

            if (iAstroNum == iCityNum)
            {
                //picCity.Visible = true;
            }
        }

        private void setSytle4UserInfo(TextBox tmpText, string inStr)
        {
            switch (inStr)
            {
                case "מאוזן":
                case "Balanced":
                    txtInfo1.BackColor = mGreen;
                    break;
                case "חצי מאוזן":
                case "Half Balanced":
                    txtInfo1.BackColor = mBlue;
                    break;
                case "לא מאוזן":
                case "Not Balanced":
                    txtInfo1.BackColor = mBlack;
                    txtInfo1.ForeColor = mRegularColor;
                    break;
            }
        }

        private void ChangeStyleTextboxAndPictureBox(TextBox tmpTextBox, PictureBox pic)
        {
            string[] txtNames;

            if (pic == null)
            {
                pic = new PictureBox();
            }

            if (tmpTextBox.Text == "") return;

            string[] s = tmpTextBox.Text.Split(Calculator.Delimiter);
            int testedNum = Calculator.GetCorrectNumberFromSplitedString(tmpTextBox.Text.Split(Calculator.Delimiter));

            int testIndex = 0;
            if (s.Length > 1) // there are more than 1 number -> ## = #
            {
                testIndex = s.Length - 1;
            }

            //int tmpnum = Convert.ToInt16(s[testIndex]);
            //if ((Calculator.isMaterNumber(tmpnum) == false) && (Calculator.isCarmaticNumber(tmpnum) == false) && (Calculator.isHalfCarmaticNumber(tmpnum) == false))
            //{
            //    if (mLuckNumber == tmpnum)
            //    pic.Visible = true;
            //}

            //for (int i = 0; i < s.Length; i++)
            //{
            bool rem = false;
            if (mLuckNumber == testedNum)//Calculator.GetCorrectNumberFromSplitedString(Calculator.CalcSum(Convert.ToInt16(s[i])).Split(Calculator.Delimiter)))
            {
                //pic.Visible = true;
                rem = true;
            }

            if (Calculator.GetCorrectNumberFromSplitedString(txtNum5.Text.Split(Calculator.Delimiter)) == testedNum)
            {
                if (rem == false)
                {
                    pic.Visible = true;
                    pic.Image = Omega.Properties.Resources.Smiely9;
                }

                if (rem == true)
                {
                    //pic.Visible = true;
                    pic.Image = Omega.Properties.Resources.Smiely_Half;
                }
            }
            //}

            //if (s.Length == 1) // only one number
            //{
            if (Calculator.isMaterNumber(testedNum))//Convert.ToInt16(s[0])))
            {
                tmpTextBox.BackColor = mMasterColor;
            }
            if (Calculator.isCarmaticNumber(testedNum))//Convert.ToInt16(s[0])))
            {
                tmpTextBox.BackColor = mBlack;
                tmpTextBox.ForeColor = mRegularColor;
            }
            if (Calculator.isHalfCarmaticNumber(testedNum))//Convert.ToInt16(s[0])))
            {
                tmpTextBox.BackColor = mHalfCarmaticColor;
            }

            if (cbPersonMaster.Checked && s.Length == 2 && s[1] == "2" && s[0] == "11")
            {
                // If person masterized, Set azure color to those chakras if contains 2                
                txtNames = new string[] { "txtNum1", "txtNum3", "txtNum5", "txtNum9" };

                if (txtNames.Any(t => t.Equals(tmpTextBox.Name, StringComparison.CurrentCultureIgnoreCase)))
                {
                    tmpTextBox.BackColor = mBlue;

                }

            }
            //else if (cbPersonMaster.Checked && txtNum9.Text.Contains("דלי") && s[0] == "11" && tmpTextBox.Name.Contains("txtNum9"))
            //{
            //    //// If person masterized, Set azure color to those chakras if contains 11                
            //    //txtNames = new string[] { "txtNum1", "txtNum3", "txtNum5", "txtNum9" };

            //    //if (txtNames.Any(t => t.Equals(tmpTextBox.Name, StringComparison.CurrentCultureIgnoreCase)))
            //    //{
            //        tmpTextBox.BackColor = mBlue;

            //    //}

            //}

            // Set azure color to those chakras if contains 7
            txtNames = new string[] { "txtNum1", "txtNum3", "txtNum5", "txtNum9" };
            int splittedVal = Calculator.GetCorrectNumberFromSplitedString(tmpTextBox.Text.Split(Calculator.Delimiter));

            if (txtNames.Any(t => t.Equals(tmpTextBox.Name, StringComparison.CurrentCultureIgnoreCase)) && new int[] { 2, 20, 7, 25 }.Contains(splittedVal))
            {
                tmpTextBox.BackColor = mBlue;

            }

            txtNames = new string[] { "txtNum2" };
            splittedVal = Calculator.GetCorrectNumberFromSplitedString(tmpTextBox.Text.Split(Calculator.Delimiter));

            if (txtNames.Any(t => t.Equals(tmpTextBox.Name, StringComparison.CurrentCultureIgnoreCase)) && new int[] { 7, 25 }.Contains(splittedVal))
            {
                tmpTextBox.BackColor = mBlue;

            }


            // Set azure color to those chakras if contains 2
            txtNames = new string[] { "txtNum5", "txtNum3", "txtNum1", "txtNum9" };

            if (txtNames.Any(t => t.Equals(tmpTextBox.Name, StringComparison.CurrentCultureIgnoreCase)) && tmpTextBox.Text.Equals("2"))
            {
                tmpTextBox.BackColor = mBlue;

            }

            // Set azure color to those cycles if contains 7 or 2 //txtX1_2
            txtNames = new string[] { "txt1_2", "txt2_2", "txt3_2", "txt4_2", "txt1_3", "txt2_3", "txt3_3", "txt4_3", "txtX2_2", "txtX3_2", "txtX4_2", "txtX1_3", "txtX2_3", "txtX3_3", "txtX4_3" };

            if (txtNames.Any(t => t.Equals(tmpTextBox.Name, StringComparison.CurrentCultureIgnoreCase)))
            {
                if ((tmpTextBox.Text.Equals("7")) || (tmpTextBox.Text.Equals("2")))
                {
                    tmpTextBox.BackColor = mBlue;

                }

            }

            // Set background to black if value is 0 in Heart/Throat/any of lifecycles
            txtNames = new string[] { "txtNum3", "txtNum4", "txt1_2", "txt2_2", "txt3_2", "txt4_2", "txtX1_2", "txtX2_2", "txtX3_2", "txtX4_2" };

            if (txtNames.Any(t => t.Equals(tmpTextBox.Name, StringComparison.CurrentCultureIgnoreCase)) && tmpTextBox.Text.Equals("0"))
            {
                tmpTextBox.BackColor = mBlack;
                tmpTextBox.ForeColor = mRegularColor;

            }

            //}
            //else
            //{
            //    if ( (Calculator.isMaterNumber(Convert.ToInt16(s[testIndex]))) || (Calculator.isMaterNumber(Convert.ToInt16(s[testIndex-1]))) )
            //    {
            //        tmpTextBox.BackColor = mMasterColor;
            //    }
            //    if ( (Calculator.isCarmaticNumber(Convert.ToInt16(s[testIndex]))) || (Calculator.isCarmaticNumber(Convert.ToInt16(s[testIndex-1]))) || (tmpTextBox.Text == "19=10=1"))
            //    {
            //        tmpTextBox.BackColor = mBlack;
            //        tmpTextBox.ForeColor = mRegularColor;
            //    }
            //    if ((Calculator.isHalfCarmaticNumber(Convert.ToInt16(s[testIndex]))) || (Calculator.isHalfCarmaticNumber(Convert.ToInt16(s[testIndex - 1]))))
            //    {
            //        tmpTextBox.BackColor = mHalfCarmaticColor;
            //    }
            //}

        }

        private void ChangeStyleForPicsOnly(TextBox tmpTextBox, PictureBox pic)
        {
            if (tmpTextBox.Text == "")
            {
                return;
            }

            int num = Convert.ToInt16(tmpTextBox.Name.Substring(tmpTextBox.Name.Length - 1, 1));
            if (num == mLuckNumber)
            {
                //pic.Visible = true;
            }
        }

        // For Astro_Luck Only
        private void ChangeStyleForTextbox_Astro(TextBox tmpTextBox)
        {
            if (tmpTextBox.Text == "")
                return;

            string[] s = tmpTextBox.Text.Split("(".ToCharArray()[0]);
            string[] s2 = s[1].Split(")".ToCharArray()[0]);
            int num = Convert.ToInt16(s2[0]);

            if (Calculator.isMaterNumber(num) == true)
            {
                tmpTextBox.BackColor = mMasterColor;
            }

            if ((Calculator.isDelayingNumber(num)) || ((num == 11) && (cbPersonMaster.Checked)))
            {
                tmpTextBox.BackColor = mBlue;

            }


        }

        private void ChangeChakraCircle(TextBox tmpTextBox, PictureBox tempPic, Bitmap CurImage)
        {
            if (tmpTextBox.Text != "")
            {
                tempPic.Image = CurImage;
            }
            else
            {
                tempPic.Image = Omega.Properties.Resources.NotOpen1;
            }
        }
        #endregion

        private void SetINtensiveMapValues(string[] Values)
        {
            txtIMap1.Text = Values[0];
            txtIMap2.Text = Values[1];
            txtIMap3.Text = Values[2];
            txtIMap4.Text = Values[3];
            txtIMap5.Text = Values[4];
            txtIMap6.Text = Values[5];
            txtIMap7.Text = Values[6];
            txtIMap8.Text = Values[7];
            txtIMap9.Text = Values[8];
        }

        private void SetPitagorasSquaresValues(string[] Values)
        {
            txtPit1.Text = Values[0];
            txtPit2.Text = Values[1];
            txtPit3.Text = Values[2];
            txtPit4.Text = Values[3];
            txtPit5.Text = Values[4];
            txtPit6.Text = Values[5];
            txtPit7.Text = Values[6];
            txtPit8.Text = Values[7];
            txtPit9.Text = Values[8];
        }

        private void SetLifeCycleValues(int[] Values)
        {
            txt1_1.Text = "0 - " + Values[0].ToString();
            txt2_1.Text = Values[0].ToString() + " - " + Values[4].ToString();
            txt3_1.Text = Values[4].ToString() + " - " + Values[8].ToString();
            txt4_1.Text = Values[8].ToString() + " - " + Values[12].ToString();

            txtX1_1.Text = txt1_1.Text;
            txtX2_1.Text = txt2_1.Text;
            txtX3_1.Text = txt3_1.Text;
            txtX4_1.Text = txt4_1.Text;

            txt1_2.Text = Calculator.CalcSum(Values[1]);
            txt2_2.Text = Calculator.CalcSum(Values[5]);
            txt3_2.Text = Calculator.CalcSum(Values[9]);
            txt4_2.Text = Calculator.CalcSum(Values[13]);

            txtX1_2.Text = txt1_2.Text;
            txtX2_2.Text = txt2_2.Text;
            txtX3_2.Text = txt3_2.Text;
            txtX4_2.Text = txt4_2.Text;

            txt1_3.Text = Calculator.CalcSum(Values[2]);
            txt2_3.Text = Calculator.CalcSum(Values[6]);

            txtX1_3.Text = txt1_3.Text;
            txtX2_3.Text = txt2_3.Text;

            txt3_3.Text = Calculator.CalcSum(Calculator.GetCorrectNumberFromSplitedString(txt2_3.Text.Split(Calculator.Delimiter)) + Calculator.GetCorrectNumberFromSplitedString(txt1_3.Text.Split(Calculator.Delimiter)));//Calculator.CalcSum(Values[10]);
            txt4_3.Text = Calculator.CalcSum(Values[14]);

            txtX3_3.Text = txt3_3.Text;
            txtX4_3.Text = txt4_3.Text;

            txt1_4.Text = Values[3].ToString();
            txt2_4.Text = Values[7].ToString();
            txt3_4.Text = Values[11].ToString();
            txt4_4.Text = Values[15].ToString();

            txtX1_4.Text = txt1_4.Text;
            txtX2_4.Text = txt2_4.Text;
            txtX3_4.Text = txt3_4.Text;
            txtX4_4.Text = txt4_4.Text;

        }

        private void SetSumTableFromTextBoxes(TextBox[] arrTB, DataGridView curDGV)
        {
            curDGV.Rows.Clear();

            object[] objStrings = new string[2];
            objStrings.Initialize();

            int sum;
            int[] CurIndex;
            string[] sSplit;

            #region Physical
            CurIndex = new int[2] { 4, 5 };
            sum = 0;
            for (int i = 0; i < CurIndex.Length; i++)
            {
                if (arrTB[CurIndex[i] - 1].Text != "")
                {
                    sSplit = arrTB[CurIndex[i] - 1].Text.Split(",".ToCharArray()[0]);
                    sum += sSplit.Length;
                }
            }
            if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
            {
                objStrings.SetValue("פיזי", 0);
            }
            if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
            {
                objStrings.SetValue("Physical", 0);
            }
            objStrings.SetValue(sum.ToString(), 1);
            curDGV.Rows.Insert(curDGV.Rows.Count, objStrings);
            #endregion

            #region Emotional
            CurIndex = new int[3] { 2, 3, 6 };
            sum = 0;
            for (int i = 0; i < CurIndex.Length; i++)
            {
                if (arrTB[CurIndex[i] - 1].Text != "")
                {
                    sSplit = arrTB[CurIndex[i] - 1].Text.Split(",".ToCharArray()[0]);
                    sum += sSplit.Length;
                }
            }
            if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
            {
                objStrings.SetValue("רגשי", 0);
            }
            if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
            {
                objStrings.SetValue("Emotional", 0);
            }
            objStrings.SetValue(sum.ToString(), 1);
            curDGV.Rows.Insert(curDGV.Rows.Count, objStrings);
            #endregion

            #region Mental
            // 1 + 8
            CurIndex = new int[2] { 1, 8 };
            sum = 0;
            for (int i = 0; i < CurIndex.Length; i++)
            {
                if (arrTB[CurIndex[i] - 1].Text != "")
                {
                    sSplit = arrTB[CurIndex[i] - 1].Text.Split(",".ToCharArray()[0]);
                    sum += sSplit.Length;
                }
            }
            if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
            {
                objStrings.SetValue("מנטאלי", 0);
            }
            if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
            {
                objStrings.SetValue("Mental", 0);
            }
            objStrings.SetValue(sum.ToString(), 1);
            curDGV.Rows.Insert(curDGV.Rows.Count, objStrings);
            #endregion

            #region Enrgettic
            // 7 + 9
            CurIndex = new int[2] { 7, 9 };
            sum = 0;
            for (int i = 0; i < CurIndex.Length; i++)
            {
                if (arrTB[CurIndex[i] - 1].Text != "")
                {
                    sSplit = arrTB[CurIndex[i] - 1].Text.Split(",".ToCharArray()[0]);
                    sum += sSplit.Length;
                }
            }
            if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
            {
                objStrings.SetValue("אנרגטי", 0);
            }
            if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
            {
                objStrings.SetValue("Energetic", 0);
            }
            objStrings.SetValue(sum.ToString(), 1);
            curDGV.Rows.Insert(curDGV.Rows.Count, objStrings);
            #endregion



            //curDGV.Rows.Insert(curDGV.Rows.Count, objStrings);
            curDGV.Update();

            curDGV.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            curDGV.Columns[1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

            curDGV.Columns[0].Width = 70;
            curDGV.Columns[1].Width = 40;

            curDGV.Update();
            curDGV.Refresh();
            curDGV.ResumeLayout(true);
            curDGV.Show();
        }

        private void SingleBusinessMatchCalc(string BonusYesOrNo, string SelfFixYesOrNo)
        {
            List<string> inBusinessValues = new List<string>();
            int age = Convert.ToInt16(txtAge.Text.Split("(".ToCharArray()[0])[0]);
            age = Convert.ToInt16(DateTimePickerTo.Value.Subtract(mB_Date).TotalDays / 365);

            #region From Chakra Map
            int upperBndry = Convert.ToInt16(txt1_1.Text.Split("-".ToCharArray()[0])[1].Trim());

            int tmpVal = 0;
            if (age <= upperBndry)
            {
                // Crown
                tmpVal = Calculator.GetCorrectNumberFromSplitedString(txtNum1.Text.Split(Calculator.Delimiter));
            }
            else
            {
                //Third Eye
                tmpVal = Calculator.GetCorrectNumberFromSplitedString(txtNum2.Text.Split(Calculator.Delimiter));
            }
            inBusinessValues.Add(tmpVal.ToString());


            tmpVal = Calculator.GetCorrectNumberFromSplitedString(txtNum3.Text.Split(Calculator.Delimiter));
            inBusinessValues.Add(tmpVal.ToString());

            tmpVal = Calculator.GetCorrectNumberFromSplitedString(txtNum5.Text.Split(Calculator.Delimiter));
            inBusinessValues.Add(tmpVal.ToString());

            tmpVal = Calculator.GetCorrectNumberFromSplitedString(txtNum6.Text.Split(Calculator.Delimiter));
            inBusinessValues.Add(tmpVal.ToString());

            tmpVal = Calculator.GetCorrectNumberFromSplitedString(txtPName_Num.Text.Split(Calculator.Delimiter));
            inBusinessValues.Add(tmpVal.ToString());

            tmpVal = Convert.ToInt16(txtAstroName.Text.Split("(".ToCharArray()[0])[1].Split(")".ToCharArray()[0])[0]);
            int tmpastronum = tmpVal;
            inBusinessValues.Add(tmpVal.ToString());

            string tmpLvl = txtInfo1.Text.Split(System.Environment.NewLine.ToCharArray()[0])[1].Trim();
            if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
            {
            }
            if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
            {
                if (tmpLvl == "Balanced")
                {
                    tmpLvl = "מאוזן";
                }
                if (tmpLvl == "Not Balanced")
                {
                    tmpLvl = "לא מאוזן";
                }
                if (tmpLvl == "Half Balanced")
                {
                    tmpLvl = "חצי מאוזן";
                }
            }
            inBusinessValues.Add(tmpLvl);
            #endregion
            if (isMultiBusineesCalc == true)
            {
                for (int i = 0; i < inBusinessValues.Count; i++)
                {
                    if (inBusinessValues[i] == "2")
                    {
                        // fix to 7 in the business Chakra chart
                        inBusinessValues.RemoveAt(i);
                        inBusinessValues.Insert(i, "33"); // fir grade 7
                    }
                    if (inBusinessValues[i] == "4")
                    {
                        // fix to 7 in the business Chakra chart
                        inBusinessValues.RemoveAt(i);
                        inBusinessValues.Insert(i, "3");// fir grade 8
                    }
                }
            }

            double ChakraMapBusinessValue = Calculator.ChakraMap2BusinessValue(inBusinessValues);

            Control.ControlCollection cntrlCls = this.Controls;

            #region Life Cycles
            int curCycle = 0;
            int[] CycleBoundries = new int[4];
            CycleBoundries[0] = Convert.ToInt16(txt1_1.Text.Split("-".ToCharArray()[0])[1].Trim());
            CycleBoundries[1] = Convert.ToInt16(txt2_1.Text.Split("-".ToCharArray()[0])[1].Trim());
            CycleBoundries[2] = Convert.ToInt16(txt3_1.Text.Split("-".ToCharArray()[0])[1].Trim());
            CycleBoundries[3] = Convert.ToInt16(txt4_1.Text.Split("-".ToCharArray()[0])[1].Trim());

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

            int chlng, clmx, ccl;
            chlng = Calculator.GetCorrectNumberFromSplitedString(cntrlCls.Find("txt" + curCycle.ToString() + "_4", true)[0].Text.Split(Calculator.Delimiter));
            clmx = Calculator.GetCorrectNumberFromSplitedString(cntrlCls.Find("txt" + curCycle.ToString() + "_3", true)[0].Text.Split(Calculator.Delimiter));
            ccl = Calculator.GetCorrectNumberFromSplitedString(cntrlCls.Find("txt" + curCycle.ToString() + "_2", true)[0].Text.Split(Calculator.Delimiter));

            inBusinessValues = new List<string>();
            //inBusinessValues.Add(chlng.ToString());
            inBusinessValues.Add(clmx.ToString());
            inBusinessValues.Add(ccl.ToString());
            #endregion Life Cycles
            if (isMultiBusineesCalc == true)
            {
                for (int i = 0; i < inBusinessValues.Count; i++)
                {
                    if (inBusinessValues[i] == "2")
                    {
                        // fix to 7 in the business Chakra chart
                        inBusinessValues.RemoveAt(i);
                        inBusinessValues.Insert(i, "33");// for grade 7
                    }
                    if (inBusinessValues[i] == "4")
                    {
                        // fix to 7 in the business Chakra chart
                        inBusinessValues.RemoveAt(i);
                        inBusinessValues.Insert(i, "3");// for grade 8
                    }
                }
            }
            double LifeCycleBusinessValue = Calculator.LifeCycle2BusinessValue(inBusinessValues, "כן" == SelfFixYesOrNo);

            double sum = (ChakraMapBusinessValue + LifeCycleBusinessValue) / 2;

            if (((AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew) && (BonusYesOrNo == "כן")) || ((AppSettings.Instance.ProgramLanguage == AppSettings.Language.English) && (BonusYesOrNo == "Yes")))
            {
                sum += 1;
            }
            else // cmbBusinessBonus.SelectedItem.ToString() == "לא"
            {
                sum += 0;
            }


            if (Calculator.GetCorrectNumberFromSplitedString(txtNum6.Text.Split(Calculator.Delimiter)) == 8)
            {
                if (Calculator.GetCorrectNumberFromSplitedString(txtNum5.Text.Split(Calculator.Delimiter)) != 8)
                {
                    sum += 1;
                }
            }

            if (Calculator.GetCorrectNumberFromSplitedString(txtNum3.Text.Split(Calculator.Delimiter)) == 8)
            {
                sum += 0.5;
            }

            // If sex & creation equals 1, add 0.5 point as bonus
            if (Calculator.GetCorrectNumberFromSplitedString(txtNum6.Text.Split(Calculator.Delimiter)) == 1)
            {
                sum += 0.5;

            }
            #region Bunos of SelfFixing

            if (SelfFixYesOrNo == "כן")
            {
                bool resSelfFix = false;
                #region From LifeCycles
                for (int i = 1; i <= curCycle; i++)
                {
                    for (int c = 2; c < 5; c++)
                    {
                        int tmpNum = Calculator.GetCorrectNumberFromSplitedString(cntrlCls.Find("txt" + i.ToString() + "_" + c.ToString(), true)[0].Text.Split(Calculator.Delimiter));
                        if ((tmpNum == 0) || (Calculator.isMaterNumber(tmpNum)) || ((Calculator.isCarmaticNumber(tmpNum)) && (tmpNum != 16)))
                        {
                            resSelfFix = true;
                        }
                    }
                }

                for (int i = 1; i <= curCycle; i++)
                {
                    int tmpNum = Calculator.GetCorrectNumberFromSplitedString(cntrlCls.Find("txt" + i.ToString() + "_2", true)[0].Text.Split(Calculator.Delimiter));

                    if (tmpNum == 16)
                    {
                        sum = sum + 2;
                        i = 11; // exit loop
                    }
                }
                #endregion From LifeCycles

                if (resSelfFix == true)
                {
                    sum += 0.5;
                }
            }
            else // cmbSelfFix.SelectedItem.ToString() == "לא"
            {
                #region xondt
                //bool isFound = false;
                //#region From MapChakra
                //List<int> nums2check = new List<int>();
                //if (curCycle == 1)
                //{
                //    nums2check.Add(Calculator.GetCorrectNumberFromSplitedString(txtNum1.Text.Split(Calculator.Delimiter)));
                //}
                //else
                //{
                //    nums2check.Add(Calculator.GetCorrectNumberFromSplitedString(txtNum2.Text.Split(Calculator.Delimiter)));
                //}

                //nums2check.Add(Calculator.GetCorrectNumberFromSplitedString(txtNum3.Text.Split(Calculator.Delimiter)));

                //nums2check.Add(Calculator.GetCorrectNumberFromSplitedString(txtNum5.Text.Split(Calculator.Delimiter)));

                //nums2check.Add(Calculator.GetCorrectNumberFromSplitedString(txtNum6.Text.Split(Calculator.Delimiter)));

                //nums2check.Add(Calculator.GetCorrectNumberFromSplitedString(txtPName_Num.Text.Split(Calculator.Delimiter)));

                //for (int i = 0; i < nums2check.Count; i++)
                //{
                //    if (nums2check[i] == 16)
                //    {
                //        isFound = true;
                //    }
                //}
                //#endregion From MapChakra

                //if (isFound == true)
                //{
                //    sum -= 0.5;
                //}
                #endregion xondt
            }

            #endregion Bunos of SelfFixing

            #region Additional Msala
            if (isNeedToCalculateWithMsala)
            {
                var isFound = false;
                for (var i = 1; i < 5; i++) {
                    if (Calculator.GetCorrectNumberFromSplitedString(cntrlCls.Find("txtX" + i +"_2", true)[0].Text.Split(Calculator.Delimiter)) == 1 ||
                        Calculator.GetCorrectNumberFromSplitedString(cntrlCls.Find("txtX" + i + "_2", true)[0].Text.Split(Calculator.Delimiter)) == 8 ||
                        Calculator.GetCorrectNumberFromSplitedString(cntrlCls.Find("txtX" + i + "_2", true)[0].Text.Split(Calculator.Delimiter)) == 9 ||
                        Calculator.GetCorrectNumberFromSplitedString(cntrlCls.Find("txtX" + i + "_2", true)[0].Text.Split(Calculator.Delimiter)) == 11 ||
                        Calculator.GetCorrectNumberFromSplitedString(cntrlCls.Find("txtX" + i + "_2", true)[0].Text.Split(Calculator.Delimiter)) == 22)
                    {
                        sum += 0.5;
                        isFound = true;
                        break;
                    }
                }
                if (!isFound)
                {
                    for (var i = 1; i < 5; i++)
                    {
                        if (Calculator.GetCorrectNumberFromSplitedString(cntrlCls.Find("txtX" + i + "_3", true)[0].Text.Split(Calculator.Delimiter)) == 1 ||
                        Calculator.GetCorrectNumberFromSplitedString(cntrlCls.Find("txtX" + i + "_3", true)[0].Text.Split(Calculator.Delimiter)) == 8 ||
                        Calculator.GetCorrectNumberFromSplitedString(cntrlCls.Find("txtX" + i + "_3", true)[0].Text.Split(Calculator.Delimiter)) == 9 ||
                        Calculator.GetCorrectNumberFromSplitedString(cntrlCls.Find("txtX" + i + "_3", true)[0].Text.Split(Calculator.Delimiter)) == 11 ||
                        Calculator.GetCorrectNumberFromSplitedString(cntrlCls.Find("txtX" + i + "_3", true)[0].Text.Split(Calculator.Delimiter)) == 22)
                        {
                            sum += 0.5;
                            isFound = true;
                            break;
                        }
                    }
                }
            }
            #endregion

            if (sum > 10)
            {
                sum = 10;
            }
            txtFinalBusinessMark.Text = Math.Round(sum, 2).ToString();

            string story = "";
            if ((sum >= 9) & (sum <= 10))
            {
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                {
                    story = "סיכוי מצויין להצלחה עסקית ויכולת לעסק לשרוד מעל חמש שנים";
                }
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                {
                    story = "Excellent chance of business success and business ability to survive over five years";
                }
            }

            if ((sum >= 8) & (sum < 9))
            {
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                {
                    story = "סיכוי טוב מאוד להצלחה עיסקית ויכולת לעסק לשרוד מעל חמש שנים";
                }
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                {
                    story = "Very good chance of business success and business ability to survive over five years";
                }
            }

            if ((sum >= 7) & (sum < 8))
            {
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                {
                    story = "סיכוי בינוני להצלחה עסקית ויכולת לעסק לשרוד מעל חמש שנים";
                }
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                {
                    story = "Medium chance of business success and business ability to survive over five years";
                }
            }

            if ((sum >= 6) & (sum < 7))
            {
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                {
                    story = "סיכוי בינוני להצלחה עסקית ויכולת לעסק לשרוד עד שנתיים";
                }
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                {
                    story = "Medium chance of business success and business ability to survive up to two years";
                }
            }

            if ((sum >= 2) & (sum < 6))
            {
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                {
                    story = "סיכוי קטן מאוד להצלחה עסקית ויכולת לעסק לשרוד עד מספר חודשים";
                }
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                {
                    story = "Very little chance of business success and business ability to survive up to several months";
                }
            }

            if ((sum >= 0) & (sum < 2))
            {
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                {
                    story = "אין סיכוי להצלחה עסקית";
                }
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                {
                    story = "no chance for business success";
                }
            }

            if (tmpastronum == 22)
            {
                //if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                //{
                //    story += Environment.NewLine + "היות ומזלך האסטרולוגי עקרב, קל לך להיות עצמאי מצליח או עובד מוערך.";
                //}
                //if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                //{
                //    story += Environment.NewLine + "As astrological sign is Scorpio it is easy for you to be a successful indipendent bussiness or a highly estimated worker.";
                //}
            }

            txtBusinessStory.Text += "(" + txtFinalBusinessMark.Text + ") " + Environment.NewLine + story;
        }


        private void SingleSecreteryMatchCalc(bool isSecreatry, bool isAccounting)
        {
            List<string> inBusinessValues = new List<string>();
            int age = Convert.ToInt16(txtAge.Text.Split("(".ToCharArray()[0])[0]);
            age = Convert.ToInt16(DateTimePickerTo.Value.Subtract(mB_Date).TotalDays / 365);

            #region From Chakra Map
            int upperBndry = Convert.ToInt16(txt1_1.Text.Split("-".ToCharArray()[0])[1].Trim());

            int tmpVal = 0;
            if (age <= upperBndry)
            {
                // Crown
                tmpVal = Calculator.GetCorrectNumberFromSplitedString(txtNum1.Text.Split(Calculator.Delimiter));
            }
            else
            {
                //Third Eye
                tmpVal = Calculator.GetCorrectNumberFromSplitedString(txtNum2.Text.Split(Calculator.Delimiter));
            }
            inBusinessValues.Add(tmpVal.ToString());


            tmpVal = Calculator.GetCorrectNumberFromSplitedString(txtNum3.Text.Split(Calculator.Delimiter));
            inBusinessValues.Add(tmpVal.ToString());

            tmpVal = Calculator.GetCorrectNumberFromSplitedString(txtNum5.Text.Split(Calculator.Delimiter));
            inBusinessValues.Add(tmpVal.ToString());

            tmpVal = Calculator.GetCorrectNumberFromSplitedString(txtNum6.Text.Split(Calculator.Delimiter));
            inBusinessValues.Add(tmpVal.ToString());

            tmpVal = Calculator.GetCorrectNumberFromSplitedString(txtPName_Num.Text.Split(Calculator.Delimiter));
            inBusinessValues.Add(tmpVal.ToString());

            tmpVal = Convert.ToInt16(txtAstroName.Text.Split("(".ToCharArray()[0])[1].Split(")".ToCharArray()[0])[0]);
            int tmpastronum = tmpVal;
            inBusinessValues.Add(tmpVal.ToString());

            string tmpLvl = txtInfo1.Text.Split(System.Environment.NewLine.ToCharArray()[0])[1].Trim();
            if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
            {
            }
            if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
            {
                if (tmpLvl == "Balanced")
                {
                    tmpLvl = "מאוזן";
                }
                if (tmpLvl == "Not Balanced")
                {
                    tmpLvl = "לא מאוזן";
                }
                if (tmpLvl == "Half Balanced")
                {
                    tmpLvl = "חצי מאוזן";
                }
            }
            inBusinessValues.Add(tmpLvl);
            #endregion
            if (isMultiSecereteryCalc == true)
            {
                for (int i = 0; i < inBusinessValues.Count; i++)
                {
                    if (inBusinessValues[i] == "2")
                    {
                        // fix to 7 in the business Chakra chart
                        inBusinessValues.RemoveAt(i);
                        inBusinessValues.Insert(i, "33"); // fir grade 7
                    }
                    if (inBusinessValues[i] == "4")
                    {
                        // fix to 7 in the business Chakra chart
                        inBusinessValues.RemoveAt(i);
                        inBusinessValues.Insert(i, "3");// fir grade 8
                    }
                }
            }

            double ChakraMapBusinessValue = Calculator.ChakraMap2SecretaryAndAccountingValue(inBusinessValues);

            Control.ControlCollection cntrlCls = this.Controls;

            #region Life Cycles
            int curCycle = 0;
            int[] CycleBoundries = new int[4];
            CycleBoundries[0] = Convert.ToInt16(txt1_1.Text.Split("-".ToCharArray()[0])[1].Trim());
            CycleBoundries[1] = Convert.ToInt16(txt2_1.Text.Split("-".ToCharArray()[0])[1].Trim());
            CycleBoundries[2] = Convert.ToInt16(txt3_1.Text.Split("-".ToCharArray()[0])[1].Trim());
            CycleBoundries[3] = Convert.ToInt16(txt4_1.Text.Split("-".ToCharArray()[0])[1].Trim());

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

            int chlng, clmx, ccl;
            chlng = Calculator.GetCorrectNumberFromSplitedString(cntrlCls.Find("txt" + curCycle.ToString() + "_4", true)[0].Text.Split(Calculator.Delimiter));
            clmx = Calculator.GetCorrectNumberFromSplitedString(cntrlCls.Find("txt" + curCycle.ToString() + "_3", true)[0].Text.Split(Calculator.Delimiter));
            ccl = Calculator.GetCorrectNumberFromSplitedString(cntrlCls.Find("txt" + curCycle.ToString() + "_2", true)[0].Text.Split(Calculator.Delimiter));

            inBusinessValues = new List<string>();
            //inBusinessValues.Add(chlng.ToString());
            inBusinessValues.Add(clmx.ToString());
            inBusinessValues.Add(ccl.ToString());
            #endregion Life Cycles
            if (isMultiSecereteryCalc == true)
            {
                for (int i = 0; i < inBusinessValues.Count; i++)
                {
                    if (inBusinessValues[i] == "2")
                    {
                        // fix to 7 in the business Chakra chart
                        inBusinessValues.RemoveAt(i);
                        inBusinessValues.Insert(i, "33");// for grade 7
                    }
                    if (inBusinessValues[i] == "4")
                    {
                        // fix to 7 in the business Chakra chart
                        inBusinessValues.RemoveAt(i);
                        inBusinessValues.Insert(i, "3");// for grade 8
                    }
                }
            }
            double LifeCycleBusinessValue = Calculator.LifeCycle2SecretaryAndAccountingValue(inBusinessValues);

            double sum = (ChakraMapBusinessValue + LifeCycleBusinessValue) / 2;

            //if (((AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew) && (BonusYesOrNo == "כן")) || ((AppSettings.Instance.ProgramLanguage == AppSettings.Language.English) && (BonusYesOrNo == "Yes")))
            //{
            //    sum += 1;
            //}
            //else // cmbBusinessBonus.SelectedItem.ToString() == "לא"
            //{
            //    sum += 0;
            //}

            //Additional points
            // If throat equals 1 or 8, minus 0.5 point as bonus
            if (Calculator.GetCorrectNumberFromSplitedString(txtNum3.Text.Split(Calculator.Delimiter)) == 1 ||
                Calculator.GetCorrectNumberFromSplitedString(txtNum3.Text.Split(Calculator.Delimiter)) == 8)
            {
                sum -= 0.5;
            }

            // If sex & creation equals 1 or 8, minus 1 point as bonus
            if (Calculator.GetCorrectNumberFromSplitedString(txtNum6.Text.Split(Calculator.Delimiter)) == 1 ||
                Calculator.GetCorrectNumberFromSplitedString(txtNum6.Text.Split(Calculator.Delimiter)) == 8)
            {
                sum -= 1;
            }
            // If sex & creation equals 2 or 4 or 6 or 9 or 33
            if (Calculator.GetCorrectNumberFromSplitedString(txtNum6.Text.Split(Calculator.Delimiter)) == 2 ||
                Calculator.GetCorrectNumberFromSplitedString(txtNum6.Text.Split(Calculator.Delimiter)) == 4 ||
                Calculator.GetCorrectNumberFromSplitedString(txtNum6.Text.Split(Calculator.Delimiter)) == 6 ||
                Calculator.GetCorrectNumberFromSplitedString(txtNum6.Text.Split(Calculator.Delimiter)) == 9 ||
                Calculator.GetCorrectNumberFromSplitedString(txtNum6.Text.Split(Calculator.Delimiter)) == 33)
            {
                sum += 0.5;
            }
            // If sex & creation equals 3 or 5
            if (Calculator.GetCorrectNumberFromSplitedString(txtNum6.Text.Split(Calculator.Delimiter)) == 3 ||
                Calculator.GetCorrectNumberFromSplitedString(txtNum6.Text.Split(Calculator.Delimiter)) == 5)
            {
                sum -= 0.5;
            }

            // If universe chakra & creation equals 1 or 7 or 9 or 22
            int universeNumber = Calculator.GetUniverseNumberFromChakra(txtNum9.Text);
            if (universeNumber == 1 ||
                universeNumber == 7 ||
                universeNumber == 9 ||
                universeNumber == 22)
            {
                sum -= 0.5;
            }

            // If universe chakra & creation equals 2 or 4 or 6 or 33
            //Calculator.GetCorrectNumberFromSplitedString(txtNum9.Text.Split(Calculator.Delimiter)) == 2
            //isAccounting && universeNumber == 8 // only for accounting tab
            if (universeNumber == 2 ||
                universeNumber == 4 ||
                universeNumber == 6 ||
                universeNumber == 33 ||
                isAccounting && universeNumber == 8)
            {
                sum += 0.5;
            }

            if (sum > 10)
            {
                sum = 10;
            }
            if (isSecreatry)
            {
                txtFinalSeceretyMark.Text = Math.Round(sum, 2).ToString();
            }
            else if (isAccounting)
            {
                txtFinalAccountingMark.Text = Math.Round(sum, 2).ToString();
            }

            string story = "";
            if ((sum >= 8) & (sum <= 10))
            {
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                {
                    story = "התאמה טובה מאוד";
                }
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                {
                    story = "Very good match";
                }
            }

            if ((sum >= 7) & (sum < 8))
            {
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                {
                    story = "התאמה טובה";
                }
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                {
                    story = "Good match";
                }
            }

            if ((sum >= 0) & (sum < 7))
            {
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                {
                    story = "התאמה חלשה";
                }
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                {
                    story = "Poor match";
                }
            }

            if (tmpastronum == 22)
            {
                //if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                //{
                //    story += Environment.NewLine + "היות ומזלך האסטרולוגי עקרב, קל לך להיות עובד מוערך.";
                //}
                //if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                //{
                //    story += Environment.NewLine + "As astrological sign is Scorpio it is easy for you to be a successful indipendent bussiness or a highly estimated worker.";
                //}
            }
            if (isSecreatry)
            {
                txtSecretaryStory.Text += "(" + txtFinalSeceretyMark.Text + ") " + Environment.NewLine + story;
            }
            else if (isAccounting)
            {
                txtAccountingStory.Text += "(" + txtFinalAccountingMark.Text + ") " + Environment.NewLine + story;
            }
        }
        // **********

        // **********

        private List<string> RelationshipMatchGatherInfo()
        {
            List<string> pPersonalValues = new List<string>();
            int age = Convert.ToInt16(txtAge.Text.Split("(".ToCharArray()[0])[0]);

            #region From Chakra Map
            int upperBndry = Convert.ToInt16(txt1_1.Text.Split("-".ToCharArray()[0])[1].Trim());

            int tmpVal = 0;
            if (age <= upperBndry)
            {
                // Crown
                tmpVal = Calculator.GetCorrectNumberFromSplitedString(txtNum1.Text.Split(Calculator.Delimiter));
            }
            else
            {
                //Third Eye
                tmpVal = Calculator.GetCorrectNumberFromSplitedString(txtNum2.Text.Split(Calculator.Delimiter));
            }
            pPersonalValues.Add(tmpVal.ToString());

            tmpVal = Calculator.GetCorrectNumberFromSplitedString(txtNum3.Text.Split(Calculator.Delimiter));
            pPersonalValues.Add(tmpVal.ToString());

            tmpVal = Calculator.GetCorrectNumberFromSplitedString(txtNum5.Text.Split(Calculator.Delimiter));
            pPersonalValues.Add(tmpVal.ToString());

            tmpVal = Calculator.GetCorrectNumberFromSplitedString(txtNum7.Text.Split(Calculator.Delimiter));
            pPersonalValues.Add(tmpVal.ToString());

            tmpVal = Calculator.GetCorrectNumberFromSplitedString(txtPName_Num.Text.Split(Calculator.Delimiter));
            pPersonalValues.Add(tmpVal.ToString());

            tmpVal = Convert.ToInt16(txtAstroName.Text.Split("(".ToCharArray()[0])[1].Split(")".ToCharArray()[0])[0]);
            pPersonalValues.Add(tmpVal.ToString());
            #endregion

            //double ChakraMapBusinessValue = Calculator.ChakraMap2BusinessValue(pPersonalValues);

            Control.ControlCollection cntrlCls = this.Controls;

            #region Life Cycles
            int curCycle = 0;
            int[] CycleBoundries = new int[4];
            CycleBoundries[0] = Convert.ToInt16(txt1_1.Text.Split("-".ToCharArray()[0])[1].Trim());
            CycleBoundries[1] = Convert.ToInt16(txt2_1.Text.Split("-".ToCharArray()[0])[1].Trim());
            CycleBoundries[2] = Convert.ToInt16(txt3_1.Text.Split("-".ToCharArray()[0])[1].Trim());
            CycleBoundries[3] = Convert.ToInt16(txt4_1.Text.Split("-".ToCharArray()[0])[1].Trim());

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

            int clmx, ccl;
            clmx = Calculator.GetCorrectNumberFromSplitedString(cntrlCls.Find("txt" + curCycle.ToString() + "_3", true)[0].Text.Split(Calculator.Delimiter));
            ccl = Calculator.GetCorrectNumberFromSplitedString(cntrlCls.Find("txt" + curCycle.ToString() + "_2", true)[0].Text.Split(Calculator.Delimiter));

            //pPersonalValues = new List<string>();
            pPersonalValues.Add(clmx.ToString());
            pPersonalValues.Add(ccl.ToString());


            // converting nums


            List<int> numsInLifeCycles = new List<int>();
            for (int i = curCycle; i < 5; i++)
            {
                numsInLifeCycles.Add(Calculator.GetCorrectNumberFromSplitedString(cntrlCls.Find("txt" + i.ToString() + "_2", true)[0].Text.Split(Calculator.Delimiter)));
                numsInLifeCycles.Add(Calculator.GetCorrectNumberFromSplitedString(cntrlCls.Find("txt" + i.ToString() + "_3", true)[0].Text.Split(Calculator.Delimiter)));
            }

            bool test1 = false;
            for (int i = 0; i < numsInLifeCycles.Count; i++)
            {
                if ((Calculator.isCarmaticNumber(numsInLifeCycles[i]) == true) || (numsInLifeCycles[i] == 9))
                {
                    test1 = true;
                }
            }

            int Test1Val = 2;
            if (test1 == false)
            {
                Test1Val = 10;
            }

            #endregion Life Cycles

            //double LifeCycleBusinessValue = Calculator.LifeCycle2BusinessValue(pPersonalValues, "כן" == SelfFixYesOrNo);

            pPersonalValues.Add(Test1Val.ToString());

            string tmpLvl = txtInfo1.Text.Split(System.Environment.NewLine.ToCharArray()[0])[1].Trim();
            if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
            {
            }
            if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
            {
                if (tmpLvl == "Balanced")
                {
                    tmpLvl = "מאוזן";
                }
                if (tmpLvl == "Not Balanced")
                {
                    tmpLvl = "לא מאוזן";
                }
                if (tmpLvl == "Half Balanced")
                {
                    tmpLvl = "חצי מאוזן";
                }
            }
            pPersonalValues.Add(tmpLvl);

            return pPersonalValues;
        }

        // **********

        private UserInfo CollectCurrentUserInfo()
        {
            UserInfo curPerson = new UserInfo();
            curPerson.mFirstName = txtPrivateName.Text.Trim();
            curPerson.mLastName = txtFamilyName.Text.Trim();
            curPerson.mFatherName = txtFatherName.Text.Trim();
            curPerson.mMotherName = txtMotherName.Text.Trim();
            curPerson.mB_Date = DateTimePickerFrom.Value;
            curPerson.mCity = txtCity.Text.Trim();
            curPerson.mStreet = txtStreet.Text.Trim();
            curPerson.mBuildingNum = Convert.ToDouble(txtBiuldingNum.Text);
            curPerson.mAppNum = Convert.ToDouble(txtAppNum.Text);


            return curPerson;
        }

        // **********

        private string[] CalcCombinations()
        {
            Dictionary<string, int> chakras = new Dictionary<string, int>();
            Dictionary<string, int> lifePeriods = new Dictionary<string, int>();
            int curLifeCycle = CalcCurrentCycle();

            FillChakrasDictionary(chakras, lifePeriods);

            return Calculator.GetCombinationResults(chakras, lifePeriods, curLifeCycle);

        }

        // **********

        private void FillChakrasDictionary(Dictionary<string, int> chakras, Dictionary<string, int> lifePeriods)
        {
            if (chakras == null) chakras = new Dictionary<string, int>();
            if (lifePeriods == null) lifePeriods = new Dictionary<string, int>();

            try
            {
                chakras.Clear();
                lifePeriods.Clear();

                for (int i = 1; i <= 9; i++)
                {
                    Control tb = Controls.Find("txtNum" + i.ToString(), true).First();

                    if (tb is TextBox)
                    {
                        if (i < 9)
                        {
                            chakras.Add(i.ToString(), Calculator.GetCorrectNumberFromSplitedString(tb.Text.Split(Calculator.Delimiter)));

                        }
                        else
                        {
                            chakras.Add(i.ToString(), Convert.ToInt16(tb.Text.Split("(".ToCharArray()[0])[1].Split(")".ToCharArray()[0])[0]));

                        }

                    }

                }

                for (int i = 1; i <= 4; i++)
                    for (int j = 2; j <= 4; j++)
                    {
                        Control tb = Controls.Find("txt" + i.ToString() + "_" + j.ToString(), true).First();

                        if (tb is TextBox)
                        {
                            lifePeriods.Add(i.ToString() + j.ToString(), Calculator.GetCorrectNumberFromSplitedString(tb.Text.Split(Calculator.Delimiter)));
                        }
                    }

            }
            catch (Exception ex)
            {
                throw new Exception("Problem in one of the parameters of input:" + Environment.NewLine, ex);

            }

        }

        #endregion

        #region Bubble Sort
        // sort the items of an array using bubble sort
        public void BubbleSort(ref int[] ar)
        {
            for (int pass = 1; pass < ar.Length; pass++)
                for (int i = 0; i < ar.Length - 1; i++)
                    if (ar[i] < ar[i + 1])
                        Swap(ref ar, i);
        }

        // swap two items of an array
        private void Swap(ref int[] ar, int first)
        {
            int hold;

            hold = ar[first];
            ar[first] = ar[first + 1];
            ar[first + 1] = hold;
        }

        private void GetMinMax4CombMap(int[] indexSum, out List<int> max, out List<int> nun)
        {
            List<int> tMax = new List<int>();

            max = new List<int>();
            nun = new List<int>();
            for (int i = 0; i < indexSum.Length; i++)
            {
                if (indexSum[i] == 0)
                {
                    nun.Add(i);
                }
            }

            int[] Sorted = new int[9];
            indexSum.CopyTo(Sorted, 0);

            BubbleSort(ref Sorted);
            if (Sorted[0] == 1) return;

            tMax.Add(Sorted[0]);
            for (int j = 1; j < Sorted.Length; j++)
            {
                if ((Sorted[j] < tMax[0]) && (Sorted[j] > 1)) //&& !inList(nun,j))
                {
                    tMax.Add(Sorted[j]);
                    break;
                }
            }

            foreach (int tmpMax in tMax)
            {
                for (int i = 0; i < indexSum.Length; i++)
                {
                    if (indexSum[i] == tmpMax)
                        max.Add(i);
                }
            }

            for (int i = 0; i < max.Count; i++)
            {
                if (max[i] == 1)
                {
                    max.Remove(i);
                }

            }

        }

        private bool inList(List<int> arr, int j)
        {
            foreach (int a in arr)
            {
                if (a == j) return true;
            }

            return false;
        }
        #endregion

        #region Marraige
        private void rbtnYesMarriage_CheckedChanged(object sender, EventArgs e)
        {
            if (rbtnYesMarriage.Checked == true)
            {
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                {
                    lblMrrgMark.Text = "ציון " + rbtnYesMarriage.Text;
                    lblMrrgInfo.Text = "על מנת שזוגיות המאפשרת קשרי משפחה" + System.Environment.NewLine + "תתקיים לאורך זמן - צריך ציון 7 ומעלה";
                }
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                {
                    lblMrrgMark.Text = "Mark " + rbtnYesMarriage.Text;
                    lblMrrgInfo.Text = "In order to achieve intimacy that allows family relationships" + System.Environment.NewLine + "will take place over time - index must be 7 or higher.";
                }
            }
            else
            {
                lblMrrgMark.Text = "ציון " + rbtnNoMarriage.Text;
                lblMrrgInfo.Text = "על מנת שזוגיות תצליח - צריך ציון 8 ומעלה ";
            }
        }

        private void btnCalcMarriage_Click(object sender, EventArgs e)
        {
            bool calcType = rbtnYesMarriage.Checked;

            if (!isCustomerLoaded())
            {
                return;

            }

            #region Gather Data Of Spouce

            DataGridViewRow r = SpouceData.Rows[0];
            //DataGridViewRow sexRow = SpouceSexData.Rows[0];

            if (!isSpouceDataValid(r))
            {
                return;

            }


            curPartner = new UserInfo();

            curPartner.mFirstName = CellValue2String(r.Cells[0]);
            curPartner.mLastName = CellValue2String(r.Cells[1]);
            curPartner.mFatherName = CellValue2String(r.Cells[3]);
            curPartner.mMotherName = CellValue2String(r.Cells[4]);

            //sexRow.Cells[0].Value = r.Cells[0].Value;
            //sexRow.Cells[1].Value = r.Cells[1].Value;
            //sexRow.Cells[3].Value = r.Cells[3].Value;
            //sexRow.Cells[4].Value = r.Cells[4].Value;

            bool resDate = DateTime.TryParse(r.Cells[2].Value.ToString().Trim(), out curPartner.mB_Date);
            if (resDate == false)
            {
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                {
                    MessageBox.Show("שגיאה בכתיבת תאריך", "שגיאת נתונים", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                {
                    MessageBox.Show("Error in date input", "Input Data Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                return;
            }

            //sexRow.Cells[2].Value = curPartner.mB_Date;
            curPartner.mCity = CellValue2String(r.Cells[6]);
            curPartner.mStreet = CellValue2String(r.Cells[7]);

            curPartner.mBuildingNum = CellValue2Double(r.Cells[8]);
            curPartner.mAppNum = CellValue2Double(r.Cells[9]);

            curPartner.mReachMaster = EnumProvider.ReachedMaster.Yes;

            string yesno = CellValue2String(r.Cells[5]);
            if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
            {
                if (yesno.Trim().Contains("כ") == true)
                    curPartner.mPassedRect = EnumProvider.PassedRectification.Passed;
                else
                    curPartner.mPassedRect = EnumProvider.PassedRectification.NotPassed;
            }
            if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
            {
                if (yesno.Trim().Contains("y") == true)
                    curPartner.mPassedRect = EnumProvider.PassedRectification.Passed;
                else
                    curPartner.mPassedRect = EnumProvider.PassedRectification.NotPassed;
            }

            #endregion
            #region Save Personal Data
            UserInfo curPerson = new UserInfo();
            curPerson.mFirstName = txtPrivateName.Text.Trim();
            curPerson.mLastName = txtFamilyName.Text.Trim();
            curPerson.mFatherName = txtFatherName.Text.Trim();
            curPerson.mMotherName = txtMotherName.Text.Trim();
            curPerson.mB_Date = DateTimePickerFrom.Value;
            curPerson.mEMail = txtEMail.Text;
            curPerson.mPhones = txtPhones.Text;

            string personalYesNo = cmbIsFix.SelectedItem.ToString();

            curPerson.mCity = txtCity.Text.Trim();
            curPerson.mStreet = txtStreet.Text.Trim();
            curPerson.mBuildingNum = Convert.ToDouble(txtBiuldingNum.Text);
            curPerson.mAppNum = Convert.ToDouble(txtAppNum.Text);

            if (cbMainCorrectionDone.Checked == true)
                curPerson.mPassedRect = EnumProvider.PassedRectification.Passed;
            else
                curPerson.mPassedRect = EnumProvider.PassedRectification.NotPassed;

            if (cbPersonMaster.Checked == false)
                curPerson.mReachMaster = EnumProvider.ReachedMaster.Yes;
            else
                curPerson.mReachMaster = EnumProvider.ReachedMaster.No;
            #endregion
            DateTime toCalcDay = DateTimePickerTo.Value;


            #region Spouce  Values
            txtPrivateName.Text = curPartner.mFirstName;
            txtFamilyName.Text = curPartner.mLastName;
            txtFatherName.Text = curPartner.mFatherName;
            txtMotherName.Text = curPartner.mMotherName;
            DateTimePickerFrom.Value = curPartner.mB_Date;
            txtCity.Text = curPartner.mCity;
            txtStreet.Text = curPartner.mStreet;
            txtBiuldingNum.Text = curPartner.mBuildingNum.ToString();
            txtAppNum.Text = curPartner.mAppNum.ToString();

            DateTimePickerTo.Value = toCalcDay;

            if (curPartner.mPassedRect == EnumProvider.PassedRectification.Passed)
            {
                cbMainCorrectionDone.Checked = true;
                cbMainCorrectionDone.CheckState = CheckState.Checked;
            }
            else
            {
                cbMainCorrectionDone.Checked = false;
                cbMainCorrectionDone.CheckState = CheckState.Unchecked;
            }

            if (curPartner.mReachMaster == EnumProvider.ReachedMaster.Yes)
            {
                cbPersonMaster.Checked = false;
                cbPersonMaster.CheckState = CheckState.Unchecked;
            }
            else
            {
                cbPersonMaster.Checked = true;
                cbPersonMaster.CheckState = CheckState.Checked;
            }
            #endregion
            checkBox1.Checked = true;
            checkBox1.CheckState = CheckState.Checked;

            runCalc();

            #region QnA
            //if ((isRunningQnA == true) && (isRunningQnAspouce == true))
            if (isRunningQnA == true)
            {
                SpouceResult = new UserResult();
                SpouceResult.PrivateNameNum = txtPName_Num.Text;
                SpouceResult.Astro = txtAstroName.Text;
                SpouceResult.birthday = DateTimePickerFrom.Value;

                SpouceResult.Base = txtNum7.Text;
                SpouceResult.Crown = txtNum1.Text;
                SpouceResult.Heart = txtNum4.Text;
                SpouceResult.SexAndCreation = txtNum6.Text;
                SpouceResult.Sun = txtNum5.Text;
                SpouceResult.ThirdEye = txtNum2.Text;
                SpouceResult.Throught = txtNum3.Text;

                SpouceResult.LC_cycle = this.Controls.Find("txt" + CalcCurrentCycle().ToString() + "_2", true)[0].Text;
                SpouceResult.LC_Climax = this.Controls.Find("txt" + CalcCurrentCycle().ToString() + "_3", true)[0].Text;
                SpouceResult.LC_Chalange = this.Controls.Find("txt" + CalcCurrentCycle().ToString() + "_4", true)[0].Text;
            }
            #endregion QnA
            SpouceResult = new UserResult();
            SpouceResult.GatherDataFromGUI(this.Controls, CalcCurrentCycle());

            bool tstSp1 = Calculator.isCarmaticNumber(Calculator.GetCorrectNumberFromSplitedString(txtNum1.Text.Split(Calculator.Delimiter)));
            bool tstSp2 = Calculator.isCarmaticNumber(Calculator.GetCorrectNumberFromSplitedString(txtNum5.Text.Split(Calculator.Delimiter)));
            bool tstSp3 = Calculator.isCarmaticNumber(Calculator.GetCorrectNumberFromSplitedString(txtPName_Num.Text.Split(Calculator.Delimiter)));
            bool tstSp4 = Calculator.isCarmaticNumber(Calculator.GetCorrectNumberFromSplitedString(txtNum2.Text.Split(Calculator.Delimiter)));
            bool tstSp5 = Calculator.isCarmaticNumber(Calculator.GetCorrectNumberFromSplitedString(txtNum3.Text.Split(Calculator.Delimiter)));
            bool tstSp6 = Calculator.isCarmaticNumber(Calculator.GetCorrectNumberFromSplitedString(txtNum7.Text.Split(Calculator.Delimiter)));

            int n1, n2;
            n1 = Calculator.GetCorrectNumberFromSplitedString(txt1_2.Text.Split(Calculator.Delimiter));
            n2 = Calculator.GetCorrectNumberFromSplitedString(txt1_3.Text.Split(Calculator.Delimiter));
            bool SpCrmCycle1 = Calculator.isCarmaticNumber(n1) || Calculator.isCarmaticNumber(n2);

            n1 = Calculator.GetCorrectNumberFromSplitedString(txt2_2.Text.Split(Calculator.Delimiter));
            n2 = Calculator.GetCorrectNumberFromSplitedString(txt2_3.Text.Split(Calculator.Delimiter));
            bool SpCrmCycle2 = Calculator.isCarmaticNumber(n1) || Calculator.isCarmaticNumber(n2);

            n1 = Calculator.GetCorrectNumberFromSplitedString(txt3_2.Text.Split(Calculator.Delimiter));
            n2 = Calculator.GetCorrectNumberFromSplitedString(txt3_3.Text.Split(Calculator.Delimiter));
            bool SpCrmCycle3 = Calculator.isCarmaticNumber(n1) || Calculator.isCarmaticNumber(n2);

            n1 = Calculator.GetCorrectNumberFromSplitedString(txt4_2.Text.Split(Calculator.Delimiter));
            n2 = Calculator.GetCorrectNumberFromSplitedString(txt4_3.Text.Split(Calculator.Delimiter));
            bool SpCrmCycle4 = Calculator.isCarmaticNumber(n1) || Calculator.isCarmaticNumber(n2);

            int SpCurrentCycle = CalcCurrentCycle();
            List<string> SpouceInfoValues = RelationshipMatchGatherInfo();

            #region Original User
            txtPrivateName.Text = curPerson.mFirstName;
            txtFamilyName.Text = curPerson.mLastName;
            txtFatherName.Text = curPerson.mFatherName;
            txtMotherName.Text = curPerson.mMotherName;
            DateTimePickerFrom.Value = curPerson.mB_Date;
            txtCity.Text = curPerson.mCity;
            txtStreet.Text = curPerson.mStreet;
            txtBiuldingNum.Text = curPerson.mBuildingNum.ToString();
            txtAppNum.Text = curPerson.mAppNum.ToString();
            cmbIsFix.SelectedItem = personalYesNo;

            txtEMail.Text = curPerson.mEMail;
            txtPhones.Text = curPerson.mPhones;

            if (curPerson.mPassedRect == EnumProvider.PassedRectification.Passed)
            {
                cbMainCorrectionDone.Checked = true;
                cbMainCorrectionDone.CheckState = CheckState.Checked;
            }
            else
            {
                cbMainCorrectionDone.Checked = false;
                cbMainCorrectionDone.CheckState = CheckState.Unchecked;
            }

            if (curPerson.mReachMaster == EnumProvider.ReachedMaster.Yes)
            {
                cbPersonMaster.Checked = false;
                cbPersonMaster.CheckState = CheckState.Unchecked;
            }
            else
            {
                cbPersonMaster.Checked = true;
                cbPersonMaster.CheckState = CheckState.Checked;
            }
            #endregion
            checkBox1.Checked = false;
            checkBox1.CheckState = CheckState.Unchecked;

            runCalc();

            bool tstPrsn1 = Calculator.isCarmaticNumber(Calculator.GetCorrectNumberFromSplitedString(txtNum1.Text.Split(Calculator.Delimiter)));
            bool tstPrsn2 = Calculator.isCarmaticNumber(Calculator.GetCorrectNumberFromSplitedString(txtNum5.Text.Split(Calculator.Delimiter)));
            bool tstPrsn3 = Calculator.isCarmaticNumber(Calculator.GetCorrectNumberFromSplitedString(txtPName_Num.Text.Split(Calculator.Delimiter)));
            bool tstPrsn4 = Calculator.isCarmaticNumber(Calculator.GetCorrectNumberFromSplitedString(txtNum2.Text.Split(Calculator.Delimiter)));
            bool tstPrsn5 = Calculator.isCarmaticNumber(Calculator.GetCorrectNumberFromSplitedString(txtNum3.Text.Split(Calculator.Delimiter)));
            bool tstPrsn6 = Calculator.isCarmaticNumber(Calculator.GetCorrectNumberFromSplitedString(txtNum7.Text.Split(Calculator.Delimiter)));

            n1 = Calculator.GetCorrectNumberFromSplitedString(txt1_2.Text.Split(Calculator.Delimiter));
            n2 = Calculator.GetCorrectNumberFromSplitedString(txt1_3.Text.Split(Calculator.Delimiter));
            bool PrnCrmCycle1 = Calculator.isCarmaticNumber(n1) || Calculator.isCarmaticNumber(n2);

            n1 = Calculator.GetCorrectNumberFromSplitedString(txt2_2.Text.Split(Calculator.Delimiter));
            n2 = Calculator.GetCorrectNumberFromSplitedString(txt2_3.Text.Split(Calculator.Delimiter));
            bool PrnCrmCycle2 = Calculator.isCarmaticNumber(n1) || Calculator.isCarmaticNumber(n2);

            n1 = Calculator.GetCorrectNumberFromSplitedString(txt3_2.Text.Split(Calculator.Delimiter));
            n2 = Calculator.GetCorrectNumberFromSplitedString(txt3_3.Text.Split(Calculator.Delimiter));
            bool PrnCrmCycle3 = Calculator.isCarmaticNumber(n1) || Calculator.isCarmaticNumber(n2);

            n1 = Calculator.GetCorrectNumberFromSplitedString(txt4_2.Text.Split(Calculator.Delimiter));
            n2 = Calculator.GetCorrectNumberFromSplitedString(txt4_3.Text.Split(Calculator.Delimiter));
            bool PrnCrmCycle4 = Calculator.isCarmaticNumber(n1) || Calculator.isCarmaticNumber(n2);

            int PrsnCurrentCycle = CalcCurrentCycle();

            List<string> UserInfoValues = RelationshipMatchGatherInfo();

            SpouceData.Rows.Clear();
            SpouceData.Rows.Add(r);


            rbtnYesMarriage.Checked = calcType;
            rbtnNoMarriage.Checked = !calcType;

            #region Show Resluts
            double ans = Calculator.MarriageCompetabilityCalc((rbtnYesMarriage.Checked == true), UserInfoValues, (personalYesNo == "כן"), SpouceInfoValues, (yesno == "כן"));

            txtMrrgMark.Text = Math.Round(ans, 2).ToString();
            if (rbtnYesMarriage.Checked == true)
            {
                if ((ans > 9) && (ans <= 10))
                {
                    if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                    {
                        txtMrrgStory.Text = "התאמה מעולה וארוכת טווח - סיכוי מצויין לקשרי משפחה ארוכי טווח (10 שנים ומעלה) ולפחות עד לסיום המחזור הנבדק";
                    }
                    if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                    {
                        txtMrrgStory.Text = "Excellent fit and long-lasting - excellent chance of long-term family ties (10 years and above) and at least until the end of the cycle the patient";
                    }
                }

                if ((ans > 8) && (ans <= 9))
                {
                    if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                    {
                        txtMrrgStory.Text = "התאמה טובה מאוד וארוכת טווח - סיכוי טוב מאוד לקשרי משפחה ארוכי טווח (10 שנים ומעלה) ולפחות עד לסיום המחזור הנבדק";
                    }
                    if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                    {
                        txtMrrgStory.Text = "Very good fit and long-lasting - a very good chance of long-term family ties (10 years and above) and at least until the end of the cycle the patient";
                    }
                }

                if ((ans > 7) && (ans <= 8))
                {
                    if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                    {
                        txtMrrgStory.Text = "התאמה בינונית לתקופה ארוכת טווח - סיכוי בינוני לקשרי משפחה לאורך זמן (10 שניםו ומעלה) ולפחות עד סיום המחזור הנבדק";
                    }
                    if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                    {
                        txtMrrgStory.Text = "Suitability for long-term moderate - moderate chance Family Relations over time (10 aphids and above) and at least until the end of the cycle the patient";
                    }
                }

                if ((ans > 6.5) && (ans <= 7))
                {
                    if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                    {
                        txtMrrgStory.Text = "התאמה חלשה לתקופה ארוכת טווח - סיכוי חלש לקשרי משפחה לאורך זמן ולפחות עד לסיום המחזור הנבדק. סיכוי לקשר עד 7 שנים";
                    }
                    if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                    {
                        txtMrrgStory.Text = "Match weak long-term period - likely weak family ties over time, at least until the period under review. Chance to link up to 7 years";
                    }
                }

                if ((ans > 6) && (ans <= 6.5))
                {
                    if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                    {
                        txtMrrgStory.Text = "איכות חלשה מאוד לקשרי משפחה לתקופה ארוכת טווח - סיכוי חלש מאוד לקשר ארוך טווח. סיכוי לקשר עד 3 שנים או עד לסיום המחזור הנבדק";
                        if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                        {
                            txtMrrgStory.Text = "Quality is very weak family ties for long term - likely very weak long term relationship. Chance to link up to three years or until the end of the current cycle";
                        }
                    }

                    if ((ans > 2) && (ans <= 6))
                    {
                        if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                        {
                            txtMrrgStory.Text = "איכות חלשה מאוד לקשרי משפחה לתקופה ארוכת טווח - סיכוי חלש מאוד לקשר ארוך טווח. סיכוי לקשר עד 3 שנים או עד לסיום המחזור הנבדק";
                        }
                        if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                        {
                            txtMrrgStory.Text = "Quality is very weak family ties for long term - likely very weak long term relationship. Chance to link up to three years or until the end of the current cycle";
                        }
                    }

                    if ((ans > 0) && (ans <= 2))
                    {
                        if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                        {
                            txtMrrgStory.Text = "אין סיכוי לקשר ארוך טווח - סיכוי לקשר של מספר חודשים";
                        }
                        if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                        {
                            txtMrrgStory.Text = "No chance of a long term relationship - a chance to connect a number of months";
                        }
                    }
                }
                else
                {
                    if ((ans > 8) && (ans <= 10))
                    {
                        if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                        {
                            txtMrrgStory.Text = "התאמה זוגית מעולה - סיכוי מצויין לזוגיות";
                        }
                        if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                        {
                            txtMrrgStory.Text = "Superb double match - Excellent chance relationship";
                        }
                    }

                    if ((ans > 6) && (ans <= 8))
                    {
                        if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                        {
                            txtMrrgStory.Text = "התאמה זוגית בינונית - סיכוי בינוני לזוגיות";
                        }
                        if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                        {
                            txtMrrgStory.Text = "Double match a medium - a medium chance relationship";
                        }
                    }

                    if ((ans > 0) && (ans <= 6))
                    {
                        if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                        {
                            txtMrrgStory.Text = "התאמה זוגית חלשה - סיכוי נמוך מאוד לזוגיות";
                        }
                        if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                        {
                            txtMrrgStory.Text = "Weak double match - a very low chance relationship";
                        }
                    }
                }



                #endregion Show Resluts

                txtMrrgStory.Text += Environment.NewLine;
                bool res = (tstPrsn1 || tstPrsn2 || tstPrsn3);
                if (PrnCrmCycle1 == true)
                {
                    res = true;
                }
                else
                {
                    if ((PrsnCurrentCycle == 2) && (PrnCrmCycle2 == true))
                    {
                        res = true;
                    }
                    if ((PrsnCurrentCycle == 3) && (PrnCrmCycle3 == true))
                    {
                        res = true;
                    }
                    if ((PrsnCurrentCycle == 4) && (PrnCrmCycle4 == true))
                    {
                        res = true;
                    }
                }

                if (res == true)
                {
                    if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                    {
                        txtMrrgStory.Text += Environment.NewLine + "הסיכוי לזוגיות יתאפשר בעקבות הבנת התיקון של" + " " + curPerson.mFirstName;
                    }
                    if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                    {
                        txtMrrgStory.Text += Environment.NewLine + "Chance possible relationship following the amendment of Understanding" + " " + curPerson.mFirstName;
                    }
                }

                res = (tstSp1 || tstSp2 || tstSp3);
                if (SpCrmCycle1 == true)
                {
                    res = true;
                }
                else
                {
                    if ((SpCurrentCycle == 2) && (SpCrmCycle2 == true))
                    {
                        res = true;
                    }
                    if ((SpCurrentCycle == 3) && (SpCrmCycle3 == true))
                    {
                        res = true;
                    }
                    if ((SpCurrentCycle == 4) && (SpCrmCycle4 == true))
                    {
                        res = true;
                    }
                }
                if (res == true)
                {
                    if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                    {
                        txtMrrgStory.Text += Environment.NewLine + "הסיכוי לזוגיות יתאפשר בעקבות הבנת התיקון של" + " " + curPartner.mFirstName;
                    }
                    if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                    {
                        txtMrrgStory.Text += Environment.NewLine + "Chance possible relationship following the amendment of Understanding" + " " + curPartner.mFirstName;
                    }
                }
            }

            if (isRunningQnA == true)
            {
                cmbQ.SelectedItem = 12;
                cmbQF.SelectedItem = 12;
                cmbQ.SelectedIndex = 12;
                cmdQnAcalc_Click(null, null);
                mTabs.SelectTab("tabQnA");
                mTabs.Show();
            }

        }

        // **********

        private bool isSpouceDataValid(DataGridViewRow r)
        {
            DateTime res;
            bool retVal = false;

            retVal = r.Cells[0].Value != null &&
                     r.Cells[1].Value != null &&
                     DateTime.TryParse(r.Cells[2].Value.ToString().Trim(), out res);

            if (!retVal)
            {
                string msg = "אחד או יותר מפרטי בן/ת הזוג חסר או לא תקין (שם פרטי/משפחה/תאריך לידה). הזן פרטי בן/ת זוג תחילה, ובצע חישוב .";

                MessageBox.Show(msg, string.Empty, MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1, MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);

            }

            return retVal;


        }

        // **********

        private void btnCalcSexMatch_Click(object sender, EventArgs e)
        {
            Dictionary<string, int> chakras = new Dictionary<string, int>();
            Dictionary<string, int> lifePeriods = new Dictionary<string, int>();
            Dictionary<string, int> chakrasSP = new Dictionary<string, int>();
            Dictionary<string, int> lifePeriodsSP = new Dictionary<string, int>();
            int addedBonus = 0;

            if (!isCustomerLoaded())
            {
                return;

            }



            #region Gather Data Of Spouce

            DataGridViewRow r = SpouceSexData.Rows[0];

            curPartner = new UserInfo();

            if (!isSpouceDataValid(r))
            {

                return;

            }

            curPartner.mFirstName = CellValue2String(r.Cells[0]);
            curPartner.mLastName = CellValue2String(r.Cells[1]);
            //curPartner.mSex = EnumProvider.Instance.GetSexEnumFromString(r.Cells[2].Value.ToString());
            curPartner.mFatherName = CellValue2String(r.Cells[3]);
            curPartner.mMotherName = CellValue2String(r.Cells[4]);
            bool resDate = DateTime.TryParse(r.Cells[2].Value.ToString().Trim(), out curPartner.mB_Date);
            if (resDate == false)
            {
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                {
                    MessageBox.Show("שגיאה בכתיבת תאריך", "שגיאת נתונים", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                {
                    MessageBox.Show("Error in date input", "Input Data Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                return;
            }

            curPartner.mCity = CellValue2String(r.Cells[6]);
            curPartner.mStreet = CellValue2String(r.Cells[7]);

            curPartner.mBuildingNum = CellValue2Double(r.Cells[8]);
            curPartner.mAppNum = CellValue2Double(r.Cells[9]);

            curPartner.mReachMaster = EnumProvider.ReachedMaster.Yes;

            string yesno = CellValue2String(r.Cells[5]);
            if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
            {
                if (yesno.Trim().Contains("כ") == true)
                    curPartner.mPassedRect = EnumProvider.PassedRectification.Passed;
                else
                    curPartner.mPassedRect = EnumProvider.PassedRectification.NotPassed;
            }
            if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
            {
                if (yesno.Trim().Contains("y") == true)
                    curPartner.mPassedRect = EnumProvider.PassedRectification.Passed;
                else
                    curPartner.mPassedRect = EnumProvider.PassedRectification.NotPassed;
            }

            #endregion
            #region Save Personal Data
            UserInfo curPerson = new UserInfo();
            curPerson.mFirstName = txtPrivateName.Text.Trim();
            curPerson.mLastName = txtFamilyName.Text.Trim();
            curPerson.mFatherName = txtFatherName.Text.Trim();
            curPerson.mMotherName = txtMotherName.Text.Trim();
            curPerson.mB_Date = DateTimePickerFrom.Value;
            curPerson.mEMail = txtEMail.Text;
            curPerson.mPhones = txtPhones.Text;

            curPerson.mCity = txtCity.Text.Trim();
            curPerson.mStreet = txtStreet.Text.Trim();
            curPerson.mBuildingNum = Convert.ToDouble(txtBiuldingNum.Text);
            curPerson.mAppNum = Convert.ToDouble(txtAppNum.Text);

            if (cbMainCorrectionDone.Checked == true)
                curPerson.mPassedRect = EnumProvider.PassedRectification.Passed;
            else
                curPerson.mPassedRect = EnumProvider.PassedRectification.NotPassed;

            if (cbPersonMaster.Checked == false)
                curPerson.mReachMaster = EnumProvider.ReachedMaster.Yes;
            else
                curPerson.mReachMaster = EnumProvider.ReachedMaster.No;
            #endregion
            DateTime toCalcDay = DateTimePickerTo.Value;

            #region Spouce  Values
            txtPrivateName.Text = curPartner.mFirstName;
            txtFamilyName.Text = curPartner.mLastName;
            txtFatherName.Text = curPartner.mFatherName;
            txtMotherName.Text = curPartner.mMotherName;
            DateTimePickerFrom.Value = curPartner.mB_Date;
            txtCity.Text = curPartner.mCity;
            txtStreet.Text = curPartner.mStreet;
            txtBiuldingNum.Text = curPartner.mBuildingNum.ToString();
            txtAppNum.Text = curPartner.mAppNum.ToString();

            DateTimePickerTo.Value = toCalcDay;

            if (curPartner.mPassedRect == EnumProvider.PassedRectification.Passed)
            {
                cbMainCorrectionDone.Checked = true;
                cbMainCorrectionDone.CheckState = CheckState.Checked;
            }
            else
            {
                cbMainCorrectionDone.Checked = false;
                cbMainCorrectionDone.CheckState = CheckState.Unchecked;
            }

            if (curPartner.mReachMaster == EnumProvider.ReachedMaster.Yes)
            {
                cbPersonMaster.Checked = false;
                cbPersonMaster.CheckState = CheckState.Unchecked;
            }
            else
            {
                cbPersonMaster.Checked = true;
                cbPersonMaster.CheckState = CheckState.Checked;
            }
            #endregion
            checkBox1.Checked = true;
            checkBox1.CheckState = CheckState.Checked;

            runCalc();

            FillChakrasDictionary(chakrasSP, lifePeriodsSP);

            #region QnA
            //if ((isRunningQnA == true) && (isRunningQnAspouce == true))
            if (isRunningQnA == true)
            {
                SpouceResult = new UserResult();
                SpouceResult.PrivateNameNum = txtPName_Num.Text;
                SpouceResult.Astro = txtAstroName.Text;
                SpouceResult.birthday = DateTimePickerFrom.Value;

                SpouceResult.Base = txtNum7.Text;
                SpouceResult.Crown = txtNum1.Text;
                SpouceResult.Heart = txtNum4.Text;
                SpouceResult.SexAndCreation = txtNum6.Text;
                SpouceResult.Sun = txtNum5.Text;
                SpouceResult.ThirdEye = txtNum2.Text;
                SpouceResult.Throught = txtNum3.Text;

                SpouceResult.LC_cycle = this.Controls.Find("txt" + CalcCurrentCycle().ToString() + "_2", true)[0].Text;
                SpouceResult.LC_Climax = this.Controls.Find("txt" + CalcCurrentCycle().ToString() + "_3", true)[0].Text;
                SpouceResult.LC_Chalange = this.Controls.Find("txt" + CalcCurrentCycle().ToString() + "_4", true)[0].Text;
            }
            #endregion QnA
            SpouceResult = new UserResult();
            SpouceResult.GatherDataFromGUI(this.Controls, CalcCurrentCycle());

            bool tstSp1 = Calculator.isCarmaticNumber(Calculator.GetCorrectNumberFromSplitedString(txtNum1.Text.Split(Calculator.Delimiter)));
            bool tstSp2 = Calculator.isCarmaticNumber(Calculator.GetCorrectNumberFromSplitedString(txtNum5.Text.Split(Calculator.Delimiter)));
            bool tstSp3 = Calculator.isCarmaticNumber(Calculator.GetCorrectNumberFromSplitedString(txtPName_Num.Text.Split(Calculator.Delimiter)));
            bool tstSp4 = Calculator.isCarmaticNumber(Calculator.GetCorrectNumberFromSplitedString(txtNum2.Text.Split(Calculator.Delimiter)));
            bool tstSp5 = Calculator.isCarmaticNumber(Calculator.GetCorrectNumberFromSplitedString(txtNum3.Text.Split(Calculator.Delimiter)));
            bool tstSp6 = Calculator.isCarmaticNumber(Calculator.GetCorrectNumberFromSplitedString(txtNum7.Text.Split(Calculator.Delimiter)));

            int n1, n2;
            n1 = Calculator.GetCorrectNumberFromSplitedString(txt1_2.Text.Split(Calculator.Delimiter));
            n2 = Calculator.GetCorrectNumberFromSplitedString(txt1_3.Text.Split(Calculator.Delimiter));
            bool SpCrmCycle1 = Calculator.isCarmaticNumber(n1) || Calculator.isCarmaticNumber(n2);

            n1 = Calculator.GetCorrectNumberFromSplitedString(txt2_2.Text.Split(Calculator.Delimiter));
            n2 = Calculator.GetCorrectNumberFromSplitedString(txt2_3.Text.Split(Calculator.Delimiter));
            bool SpCrmCycle2 = Calculator.isCarmaticNumber(n1) || Calculator.isCarmaticNumber(n2);

            n1 = Calculator.GetCorrectNumberFromSplitedString(txt3_2.Text.Split(Calculator.Delimiter));
            n2 = Calculator.GetCorrectNumberFromSplitedString(txt3_3.Text.Split(Calculator.Delimiter));
            bool SpCrmCycle3 = Calculator.isCarmaticNumber(n1) || Calculator.isCarmaticNumber(n2);

            n1 = Calculator.GetCorrectNumberFromSplitedString(txt4_2.Text.Split(Calculator.Delimiter));
            n2 = Calculator.GetCorrectNumberFromSplitedString(txt4_3.Text.Split(Calculator.Delimiter));
            bool SpCrmCycle4 = Calculator.isCarmaticNumber(n1) || Calculator.isCarmaticNumber(n2);

            int SpCurrentCycle = CalcCurrentCycle();
            List<string> SpouceInfoValues = RelationshipMatchGatherInfo();

            #region Original User
            txtPrivateName.Text = curPerson.mFirstName;
            txtFamilyName.Text = curPerson.mLastName;
            txtFatherName.Text = curPerson.mFatherName;
            txtMotherName.Text = curPerson.mMotherName;
            DateTimePickerFrom.Value = curPerson.mB_Date;
            txtCity.Text = curPerson.mCity;
            txtStreet.Text = curPerson.mStreet;
            txtBiuldingNum.Text = curPerson.mBuildingNum.ToString();
            txtAppNum.Text = curPerson.mAppNum.ToString();

            txtEMail.Text = curPerson.mEMail;
            txtPhones.Text = curPerson.mPhones;

            if (curPerson.mPassedRect == EnumProvider.PassedRectification.Passed)
            {
                cbMainCorrectionDone.Checked = true;
                cbMainCorrectionDone.CheckState = CheckState.Checked;
            }
            else
            {
                cbMainCorrectionDone.Checked = false;
                cbMainCorrectionDone.CheckState = CheckState.Unchecked;
            }

            if (curPerson.mReachMaster == EnumProvider.ReachedMaster.Yes)
            {
                cbPersonMaster.Checked = false;
                cbPersonMaster.CheckState = CheckState.Unchecked;
            }
            else
            {
                cbPersonMaster.Checked = true;
                cbPersonMaster.CheckState = CheckState.Checked;
            }
            #endregion
            checkBox1.Checked = false;
            checkBox1.CheckState = CheckState.Unchecked;

            runCalc();
            FillChakrasDictionary(chakras, lifePeriods);

            bool tstPrsn1 = Calculator.isCarmaticNumber(Calculator.GetCorrectNumberFromSplitedString(txtNum1.Text.Split(Calculator.Delimiter)));
            bool tstPrsn2 = Calculator.isCarmaticNumber(Calculator.GetCorrectNumberFromSplitedString(txtNum5.Text.Split(Calculator.Delimiter)));
            bool tstPrsn3 = Calculator.isCarmaticNumber(Calculator.GetCorrectNumberFromSplitedString(txtPName_Num.Text.Split(Calculator.Delimiter)));
            bool tstPrsn4 = Calculator.isCarmaticNumber(Calculator.GetCorrectNumberFromSplitedString(txtNum2.Text.Split(Calculator.Delimiter)));
            bool tstPrsn5 = Calculator.isCarmaticNumber(Calculator.GetCorrectNumberFromSplitedString(txtNum3.Text.Split(Calculator.Delimiter)));
            bool tstPrsn6 = Calculator.isCarmaticNumber(Calculator.GetCorrectNumberFromSplitedString(txtNum7.Text.Split(Calculator.Delimiter)));

            n1 = Calculator.GetCorrectNumberFromSplitedString(txt1_2.Text.Split(Calculator.Delimiter));
            n2 = Calculator.GetCorrectNumberFromSplitedString(txt1_3.Text.Split(Calculator.Delimiter));
            bool PrnCrmCycle1 = Calculator.isCarmaticNumber(n1) || Calculator.isCarmaticNumber(n2);

            n1 = Calculator.GetCorrectNumberFromSplitedString(txt2_2.Text.Split(Calculator.Delimiter));
            n2 = Calculator.GetCorrectNumberFromSplitedString(txt2_3.Text.Split(Calculator.Delimiter));
            bool PrnCrmCycle2 = Calculator.isCarmaticNumber(n1) || Calculator.isCarmaticNumber(n2);

            n1 = Calculator.GetCorrectNumberFromSplitedString(txt3_2.Text.Split(Calculator.Delimiter));
            n2 = Calculator.GetCorrectNumberFromSplitedString(txt3_3.Text.Split(Calculator.Delimiter));
            bool PrnCrmCycle3 = Calculator.isCarmaticNumber(n1) || Calculator.isCarmaticNumber(n2);

            n1 = Calculator.GetCorrectNumberFromSplitedString(txt4_2.Text.Split(Calculator.Delimiter));
            n2 = Calculator.GetCorrectNumberFromSplitedString(txt4_3.Text.Split(Calculator.Delimiter));
            bool PrnCrmCycle4 = Calculator.isCarmaticNumber(n1) || Calculator.isCarmaticNumber(n2);

            int PrsnCurrentCycle = CalcCurrentCycle();

            List<string> UserInfoValues = RelationshipMatchGatherInfo();

            SpouceSexData.Rows.Clear();
            SpouceSexData.Rows.Add(r);

            #region Show Resluts

            Calc cl = Calculator;

            // If one of the partners age is older than 40 and female
            #region Calc bonus for women

            if ((curPartner.Age > 40) && (curPartner.mSex == EnumProvider.Sex.Female) ||
                (curPerson.Age > 40) && (curPerson.mSex == EnumProvider.Sex.Female))
            {
                if (cl.ChakraContainsValues(chakras, new string[] { "1", "2", "3", "5", "8" }, new int[] { 16 }) ||
                    cl.LifePeriodsContainsValues(lifePeriods, new int[] { 16 }, new int[] { 0, 1, 1, 0 }, PrsnCurrentCycle) ||
                    cl.ChakraContainsValues(chakrasSP, new string[] { "1", "2", "3", "5", "8" }, new int[] { 16 }) ||
                    cl.LifePeriodsContainsValues(lifePeriodsSP, new int[] { 16 }, new int[] { 0, 1, 1, 0 }, SpCurrentCycle))
                {
                    addedBonus -= 2;
                }

                if (cl.ChakraContainsValues(chakras, new string[] { "1", "2", "3", "5", "9" }, new int[] { 7 }) ||
                    cl.LifePeriodsContainsValues(lifePeriods, new int[] { 7 }, new int[] { 0, 1, 1, 0 }, PrsnCurrentCycle) ||
                    cl.ChakraContainsValues(chakrasSP, new string[] { "1", "2", "3", "5", "9" }, new int[] { 7 }) ||
                    cl.LifePeriodsContainsValues(lifePeriodsSP, new int[] { 7 }, new int[] { 0, 1, 1, 0 }, SpCurrentCycle))
                {
                    addedBonus -= 2;
                }

            }

            #endregion Calc bonus for women

            // If one of the partners age is older than 40 and male

            #region Calc bonus for men

            if ((curPartner.Age > 40) && (curPartner.mSex == EnumProvider.Sex.Male) ||
                (curPerson.Age > 40) && (curPerson.mSex == EnumProvider.Sex.Male))
            {
                if (cl.ChakraContainsValues(chakras, new string[] { "1", "2", "3", "5", "8" }, new int[] { 16 }) ||
                    cl.LifePeriodsContainsValues(lifePeriods, new int[] { 16 }, new int[] { 0, 1, 1, 0 }, 1) ||
                    cl.ChakraContainsValues(chakrasSP, new string[] { "1", "2", "3", "5", "8" }, new int[] { 16 }) ||
                    cl.LifePeriodsContainsValues(lifePeriodsSP, new int[] { 16 }, new int[] { 0, 1, 1, 0 }, 1))
                {
                    addedBonus += 1;

                }

                if (cl.ChakraContainsValues(chakras, new string[] { "1", "2", "3", "5", "9" }, new int[] { 7 }) ||
                    cl.LifePeriodsContainsValues(lifePeriods, new int[] { 7 }, new int[] { 0, 1, 1, 0 }, 1) ||
                    cl.ChakraContainsValues(chakrasSP, new string[] { "1", "2", "3", "5", "9" }, new int[] { 7 }) ||
                    cl.LifePeriodsContainsValues(lifePeriodsSP, new int[] { 7 }, new int[] { 0, 1, 1, 0 }, 1))
                {
                    addedBonus += 1;

                }

            }

            if (curPerson.mSex == EnumProvider.Sex.Male)
            {
                if (cl.LifePeriodsContainsValues(lifePeriods, new int[] { 7, 16 }, new int[] { 0, 0, 1, 0 }, 2) ||
                    cl.LifePeriodsContainsValues(lifePeriods, new int[] { 7, 16 }, new int[] { 0, 0, 1, 0 }, 3) ||
                    cl.LifePeriodsContainsValues(lifePeriods, new int[] { 7, 16 }, new int[] { 0, 0, 1, 0 }, 4))
                {
                    addedBonus -= 2;

                }

                if (cl.LifePeriodsContainsValues(lifePeriods, new int[] { 7, 16 }, new int[] { 0, 1, 0, 0 }, 2) ||
                    cl.LifePeriodsContainsValues(lifePeriods, new int[] { 7, 16 }, new int[] { 0, 1, 0, 0 }, 3) ||
                    cl.LifePeriodsContainsValues(lifePeriods, new int[] { 7, 16 }, new int[] { 0, 1, 0, 0 }, 4))
                {
                    addedBonus -= 1;

                }

                // If age is older than 35, and 7,13,14,16,19 in Crown/Meta chakras
                if (curPerson.Age > 35 && cl.ChakraContainsValues(chakras, new string[] { "1", "8" }, new int[] { 7, 13, 14, 16, 19 }))
                {
                    addedBonus -= 1;

                }

            }

            if (curPartner.mSex == EnumProvider.Sex.Male)
            {
                if (cl.LifePeriodsContainsValues(lifePeriodsSP, new int[] { 7, 16 }, new int[] { 0, 0, 1, 0 }, 2) ||
                    cl.LifePeriodsContainsValues(lifePeriodsSP, new int[] { 7, 16 }, new int[] { 0, 0, 1, 0 }, 3) ||
                    cl.LifePeriodsContainsValues(lifePeriodsSP, new int[] { 7, 16 }, new int[] { 0, 0, 1, 0 }, 4))
                {
                    addedBonus -= 2;

                }

                if (cl.LifePeriodsContainsValues(lifePeriodsSP, new int[] { 7, 16 }, new int[] { 0, 1, 0, 0 }, 2) ||
                    cl.LifePeriodsContainsValues(lifePeriodsSP, new int[] { 7, 16 }, new int[] { 0, 1, 0, 0 }, 3) ||
                    cl.LifePeriodsContainsValues(lifePeriodsSP, new int[] { 7, 16 }, new int[] { 0, 1, 0, 0 }, 4))
                {
                    addedBonus -= 1;

                }

                // If age is older than 35, and 7,13,14,16,19 in Crown/Meta chakras
                if (curPartner.Age > 35 && cl.ChakraContainsValues(chakrasSP, new string[] { "1", "8" }, new int[] { 7, 13, 14, 16, 19 }))
                {
                    addedBonus -= 1;

                }

            }

            #endregion Calc bonus for men

            double ans = Calculator.SexualMatchCalc(UserInfoValues, SpouceInfoValues) + addedBonus;

            if (ans > 10)
            {
                ans = 10;
                txtSexMatchStory.Text = "התאמה מינית טובה מאוד";

            }
            else if (ans > 8 && ans <= 10)
            {
                txtSexMatchStory.Text = "התאמה מינית טובה מאוד";

            }
            else if (ans > 7 && ans <= 8)
            {
                txtSexMatchStory.Text = "התאמה מינית טובה";

            }
            else if (ans > 6 && ans <= 7)
            {
                txtSexMatchStory.Text = "התאמה מינית בעלת איכות בינונית";

            }
            else if (ans <= 6)
            {
                txtSexMatchStory.Text = "לא מומלץ. התאמה מינית בעלת איכות חלשה מאוד";

            }

            txtSexMatchMark.Text = Math.Round(ans, 2).ToString();

            #endregion Show Resluts

            if (isRunningQnA == true)
            {
                cmbQ.SelectedItem = 10;
                cmbQF.SelectedItem = 10;
                cmdQnAcalc_Click(null, null);
                mTabs.SelectTab("tabQnA");
                mTabs.Show();

            }

        }

        // **********

        private bool isCustomerLoaded()
        {
            DateTime res;
            bool retVal = false;

            retVal = !string.IsNullOrEmpty(txtPrivateName.Text.Trim()) &&
                     !string.IsNullOrEmpty(txtFamilyName.Text.Trim()) &&
                     DateTime.TryParse(DateTimePickerFrom.Value.ToLongDateString(), out res);

            if (!retVal)
            {
                string msg = "אחד או יותר מפרטי הלקוח חסר או לא תקין (שם פרטי/משפחה/תאריך לידה). טען לקוח תחילה, ובצע חישוב .";

                MessageBox.Show(msg, string.Empty, MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1, MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);

            }

            return retVal;
        }

        #endregion Marraige

        private int CalcCurrentCycle()
        {
            if (mB_Date > DateTimePickerTo.Value)
            {
                DateTimePickerTo.Value = mB_Date.AddYears(100);
            }
            int age = Convert.ToInt16(sCalcAge_wMonths(mB_Date).Split("(".ToCharArray()[0])[0]);

            int curCycle = 0;

            int[] CycleBoundries = new int[8];

            CycleBoundries[0] = Convert.ToInt16(txt1_1.Text.Split("-".ToCharArray()[0])[0].Trim());
            CycleBoundries[1] = Convert.ToInt16(txt1_1.Text.Split("-".ToCharArray()[0])[1].Trim());
            CycleBoundries[2] = Convert.ToInt16(txt2_1.Text.Split("-".ToCharArray()[0])[0].Trim());
            CycleBoundries[3] = Convert.ToInt16(txt2_1.Text.Split("-".ToCharArray()[0])[1].Trim());
            CycleBoundries[4] = Convert.ToInt16(txt3_1.Text.Split("-".ToCharArray()[0])[0].Trim());
            CycleBoundries[5] = Convert.ToInt16(txt3_1.Text.Split("-".ToCharArray()[0])[1].Trim());
            CycleBoundries[6] = Convert.ToInt16(txt4_1.Text.Split("-".ToCharArray()[0])[0].Trim());
            CycleBoundries[7] = Convert.ToInt16(txt4_1.Text.Split("-".ToCharArray()[0])[1].Trim());

            if ((age >= CycleBoundries[0]) && (age <= CycleBoundries[1]))
            {
                curCycle = 1;
            }
            else if ((age >= CycleBoundries[2]) && (age <= CycleBoundries[3]))
            {
                curCycle = 2;
            }
            else if ((age >= CycleBoundries[4]) && (age <= CycleBoundries[5]))
            {
                curCycle = 3;
            }
            else if ((age >= CycleBoundries[6]) && (age <= CycleBoundries[7]))
            {
                curCycle = 4;
            }

            return curCycle;
        }

        #region LearnSccss_AttPrbl

        private bool is14Test = false;
        private List<string> LearnSccss_GatherInfo()
        {
            List<string> pPersonalValues = new List<string>();

            int age = Convert.ToInt16(sCalcAge_wMonths(mB_Date).Split("(".ToCharArray()[0])[0]);

            #region From Chakra Map
            int upperBndry = Convert.ToInt16(txt1_1.Text.Split("-".ToCharArray()[0])[1].Trim());

            int tmpVal = 0;
            if (age <= upperBndry)
            {
                // Crown
                tmpVal = Calculator.GetCorrectNumberFromSplitedString(txtNum1.Text.Split(Calculator.Delimiter));
                if ((age < 14) && (tmpVal == 1))
                {
                    is14Test = true;
                }
                pPersonalValues.Add(txtNum1.Text);
            }
            else
            {
                //Third Eye
                tmpVal = Calculator.GetCorrectNumberFromSplitedString(txtNum2.Text.Split(Calculator.Delimiter));
                pPersonalValues.Add(txtNum2.Text);
            }
            //pPersonalValues.Add(tmpVal.ToString());


            tmpVal = Calculator.GetCorrectNumberFromSplitedString(txtNum3.Text.Split(Calculator.Delimiter));
            //pPersonalValues.Add(tmpVal.ToString());
            pPersonalValues.Add(txtNum3.Text);
            if ((age < 14) && (tmpVal == 1))
            {
                is14Test = true;
            }

            tmpVal = Calculator.GetCorrectNumberFromSplitedString(txtNum5.Text.Split(Calculator.Delimiter));
            //pPersonalValues.Add(tmpVal.ToString());
            pPersonalValues.Add(txtNum5.Text);
            if ((age < 14) && (tmpVal == 1))
            {
                is14Test = true;
            }

            tmpVal = Calculator.GetCorrectNumberFromSplitedString(txtNum7.Text.Split(Calculator.Delimiter));
            //pPersonalValues.Add(tmpVal.ToString());
            pPersonalValues.Add(txtNum7.Text);

            tmpVal = Calculator.GetCorrectNumberFromSplitedString(txtPName_Num.Text.Split(Calculator.Delimiter));
            //pPersonalValues.Add(tmpVal.ToString());
            pPersonalValues.Add(txtPName_Num.Text);

            tmpVal = Convert.ToInt16(txtAstroName.Text.Split("(".ToCharArray()[0])[1].Split(")".ToCharArray()[0])[0]);
            pPersonalValues.Add(tmpVal.ToString());
            if ((age < 14) && (tmpVal == 1))
            {
                is14Test = true;
            }
            #endregion

            //double ChakraMapBusinessValue = Calculator.ChakraMap2BusinessValue(pPersonalValues);

            Control.ControlCollection cntrlCls = this.Controls;

            #region Life Cycles
            int curCycle = 0;
            int[] CycleBoundries = new int[4];
            CycleBoundries[0] = Convert.ToInt16(txt1_1.Text.Split("-".ToCharArray()[0])[1].Trim());
            CycleBoundries[1] = Convert.ToInt16(txt2_1.Text.Split("-".ToCharArray()[0])[1].Trim());
            CycleBoundries[2] = Convert.ToInt16(txt3_1.Text.Split("-".ToCharArray()[0])[1].Trim());
            CycleBoundries[3] = Convert.ToInt16(txt4_1.Text.Split("-".ToCharArray()[0])[1].Trim());

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

            int clmx, ccl; //cclng;
            clmx = Calculator.GetCorrectNumberFromSplitedString(cntrlCls.Find("txt" + curCycle.ToString() + "_3", true)[0].Text.Split(Calculator.Delimiter));
            ccl = Calculator.GetCorrectNumberFromSplitedString(cntrlCls.Find("txt" + curCycle.ToString() + "_2", true)[0].Text.Split(Calculator.Delimiter));
            //cclng = Calculator.GetCorrectNumberFromSplitedString(cntrlCls.Find("txt" + curCycle.ToString() + "_4", true)[0].Text.Split(Calculator.Delimiter));

            //pPersonalValues = new List<string>();
            pPersonalValues.Add(clmx.ToString());
            pPersonalValues.Add(ccl.ToString());
            //pPersonalValues.Add(cclng.ToString());

            #endregion Life Cycles

            //double LifeCycleBusinessValue = Calculator.LifeCycle2BusinessValue(pPersonalValues, "כן" == SelfFixYesOrNo);

            string tmpLvl = txtInfo1.Text.Split(System.Environment.NewLine.ToCharArray()[0])[1].Trim();
            if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
            {
            }
            if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
            {
                if (tmpLvl == "Balanced")
                {
                    tmpLvl = "מאוזן";
                }
                if (tmpLvl == "Not Balanced")
                {
                    tmpLvl = "לא מאוזן";
                }
                if (tmpLvl == "Half Balanced")
                {
                    tmpLvl = "חצי מאוזן";
                }
            }
            pPersonalValues.Add(tmpLvl);

            return pPersonalValues;
        }
        private List<string> AttPrblm_GatherInfo()
        {
            List<string> pPersonalValues = new List<string>();
            int age = Convert.ToInt16(txtAge.Text.Split("(".ToCharArray()[0])[0]);

            #region From Chakra Map
            int upperBndry = Convert.ToInt16(txt1_1.Text.Split("-".ToCharArray()[0])[1].Trim());

            int tmpVal = 0;

            // Crown
            tmpVal = Calculator.GetCorrectNumberFromSplitedString(txtNum1.Text.Split(Calculator.Delimiter));
            //pPersonalValues.Add(tmpVal.ToString());
            pPersonalValues.Add(txtNum1.Text);

            //if (age >= upperBndry)
            //{
            //Third Eye
            tmpVal = Calculator.GetCorrectNumberFromSplitedString(txtNum2.Text.Split(Calculator.Delimiter));
            //pPersonalValues.Add(tmpVal.ToString());
            pPersonalValues.Add(txtNum2.Text);
            //}

            tmpVal = Calculator.GetCorrectNumberFromSplitedString(txtNum3.Text.Split(Calculator.Delimiter));
            //pPersonalValues.Add(tmpVal.ToString());
            pPersonalValues.Add(txtNum3.Text);


            tmpVal = Calculator.GetCorrectNumberFromSplitedString(txtNum5.Text.Split(Calculator.Delimiter));
            //pPersonalValues.Add(tmpVal.ToString());
            pPersonalValues.Add(txtNum5.Text);


            tmpVal = Calculator.GetCorrectNumberFromSplitedString(txtNum6.Text.Split(Calculator.Delimiter));
            //pPersonalValues.Add(tmpVal.ToString());
            pPersonalValues.Add(txtNum6.Text);

            tmpVal = Calculator.GetCorrectNumberFromSplitedString(txtNum7.Text.Split(Calculator.Delimiter));
            //pPersonalValues.Add(tmpVal.ToString());
            pPersonalValues.Add(txtNum7.Text);

            tmpVal = Calculator.GetCorrectNumberFromSplitedString(txtPName_Num.Text.Split(Calculator.Delimiter));
            //pPersonalValues.Add(tmpVal.ToString());
            pPersonalValues.Add(txtPName_Num.Text);

            tmpVal = Convert.ToInt16(txtAstroName.Text.Split("(".ToCharArray()[0])[1].Split(")".ToCharArray()[0])[0]);
            pPersonalValues.Add(tmpVal.ToString());
            #endregion

            //double ChakraMapBusinessValue = Calculator.ChakraMap2BusinessValue(pPersonalValues);

            Control.ControlCollection cntrlCls = this.Controls;

            #region Life Cycles
            int curCycle = 0;
            int[] CycleBoundries = new int[4];
            CycleBoundries[0] = Convert.ToInt16(txt1_1.Text.Split("-".ToCharArray()[0])[1].Trim());
            CycleBoundries[1] = Convert.ToInt16(txt2_1.Text.Split("-".ToCharArray()[0])[1].Trim());
            CycleBoundries[2] = Convert.ToInt16(txt3_1.Text.Split("-".ToCharArray()[0])[1].Trim());
            CycleBoundries[3] = Convert.ToInt16(txt4_1.Text.Split("-".ToCharArray()[0])[1].Trim());

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

            int clmx, ccl, cclng;
            clmx = Calculator.GetCorrectNumberFromSplitedString(cntrlCls.Find("txt" + curCycle.ToString() + "_3", true)[0].Text.Split(Calculator.Delimiter));
            ccl = Calculator.GetCorrectNumberFromSplitedString(cntrlCls.Find("txt" + curCycle.ToString() + "_2", true)[0].Text.Split(Calculator.Delimiter));
            cclng = Calculator.GetCorrectNumberFromSplitedString(cntrlCls.Find("txt" + curCycle.ToString() + "_4", true)[0].Text.Split(Calculator.Delimiter));

            //pPersonalValues = new List<string>();
            pPersonalValues.Add(clmx.ToString());
            pPersonalValues.Add(ccl.ToString());
            pPersonalValues.Add(cclng.ToString());

            #endregion Life Cycles

            //double LifeCycleBusinessValue = Calculator.LifeCycle2BusinessValue(pPersonalValues, "כן" == SelfFixYesOrNo);

            string tmpLvl = txtInfo1.Text.Split(System.Environment.NewLine.ToCharArray()[0])[1].Trim();
            if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
            {
            }
            if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
            {
                if (tmpLvl == "Balanced")
                {
                    tmpLvl = "מאוזן";
                }
                if (tmpLvl == "Not Balanced")
                {
                    tmpLvl = "לא מאוזן";
                }
                if (tmpLvl == "Half Balanced")
                {
                    tmpLvl = "חצי מאוזן";
                }
            }
            pPersonalValues.Add(tmpLvl);

            return pPersonalValues;
        }

        private void btnLrnSccssAttPrblm_Click(object sender, EventArgs e)
        {
            LrnSccssAttPrblm();
        }

        public void LrnSccssAttPrblm()
        {
            int cc = CalcCurrentCycle();
            string v = "";
            DateTime dt = DateTimePickerTo.Value;

            List<string> personalInfo;
            double res;
            bool tstPrsn1, tstPrsn2, tstPrsn3, isCarmaticFoundInData, tst71, tst72, tst73, tst74, tst75, tst76;
            bool canReadFrom5 = false;
            List<int> tmpIarr;
            string[] tmpSarr;

            int fromJC = 0, toJC = 0;
            if (isRunningNameBalance == true)// ריצה לאיזון שם  אין צורך הלתאמץ
            {
                fromJC = cc;
                toJC = cc + 1;
            }
            if (isRunningNameBalance == false)// ריצה רגיל להציג הכל
            {
                fromJC = 1;
                toJC = 5;
            }

            for (int j = fromJC; j < toJC; j++)
            {
                string[] temp = (this.Controls.Find("txt" + j.ToString() + "_1", true)[0] as TextBox).Text.Split("-".ToCharArray());
                DateTimePickerTo.Value = mB_Date.AddYears((Convert.ToInt16(temp[0]) + Convert.ToInt16(temp[1])) / 2);

                personalInfo = LearnSccss_GatherInfo();

                #region Success
                res = Calculator.CalcLearnSusccess(personalInfo, cmbtnSccssSelect.SelectedItem.ToString(), is14Test);

                if (res >= 6)
                {
                    txtLeanSccss.Text = Math.Round(res, 1).ToString();
                }
                else
                {
                    txtLeanSccss.Text = "6";
                }

                if (j == cc)
                {
                    v = txtLeanSccss.Text;
                }

                #region intro 4 each cycle
                if (j == 1)
                {
                    if (BLL.AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                    {
                        txtLearnStory.Text += "במחזור ראשון: ";
                    }
                    if (BLL.AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                    {
                        txtLearnStory.Text += "First Cycle: ";
                    }
                }
                if (j == 2)
                {
                    if (BLL.AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                    {
                        txtLearnStory.Text += "במחזור שני: ";
                    }
                    if (BLL.AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                    {
                        txtLearnStory.Text += "Second Cycle: ";
                    }
                }
                if (j == 3)
                {
                    if (BLL.AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                    {
                        txtLearnStory.Text += "במחזור שלישי: ";
                    }
                    if (BLL.AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                    {
                        txtLearnStory.Text += "Third Cycle: ";
                    }
                }
                if (j == 4)
                {
                    if (BLL.AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                    {
                        txtLearnStory.Text += "במחזור רביעי: ";
                    }
                    if (BLL.AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                    {
                        txtLearnStory.Text += "Fourth Cycle: ";
                    }
                }

                txtLearnStory.Text += "(" + txtLeanSccss.Text + ")" + Environment.NewLine;
                #endregion intro 4 each cycle

                #region Sccss Story
                if (res <= 5.9)
                {
                    if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                    {
                        txtLearnStory.Text += "יכולת לימוד חלשה מאוד";
                    }
                    if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                    {
                        txtLearnStory.Text += "Very low learning ability";
                    }
                }
                if ((res >= 6.0) && (res <= 6.9))
                {
                    if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                    {
                        txtLearnStory.Text += "יכולת לימוד בינונית ומטה";
                    }
                    if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                    {
                        txtLearnStory.Text += "Learning capability mediocre";
                    }
                }
                if ((res >= 7.0) && (res <= 7.9))
                {
                    if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                    {
                        txtLearnStory.Text += "יכולת לימוד בינונית ומעלה";
                    }
                    if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                    {
                        txtLearnStory.Text += "Medium-learning ability and up";
                    }
                }
                if ((res >= 8.0) && (res <= 8.5))
                {
                    if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                    {
                        txtLearnStory.Text += "יכולת לימוד טובה";
                    }
                    if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                    {
                        txtLearnStory.Text += "Good learning capability";
                    }
                }
                if ((res >= 8.6) && (res <= 8.9))
                {
                    if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                    {
                        txtLearnStory.Text += "יכולת לימוד טובה מאוד";
                    }
                    if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                    {
                        txtLearnStory.Text += "Very good learning ability";
                    }
                }
                if ((res >= 9.0))
                {
                    if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                    {
                        txtLearnStory.Text += "יכולת הצטיינות בלימודים";
                    }
                    if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                    {
                        txtLearnStory.Text += "Capacity for academic excellence";
                    }
                }
                #endregion Sccss Story

                tstPrsn1 = Calculator.isCarmaticNumber(Calculator.GetCorrectNumberFromSplitedString(txtNum1.Text.Split(Calculator.Delimiter)));
                tstPrsn2 = Calculator.isCarmaticNumber(Calculator.GetCorrectNumberFromSplitedString(txtNum5.Text.Split(Calculator.Delimiter)));
                tstPrsn3 = Calculator.isCarmaticNumber(Calculator.GetCorrectNumberFromSplitedString(txtPName_Num.Text.Split(Calculator.Delimiter)));

                isCarmaticFoundInData = false;
                for (int i = 0; i < personalInfo.Count - 2; i++)
                {
                    string tmpS1 = personalInfo[i];
                    isCarmaticFoundInData |= Calculator.isCarmaticNumber(Calculator.GetCorrectNumberFromSplitedString(tmpS1.Split(Calculator.Delimiter)));
                }

                if ((res < 8.0) && (isCarmaticFoundInData == true))
                {
                    if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                    {
                        txtLearnStory.Text += Environment.NewLine + "השגת יעדי הלמידה יתאפשרו אם התלמיד ישקיע בהכנת שיעורים מעל שעתיים ביום.";
                    }
                    if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                    {
                        txtLearnStory.Text += Environment.NewLine + "Achievement of the learning goals will be possible if the student invest in homework over two hours a day.";
                    }
                }

                tst71 = Calculator.isNumInArray(txtNum1.Text.Split(Calculator.Delimiter), 7);
                tst72 = Calculator.isNumInArray(txtNum3.Text.Split(Calculator.Delimiter), 7);
                tst73 = Calculator.isNumInArray(txtNum5.Text.Split(Calculator.Delimiter), 7);
                tst74 = (Convert.ToInt16(txtAstroName.Text.Split("(".ToCharArray()[0])[1].Split(")".ToCharArray()[0])[0]) == 7);

                tmpIarr = new List<int>();
                tmpSarr = personalInfo[personalInfo.Count - 2].Split(Calculator.Delimiter);
                for (int i = 0; i < tmpSarr.Length; i++)
                {
                    tmpIarr.Add(Convert.ToInt16(tmpSarr[i]));
                }
                tmpSarr = personalInfo[personalInfo.Count - 3].Split(Calculator.Delimiter);
                for (int i = 0; i < tmpSarr.Length; i++)
                {
                    tmpIarr.Add(Convert.ToInt16(tmpSarr[i]));
                }
                tst75 = Calculator.isNumInList(7, tmpIarr);

                canReadFrom5 = false;
                if (tst71 || tst72 | tst73 || tst74 || tst75)
                {
                    if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                    {
                        txtLearnStory.Text += Environment.NewLine + "יכולת קריאה כבר מגיל 5 (גן ילדים)";
                    }
                    if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                    {
                        txtLearnStory.Text += Environment.NewLine + "The ability to read from the age of 5 (Kindergarten)";
                    }
                    canReadFrom5 = true;
                }

                txtLearnStory.Text += Environment.NewLine + Environment.NewLine;
                #endregion Success
            }
            txtLeanSccss.Text = v;
            DateTimePickerTo.Value = dt;
            // ------------------------------------------------------------------------

            #region Attension Problems
            txtAttStory.Clear();

            personalInfo = AttPrblm_GatherInfo();
            int[] MjrCarmaticValues, MnrCarmaticValues;
            res = Calculator.CalcAttentionProblems(personalInfo, out MjrCarmaticValues, out MnrCarmaticValues);

            if (MjrCarmaticValues.Length > 1)
            {
                BubbleSort(ref MjrCarmaticValues);
            }

            if (MnrCarmaticValues.Length > 1)
            {
                BubbleSort(ref MnrCarmaticValues);
            }

            txtAtt.Text = Math.Round(res, 1).ToString();

            #region AttentionProblems Story
            #region numbers
            string out2txt = "";
            if (MjrCarmaticValues[0] != -1)
            {
                for (int i = 0; i < MjrCarmaticValues.Length; i++)
                {
                    out2txt += MjrCarmaticValues[i].ToString() + ",";
                }
                txtAttMajor.Text = out2txt.Substring(0, out2txt.Length - 1);
            }
            else
            {
                txtAttMajor.Text = "";
            }

            out2txt = "";
            if (MnrCarmaticValues[0] != -1)
            {
                for (int i = 0; i < MnrCarmaticValues.Length; i++)
                {
                    out2txt += MnrCarmaticValues[i].ToString() + ",";
                }
                txtAttMinor.Text = out2txt.Substring(0, out2txt.Length - 1);
            }
            else
            {
                txtAttMinor.Text = "";
            }

            if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
            {
                if (personalInfo[personalInfo.Count - 1] == "לא מאוזן")
                {
                    txtAttMinor.Text = out2txt + personalInfo[personalInfo.Count - 1];
                }
            }
            if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
            {
                if (personalInfo[personalInfo.Count - 1] == "Not Balanced")
                {
                    txtAttMinor.Text = out2txt + personalInfo[personalInfo.Count - 1];
                }
            }
            #endregion numbers

            if ((txtAttMajor.Text.Trim() == "") && (txtAttMinor.Text.Trim() == ""))
            {
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                {
                    txtAttStory.Text = "אין בעיות של קשב וריכוז";
                }
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                {
                    txtAttStory.Text = "No issues of ADHD";
                }
            }


            tst71 = (Calculator.GetCorrectNumberFromSplitedString(txtPName_Num.Text.Split(Calculator.Delimiter)) == 13);
            tst72 = (Calculator.GetCorrectNumberFromSplitedString(txtNum5.Text.Split(Calculator.Delimiter)) == 13);
            tst73 = (Calculator.GetCorrectNumberFromSplitedString(txtNum1.Text.Split(Calculator.Delimiter)) == 13);
            tst74 = (Calculator.GetCorrectNumberFromSplitedString(personalInfo[personalInfo.Count - 2].Split(Calculator.Delimiter)) == 13);
            tst75 = (Calculator.GetCorrectNumberFromSplitedString(personalInfo[personalInfo.Count - 3].Split(Calculator.Delimiter)) == 13);
            tst76 = (Calculator.GetCorrectNumberFromSplitedString(txtNum3.Text.Split(Calculator.Delimiter)) == 13);

            if (tst71 || tst72 | tst73 || tst74 || tst75 || tst76)
            {
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                {
                    txtAttStory.Text += Environment.NewLine + "נטייה לאלימות פיזית ו/או מילולית";
                }
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                {
                    txtAttStory.Text += Environment.NewLine + "Tendency for Physical and\\or literal Violence";
                }
            }


            tst71 = (Calculator.GetCorrectNumberFromSplitedString(txtPName_Num.Text.Split(Calculator.Delimiter)) == 14);
            tst72 = (Calculator.GetCorrectNumberFromSplitedString(txtNum5.Text.Split(Calculator.Delimiter)) == 14);
            tst73 = (Calculator.GetCorrectNumberFromSplitedString(txtNum1.Text.Split(Calculator.Delimiter)) == 14);
            tst74 = (Calculator.GetCorrectNumberFromSplitedString(personalInfo[personalInfo.Count - 2].Split(Calculator.Delimiter)) == 14);
            tst75 = (Calculator.GetCorrectNumberFromSplitedString(personalInfo[personalInfo.Count - 3].Split(Calculator.Delimiter)) == 14);
            tst76 = (Calculator.GetCorrectNumberFromSplitedString(txtNum3.Text.Split(Calculator.Delimiter)) == 14);

            if (tst71 || tst72 | tst73 || tst74 || tst75 || tst76)
            {
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                {
                    txtAttStory.Text += Environment.NewLine + "נטייה לסכסוכים והמרדה אחד כלפי השני וכלפי הזולת";
                }
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                {
                    txtAttStory.Text += Environment.NewLine + "Tendency to conflict and rebellion towards each other and towards others";
                }
            }


            tst71 = (Calculator.GetCorrectNumberFromSplitedString(txtPName_Num.Text.Split(Calculator.Delimiter)) == 16);
            tst72 = (Calculator.GetCorrectNumberFromSplitedString(txtNum5.Text.Split(Calculator.Delimiter)) == 16);
            tst73 = (Calculator.GetCorrectNumberFromSplitedString(txtNum1.Text.Split(Calculator.Delimiter)) == 16);
            tst74 = (Calculator.GetCorrectNumberFromSplitedString(personalInfo[personalInfo.Count - 2].Split(Calculator.Delimiter)) == 16);
            tst75 = (Calculator.GetCorrectNumberFromSplitedString(personalInfo[personalInfo.Count - 3].Split(Calculator.Delimiter)) == 16);
            tst76 = (Calculator.GetCorrectNumberFromSplitedString(txtNum3.Text.Split(Calculator.Delimiter)) == 16);

            if (tst71 || tst72 | tst73 || tst74 || tst75 || tst76)
            {
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                {
                    txtAttStory.Text += Environment.NewLine + "נטייה למופנמות ודחף עז ללימודים והתרחקות חברתית";
                }
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                {
                    txtAttStory.Text += Environment.NewLine + "Inward orientation and a strong urge to study and social distancing";
                }

            }


            tst71 = (Calculator.GetCorrectNumberFromSplitedString(txtPName_Num.Text.Split(Calculator.Delimiter)) == 19);
            tst72 = (Calculator.GetCorrectNumberFromSplitedString(txtNum5.Text.Split(Calculator.Delimiter)) == 19);
            tst73 = (Calculator.GetCorrectNumberFromSplitedString(txtNum1.Text.Split(Calculator.Delimiter)) == 19);
            tst74 = (Calculator.GetCorrectNumberFromSplitedString(personalInfo[personalInfo.Count - 2].Split(Calculator.Delimiter)) == 19);
            tst75 = (Calculator.GetCorrectNumberFromSplitedString(personalInfo[personalInfo.Count - 3].Split(Calculator.Delimiter)) == 19);
            tst76 = (Calculator.GetCorrectNumberFromSplitedString(txtNum3.Text.Split(Calculator.Delimiter)) == 19);

            if (tst71 || tst72 | tst73 || tst74 || tst75 || tst76)
            {
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                {
                    txtAttStory.Text += Environment.NewLine + "נטייה לקורבנות (להיות קורבן) או לקרבן (בריונות)";
                }
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                {
                    txtAttStory.Text += Environment.NewLine + "Tendency to being a victim or bullying";
                }
            }


            tst71 = Calculator.isNumInArray(txtNum1.Text.Split(Calculator.Delimiter), 1);
            tst72 = Calculator.isNumInArray(txtNum3.Text.Split(Calculator.Delimiter), 1);
            tst73 = Calculator.isNumInArray(txtNum5.Text.Split(Calculator.Delimiter), 1);
            tst74 = (Convert.ToInt16(txtAstroName.Text.Split("(".ToCharArray()[0])[1].Split(")".ToCharArray()[0])[0]) == 1);

            tmpIarr = new List<int>();
            tmpSarr = personalInfo[personalInfo.Count - 2].Split(Calculator.Delimiter);
            for (int i = 0; i < tmpSarr.Length; i++)
            {
                tmpIarr.Add(Convert.ToInt16(tmpSarr[i]));
            }
            tmpSarr = personalInfo[personalInfo.Count - 3].Split(Calculator.Delimiter);
            for (int i = 0; i < tmpSarr.Length; i++)
            {
                tmpIarr.Add(Convert.ToInt16(tmpSarr[i]));
            }
            tst75 = Calculator.isNumInList(1, tmpIarr);

            List<int> cv = new List<int> { Calculator.GetCorrectNumberFromSplitedString(txtNum1.Text.Split(Calculator.Delimiter)),
                                           Calculator.GetCorrectNumberFromSplitedString(txtNum2.Text.Split(Calculator.Delimiter)),
                                           Calculator.GetCorrectNumberFromSplitedString(txtNum3.Text.Split(Calculator.Delimiter)),
                                           Calculator.GetCorrectNumberFromSplitedString(txtNum5.Text.Split(Calculator.Delimiter)),
                                           Convert.ToInt16(txtAstroName.Text.Split(" ".ToCharArray()[0])[1].Replace("("," ").Replace (")"," ").Trim()),
                                           Calculator.GetCorrectNumberFromSplitedString(this.Controls.Find("txt" + CalcCurrentCycle().ToString()  + "_2",true)[0].Text.Split(Calculator.Delimiter)),
                                           Calculator.GetCorrectNumberFromSplitedString(this.Controls.Find("txt" + CalcCurrentCycle().ToString()  + "_3",true)[0].Text.Split(Calculator.Delimiter))};
            if (tst71 || tst72 | tst73 || tst74 || tst75)
            {
                if ((canReadFrom5 == false) && (cv.Contains(7) == false))
                {
                    if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                    {
                        txtLearnStory.Text += Environment.NewLine + "קושי בקריאה עד גיל עשר";
                    }
                    if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                    {
                        txtLearnStory.Text += Environment.NewLine + "Difficulty reading until the age of ten";
                    }
                }
            }
            #endregion AttentionProblems Story

            #endregion Attension Problems
        }
        #endregion

        private void btnSearchDB_Click(object sender, EventArgs e)
        {

        }

        private void searchDB()
        {
            dgvSearch.Rows.Clear();

            //DataGridViewRow row2insert;
            foreach (DataGridViewRow curRow in gvDBview.Rows)
            {
                if (isWordInDataRow(curRow, txtSearchText.Text) == true)
                {
                    object[] objs = new string[16];
                    objs.Initialize();

                    for (int i = 0; i < curRow.Cells.Count; i++)
                    {
                        objs.SetValue(curRow.Cells[i].Value.ToString(), i);
                    }


                    dgvSearch.Rows.Insert(dgvSearch.Rows.Count, objs);
                }
            }
        }

        private void btnResetSearch_Click(object sender, EventArgs e)
        {
            dgvSearch.Rows.Clear();
            txtSearchText.Text = "";
        }

        private bool isWordInDataRow(DataGridViewRow row, string word2Find)
        {
            bool res = false;

            foreach (DataGridViewCell cell in row.Cells)
            {
                res |= isSectionInWord(word2Find, cell.Value.ToString());
            }

            return res;
        }

        private bool isSectionInWord(string Section, string inWord)
        {
            bool res = false;

            if (Section.Length > inWord.Length)
            {
                return false;
            }

            if (Section.Length == inWord.Length)
            {
                return (Section == inWord);
            }

            if (Section.Length < inWord.Length)
            {
                for (int i = 0; i < inWord.Length - Section.Length + 1; i++)
                {
                    if (Section == inWord.Substring(i, Section.Length))
                    {
                        res = true;
                    }
                }
            }

            return res;
        }

        private void btnLoad2Project2_Click(object sender, EventArgs e)
        {
            //if (gvDBview.SelectedRows.Count == 0)
            //{
            //    MessageBox.Show("צריך לבחור שורה שלמה", "שגיאה בסימון שורה", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //    return;
            //}

            if (dgvSearch.SelectedRows.Count == 0)
            {
                return;
            }

            ClearForm();

            checkBox1.Checked = false;
            checkBox1.CheckState = CheckState.Unchecked;
            checkBox1_CheckedChanged(null, null);

            //int RowNum = gvDBview.SelectedRows[0].Index;
            int RowNum = dgvSearch.SelectedCells[0].RowIndex;
            DataGridViewRow CurClient = dgvSearch.SelectedRows[0];
            /*
            DataGridViewRow gvClientRow = gvDBview.Rows[RowNum];

            DBDataSet.ClientsRow CurClient = omegaDataSet.Clients.NewRow() as DBDataSet.ClientsRow;
            CurClient = omegaDataSet.Clients.Rows[RowNum] as DBDataSet.ClientsRow;

            txtPrivateName.Text = CurClient.PrivateName;
            txtFamilyName.Text = CurClient.LastName;
            txtMotherName.Text = CurClient.MotherName;
            txtFatherName.Text = CurClient.FatherName;
            DateTimePickerFrom.Value = CurClient.B_Date;
            txtCity.Text = CurClient.City;
            txtStreet.Text = CurClient.Street;
            txtBiuldingNum.Text = CurClient.BuildingNum.ToString();
            txtAppNum.Text = CurClient.AppNum.ToString();
            txtEMail.Text = CurClient.EMail.ToString();
            txtPhones.Text = CurClient.Phones.ToString();
            */
            txtPrivateName.Text = CurClient.Cells[1].Value.ToString();
            txtFamilyName.Text = CurClient.Cells[2].Value.ToString();
            txtMotherName.Text = CurClient.Cells[4].Value.ToString();
            txtFatherName.Text = CurClient.Cells[3].Value.ToString();
            DateTimePickerFrom.Value = DateTime.Parse(CurClient.Cells[5].Value.ToString());
            txtCity.Text = CurClient.Cells[6].Value.ToString();
            txtStreet.Text = CurClient.Cells[7].Value.ToString();
            txtBiuldingNum.Text = CurClient.Cells[8].Value.ToString();
            txtAppNum.Text = CurClient.Cells[9].Value.ToString();
            txtEMail.Text = CurClient.Cells[10].Value.ToString();
            txtPhones.Text = CurClient.Cells[11].Value.ToString();
            EnumProvider.Sex sex = EnumProvider.Sex.Male;
            EnumProvider.PassedRectification passedrect = EnumProvider.PassedRectification.NotPassed;
            EnumProvider.ReachedMaster reachmaster = EnumProvider.ReachedMaster.No;

            if (BLL.AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
            {
                sex = EnumProvider.Instance.GetSexEnumFromString(CurClient.Cells[13].Value.ToString());
                passedrect = EnumProvider.Instance.GetPassedRectificationEnumFromString(CurClient.Cells[14].Value.ToString());
                reachmaster = EnumProvider.Instance.GGetIsMasterEnumFromString(CurClient.Cells[15].Value.ToString());
            }
            if (BLL.AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
            {
                string s = CurClient.Cells[13].Value.ToString(); // SEX
                if (s.ToLower() == EnumProvider.Sex.Female.ToString().ToLower())
                {
                    sex = EnumProvider.Sex.Female;
                }
                if (s.ToLower() == EnumProvider.Sex.Male.ToString().ToLower())
                {
                    sex = EnumProvider.Sex.Male;
                }

                s = CurClient.Cells[14].Value.ToString(); // RECTIFICATION
                if (s.ToLower() == EnumProvider.PassedRectification.NotPassed.ToString().ToLower())
                {
                    passedrect = EnumProvider.PassedRectification.NotPassed;
                }
                if (s.ToLower() == EnumProvider.PassedRectification.Passed.ToString().ToLower())
                {
                    passedrect = EnumProvider.PassedRectification.Passed;
                }

                s = CurClient.Cells[15].Value.ToString(); // MASTER
                if (s.ToLower() == EnumProvider.ReachedMaster.No.ToString().ToLower())
                {
                    reachmaster = EnumProvider.ReachedMaster.No;
                }
                if (s.ToLower() == EnumProvider.ReachedMaster.Yes.ToString().ToLower())
                {
                    reachmaster = EnumProvider.ReachedMaster.Yes;
                }
            }

            if (sex == EnumProvider.Sex.Male)
            {
                cmbSexSelect.Text = cmbSexSelect.Items[0].ToString();
                //cmbSexSelect.Select(0, 1);
                cmbSexSelect.SelectedIndex = cmbSexSelect.FindStringExact(cmbSexSelect.Text);
            }
            else
            {
                cmbSexSelect.Text = cmbSexSelect.Items[1].ToString();
                //cmbSexSelect.Select(1, 1);
                cmbSexSelect.SelectedIndex = cmbSexSelect.FindStringExact(cmbSexSelect.Text);
            }
            if (passedrect == EnumProvider.PassedRectification.NotPassed)
            {
                cbMainCorrectionDone.Checked = false;
                cmbSelfFix.SelectedIndex = cmbSelfFix.FindStringExact(cmbSelfFix.Items[0].ToString());
            }
            else
            {
                cbMainCorrectionDone.Checked = true;
                cmbSelfFix.SelectedIndex = cmbSelfFix.FindStringExact(cmbSelfFix.Items[1].ToString());
            }
            //if (reachmaster == EnumProvider.ReachedMaster.No)
            //{
            //    cbPersonMaster.Checked = false;
            //}
            //else
            //{
            //    cbPersonMaster.Checked = true;
            //}

            tcInput.SelectTab("ClientDataPage");
        }

        private void txtSearchText_TextChanged(object sender, EventArgs e)
        {
            searchDB();
        }

        private void btnMailTo_Click(object sender, EventArgs e)
        {
            string Email = txtEMail.Text.Split(",".ToCharArray()[0])[0];
            if (Email.Length > 0)
            {
                string str = "mailto:" + Email + "";
                System.Diagnostics.Process.Start(str);
            }
        }

        private void dgvSearch_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void tabPage1_Click(object sender, EventArgs e)
        {

        }

        private void lblMrrgInfo_Click(object sender, EventArgs e)
        {

        }

        private void groupBox12_Enter(object sender, EventArgs e)
        {

        }

        private void button3_Click_1(object sender, EventArgs e)
        {
            DeleteRecordFromXMLDB(ref dgvSearch);
        }

        private void button86_Click(object sender, EventArgs e)
        {
            textBox48.Text = "";
        }

        private void button87_Click(object sender, EventArgs e)
        {
            textBox48.Text = ReportDataProvider.Instance.ReadFromTextFile(Reports.ReportDataProvider.Instance.ConstructFilePath2FianlySummary());
        }

        private void button85_Click(object sender, EventArgs e)
        {
            ReportDataProvider.Instance.WriteText2File(Reports.ReportDataProvider.Instance.ConstructFilePath2FianlySummary(), textBox48.Text);
        }

        private void cmbGenderSelect_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cmbGenderSelect.SelectedIndex == 0) // זכר
            {
                Reports.ReportDataProvider.Instance.Gender = EnumProvider.Sex.Male;
            }
            else // נקבה
            {
                Reports.ReportDataProvider.Instance.Gender = EnumProvider.Sex.Female;
            }
        }

        private Omega.Enums.EnumProvider.Balance GetBalanceStatus()
        {
            string[] sTexts = txtInfo1.Text.Trim().Split(Environment.NewLine.ToCharArray()[0]);
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

        private void txtQfilter_TextChanged(object sender, EventArgs e)
        {
            string sf = txtQfilter.Text;
            if (sf.Trim().Length == 0)
            {
                cmbQ.Visible = true;
                cmbQF.Visible = false;
            }
            else
            {
                cmbQ.Visible = false;

                cmbQF.Visible = true;
                cmbQF.Items.Clear();

                for (int i = 0; i < cmbQ.Items.Count; i++)
                {
                    if (cmbQ.Items[i].ToString().Contains(sf) == true)
                    {
                        cmbQF.Items.Add(cmbQ.Items[i]);
                    }
                }
            }

        }

        private void button5_Click(object sender, EventArgs e)
        {
            txtQfilter.Text = "";
            txtQfilter_TextChanged(null, null);
        }

        public void cmdQnAcalc_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(txtPrivateName.Text) ||
                string.IsNullOrEmpty(txtFamilyName.Text) ||
                string.IsNullOrEmpty(DateTimePickerFrom.Text) ||
                string.IsNullOrEmpty(txtAstroName.Text) ||
                string.IsNullOrEmpty(txtPName_Num.Text) ||
                string.IsNullOrEmpty(txtFName_Num.Text))
            {
                string msg = "על מנת לענות על שאלות, תחילה יש לטעון לקוח ולבצע חישוב";
                string caption = "נומרולוגיית הצ'אקרות הדינמיות";
                MessageBox.Show(msg, caption, MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1, MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);

                return;

            }

            #region Q_n_A
            /*
    01. עתידות קלאסי
    02. מתי כדאי לי להתחתן?
    03. מתי כדאי לי לרכשו דירה?
    04. מתי כדאי לי לפתוח עסק?
    05. מתי כדאי לי לבצע פעולה עסקית?
    06. מתי כדאי לי לגשת למבחן תיאוריה או טסט לרכב?
    07. מהו הזמן המתאים למכור דירה?
    08. מתי כדאי לי לשנות שם?
    09. שותפיות עסקיות
    10. האם הוא אוהב אותי?
    11. מתי אמצא זוגיות?
    12. האם בן הזוג שלי מתאים לי?
    13. האם אנחנו נתחתן?
    14. האם כדאי להתחתן עם בן הזוג?
    15. האם הוא יחזור אלי?
    16. האם כדאי שיחזור אלי?
    17. האם הוא בוגד בי?
    18. למה הזוגיות מתעכבת?
    19. מתי אסתדר כלכלית?
    20. מתי אמצא עבודה?
    -- Removed 19. האם לעבור לעבודה חדשה?
    -- Removed 20. האם לעשות שינוי בקריירה?
    21. Prev. 21. במה מתאים לי לעסוק?
    22. Prev. 22. מתי אקבל קידום בעבודה?
    23. Prev. 23. האם כדאי לפתוח עסק?
    24. Prev. 24. באיזה תחום כדאי לפתוח עסק?
    25. Prev. 25. מהי הדרך הנכונה להצלחה כלכלית?
    -- Removed 26. מתי דברים יתחילו להסתדר עבורי?
    -- Removed 27. האם השם שלי מתאים לי?
    -- Missed 28. מה צופן לי העתיד?
    -- Missed 29. למה כלכך קשה לי בחיים?
    -- Missed 30. מתי יפתח לי המזל?
    -- Missed 31. מתי בן\בת ימצאו זוגיות?
    -- Missed 32. האם הבן\בת יתחתנו בקרוב?
    -- Missed 33. האם אצליח להכנס להריון?
    -- Missed 34. האם כדאי לעבור דירה?
    -- Missed 35. האם עכשיו זה זמן טוב לקנות דירה?
    26. האם זה הזמן המתאים לעבור דירה?
            */
            #endregion
            // רשימת שאלות ורשימה מסוננת - שני פקדים זה מעל זה
            // switch visability by filter

            #region Test Question Selected
            //int testint = 0;
            if ((cmbQ.SelectedItem == null) && (cmbQF.SelectedItem == null))
            {
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                {
                    MessageBox.Show("Select a question", "Questions and Answers");
                }
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                {
                    MessageBox.Show("יש לבחור שאלה", "שאלות ותשובות");
                }
                return;
            }
            //if ( (int.TryParse(cmbQ.SelectedItem.ToString().Substring(0,1), out testint) == false) ||
            //    (int.TryParse(cmbQF.SelectedItem.ToString().Substring(0,1), out testint) == false) )
            //{
            //    if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
            //    {
            //        MessageBox.Show("Select a question", "Questions and Answers");
            //    }
            //    if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
            //    {
            //        MessageBox.Show("יש לבחור שאלה", "שאלות ותשובות");
            //    }
            //    return;
            //}
            #endregion Test Question Selected

            #region Q and A - Filters
            // to show in output table
            // מגנון8 תמיד בהתחלה
            List<string> slQAfilter = new List<string>() {  "-1",
                                                            ",1,2,3,4,5,6,8,9,11,22,33,",
                                                            ",1,3,4,5,8,9,11,22,",//",1,2,4,6,8,9,11,22,33,",
                                                            ",1,3,4,5,8,9,11,22,",//",1,4,8,9,11,22,",
                                                            ",1,3,4,5,8,9,11,22,",
                                                            ",1,2,3,4,5,6,8,9,11,14,19,22,33,",
                                                            ",1,2,3,4,5,6,8,9,11,14,19,22,33,",
                                                            ",1,2,3,4,5,6,8,9,11,14,19,22,33,",
                                                            ",1,2,4,6,8,9,11,22,33,",
                                                            ",7,13,14,16,19,#,13,14,16,19,", //#8 splitting 2 options
                                                            ",7,13,14,16,19,#,13,14,16,19,#,1,2,3,4,5,6,8,9,11,22,33,",
                                                            "",
                                                            "",
                                                            "",
                                                            ",7,13,14,16,19,#,13,14,16,19,",//13
                                                            "",
                                                            ",3,5,9,11,22,14,#,14,5,#,7,16,#,16,",//15
                                                            ",7,13,14,16,19,#,13,14,16,19,",//16
                                                            ",1,3,5,9,11,22,13,14,16,19,#,13,14,16,19,#,2,4,6,7,13,14,16,19,33,#,9,11,22,",//17
                                                            ",1,2,3,4,5,6,8,9,11,22,33,#,7,13,14,16,19,#,13,14,16,19,#,2,4,6,7,13,14,16,19,33,#,9,11,22,",
                                                            ",1,3,4,5,8,9,11,22,#,7,13,14,16,19,#,13,14,16,19,",
                                                            //19"",
                                                            //20"",
                                                            ",1,3,5,8,9,11,22,",//22
                                                            "",
                                                            "",
                                                            ",2,4,6,7,33,14,16,19,"/*25*/,
                                                            ",1,3,4,5,8,9,11,",
                                                            "",
                                                            ""};
            #endregion

            if (isRunningQnA == false)
            {
                txtQnAres.Text = "";
                dgvQnA.Rows.Clear();
                txtOut = new List<string>();
            }

            bool py = cbQPYmode.Checked;

            string sCurrentFilter = "";
            #region Get Correct Filter
            string sQ = "";
            if (txtQfilter.Text.Trim().Length == 0)
            {
                sQ = cmbQ.SelectedItem.ToString();
                sCurrentFilter = slQAfilter[cmbQ.SelectedIndex];
            }
            else
            {
                sQ = cmbQF.SelectedItem.ToString();
                sCurrentFilter = slQAfilter[Convert.ToInt16(sQ.Split(".".ToCharArray()[0])[0]) - 1];
            }
            #endregion Get Correct Filter

            List<DateTime> DateList = CalcDates2FutureCalc(Convert.ToInt16(cmbInterval.Text), cmbType.Text, Convert.ToInt16(cmbTimes.Text));

            //DateTime futuredate = DateList[i];//DateTime.Now.AddDays(days * i);

            //string pY, pM, pD, strDate;
            //Calculator.CalcPersonalInfo(mB_Date, futuredate, out pY, out pM, out pD);

            //strDate = futuredate.Day.ToString() + "." + futuredate.Month.ToString() + "." + futuredate.Year.ToString() + "  (" + FlipString(Calculator.FateCalc(futuredate)) + ")";

            //DataGridViewRow dR;
            //dR = new DataGridViewRow();

            string[] spltQ = "".Split(",".ToCharArray()[0]); //  init
            bool testText = false;

            currentQ = Convert.ToInt16(sQ.Split(".".ToCharArray()[0])[0]);

            UserResult thisRes = new UserResult();
            thisRes.GatherDataFromGUI(this.Controls, CalcCurrentCycle());

            dgvQnA.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgvQnA.Columns[1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgvQnA.Columns[2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgvQnA.Columns[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgvQnA.Columns[4].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgvQnA.Columns[5].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

            dgvQnA.Columns[0].Width = 72;
            dgvQnA.Columns[1].Width = 66;
            dgvQnA.Columns[2].Width = 48;
            dgvQnA.Columns[3].Width = 48;
            dgvQnA.Columns[4].Width = 48;
            dgvQnA.Columns[5].Width = 48;
            dgvQnA.Columns[5].Visible = true;

            dgvQnA.Update();

            switch (currentQ)
            {
                case 1:
                case 2:
                case 3:
                case 4:
                case 5:
                case 6:
                case 7:
                case 8:
                case 9:
                    #region 9
                    for (int i = 0; i < DateList.Count; i++)
                    {
                        DateTime futuredate = DateList[i];//DateTime.Now.AddDays(days * i);

                        string pY, pM, pD, strDate, strFate;
                        Calculator.CalcPersonalInfo(mB_Date, futuredate, out pY, out pM, out pD);

                        DataGridViewRow dR = new DataGridViewRow();
                        //dataFuture.Rows.Add(1);

                        //strDate = futuredate.Day.ToString() + "." + futuredate.Month.ToString() + "." + futuredate.Year.ToString() + "  (" + FlipString(Calculator.FateCalc(futuredate)) + ")";
                        //strDate = futuredate.ToString("dd.MM.yyyy") + "  (" + FlipString(Calculator.FateCalc(futuredate)) + ")";
                        strDate = futuredate.ToString("dd.MM.yyyy");
                        strFate = FlipString(Calculator.FateCalc(futuredate));
                        #region Climax - Future
                        DateTime newBD = futuredate;
                        int[] fLC = Calculator.CareteLifeCycles(newBD);

                        // first climax
                        string ftrClmx = Calculator.CalcSum(fLC[2]);
                        #endregion

                        #region Question Filter
                        spltQ = sQ.Split(".".ToCharArray()[0]);

                        bool passedfilter = true;
                        int tmpPy = Calculator.GetCorrectNumberFromSplitedString(pY.Split(Calculator.Delimiter));
                        int tmpPm = Calculator.GetCorrectNumberFromSplitedString(pM.Split(Calculator.Delimiter));
                        int tmpPd = Calculator.GetCorrectNumberFromSplitedString(pD.Split(Calculator.Delimiter));
                        int tmpFate = Calculator.GetCorrectNumberFromSplitedString(Calculator.FateCalc(futuredate).Split(Calculator.Delimiter));
                        int tmpFtrClmx = Calculator.GetCorrectNumberFromSplitedString(ftrClmx.Split(Calculator.Delimiter));

                        if (Convert.ToInt16(spltQ[0]) != 1) // with filter
                        {
                            if (Calculator.isCarmaticNumber(futuredate.Day) == true)
                            {
                                passedfilter = false;
                            }
                            else
                            {
                                //string sCurFilter = slQAfilter[Convert.ToInt16(spltQ[0]) - 1];
                                //cbPYtake
                                if (cbQPYmode.Checked == true) //true
                                {
                                    passedfilter &= sCurrentFilter.Contains("," + tmpPy.ToString() + ",");
                                }

                                passedfilter &= sCurrentFilter.Contains("," + tmpPm.ToString() + ",");
                                passedfilter &= sCurrentFilter.Contains("," + tmpPd.ToString() + ",");
                                passedfilter &= sCurrentFilter.Contains("," + tmpFate.ToString() + ",");
                                passedfilter &= sCurrentFilter.Contains("," + tmpFtrClmx.ToString() + ",");
                            }

                            if (Convert.ToInt16(sQ.Split(".".ToCharArray()[0])[0]) == 8) // only for Q6
                            {
                                #region 6
                                string s1 = "," + tmpPy + "," + tmpPm + "," + tmpPd + "," + tmpFate + "," + tmpFtrClmx + ",";

                                if ((s1.Contains(",14,") == true) || (s1.Contains(",19,") == true))
                                {
                                    testText = true;

                                    if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                                    {
                                        txtOut.Add("כאשר המספרים 14 או 19 מופיעים: מומלץ לבצע את הפעולה מתוך תודעה שהשינוי הוא לטובתך העליונה.");
                                    }
                                    if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                                    {
                                        txtOut.Add("When 14 or 19 apperas: It is recomended to preform the action of change with the knowledge that it is the best for you to do.");
                                    }
                                }
                                if (s1.Contains(",19,") == true)
                                {
                                    testText = true;

                                    if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                                    {
                                        txtOut.Add("כאשר המספר 19 מופיע : מומלץ להתכוונן שעליו לנצל הזדמנויות להצליח בפעולתיו.");
                                    }
                                    if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                                    {
                                        txtOut.Add("When 19 apperas: It is recomended to prepare to take adventage on every opprotunity inorder to achieve success.");
                                    }
                                }
                                #endregion 6
                            }
                            if (Convert.ToInt16(sQ.Split(".".ToCharArray()[0])[0]) == 9) // only for Q9
                            {
                                #region 9
                                testText = true;
                                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                                {
                                    txtOut.Add("בדוק את ההצלחה העסקית של השותפים בלשונית התאמה עסקית -> שותפים ");
                                }
                                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                                {
                                    txtOut.Add("Test for the Success rate value of the Business matching tab -> Combined");
                                }
                                #endregion 9
                            }
                        }


                        #endregion Question Filter

                        if (passedfilter == true)
                        {
                            object[] objects = new string[6];
                            objects.Initialize();
                            objects.SetValue(strDate, 0);
                            objects.SetValue(strFate, 1);
                            objects.SetValue(FlipString(pY), 2);
                            objects.SetValue(FlipString(pM), 3);
                            objects.SetValue(FlipString(pD), 4);
                            objects.SetValue(FlipString(ftrClmx), 5);

                            dgvQnA.Rows.Insert(0, objects);

                            //if (dgvQnA.Rows.Count == 0)
                            //{
                            //    dgvQnA.Rows.Insert(0, objects);
                            //}
                            //else
                            //{
                            //    dgvQnA.Rows.Insert(dgvQnA.Rows.Count - 1, objects);

                            //}
                            //dgvQnA.Update();
                        }
                        //txtFuture.Text += futuredate.Date.ToShortDateString() + mSpaceTab + pY + mSpaceTab + pM + mSpaceTab + pD + System.Environment.NewLine;
                    }
                    #endregion
                    break;
                case 10:
                    #region 10
                    bool test = false;
                    int count = 0;
                    string[] sqs = sCurrentFilter.Split("#".ToCharArray()[0]);
                    for (int i = 0; i < sqs.Length; i++)
                    {
                        count = 0;
                        if (i == 0)
                        {
                            #region 1
                            // chakra
                            if (sqs[i].Contains("," + txtAstroName.Text.Split(" ".ToCharArray())[1].Replace("(", "").Replace(")", "") + ",") == true)
                            {
                                count++;
                                test = true;
                            }
                            if (sqs[i].Contains("," + Calculator.GetCorrectNumberFromSplitedString(txtNum1.Text.Split(Calculator.Delimiter)) + ",") == true)
                            {
                                count++;
                                test = true;
                            }
                            if (test |= sqs[i].Contains("," + Calculator.GetCorrectNumberFromSplitedString(txtNum2.Text.Split(Calculator.Delimiter)) + ",") == true)
                            {
                                count++;
                                test = true;
                            }
                            if (sqs[i].Contains("," + Calculator.GetCorrectNumberFromSplitedString(txtNum3.Text.Split(Calculator.Delimiter)) + ",") == true)
                            {
                                count++;
                                test = true;
                            }
                            if (sqs[i].Contains("," + Calculator.GetCorrectNumberFromSplitedString(txtNum5.Text.Split(Calculator.Delimiter)) + ",") == true)
                            {
                                count++;
                                test = true;
                            }
                            // life cycles
                            if (sqs[i].Contains("," + Calculator.GetCorrectNumberFromSplitedString(this.Controls.Find("txt" + CalcCurrentCycle().ToString() + "_3", true)[0].Text.Split(Calculator.Delimiter)) + ",") == true)
                            {
                                count++;
                                test = true;
                            }
                            if (sqs[i].Contains("," + Calculator.GetCorrectNumberFromSplitedString(this.Controls.Find("txt" + CalcCurrentCycle().ToString() + "_2", true)[0].Text.Split(Calculator.Delimiter)) + ",") == true)
                            {
                                count++;
                                test = true;
                            }
                            #endregion 1
                        }
                        if (i == 1)
                        {
                            #region 2
                            if (sqs[i].Contains("," + Calculator.GetCorrectNumberFromSplitedString(txtPName_Num.Text.Split(Calculator.Delimiter)) + ",") == true)
                            {
                                count++;
                                test = true;
                            }
                            if (sqs[i].Contains("," + Calculator.GetCorrectNumberFromSplitedString(txtNum7.Text.Split(Calculator.Delimiter)) + ",") == true)
                            {
                                count++;
                                test = true;
                            }
                            #endregion 2
                        }
                        if (test == true)
                        {
                            i = 10;
                        }
                    }

                    testText = true;
                    if (currentQ == 10)
                    {
                        #region 10
                        if (test == true)
                        {
                            if (count < 3)
                            {
                                if (count == 1)
                                {
                                    #region ב
                                    if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                                    {
                                        txtOut.Add("");
                                    }
                                    if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                                    {
                                        txtOut.Add("הוא יכול לאהוב אותך אולם הדבר חופן בתוכו קשיים במערכת הזוגית ועל מנת שתהא מערכת זוגית טובה יותר עליך לעבור את התיקון האישי. מומלץ לבדוק התאמה זוגית.");
                                    }
                                    #endregion ב
                                }
                                else
                                {
                                    if ((Convert.ToInt16(txtAstroName.Text.Split(" ".ToCharArray())[1].Replace("(", "").Replace(")", "")) == 2) ||
                                         (Convert.ToInt16(Calculator.GetCorrectNumberFromSplitedString(txtNum5.Text.Split(Calculator.Delimiter))) == 2))
                                    {
                                        #region ב
                                        if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                                        {
                                            txtOut.Add("");
                                        }
                                        if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                                        {
                                            txtOut.Add("הוא יכול לאהוב אותך אולם הדבר חופן בתוכו קשיים במערכת הזוגית ועל מנת שתהא מערכת זוגית טובה יותר עליך לעבור את התיקון האישי. מומלץ לבדוק התאמה זוגית.");
                                        }
                                        #endregion ב
                                    }
                                    else
                                    {
                                        #region ג
                                        if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                                        {
                                            txtOut.Add("");
                                        }
                                        if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                                        {
                                            txtOut.Add("הוא מתקשה  לאהוב אותך שכן מפתך 'מייצרת' קשיים במערכת הזוגית ועל מנת שתהא מערכת זוגית טובה יותר עליך לעבור את התיקון האישי. מומלץ לבדוק התאמה זוגית.");
                                        }
                                        #endregion ג
                                    }
                                }
                            }
                            else //(count >2)
                            {
                                #region ג
                                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                                {
                                    txtOut.Add("");
                                }
                                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                                {
                                    txtOut.Add("הוא מתקשה  לאהוב אותך שכן מפתך 'מייצרת' קשיים במערכת הזוגית ועל מנת שתהא מערכת זוגית טובה יותר עליך לעבור את התיקון האישי. מומלץ לבדוק התאמה זוגית.");
                                }
                                #endregion
                            }
                        }
                        else // count == 0;
                        {
                            #region Count == 0
                            if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                            {
                                txtOut.Add("Your partner loves you.");
                            }
                            if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                            {
                                txtOut.Add("הוא אוהב אותך.");
                            }
                            #endregion Count == 0
                        }
                        #endregion 8
                    }
                    if (currentQ == 123456)
                    {

                    }
                    #endregion
                    break;
                case 11:
                    #region Text and Dates

                    sqs = sCurrentFilter.Split("#".ToCharArray()[0]);
                    #region Text
                    test = false;
                    count = 0;
                    for (int i = 0; i < 2; i++)
                    {
                        count = 0;
                        if (i == 0)
                        {
                            #region 1
                            // chakra
                            if (sqs[i].Contains("," + txtAstroName.Text.Split(" ".ToCharArray())[1].Replace("(", "").Replace(")", "") + ",") == true)
                            {
                                count++;
                                test = true;
                            }
                            if (sqs[i].Contains("," + Calculator.GetCorrectNumberFromSplitedString(txtNum1.Text.Split(Calculator.Delimiter)) + ",") == true)
                            {
                                count++;
                                test = true;
                            }
                            if (test |= sqs[i].Contains("," + Calculator.GetCorrectNumberFromSplitedString(txtNum2.Text.Split(Calculator.Delimiter)) + ",") == true)
                            {
                                count++;
                                test = true;
                            }
                            if (sqs[i].Contains("," + Calculator.GetCorrectNumberFromSplitedString(txtNum3.Text.Split(Calculator.Delimiter)) + ",") == true)
                            {
                                count++;
                                test = true;
                            }
                            if (sqs[i].Contains("," + Calculator.GetCorrectNumberFromSplitedString(txtNum5.Text.Split(Calculator.Delimiter)) + ",") == true)
                            {
                                count++;
                                test = true;
                            }
                            // life cycles
                            if (sqs[i].Contains("," + Calculator.GetCorrectNumberFromSplitedString(this.Controls.Find("txt" + CalcCurrentCycle().ToString() + "_3", true)[0].Text.Split(Calculator.Delimiter)) + ",") == true)
                            {
                                count++;
                                test = true;
                            }
                            if (sqs[i].Contains("," + Calculator.GetCorrectNumberFromSplitedString(this.Controls.Find("txt" + CalcCurrentCycle().ToString() + "_2", true)[0].Text.Split(Calculator.Delimiter)) + ",") == true)
                            {
                                count++;
                                test = true;
                            }
                            #endregion 1
                        }
                        if (i == 1)
                        {
                            #region 2
                            if (sqs[i].Contains("," + Calculator.GetCorrectNumberFromSplitedString(txtPName_Num.Text.Split(Calculator.Delimiter)) + ",") == true)
                            {
                                count++;
                                test = true;
                            }
                            if (sqs[i].Contains("," + Calculator.GetCorrectNumberFromSplitedString(txtNum7.Text.Split(Calculator.Delimiter)) + ",") == true)
                            {
                                count++;
                                test = true;
                            }
                            #endregion 2
                        }
                        if (test == true)
                        {
                            i = 8;
                        }
                    }

                    testText = true;
                    #endregion Text

                    #region Dates
                    dgvQnA.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    dgvQnA.Columns[1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    dgvQnA.Columns[2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    dgvQnA.Columns[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    dgvQnA.Columns[4].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

                    dgvQnA.Columns[0].Width = 110;
                    dgvQnA.Columns[1].Width = 64;
                    dgvQnA.Columns[2].Width = 52;
                    dgvQnA.Columns[3].Width = 52;
                    dgvQnA.Columns[4].Width = 52;
                    dgvQnA.Columns[5].Visible = false;

                    dgvQnA.Update();

                    for (int i = 0; i < DateList.Count; i++)
                    {
                        DateTime futuredate = DateList[i];//DateTime.Now.AddDays(days * i);

                        string pY, pM, pD, strDate;
                        Calculator.CalcPersonalInfo(mB_Date, futuredate, out pY, out pM, out pD);

                        DataGridViewRow dR = new DataGridViewRow();
                        //dataFuture.Rows.Add(1);

                        strDate = futuredate.Day.ToString() + "." + futuredate.Month.ToString() + "." + futuredate.Year.ToString() + "  (" + FlipString(Calculator.FateCalc(futuredate)) + ")";

                        #region Climax - Future
                        DateTime newBD = futuredate;
                        int[] fLC = Calculator.CareteLifeCycles(newBD);

                        // first climax
                        string ftrClmx = Calculator.CalcSum(fLC[2]);
                        #endregion

                        #region Question Filter
                        bool passedfilter = true;
                        if (Calculator.isCarmaticNumber(futuredate.Day) == true)
                        {
                            passedfilter = false;
                        }
                        else
                        {
                            int tmpPy = Calculator.GetCorrectNumberFromSplitedString(pY.Split(Calculator.Delimiter));
                            int tmpPm = Calculator.GetCorrectNumberFromSplitedString(pM.Split(Calculator.Delimiter));
                            int tmpPd = Calculator.GetCorrectNumberFromSplitedString(pD.Split(Calculator.Delimiter));
                            int tmpFate = Calculator.GetCorrectNumberFromSplitedString(Calculator.FateCalc(futuredate).Split(Calculator.Delimiter));
                            int tmpFtrClmx = Calculator.GetCorrectNumberFromSplitedString(ftrClmx.Split(Calculator.Delimiter));

                            sCurrentFilter = sqs[sqs.Length - 1];

                            if (cbPYtake.Checked == true)
                            {
                                passedfilter &= sCurrentFilter.Contains("," + tmpPy.ToString() + ",");
                            }

                            passedfilter &= sCurrentFilter.Contains("," + tmpPm.ToString() + ",");
                            passedfilter &= sCurrentFilter.Contains("," + tmpPd.ToString() + ",");
                            passedfilter &= sCurrentFilter.Contains("," + tmpFate.ToString() + ",");
                            passedfilter &= sCurrentFilter.Contains("," + tmpFtrClmx.ToString() + ",");
                        }
                        #endregion Question Filter

                        if (passedfilter == true)
                        {
                            object[] objects = new string[5];
                            objects.Initialize();
                            objects.SetValue(strDate, 0);
                            objects.SetValue(FlipString(pY), 1);
                            objects.SetValue(FlipString(pM), 2);
                            objects.SetValue(FlipString(pD), 3);
                            objects.SetValue(FlipString(ftrClmx), 4);

                            if (dataFuture.Rows.Count == 0)
                            {
                                dgvQnA.Rows.Insert(0, objects);
                            }
                            else
                            {
                                dgvQnA.Rows.Insert(dataFuture.Rows.Count - 1, objects);
                            }
                            dgvQnA.Update();
                        }
                        //txtFuture.Text += futuredate.Date.ToShortDateString() + mSpaceTab + pY + mSpaceTab + pM + mSpaceTab + pD + System.Environment.NewLine;
                    }

                    #endregion

                    if (testText == true)
                    {
                        if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                        {
                            txtOut.Add("זוגיות חדשה תתאפשר בעקבות הבנת התיקון האישי ועבודה פנימית ובמועדים המפורטים.");
                        }
                        if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                        {
                            txtOut.Add("");
                        }
                    }
                    else //(testText == false) // count == 0;
                    {
                        if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                        {
                            txtOut.Add("זוגיות חדשה תתאפשר במועדים המפורטים בטבלה.");
                        }
                        if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                        {
                            txtOut.Add("");
                        }
                    }
                    #endregion
                    break;
                case 12:
                    #region 12
                    //mTabs.TabPages[9].Show(); // tabZugiut
                    mTabs.SelectTab("tabZugiut");
                    mTabs.SelectedTab.Show();
                    #endregion
                    break;
                case 13:
                case 14:
                    #region Zugiut
                    isRunningQnA = true;
                    txtOut = new List<string>();

                    if (txtMrrgMark.Text.Trim().Length == 0)
                    {
                        mTabs.SelectTab("tabZugiut");
                        mTabs.SelectedTab.Show();
                    }
                    else
                    {
                        if (isRunningQnA == true)
                        {
                            isRunningQnA = false;

                            double m = Convert.ToDouble(txtMrrgMark.Text.Trim());

                            if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                            {
                                txtOut.Add("תוצאת חישוב הזוגיות הינה: " + m.ToString() + Environment.NewLine);
                            }
                            if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                            {
                                txtOut.Add("Relationship mark is: " + m.ToString() + Environment.NewLine);
                            }

                            testText = true;

                            if (currentQ == 13)
                            {
                                #region 13
                                if (m > 7)
                                {
                                    if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                                    {
                                        txtOut.Add("כן קיימת אפשרות טובה לחתונה ולזוגיות טובה עד לסיום המחזור הנבדק. עם זאת מומלץ לבדוק מה יהיה השינוי לאחר החתונה עם שינוי שם המשפחה של הכלה או החתן או שניהם כך שלא ישפיעו לרע על המערכת הזוגית. ");
                                    }
                                    if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                                    {
                                        txtOut.Add("");
                                    }
                                }
                                if ((m >= 6.5) && (m <= 7))
                                {
                                    if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                                    {
                                        txtOut.Add("כן קיימת אפשרות לחתונה ולזוגיות טובה עד לסיום המחזור הנבדק. עם זאת זוגיות טובה תתאפשר בעקבות הבנת התיקון האישי ועבודה פנימית אצל בני הזוג להם יש תיקון קארמתי. מומלץ לבדוק מה יהיה השינוי לאחר החתונה עם שינוי שם המשפחה של הכלה או החתן או שניהם כך שלא ישפיעו לרע על המערכת הזוגית. ");
                                    }
                                    if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                                    {
                                        txtOut.Add("");
                                    }
                                }
                                if (m < 6.5)
                                {
                                    if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                                    {
                                        txtOut.Add("התאמה זוגית חלשה לא מומלץ. עם זאת זוגיות טובה תתאפשר בעקבות הבנת התיקון האישי ועבודה פנימית אצל בני הזוג להם יש תיקון קארמתי. מומלץ לבדוק מה יהיה השפעת השינוי לאחר החתונה עם שינוי שם המשפחה של הכלה או החתן או שניהם, על המערכת הזוגית.");
                                    }
                                    if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                                    {
                                        txtOut.Add("");
                                    }
                                }
                                #endregion 13
                            }
                            if (currentQ == 14)
                            {
                                #region 14
                                if (m > 7)
                                {
                                    if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                                    {
                                        txtOut.Add("כן קיימת אפשרות טובה לחתונה ולזוגיות טובה עד לסיום המחזור הנבדק. עם זאת מומלץ לבדוק מה יהיה השינוי לאחר החתונה עם שינוי שם המשפחה של הכלה או החתן או שניהם כך שלא ישפיעו לרע על המערכת הזוגית. ");
                                    }
                                    if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                                    {
                                        txtOut.Add("");
                                    }
                                }
                                if ((m >= 6.5) && (m <= 7))
                                {
                                    if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                                    {
                                        txtOut.Add("כן קיימת אפשרות לחתונה. עם זאת זוגיות טובה תתאפשר בעקבות הבנת התיקון האישי ועבודה פנימית אצל בני הזוג להם יש תיקון קארמתי. מומלץ לבדוק מה יהיה השינוי לאחר החתונה עם שינוי שם המשפחה של הכלה או החתן או שניהם כך שלא ישפיעו לרע על המערכת הזוגית.");
                                    }
                                    if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                                    {
                                        txtOut.Add("");
                                    }
                                }
                                if (m < 6.5)
                                {
                                    if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                                    {
                                        txtOut.Add("התאמה זוגית חלשה לא מומלץ. עם זאת זוגיות טובה תתאפשר בעקבות הבנת התיקון האישי ועבודה פנימית אצל בני הזוג להם יש תיקון קארמתי. מומלץ לבדוק מה יהיה השפעת השינוי לאחר החתונה עם שינוי שם המשפחה של הכלה או החתן או שניהם, על המערכת הזוגית.");
                                    }
                                    if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                                    {
                                        txtOut.Add("");
                                    }
                                }
                                #endregion 14
                            }

                        }
                    }
                    #endregion
                    break;
                case 15:
                case 16:
                    #region 15+16
                    UserResult u = new UserResult();

                    test = false;
                    count = 0;
                    sqs = sCurrentFilter.Split("#".ToCharArray()[0]);

                    #region ZUGIUT
                    isRunningQnA = true;
                    if (currentQ == 16)
                    {
                        isRunningQnAspouce = true;
                    }

                    if (txtMrrgMark.Text.Trim().Length == 0)
                    {
                        mTabs.SelectTab("tabZugiut");
                        mTabs.SelectedTab.Show();
                    }
                    else
                    {
                        if (isRunningQnA == true)
                        {
                            isRunningQnA = false;
                            isRunningQnAspouce = false;

                            if (currentQ == 15) // isRunningQnAspouce = false;
                            {
                                u.PrivateNameNum = txtPrivateName.Text;
                                u.Astro = txtAstroName.Text;
                                u.birthday = DateTimePickerFrom.Value;

                                u.Base = txtNum7.Text;
                                u.Crown = txtNum1.Text;
                                u.Heart = txtNum4.Text;
                                u.SexAndCreation = txtNum6.Text;
                                u.Sun = txtNum5.Text;
                                u.ThirdEye = txtNum2.Text;
                                u.Throught = txtNum3.Text;

                                u.LC_cycle = this.Controls.Find("txt" + CalcCurrentCycle().ToString() + "_2", true)[0].Text;
                                u.LC_Climax = this.Controls.Find("txt" + CalcCurrentCycle().ToString() + "_3", true)[0].Text;
                                u.LC_Chalange = this.Controls.Find("txt" + CalcCurrentCycle().ToString() + "_4", true)[0].Text;
                            }
                            if (currentQ == 16) // isRunningQnAspouce = true;
                            {
                                u = new UserResult();
                                u = SpouceResult;
                            }

                            for (int i = 0; i < sqs.Length; i++)
                            {
                                count = 0;
                                if (i == 0)
                                {
                                    #region 1
                                    // chakra
                                    if (sqs[i].Contains("," + u.Astro.Split(" ".ToCharArray())[1].Replace("(", "").Replace(")", "") + ",") == true)
                                    {
                                        count++;
                                        test = true;
                                    }
                                    if (sqs[i].Contains("," + Calculator.GetCorrectNumberFromSplitedString(u.Crown.Split(Calculator.Delimiter)) + ",") == true)
                                    {
                                        count++;
                                        test = true;
                                    }
                                    if (test |= sqs[i].Contains("," + Calculator.GetCorrectNumberFromSplitedString(u.ThirdEye.Split(Calculator.Delimiter)) + ",") == true)
                                    {
                                        count++;
                                        test = true;
                                    }
                                    if (sqs[i].Contains("," + Calculator.GetCorrectNumberFromSplitedString(u.Throught.Split(Calculator.Delimiter)) + ",") == true)
                                    {
                                        count++;
                                        test = true;
                                    }
                                    if (sqs[i].Contains("," + Calculator.GetCorrectNumberFromSplitedString(u.Sun.Split(Calculator.Delimiter)) + ",") == true)
                                    {
                                        count++;
                                        test = true;
                                    }
                                    // life cycles
                                    if (sqs[i].Contains("," + Calculator.GetCorrectNumberFromSplitedString(u.LC_cycle.Split(Calculator.Delimiter)) + ",") == true)
                                    {
                                        count++;
                                        test = true;
                                    }
                                    if (sqs[i].Contains("," + Calculator.GetCorrectNumberFromSplitedString(u.LC_Climax.Split(Calculator.Delimiter)) + ",") == true)
                                    {
                                        count++;
                                        test = true;
                                    }
                                    #endregion 1
                                }
                                if (i == 1)
                                {
                                    #region 2
                                    if (sqs[i].Contains("," + Calculator.GetCorrectNumberFromSplitedString(u.PrivateNameNum.Split(Calculator.Delimiter)) + ",") == true)
                                    {
                                        count++;
                                        test = true;
                                    }
                                    if (sqs[i].Contains("," + Calculator.GetCorrectNumberFromSplitedString(u.Base.Split(Calculator.Delimiter)) + ",") == true)
                                    {
                                        count++;
                                        test = true;
                                    }
                                    #endregion 2
                                }
                                if (test == true)
                                {
                                    i = 8;
                                }
                            }
                        }
                    }
                    #endregion

                    testText = true;

                    if (currentQ == 15)
                    {
                        #region 15
                        if (isRunningQnA == false)
                        {
                            if (test == true)
                            {
                                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                                {
                                    txtOut.Add("סביר להניח שהוא לא יחזור. אם ברצונך להשיבו אליך מומלץ לך להבין את התיקון האישי הקארמתי, לבצעה עבודה פנימית ולהגיע להבנה מעמיקה, טהורה וגבוה.");
                                }
                                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                                {
                                    txtOut.Add("");
                                }
                            }
                            else //(test == false) count == 0
                            {
                                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                                {
                                    txtOut.Add("קיימת אפשרות טובה שהוא יחזור.");
                                }
                                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                                {
                                    txtOut.Add("");
                                }
                            }
                        }
                        #endregion 15
                    }
                    if (currentQ == 16)
                    {
                        #region 16
                        if (isRunningQnA == false)
                        {
                            double mm = Convert.ToDouble(txtMrrgMark.Text);
                            if (mm >= 6.5)
                            {
                                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                                {
                                    txtOut.Add("כן מומלץ. עם זאת מומלץ לכוון את בן הזוג להבנת התיקון האישי שלו ולעבודה פנימית בתחום ההתנהגות הזוגית ובמיוחד לעבוד על התיקון קארמתי.  גם לגבי השואל/ת.");
                                }
                                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                                {
                                    txtOut.Add("");
                                }
                            }
                            else
                            {
                                //if (count > 2) // 3++
                                //{
                                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                                {
                                    txtOut.Add("לא מומלץ. עם זאת מומלץ לכוון את בן הזוג להבנת התיקון האישי שלו ולעבודה פנימית בתחום ההתנהגות הזוגית ובמיוחד לעבוד על התיקון קארמתי. גם לגבי השואל/ת");
                                }
                                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                                {
                                    txtOut.Add("");
                                }

                            }
                        }
                        #endregion 16
                    }
                    #endregion
                    break;
                case 17:
                    #region 17
                    u = new UserResult();

                    sqs = sCurrentFilter.Split("#".ToCharArray()[0]);
                    test = false;
                    count = 0;
                    bool test1 = false;
                    int count1 = 0;

                    #region ZUGIUT
                    isRunningQnA = true;
                    if (currentQ == 17)
                    {
                        isRunningQnAspouce = true;
                    }

                    if (txtMrrgMark.Text.Trim().Length == 0)
                    {
                        mTabs.SelectTab("tabZugiut");
                        mTabs.SelectedTab.Show();
                    }
                    else
                    {
                        if (isRunningQnA == true)
                        {
                            isRunningQnA = false;
                            isRunningQnAspouce = false;

                            u = SpouceResult;

                            for (int i = 0; i < 2; i++)
                            {
                                count = 0;
                                if (i == 0)
                                {
                                    #region 1
                                    // chakra
                                    if (sqs[i].Contains("," + u.Astro.Split(" ".ToCharArray())[1].Replace("(", "").Replace(")", "") + ",") == true)
                                    {
                                        count++;
                                        test = true;
                                    }
                                    if (sqs[i].Contains("," + Calculator.GetCorrectNumberFromSplitedString(u.Crown.Split(Calculator.Delimiter)) + ",") == true)
                                    {
                                        count++;
                                        test = true;
                                    }
                                    if (test |= sqs[i].Contains("," + Calculator.GetCorrectNumberFromSplitedString(u.ThirdEye.Split(Calculator.Delimiter)) + ",") == true)
                                    {
                                        count++;
                                        test = true;
                                    }
                                    if (sqs[i].Contains("," + Calculator.GetCorrectNumberFromSplitedString(u.Throught.Split(Calculator.Delimiter)) + ",") == true)
                                    {
                                        count++;
                                        test = true;
                                    }
                                    if (sqs[i].Contains("," + Calculator.GetCorrectNumberFromSplitedString(u.Sun.Split(Calculator.Delimiter)) + ",") == true)
                                    {
                                        count++;
                                        test = true;
                                    }
                                    // life cycles
                                    if (sqs[i].Contains("," + Calculator.GetCorrectNumberFromSplitedString(u.LC_cycle.Split(Calculator.Delimiter)) + ",") == true)
                                    {
                                        count++;
                                        test = true;
                                    }
                                    if (sqs[i].Contains("," + Calculator.GetCorrectNumberFromSplitedString(u.LC_Climax.Split(Calculator.Delimiter)) + ",") == true)
                                    {
                                        count++;
                                        test = true;
                                    }
                                    #endregion 1
                                }
                                if (i == 1)
                                {
                                    #region 2
                                    if (sqs[i].Contains("," + Calculator.GetCorrectNumberFromSplitedString(u.PrivateNameNum.Split(Calculator.Delimiter)) + ",") == true)
                                    {
                                        count++;
                                        test = true;
                                    }
                                    if (sqs[i].Contains("," + Calculator.GetCorrectNumberFromSplitedString(u.Base.Split(Calculator.Delimiter)) + ",") == true)
                                    {
                                        count++;
                                        test = true;
                                    }
                                    #endregion 2
                                }
                                if (test == true)
                                {
                                    i = 8;
                                }
                            }

                            u = new UserResult();
                            u.GatherDataFromGUI(this.Controls, CalcCurrentCycle());

                            for (int i = 2; i < sqs.Length; i++)
                            {
                                count1 = 0;
                                if (i == 0)
                                {
                                    #region 1
                                    // chakra
                                    if (sqs[i].Contains("," + u.Astro.Split(" ".ToCharArray())[1].Replace("(", "").Replace(")", "") + ",") == true)
                                    {
                                        count1++;
                                        test1 = true;
                                    }
                                    if (sqs[i].Contains("," + Calculator.GetCorrectNumberFromSplitedString(u.Crown.Split(Calculator.Delimiter)) + ",") == true)
                                    {
                                        count1++;
                                        test = true;
                                    }
                                    if (test1 |= sqs[i].Contains("," + Calculator.GetCorrectNumberFromSplitedString(u.ThirdEye.Split(Calculator.Delimiter)) + ",") == true)
                                    {
                                        count++;
                                        test = true;
                                    }
                                    if (sqs[i].Contains("," + Calculator.GetCorrectNumberFromSplitedString(u.Throught.Split(Calculator.Delimiter)) + ",") == true)
                                    {
                                        count1++;
                                        test1 = true;
                                    }
                                    if (sqs[i].Contains("," + Calculator.GetCorrectNumberFromSplitedString(u.Sun.Split(Calculator.Delimiter)) + ",") == true)
                                    {
                                        count1++;
                                        test1 = true;
                                    }
                                    // life cycles
                                    if (sqs[i].Contains("," + Calculator.GetCorrectNumberFromSplitedString(u.LC_cycle.Split(Calculator.Delimiter)) + ",") == true)
                                    {
                                        count1++;
                                        test1 = true;
                                    }
                                    if (sqs[i].Contains("," + Calculator.GetCorrectNumberFromSplitedString(u.LC_Climax.Split(Calculator.Delimiter)) + ",") == true)
                                    {
                                        count1++;
                                        test1 = true;
                                    }
                                    #endregion 1
                                }
                                if (i == 1)
                                {
                                    #region 2
                                    if (sqs[i].Contains("," + Calculator.GetCorrectNumberFromSplitedString(u.PrivateNameNum.Split(Calculator.Delimiter)) + ",") == true)
                                    {
                                        count1++;
                                        test1 = true;
                                    }
                                    if (sqs[i].Contains("," + Calculator.GetCorrectNumberFromSplitedString(u.Base.Split(Calculator.Delimiter)) + ",") == true)
                                    {
                                        count1++;
                                        test1 = true;
                                    }
                                    #endregion 2
                                }
                                if (test1 == true)
                                {
                                    i = 8;
                                }
                            }
                        }
                    }
                    #endregion

                    testText = true;

                    #region spouce
                    if ((test == true) && (count > 1))
                    {
                        if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                        {
                            txtOut.Add("מפתו של בן הזוג מתארת  שיש באפשרותו לבגוד.");
                        }
                        if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                        {
                            txtOut.Add("");
                        }
                    }
                    if ((test == true) && (count < 2))
                    {
                        if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                        {
                            txtOut.Add("אין לבן הזוג נטייה לבגוד.");
                        }
                        if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                        {
                            txtOut.Add("");
                        }
                    }
                    if (test == false)
                    {
                        if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                        {
                            txtOut.Add("אין לבן הזוג נטייה לבגוד.");
                        }
                        if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                        {
                            txtOut.Add("");
                        }
                    }
                    #endregion spouce

                    #region current user
                    if (test1 == true)
                    {
                        if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                        {
                            txtOut.Add("מפתך מתארת אדם המקשה על מערכת זוגית במיוחד מהפן המיני באופן שדוחק את בן הזוג לבגוד.");
                        }
                        if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                        {
                            txtOut.Add("");
                        }
                    }
                    else
                    {
                        if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                        {
                            txtOut.Add("");
                        }
                        if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                        {
                            txtOut.Add("");
                        }
                    }
                    #endregion current user

                    #endregion
                    break;
                case 18:
                    #region 18
                    u = new UserResult();

                    sqs = sCurrentFilter.Split("#".ToCharArray()[0]);
                    test = false;
                    count = 0;

                    u.GatherDataFromGUI(this.Controls, CalcCurrentCycle());

                    for (int i = 0; i < sqs.Length; i++)
                    {
                        count = 0;
                        if (i == 0)
                        {
                            #region 1
                            // chakra
                            if (sqs[i].Contains("," + u.Astro.Split(" ".ToCharArray())[1].Replace("(", "").Replace(")", "") + ",") == true)
                            {
                                count++;
                                test = true;
                            }
                            if (sqs[i].Contains("," + Calculator.GetCorrectNumberFromSplitedString(u.Crown.Split(Calculator.Delimiter)) + ",") == true)
                            {
                                count++;
                                test = true;
                            }
                            if (test |= sqs[i].Contains("," + Calculator.GetCorrectNumberFromSplitedString(u.ThirdEye.Split(Calculator.Delimiter)) + ",") == true)
                            {
                                count++;
                                test = true;
                            }
                            if (sqs[i].Contains("," + Calculator.GetCorrectNumberFromSplitedString(u.Throught.Split(Calculator.Delimiter)) + ",") == true)
                            {
                                count++;
                                test = true;
                            }
                            if (sqs[i].Contains("," + Calculator.GetCorrectNumberFromSplitedString(u.Sun.Split(Calculator.Delimiter)) + ",") == true)
                            {
                                count++;
                                test = true;
                            }
                            // life cycles
                            if (sqs[i].Contains("," + Calculator.GetCorrectNumberFromSplitedString(u.LC_cycle.Split(Calculator.Delimiter)) + ",") == true)
                            {
                                count++;
                                test = true;
                            }
                            if (sqs[i].Contains("," + Calculator.GetCorrectNumberFromSplitedString(u.LC_Climax.Split(Calculator.Delimiter)) + ",") == true)
                            {
                                count++;
                                test = true;
                            }
                            #endregion 1
                        }
                        if (i == 1)
                        {
                            #region 2
                            if (sqs[i].Contains("," + Calculator.GetCorrectNumberFromSplitedString(u.PrivateNameNum.Split(Calculator.Delimiter)) + ",") == true)
                            {
                                count++;
                                test = true;
                            }
                            if (sqs[i].Contains("," + Calculator.GetCorrectNumberFromSplitedString(u.Base.Split(Calculator.Delimiter)) + ",") == true)
                            {
                                count++;
                                test = true;
                            }
                            #endregion 2
                        }
                        if (test == true)
                        {
                            i = 8;
                        }
                    }

                    testText = true;
                    if (test == false)
                    {
                        if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                        {
                            txtOut.Add("מפתך מתארת אדם שאין לו בעיה אישית למשוך זוגיות. עם זאת הבעיה היא שטרם מצאת את האדם אליו את כמהה/משתוקקת. מומלץ לפעול בדרכים המקובלות על מנת להכיר בני זוג פוטנציאליים.");
                        }
                        if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                        {
                            txtOut.Add("");
                        }
                    }
                    else
                    {
                        if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                        {
                            txtOut.Add("מפתך מתארת אדם קשה ונוקשה, בעל דרישות לא שגרתיות ממערכת זוגית. אדם הנוטה להכתיב לאחר את דרך חייו ורצונותיו. אדם השוכח כי מערכת זוגית מושתתת על שיתופי פעולה ויחס הדדי בפן החיובי ובכל הקשור בניהול מערכת זוגית יציבה.");
                        }
                        if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                        {
                            txtOut.Add("");
                        }
                    }

                    if ((new List<string> { "16", "7" }.Contains(Calculator.GetCorrectNumberFromSplitedString(u.LC_cycle.Split(Calculator.Delimiter)).ToString())) ||
                        (new List<string> { "16", "7" }.Contains(Calculator.GetCorrectNumberFromSplitedString(u.LC_Climax.Split(Calculator.Delimiter)).ToString())))
                    {
                        if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                        {
                            txtOut.Add("הינך נמצאת בתקופה המעכבת ומקשה על הזוגית. ");
                        }
                        if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                        {
                            txtOut.Add("");
                        }
                    }

                    if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                    {
                        txtOut.Add("עם זאת זוגיות טובה תתאפשר בעקבות הבנת התיקון הקארמתי האישי ועבודה פנימית מעמיקה.");
                    }
                    if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                    {
                        txtOut.Add("");
                    }
                    #endregion
                    break;
                case 19:
                    #region 19
                    u = new UserResult();
                    u.GatherDataFromGUI(this.Controls, CalcCurrentCycle());

                    sqs = sCurrentFilter.Split("#".ToCharArray()[0]);
                    test = false;
                    count = 0;

                    for (int i = 0; i < 2; i++)
                    {
                        count = 0;
                        if (i == 0)
                        {
                            #region 1
                            // chakra
                            if (sqs[i].Contains("," + u.Astro.Split(" ".ToCharArray())[1].Replace("(", "").Replace(")", "") + ",") == true)
                            {
                                count++;
                                test = true;
                            }
                            if (sqs[i].Contains("," + Calculator.GetCorrectNumberFromSplitedString(u.Crown.Split(Calculator.Delimiter)) + ",") == true)
                            {
                                count++;
                                test = true;
                            }
                            if (test |= sqs[i].Contains("," + Calculator.GetCorrectNumberFromSplitedString(u.ThirdEye.Split(Calculator.Delimiter)) + ",") == true)
                            {
                                count++;
                                test = true;
                            }
                            if (sqs[i].Contains("," + Calculator.GetCorrectNumberFromSplitedString(u.Throught.Split(Calculator.Delimiter)) + ",") == true)
                            {
                                count++;
                                test = true;
                            }
                            if (sqs[i].Contains("," + Calculator.GetCorrectNumberFromSplitedString(u.Sun.Split(Calculator.Delimiter)) + ",") == true)
                            {
                                count++;
                                test = true;
                            }
                            // life cycles
                            if (sqs[i].Contains("," + Calculator.GetCorrectNumberFromSplitedString(u.LC_cycle.Split(Calculator.Delimiter)) + ",") == true)
                            {
                                count++;
                                test = true;
                            }
                            if (sqs[i].Contains("," + Calculator.GetCorrectNumberFromSplitedString(u.LC_Climax.Split(Calculator.Delimiter)) + ",") == true)
                            {
                                count++;
                                test = true;
                            }
                            #endregion 1
                        }
                        if (i == 1)
                        {
                            #region 2
                            if (sqs[i].Contains("," + Calculator.GetCorrectNumberFromSplitedString(u.PrivateNameNum.Split(Calculator.Delimiter)) + ",") == true)
                            {
                                count++;
                                test = true;
                            }
                            if (sqs[i].Contains("," + Calculator.GetCorrectNumberFromSplitedString(u.Base.Split(Calculator.Delimiter)) + ",") == true)
                            {
                                count++;
                                test = true;
                            }
                            #endregion 2
                        }
                        if (test == true)
                        {
                            i = 8;
                        }
                    }

                    testText = true;
                    if (u.isMapStrong == false) // מפה חלשה
                    {
                        if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                        {
                            txtOut.Add("מפתך חלשה מומלץ לחזק את המפה. ");
                        }
                        if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                        {
                            txtOut.Add("");
                        }
                    }
                    if (u.mrkBssnss < 8)
                    {
                        if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                        {
                            txtOut.Add("מפתך בעלת סיכוי קטן להצלחה כלכלית ובעל יכולת חלשה להרוויח סכומי כסף גדולים ולנהל אותם באופן נכון וכזה שיוביל אותך לעצמאות כלכלית. מומלץ לחזק את מפתך.");
                        }
                        if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                        {
                            txtOut.Add("");
                        }
                    }
                    if ((test == true) && (count > 2) && ((Calculator.GetCorrectNumberFromSplitedString(u.Throught.Split(Calculator.Delimiter)) != 8) || (Calculator.GetCorrectNumberFromSplitedString(u.SexAndCreation.Split(Calculator.Delimiter)) != 8)))
                    {
                        if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                        {
                            txtOut.Add("מפתך מתארת אדם הנוטה להיות בזבזן.");
                        }
                        if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                        {
                            txtOut.Add("");
                        }
                    }

                    test = false;
                    for (int i = 0; i < u.SexAndCreation.Split(Calculator.Delimiter).Length; i++)
                    {
                        if (sqs[2].Contains("," + u.SexAndCreation.Split(Calculator.Delimiter)[i] + ",") || (sqs[3].Contains("," + u.SexAndCreation.Split(Calculator.Delimiter)[i] + ",") && (u.isMapStrong == false)))
                        {
                            test = true;
                        }
                    }
                    if (test == true)
                    {
                        if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                        {
                            txtOut.Add("כישורי התעסוקה שלך מתקשים לממש משכורות או עסקים בעלי הכנסה גדולה המובילים לרווחה כלכלית. מומלץ לשפר את כישורי התעסוקה.");
                        }
                        if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                        {
                            txtOut.Add("");
                        }
                    }
                    if ((test == true) && (u.mrkBssnss >= 8))
                    {

                        if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                        {
                            txtOut.Add("מומלץ לחזק את כישורי התעסוקה.");
                        }
                        if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                        {
                            txtOut.Add("");
                        }
                    }

                    #endregion
                    break;
                case 20:
                    #region 20
                    u = new UserResult();
                    u.GatherDataFromGUI(this.Controls, CalcCurrentCycle());

                    sqs = sCurrentFilter.Split("#".ToCharArray()[0]);
                    test = false;
                    count = 0;

                    #region Dates
                    dgvQnA.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    dgvQnA.Columns[1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    dgvQnA.Columns[2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    dgvQnA.Columns[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    dgvQnA.Columns[4].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

                    dgvQnA.Columns[0].Width = 110;
                    dgvQnA.Columns[1].Width = 64;
                    dgvQnA.Columns[2].Width = 52;
                    dgvQnA.Columns[3].Width = 52;
                    dgvQnA.Columns[4].Width = 52;
                    dgvQnA.Columns[5].Visible = false;

                    dgvQnA.Update();

                    for (int i = 0; i < DateList.Count; i++)
                    {
                        DateTime futuredate = DateList[i];//DateTime.Now.AddDays(days * i);

                        string pY, pM, pD, strDate;
                        Calculator.CalcPersonalInfo(mB_Date, futuredate, out pY, out pM, out pD);

                        DataGridViewRow dR = new DataGridViewRow();
                        //dataFuture.Rows.Add(1);

                        strDate = futuredate.Day.ToString() + "." + futuredate.Month.ToString() + "." + futuredate.Year.ToString() + "  (" + FlipString(Calculator.FateCalc(futuredate)) + ")";

                        #region Climax - Future
                        DateTime newBD = futuredate;
                        int[] fLC = Calculator.CareteLifeCycles(newBD);

                        // first climax
                        string ftrClmx = Calculator.CalcSum(fLC[2]);
                        #endregion

                        #region Question Filter
                        bool passedfilter = true;
                        if (Calculator.isCarmaticNumber(futuredate.Day) == true)
                        {
                            passedfilter = false;
                        }
                        else
                        {
                            int tmpPy = Calculator.GetCorrectNumberFromSplitedString(pY.Split(Calculator.Delimiter));
                            int tmpPm = Calculator.GetCorrectNumberFromSplitedString(pM.Split(Calculator.Delimiter));
                            int tmpPd = Calculator.GetCorrectNumberFromSplitedString(pD.Split(Calculator.Delimiter));
                            int tmpFate = Calculator.GetCorrectNumberFromSplitedString(Calculator.FateCalc(futuredate).Split(Calculator.Delimiter));
                            int tmpFtrClmx = Calculator.GetCorrectNumberFromSplitedString(ftrClmx.Split(Calculator.Delimiter));

                            sCurrentFilter = sqs[0];

                            if (cbPYtake.Checked == true)
                            {
                                passedfilter &= sCurrentFilter.Contains("," + tmpPy.ToString() + ",");
                            }

                            passedfilter &= sCurrentFilter.Contains("," + tmpPm.ToString() + ",");
                            passedfilter &= sCurrentFilter.Contains("," + tmpPd.ToString() + ",");
                            passedfilter &= sCurrentFilter.Contains("," + tmpFate.ToString() + ",");
                            passedfilter &= sCurrentFilter.Contains("," + tmpFtrClmx.ToString() + ",");
                        }
                        #endregion Question Filter

                        if (passedfilter == true)
                        {
                            object[] objects = new string[5];
                            objects.Initialize();
                            objects.SetValue(strDate, 0);
                            objects.SetValue(FlipString(pY), 1);
                            objects.SetValue(FlipString(pM), 2);
                            objects.SetValue(FlipString(pD), 3);
                            objects.SetValue(FlipString(ftrClmx), 4);

                            if (dataFuture.Rows.Count == 0)
                            {
                                dgvQnA.Rows.Insert(0, objects);
                            }
                            else
                            {
                                dgvQnA.Rows.Insert(dataFuture.Rows.Count - 1, objects);
                            }
                            dgvQnA.Update();
                        }
                        //txtFuture.Text += futuredate.Date.ToShortDateString() + mSpaceTab + pY + mSpaceTab + pM + mSpaceTab + pD + System.Environment.NewLine;
                    }
                    #endregion

                    for (int i = 1; i < 3; i++)
                    {
                        count = 0;
                        if (i == 0)
                        {
                            #region 1
                            // chakra
                            if (sqs[i].Contains("," + u.Astro.Split(" ".ToCharArray())[1].Replace("(", "").Replace(")", "") + ",") == true)
                            {
                                count++;
                                test = true;
                            }
                            if (sqs[i].Contains("," + Calculator.GetCorrectNumberFromSplitedString(u.Crown.Split(Calculator.Delimiter)) + ",") == true)
                            {
                                count++;
                                test = true;
                            }
                            if (test |= sqs[i].Contains("," + Calculator.GetCorrectNumberFromSplitedString(u.ThirdEye.Split(Calculator.Delimiter)) + ",") == true)
                            {
                                count++;
                                test = true;
                            }
                            if (sqs[i].Contains("," + Calculator.GetCorrectNumberFromSplitedString(u.Throught.Split(Calculator.Delimiter)) + ",") == true)
                            {
                                count++;
                                test = true;
                            }
                            if (sqs[i].Contains("," + Calculator.GetCorrectNumberFromSplitedString(u.Sun.Split(Calculator.Delimiter)) + ",") == true)
                            {
                                count++;
                                test = true;
                            }
                            // life cycles
                            if (sqs[i].Contains("," + Calculator.GetCorrectNumberFromSplitedString(u.LC_cycle.Split(Calculator.Delimiter)) + ",") == true)
                            {
                                count++;
                                test = true;
                            }
                            if (sqs[i].Contains("," + Calculator.GetCorrectNumberFromSplitedString(u.LC_Climax.Split(Calculator.Delimiter)) + ",") == true)
                            {
                                count++;
                                test = true;
                            }
                            #endregion 1
                        }
                        if (i == 1)
                        {
                            #region 2
                            if (sqs[i].Contains("," + Calculator.GetCorrectNumberFromSplitedString(u.PrivateNameNum.Split(Calculator.Delimiter)) + ",") == true)
                            {
                                count++;
                                test = true;
                            }
                            if (sqs[i].Contains("," + Calculator.GetCorrectNumberFromSplitedString(u.Base.Split(Calculator.Delimiter)) + ",") == true)
                            {
                                count++;
                                test = true;
                            }
                            #endregion 2
                        }
                        if (test == true)
                        {
                            i = 800;
                        }
                    }

                    testText = true;
                    if (test == true)
                    {
                        if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                        {
                            txtOut.Add("מפתך מתארת אדם הנוטה לייצר בעיות במסגרת העבודה. אדם קשה ונוקשה, בעל דרישות מוגזמות ולא שגרתיות בעבודה. אדם ביקורתי, שאינו מתחשב בזולת ומתקשה להשתלב במערכת יחסי עבודה. מומלץ לשנות גישה. מומלץ לאזן מפתך על מנת להתגבר על הקשיים.");
                        }
                        if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                        {
                            txtOut.Add("");
                        }
                    }

                    test = false;
                    for (int i = 0; i < u.SexAndCreation.Split(Calculator.Delimiter).Length; i++)
                    {
                        if (sqs[3].Contains("," + u.SexAndCreation.Split(Calculator.Delimiter)[i] + ",") || (sqs[4].Contains("," + u.SexAndCreation.Split(Calculator.Delimiter)[i] + ",") && (u.isMapStrong == false)))
                        {
                            test = true;
                        }
                    }
                    if (test == true)
                    {
                        if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                        {
                            txtOut.Add("כישורי התעסוקה שלך מתקשים על מציאת עבודה ומקשים לממש משכורות בעלת הכנסה גדולה המובילה לרווחה כלכלית. מומלץ לשפר את כישורי התעסוקה.");
                        }
                        if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                        {
                            txtOut.Add("");
                        }
                    }

                    #endregion
                    break;
                //case 19:          // TBD  stuff....
                //    MessageBox.Show("לא פעיל כרגע");
                //    break;
                //    #region 19

                //    if (isRunningQnA == false)
                //    {
                //        isRunningQnA = true;

                //        u = new UserResult();
                //        u.GatherDataFromGUI(this.Controls, CalcCurrentCycle());

                //        sqs = sCurrentFilter.Split("#".ToCharArray()[0]);
                //        test = false;
                //        count = 0;

                //        #region Dates
                //        for (int i = 0; i < DateList.Count; i++)
                //        {
                //            DateTime futuredate = DateList[i];//DateTime.Now.AddDays(days * i);

                //            string pY, pM, pD, strDate;
                //            Calculator.CalcPersonalInfo(mB_Date, futuredate, out pY, out pM, out pD);

                //            DataGridViewRow dR = new DataGridViewRow();
                //            //dataFuture.Rows.Add(1);

                //            strDate = futuredate.Day.ToString() + "." + futuredate.Month.ToString() + "." + futuredate.Year.ToString() + "  (" + FlipString(Calculator.FateCalc(futuredate)) + ")";

                //            #region Climax - Future
                //            DateTime newBD = futuredate;
                //            int[] fLC = Calculator.CareteLifeCycles(newBD);

                //            // first climax
                //            string ftrClmx = Calculator.CalcSum(fLC[2]);
                //            #endregion

                //            #region Question Filter
                //            bool passedfilter = true;
                //            if (Calculator.isCarmaticNumber(futuredate.Day) == true)
                //            {
                //                passedfilter = false;
                //            }
                //            else
                //            {
                //                int tmpPy = Calculator.GetCorrectNumberFromSplitedString(pY.Split(Calculator.Delimiter));
                //                int tmpPm = Calculator.GetCorrectNumberFromSplitedString(pM.Split(Calculator.Delimiter));
                //                int tmpPd = Calculator.GetCorrectNumberFromSplitedString(pD.Split(Calculator.Delimiter));
                //                int tmpFate = Calculator.GetCorrectNumberFromSplitedString(Calculator.FateCalc(futuredate).Split(Calculator.Delimiter));
                //                int tmpFtrClmx = Calculator.GetCorrectNumberFromSplitedString(ftrClmx.Split(Calculator.Delimiter));

                //                sCurrentFilter = sqs[0];

                //                if (cbPYtake.Checked == true)
                //                {
                //                    passedfilter &= sCurrentFilter.Contains("," + tmpPy.ToString() + ",");
                //                }

                //                passedfilter &= sCurrentFilter.Contains("," + tmpPm.ToString() + ",");
                //                passedfilter &= sCurrentFilter.Contains("," + tmpPd.ToString() + ",");
                //                passedfilter &= sCurrentFilter.Contains("," + tmpFate.ToString() + ",");
                //                passedfilter &= sCurrentFilter.Contains("," + tmpFtrClmx.ToString() + ",");
                //            }
                //            #endregion Question Filter

                //            if (passedfilter == true)
                //            {
                //                object[] objects = new string[5];
                //                objects.Initialize();
                //                objects.SetValue(strDate, 0);
                //                objects.SetValue(FlipString(pY), 1);
                //                objects.SetValue(FlipString(pM), 2);
                //                objects.SetValue(FlipString(pD), 3);
                //                objects.SetValue(FlipString(ftrClmx), 4);

                //                if (dataFuture.Rows.Count == 0)
                //                {
                //                    dgvQnA.Rows.Insert(0, objects);
                //                }
                //                else
                //                {
                //                    dgvQnA.Rows.Insert(dataFuture.Rows.Count - 1, objects);
                //                }
                //                dgvQnA.Update();
                //            }
                //            //txtFuture.Text += futuredate.Date.ToShortDateString() + mSpaceTab + pY + mSpaceTab + pM + mSpaceTab + pD + System.Environment.NewLine;
                //        }
                //        dgvQnA.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                //        dgvQnA.Columns[1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                //        dgvQnA.Columns[2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                //        dgvQnA.Columns[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                //        dgvQnA.Columns[4].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

                //        dgvQnA.Columns[0].Width = 115;
                //        dgvQnA.Columns[1].Width = 50;
                //        dgvQnA.Columns[2].Width = 50;
                //        dgvQnA.Columns[3].Width = 50;
                //        dgvQnA.Columns[4].Width = 50;

                //        //dgvQnA.Sort(dgvQnA.Columns[0], ListSortDirection.Ascending);
                //        dgvQnA.Update();
                //        #endregion

                //        for (int i = 1; i < 3; i++)
                //        {
                //            count = 0;
                //            if (i == 0)
                //            {
                //                #region 1
                //                // chakra
                //                if (sqs[i].Contains("," + u.Astro.Split(" ".ToCharArray())[1].Replace("(", "").Replace(")", "") + ",") == true)
                //                {
                //                    count++;
                //                    test = true;
                //                }
                //                if (sqs[i].Contains("," + Calculator.GetCorrectNumberFromSplitedString(u.Crown.Split(Calculator.Delimiter)) + ",") == true)
                //                {
                //                    count++;
                //                    test = true;
                //                }
                //                if (test |= sqs[i].Contains("," + Calculator.GetCorrectNumberFromSplitedString(u.ThirdEye.Split(Calculator.Delimiter)) + ",") == true)
                //                {
                //                    count++;
                //                    test = true;
                //                }
                //                if (sqs[i].Contains("," + Calculator.GetCorrectNumberFromSplitedString(u.Throught.Split(Calculator.Delimiter)) + ",") == true)
                //                {
                //                    count++;
                //                    test = true;
                //                }
                //                if (sqs[i].Contains("," + Calculator.GetCorrectNumberFromSplitedString(u.Sun.Split(Calculator.Delimiter)) + ",") == true)
                //                {
                //                    count++;
                //                    test = true;
                //                }
                //                // life cycles
                //                if (sqs[i].Contains("," + Calculator.GetCorrectNumberFromSplitedString(u.LC_cycle.Split(Calculator.Delimiter)) + ",") == true)
                //                {
                //                    count++;
                //                    test = true;
                //                }
                //                if (sqs[i].Contains("," + Calculator.GetCorrectNumberFromSplitedString(u.LC_Climax.Split(Calculator.Delimiter)) + ",") == true)
                //                {
                //                    count++;
                //                    test = true;
                //                }
                //                #endregion 1
                //            }
                //            if (i == 1)
                //            {
                //                #region 2
                //                if (sqs[i].Contains("," + Calculator.GetCorrectNumberFromSplitedString(u.PrivateNameNum.Split(Calculator.Delimiter)) + ",") == true)
                //                {
                //                    count++;
                //                    test = true;
                //                }
                //                if (sqs[i].Contains("," + Calculator.GetCorrectNumberFromSplitedString(u.Base.Split(Calculator.Delimiter)) + ",") == true)
                //                {
                //                    count++;
                //                    test = true;
                //                }
                //                #endregion 2
                //            }
                //            if (test == true)
                //            {
                //                i = 800;
                //            }
                //        }

                //        testText = true;

                //        #region Test by MessageBox Answer
                //        string strStr = "", strCap = "";
                //        if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                //        {
                //            strStr = "האם את/ה מרוצה בעבודה";
                //            strCap = "שאלות ותשובות";
                //        }
                //        if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                //        {
                //            strStr = "";
                //            strCap = "";
                //        }
                //        DialogResult resDeg = MessageBox.Show(strStr, strCap, MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                //        strStr = "";
                //        switch (resDeg)
                //        {
                //            case DialogResult.Abort:
                //            case DialogResult.Cancel:
                //            case DialogResult.Ignore:
                //            case DialogResult.None:
                //                return;
                //                break;
                //            case System.Windows.Forms.DialogResult.Yes:
                //                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                //                {
                //                    strStr = "מדוע את/ה רוצה להחליף עבודה?";
                //                    strCap = "שאלות ותשובות";
                //                }
                //                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                //                {
                //                    strStr = "";
                //                    strCap = "";
                //                }
                //                MessageBox.Show(strStr, strCap, MessageBoxButtons.OK, MessageBoxIcon.Question);
                //                txtQnAres.Text = "   ";
                //                break;
                //            case System.Windows.Forms.DialogResult.No:
                //                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                //                {
                //                    txtOut.Add("מומלץ עבורך להחליף עבודה בתאריכים הנקובים.");
                //                }
                //                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                //                {
                //                    txtOut.Add("");
                //                }
                //                break;
                //        }
                //        #endregion Test by MessageBox Answer

                //        #region 2
                //        if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                //        {
                //            strStr = "האם את/ה מרגיש/ה שרוצים לפטר אותך?";
                //            strCap = "שאלות ותשובות";
                //        }
                //        if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                //        {
                //            strStr = "";
                //            strCap = "";
                //        }
                //        resDeg = MessageBox.Show(strStr, strCap, MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                //        strStr = "";
                //        switch (resDeg)
                //        {
                //            case DialogResult.Abort:
                //            case DialogResult.Cancel:
                //            case DialogResult.Ignore:
                //            case DialogResult.None:
                //                return;
                //                break;
                //            case System.Windows.Forms.DialogResult.Yes:
                //                if (test == true)
                //                {
                //                    if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                //                    {
                //                        txtOut.Add("מפתך מתארת אדם הנוטה לייצר בעיות במסגרת העבודה. אדם קשה ונוקשה, בעל דרישות מוגזמות ולא שגרתיות בעבודה. אדם ביקורתי, שאינו מתחשב בזולת ומתקשה להשתלב במערכת יחסי עבודה. מומלץ לשנות גישה. מומלץ לאזן מפתך על מנת להתגבר על הקשיים.");
                //                        txtOut.Add(Environment.NewLine);
                //                        txtOut.Add("בדוק האם כישורי התעסוקה העיקריים מתאימים לעבודה החדשה.");
                //                        txtOut.Add(Environment.NewLine);
                //                        txtOut.Add(ReportDataProvider.Instance.ReadFromTextFile(ReportDataProvider.Instance.ConstructFilePath2WorkInfo(EnumProvider.ReservedWork._תעסוקה_ראשי_.ToString())));
                //                        isRunningQnA = false;
                //                    }
                //                    if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                //                    {
                //                        txtOut.Add("");
                //                        txtOut.Add(ReportDataProvider.Instance.ReadFromTextFile(ReportDataProvider.Instance.ConstructFilePath2WorkInfo(EnumProvider.ReservedWork._תעסוקה_ראשי_.ToString())));
                //                    }
                //                }
                //                break;
                //            case System.Windows.Forms.DialogResult.No:
                //                isRunningQnA = true;
                //                Omega.Objects.qnaPickWorkType frm = new Omega.Objects.qnaPickWorkType(this);
                //                frm.Show();
                //                break;
                //        }

                //        isRunningQnA = true;
                //        Omega.Objects.qnaPickWorkType frm1 = new Omega.Objects.qnaPickWorkType(this);
                //        frm1.Show();

                //        #endregion 2
                //    }
                //    else
                //    {
                //        isRunningQnA = false;
                //        testText = true;
                //        string strOptionalWorkTypes = ReportDataProvider.Instance.ReadFromTextFile(ReportDataProvider.Instance.ConstructFilePath2WorkInfo(EnumProvider.ReservedWork._תעסוקה_ראשי_.ToString()));

                //        if (sQnAresWorkType == null) return;

                //        if (strOptionalWorkTypes.Contains(sQnAresWorkType))
                //        {
                //            if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                //            {
                //                txtOut.Add("בחרת מקצוע המתאים לכישוריך - " + sQnAresWorkType);
                //            }
                //            if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                //            {
                //                txtOut.Add("");
                //            }
                //        }
                //        else
                //        {
                //            strOptionalWorkTypes = ReportDataProvider.Instance.ReadFromTextFile(ReportDataProvider.Instance.ConstructFilePath2WorkInfo(EnumProvider.ReservedWork._תעסוקה_משני_.ToString()));
                //            strOptionalWorkTypes += ReportDataProvider.Instance.ReadFromTextFile(ReportDataProvider.Instance.ConstructFilePath2WorkInfo(EnumProvider.ReservedWork._תעסוקה_שם_פרטי_.ToString()));
                //            if (strOptionalWorkTypes.Contains(sQnAresWorkType))
                //            {
                //                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                //                {
                //                    txtOut.Add("הינך מנצל רק חלק מכישורי התעסוקה שלך.");
                //                }
                //                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                //                {
                //                    txtOut.Add("");
                //                }
                //            }
                //            else
                //            {
                //                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                //                {
                //                    txtOut.Add("המקצוע אינו מתאים עבורך - לפי בדיקת צ'אקרת הגרון - השאיפות שלך מתאימות למקצוע" + " " + sQnAresWorkType + " "+ "אולם, כישורי התעסוקה שלך אינם מתאימים במלואם.");
                //                    //int v = Calculator.GetCorrectNumberFromSplitedString(u.Throught.Split(Calculator.Delimiter));
                //                    //txtOut.Add(ReportDataProvider.Instance.ReadFromTextFile(ReportDataProvider.Instance.ConstructFilePath2WorkInfo(EnumProvider.ReservedWork._תעסוקה_ראשי_.ToString(),v)));
                //                }
                //                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                //                {
                //                    txtOut.Add("");
                //                }
                //            }

                //        }


                //        // end of...
                //        if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                //        {
                //            txtOut.Add("כידוע לך מעבר עבודה כרוך בתחומים רבים אותם יש לקחת בחשבון כגון: קירבה למקום העבודה החדש, שכר, התנהגות המעביד, הענף בו מתקיים העסק האם יש בו סיכונים ומהם, אך ישפיע המעבר לעבודה חדשה על המשפחה ועוד ....");
                //        }
                //        if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                //        {
                //            txtOut.Add("");
                //        }
                //    }

                //    #endregion
                //    break;
                //case 20:
                //    #region 20

                //    #endregion
                //    break;
                case 21: //21
                    #region 21
                    testText = true;
                    txtOut.Add(ReportDataProvider.Instance.ReadFromTextFile(ReportDataProvider.Instance.ConstructFilePath2WorkInfo(EnumProvider.ReservedWork._תעסוקה_ראשי_.ToString())));
                    #endregion
                    break;
                case 22: //22
                    #region 22

                    dgvQnA.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    dgvQnA.Columns[1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    dgvQnA.Columns[2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    dgvQnA.Columns[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    dgvQnA.Columns[4].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

                    dgvQnA.Columns[0].Width = 110;
                    dgvQnA.Columns[1].Width = 64;
                    dgvQnA.Columns[2].Width = 52;
                    dgvQnA.Columns[3].Width = 52;
                    dgvQnA.Columns[4].Width = 52;
                    dgvQnA.Columns[5].Visible = false;
                    dgvQnA.Update();

                    sqs = sCurrentFilter.Split("#".ToCharArray()[0]);

                    #region Dates
                    for (int i = 0; i < DateList.Count; i++)
                    {
                        DateTime futuredate = DateList[i];//DateTime.Now.AddDays(days * i);

                        string pY, pM, pD, strDate;
                        Calculator.CalcPersonalInfo(mB_Date, futuredate, out pY, out pM, out pD);

                        DataGridViewRow dR = new DataGridViewRow();
                        //dataFuture.Rows.Add(1);

                        strDate = futuredate.Day.ToString() + "." + futuredate.Month.ToString() + "." + futuredate.Year.ToString() + "  (" + FlipString(Calculator.FateCalc(futuredate)) + ")";

                        #region Climax - Future
                        DateTime newBD = futuredate;
                        int[] fLC = Calculator.CareteLifeCycles(newBD);

                        // first climax
                        string ftrClmx = Calculator.CalcSum(fLC[2]);
                        #endregion

                        #region Question Filter
                        bool passedfilter = true;
                        if (Calculator.isCarmaticNumber(futuredate.Day) == true)
                        {
                            passedfilter = false;
                        }
                        else
                        {
                            int tmpPy = Calculator.GetCorrectNumberFromSplitedString(pY.Split(Calculator.Delimiter));
                            int tmpPm = Calculator.GetCorrectNumberFromSplitedString(pM.Split(Calculator.Delimiter));
                            int tmpPd = Calculator.GetCorrectNumberFromSplitedString(pD.Split(Calculator.Delimiter));
                            int tmpFate = Calculator.GetCorrectNumberFromSplitedString(Calculator.FateCalc(futuredate).Split(Calculator.Delimiter));
                            int tmpFtrClmx = Calculator.GetCorrectNumberFromSplitedString(ftrClmx.Split(Calculator.Delimiter));

                            sCurrentFilter = sqs[0];

                            if (cbPYtake.Checked == true)
                            {
                                passedfilter &= sCurrentFilter.Contains("," + tmpPy.ToString() + ",");
                            }

                            passedfilter &= sCurrentFilter.Contains("," + tmpPm.ToString() + ",");
                            passedfilter &= sCurrentFilter.Contains("," + tmpPd.ToString() + ",");
                            passedfilter &= sCurrentFilter.Contains("," + tmpFate.ToString() + ",");
                            passedfilter &= sCurrentFilter.Contains("," + tmpFtrClmx.ToString() + ",");
                        }
                        #endregion Question Filter

                        if (passedfilter == true)
                        {
                            object[] objects = new string[5];
                            objects.Initialize();
                            objects.SetValue(strDate, 0);
                            objects.SetValue(FlipString(pY), 1);
                            objects.SetValue(FlipString(pM), 2);
                            objects.SetValue(FlipString(pD), 3);
                            objects.SetValue(FlipString(ftrClmx), 4);

                            if (dataFuture.Rows.Count == 0)
                            {
                                dgvQnA.Rows.Insert(0, objects);
                            }
                            else
                            {
                                dgvQnA.Rows.Insert(dataFuture.Rows.Count - 1, objects);
                            }
                            dgvQnA.Update();
                        }
                        //txtFuture.Text += futuredate.Date.ToShortDateString() + mSpaceTab + pY + mSpaceTab + pM + mSpaceTab + pD + System.Environment.NewLine;
                    }
                    #endregion

                    if (dgvQnA.Rows.Count > 0)
                    {
                        testText = true;
                        if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                        {
                            txtOut.Add("מומלץ ליזום את הקידום בדרכים המקובלות. הקידום יתאפשר במועדים הבאים:");
                        }
                        if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                        {
                            txtOut.Add("");
                        }
                    }
                    #endregion
                    break;
                case 23: //23
                    #region 23
                    u = new UserResult();
                    u.GatherDataFromGUI(this.Controls, CalcCurrentCycle());

                    testText = true;
                    if (u.mrkBssnss >= 8)
                    {
                        if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                        {
                            txtOut.Add(" כן מומלץ. ");
                        }
                        if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                        {
                            txtOut.Add("");
                        }
                    }
                    if ((u.mrkBssnss >= 7) && (u.mrkBssnss < 8))
                    {
                        if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                        {
                            txtOut.Add("סיכוי חלש להצלחה עסקית מומלץ לחזק את המפה על מנת להצליח עסקית וכלכלית.");
                        }
                        if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                        {
                            txtOut.Add("");
                        }
                    }
                    if (u.mrkBssnss < 7)
                    {
                        if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                        {
                            txtOut.Add("לא מומלץ.");
                        }
                        if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                        {
                            txtOut.Add("");
                        }
                    }

                    if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                    {
                        txtOut.Add("המועדים הטובים ביותר לפתיחת עסק חדש הינם:");
                    }
                    if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                    {
                        txtOut.Add("");
                    }
                    #region Dates
                    sqs = slQAfilter[3].Split("#".ToCharArray()[0]);

                    DateList = CalcDates2FutureCalc(Convert.ToInt16(cmbInterval.Text), cmbType.Text, Convert.ToInt16(cmbTimes.Items[cmbTimes.Items.Count - 1]));

                    #region Dates
                    dgvQnA.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    dgvQnA.Columns[1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    dgvQnA.Columns[2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    dgvQnA.Columns[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    dgvQnA.Columns[4].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

                    dgvQnA.Columns[0].Width = 110;
                    dgvQnA.Columns[1].Width = 64;
                    dgvQnA.Columns[2].Width = 52;
                    dgvQnA.Columns[3].Width = 52;
                    dgvQnA.Columns[4].Width = 52;
                    dgvQnA.Columns[5].Visible = false;
                    dgvQnA.Update();

                    for (int i = 0; i < DateList.Count; i++)
                    {
                        DateTime futuredate = DateList[i];//DateTime.Now.AddDays(days * i);

                        string pY, pM, pD, strDate;
                        Calculator.CalcPersonalInfo(mB_Date, futuredate, out pY, out pM, out pD);

                        DataGridViewRow dR = new DataGridViewRow();
                        //dataFuture.Rows.Add(1);

                        strDate = futuredate.Day.ToString() + "." + futuredate.Month.ToString() + "." + futuredate.Year.ToString() + "  (" + FlipString(Calculator.FateCalc(futuredate)) + ")";

                        #region Climax - Future
                        DateTime newBD = futuredate;
                        int[] fLC = Calculator.CareteLifeCycles(newBD);

                        // first climax
                        string ftrClmx = Calculator.CalcSum(fLC[2]);
                        #endregion

                        #region Question Filter
                        bool passedfilter = true;
                        if (Calculator.isCarmaticNumber(futuredate.Day) == true)
                        {
                            passedfilter = false;
                        }
                        else
                        {
                            int tmpPy = Calculator.GetCorrectNumberFromSplitedString(pY.Split(Calculator.Delimiter));
                            int tmpPm = Calculator.GetCorrectNumberFromSplitedString(pM.Split(Calculator.Delimiter));
                            int tmpPd = Calculator.GetCorrectNumberFromSplitedString(pD.Split(Calculator.Delimiter));
                            int tmpFate = Calculator.GetCorrectNumberFromSplitedString(Calculator.FateCalc(futuredate).Split(Calculator.Delimiter));
                            int tmpFtrClmx = Calculator.GetCorrectNumberFromSplitedString(ftrClmx.Split(Calculator.Delimiter));

                            sCurrentFilter = sqs[sqs.Length - 1];

                            if (cbPYtake.Checked == true)
                            {
                                passedfilter &= sCurrentFilter.Contains("," + tmpPy.ToString() + ",");
                            }

                            passedfilter &= sCurrentFilter.Contains("," + tmpPm.ToString() + ",");
                            passedfilter &= sCurrentFilter.Contains("," + tmpPd.ToString() + ",");
                            passedfilter &= sCurrentFilter.Contains("," + tmpFate.ToString() + ",");
                            passedfilter &= sCurrentFilter.Contains("," + tmpFtrClmx.ToString() + ",");
                        }
                        #endregion Question Filter

                        if (passedfilter == true)
                        {
                            object[] objects = new string[5];
                            objects.Initialize();
                            objects.SetValue(strDate, 0);
                            objects.SetValue(FlipString(pY), 1);
                            objects.SetValue(FlipString(pM), 2);
                            objects.SetValue(FlipString(pD), 3);
                            objects.SetValue(FlipString(ftrClmx), 4);

                            if (dataFuture.Rows.Count == 0)
                            {
                                dgvQnA.Rows.Insert(0, objects);
                            }
                            else
                            {
                                dgvQnA.Rows.Insert(dataFuture.Rows.Count - 1, objects);
                            }
                            dgvQnA.Update();
                        }
                        //txtFuture.Text += futuredate.Date.ToShortDateString() + mSpaceTab + pY + mSpaceTab + pM + mSpaceTab + pD + System.Environment.NewLine;
                    }

                    #endregion

                    #endregion
                    #endregion
                    break;
                case 24: //24
                    #region 24
                    testText = true;
                    u = new UserResult();
                    u.GatherDataFromGUI(this.Controls, CalcCurrentCycle());

                    if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                    {
                        txtOut.Add("העדפה ראשונה:");
                        txtOut.Add(ReportDataProvider.Instance.ReadFromTextFile(ReportDataProvider.Instance.ConstructFilePath2WorkInfo(EnumProvider.ReservedWork._תעסוקה_ראשי_.ToString())));

                        if (Calculator.GetCorrectNumberFromSplitedString(u.SexAndCreation.Split(Calculator.Delimiter)) !=
                            Convert.ToInt16(u.Astro.Split(" ".ToCharArray()[0])[1].Replace("(", "").Replace(")", "")))
                        {
                            txtOut.Add("העדפה שנייה:");
                            txtOut.Add(ReportDataProvider.Instance.ReadFromTextFile(ReportDataProvider.Instance.ConstructFilePath2WorkInfo(EnumProvider.ReservedWork._תעסוקה_משני_.ToString())));
                        }

                        if ((Calculator.GetCorrectNumberFromSplitedString(u.SexAndCreation.Split(Calculator.Delimiter)) != (Calculator.GetCorrectNumberFromSplitedString(u.PrivateNameNum.Split(Calculator.Delimiter))) ||
                             (Calculator.GetCorrectNumberFromSplitedString(u.PrivateNameNum.Split(Calculator.Delimiter)) != Convert.ToInt16(u.Astro.Split(" ".ToCharArray()[0])[1].Replace("(", "").Replace(")", "")))))
                        {
                            txtOut.Add("העדפה נוספת:");
                            txtOut.Add(ReportDataProvider.Instance.ReadFromTextFile(ReportDataProvider.Instance.ConstructFilePath2WorkInfo(EnumProvider.ReservedWork._תעסוקה_שם_פרטי_.ToString())));
                        }
                    }
                    if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                    {
                        txtOut.Add("First options:");
                        txtOut.Add(ReportDataProvider.Instance.ReadFromTextFile(ReportDataProvider.Instance.ConstructFilePath2WorkInfo(EnumProvider.ReservedWork._תעסוקה_ראשי_.ToString())));

                        if (Calculator.GetCorrectNumberFromSplitedString(u.SexAndCreation.Split(Calculator.Delimiter)) !=
                            Convert.ToInt16(u.Astro.Split(" ".ToCharArray()[0])[1].Replace("(", "").Replace(")", "")))
                        {
                            txtOut.Add("Secondary options:");
                            txtOut.Add(ReportDataProvider.Instance.ReadFromTextFile(ReportDataProvider.Instance.ConstructFilePath2WorkInfo(EnumProvider.ReservedWork._תעסוקה_משני_.ToString())));
                        }

                        if ((Calculator.GetCorrectNumberFromSplitedString(u.SexAndCreation.Split(Calculator.Delimiter)) != (Calculator.GetCorrectNumberFromSplitedString(u.PrivateNameNum.Split(Calculator.Delimiter))) ||
                             (Calculator.GetCorrectNumberFromSplitedString(u.PrivateNameNum.Split(Calculator.Delimiter)) != Convert.ToInt16(u.Astro.Split(" ".ToCharArray()[0])[1].Replace("(", "").Replace(")", "")))))
                        {
                            txtOut.Add("Furthermore:");
                            txtOut.Add(ReportDataProvider.Instance.ReadFromTextFile(ReportDataProvider.Instance.ConstructFilePath2WorkInfo(EnumProvider.ReservedWork._תעסוקה_שם_פרטי_.ToString())));
                        }
                    }

                    #endregion
                    break;
                case 25: //25
                    #region 25
                    testText = true;
                    txtOut = new List<string>();
                    if (thisRes.mrkBssnss >= 8)
                    {
                        bool filterpassed = true;

                        filterpassed &= sCurrentFilter.Contains("," + thisRes.Sun + ",");
                        filterpassed &= sCurrentFilter.Contains("," + thisRes.Throught + ",");
                        filterpassed &= sCurrentFilter.Contains("," + thisRes.LC_Climax + ",");

                        if (filterpassed == true)
                        {
                            if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                            {
                                txtOut.Add("מומלץ לך לפעול בשיתופי פעולה במטרה להצליח כלכלית או לפנות למאמן אישי או עסקי. ");
                            }
                            if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                            {
                                txtOut.Add("It is advised for you to act in co-relation inorder to achieve financial success, or use a personal or business coacher.");
                            }
                        }
                        else
                        {
                            if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                            {
                                txtOut.Add("יש לך את היכולת להצליח בכוחות עצמך. מומלץ להתמקד בתחומים בהם אתה טוב ולהכין תוכנית עבודה מסודרת עם יעדים ומשימות.");
                            }
                            if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                            {
                                txtOut.Add("You have th ability to achive financial seuccess on your own. It is advised to plan your agenda with goals and assingments. You should invest on the practice you have most strength in.");
                            }
                        }
                    }
                    else
                    {
                        if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                        {
                            txtOut.Add("מומלץ לפנות למאמן אישי המתמחה בתחום הכספי או בתחום העסקי במטרה שיכוון אותך באופן הטוב ביותר.");
                        }
                        if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                        {
                            txtOut.Add("It is advised that you should approach a finance coucher or business coacher - to aim you in the best direction for your success.");
                        }
                    }
                    #endregion
                    break;
                case 26:
                    string value = this.Controls.Find("txt" + CalcCurrentCycle().ToString() + "_2", true)[0].Text;
                    int val = Calculator.GetCorrectNumberFromSplitedString(value.Split(Calculator.Delimiter));

                    if (new int[] { 1, 2, 3, 4, 5, 6, 8, 9, 11, 22, 19 }.Any(n => n == val))
                    {
                        txtQnAres.Text = "זהו זמן מתאים לעבור דירה";

                    }
                    else if (new int[] { 7, 13, 14, 16 }.Any(n => n == val))
                    {
                        txtQnAres.Text = "זהו אינו הזמן המומלץ עבורך לעבור דירה";

                    }
                    break;
                case 27:
                    txtQnAres.Text = CalcChangeWork();
                    break;
                case 28:
                    txtQnAres.Text = CalcCanReturnLoan();

                    break;

                    //case 26:
                    //    #region
                    //    testText = true;
                    //    bool timeFound = false;
                    //    DateTime tstDate = DateTime.Now;
                    //    //sCurrentFilter

                    //    DateTime saveDate = DateTimePickerTo.Value;
                    //    while (timeFound == false)
                    //    {
                    //        tstDate.AddDays(1.0);
                    //        DateTimePickerTo.Value = tstDate;

                    //        runCalc();

                    //        timeFound = true;
                    //        timeFound &= sCurrentFilter.Contains("," + Calculator.GetCorrectNumberFromSplitedString(txtPDay.Text.Split(Calculator.Delimiter)) + ",");
                    //        timeFound &= sCurrentFilter.Contains("," + Calculator.GetCorrectNumberFromSplitedString(txtPMonth.Text.Split(Calculator.Delimiter)) + ",");
                    //        timeFound &= sCurrentFilter.Contains("," + Calculator.GetCorrectNumberFromSplitedString(txtPYear.Text.Split(Calculator.Delimiter)) + ",");

                    //        if (tstDate.Subtract(saveDate).TotalDays > 5000)
                    //        {
                    //            if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                    //            {
                    //                MessageBox.Show("עבור שאלה זו אין תשובה");
                    //            }
                    //            if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                    //            {
                    //                MessageBox.Show("No Naswer for this question");
                    //            }
                    //            return;
                    //        }
                    //    }

                    //    DateTimePickerTo.Value = saveDate;
                    //    runCalc();

                    //    string stmpout = "";

                    //    double days = tstDate.Subtract(saveDate).TotalDays;
                    //    if ( days < 32.0)
                    //    {
                    //        if (AppSettings.Instance.ProgramLanguage ==AppSettings.Language.Hebrew)
                    //        {
                    //            stmpout = "דברים יתחילו להסתדר עבורך בעוד " + (Convert.ToInt32(days)).ToString() + " ימים.";
                    //        }
                    //        if (AppSettings.Instance.ProgramLanguage ==AppSettings.Language.English)
                    //        {
                    //            stmpout = "Things will start to get better in" + (Convert.ToInt32(days)).ToString() + " days.";
                    //        }
                    //    }
                    //    if ((days > 32.0) && (days < 367))
                    //    {
                    //        if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                    //        {
                    //            stmpout = "דברים יתחילו להסתדר עבורך בעוד " + (Convert.ToInt32(days/30)).ToString() + " חודשים.";
                    //        }
                    //        if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                    //        {
                    //            stmpout = "Things will start to get better in" + (Convert.ToInt32(days/30)).ToString() + " months.";
                    //        }
                    //    }
                    //    if (days > 366) 
                    //    {
                    //        if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                    //        {
                    //            stmpout = "דברים יתחילו להסתדר עבורך בעוד " + (Convert.ToInt32(days / 366)).ToString() + " שנים.";
                    //        }
                    //        if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                    //        {
                    //            stmpout = "Things will start to get better in" + (Convert.ToInt32(days / 366)).ToString() + " years.";
                    //        }
                    //    }

                    //    txtOut.Add(stmpout);

                    //    if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                    //    {
                    //        stmpout = "תאריך היעד הינו:" + tstDate.ToLongDateString();
                    //    }
                    //    if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                    //    {
                    //        stmpout = "Date found:" + tstDate.ToLongDateString();
                    //    }
                    //    txtOut.Add(stmpout);

                    //    #endregion
                    //    break;
                    //case 27:
                    //    #region
                    //    testText = true;
                    //    txtOut = new List<string>();
                    //    if ((thisRes.MapBalanceStatus == EnumProvider.Balance._מאוזן_) || (thisRes.MapBalanceStatus == EnumProvider.Balance._מאוזן_חלקית_))
                    //    {
                    //        if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                    //        {
                    //            txtOut.Add("השם מתאים לך.");
                    //        }
                    //        if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                    //        {
                    //            txtOut.Add("Your name suites you.");
                    //        }

                    //        #region mssbx
                    //        string msg = "", ttl = "";
                    //        if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                    //        {
                    //            msg = "האם הלקוח/ה הצליח/ה להשיג מטרות/יעדים?";
                    //            ttl = "שאלה";
                    //        }
                    //        if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                    //        {
                    //            msg = "Did the person acheived his/her life goals?";
                    //            ttl = "Chakra Prog Question...";
                    //        }
                    //        #endregion
                    //        DialogResult dlgres = MessageBox.Show(msg,ttl,MessageBoxButtons.YesNo,MessageBoxIcon.Question);

                    //        if (dlgres == System.Windows.Forms.DialogResult.No)
                    //        {
                    //            if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                    //            {
                    //                txtOut.Add("מומלץ להוסיף שם פרטי על מנת לאפשר השגת היעדים והמטרות האישיות.");
                    //            }
                    //            if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                    //            {
                    //                txtOut.Add("It is recommanded to add a name to your given name to help you acheive your goals and insperations.");
                    //            }
                    //        }
                    //    }
                    //    else
                    //    {
                    //        if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                    //        {
                    //            txtOut.Add("השם אינו מתאים.");
                    //        }
                    //        if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                    //        {
                    //            txtOut.Add("The name does not suite you.");
                    //        }
                    //    }
                    //    #endregion
                    //    break;
                    //case 28:
                    //    #region
                    //    MessageBox.Show("לא פעיל כרגע");
                    //    #endregion
                    //    break;
                    //case 29:
                    //    #region
                    //    MessageBox.Show("לא פעיל כרגע");
                    //    #endregion
                    //    break;
                    //case 30:
                    //    #region
                    //    MessageBox.Show("לא פעיל כרגע");
                    //    #endregion
                    //    break;
                    //case 31:
                    //    #region
                    //    MessageBox.Show("לא פעיל כרגע");
                    //    #endregion
                    //    break;
                    //case 32:
                    //    #region
                    //    MessageBox.Show("לא פעיל כרגע");
                    //    #endregion
                    //    break;
                    //case 33:
                    //    #region
                    //    MessageBox.Show("לא פעיל כרגע");
                    //    #endregion
                    //    break;
                    //case 34:
                    //    #region
                    //    MessageBox.Show("לא פעיל כרגע");
                    //    #endregion
                    //    break;
                    //case 35:
                    //    #region
                    //    MessageBox.Show("לא פעיל כרגע");
                    //    #endregion
                    //    break;

            } //switch (Convert.ToInt16(sQ.Split(".".ToCharArray()[0])[0]))

            if (testText == true)
            {
                txtQnAres.Text = "";
                foreach (string s in txtOut)
                {
                    if (txtQnAres.Text.Contains(s.Trim()) == false)
                    {
                        txtQnAres.Text += s.Trim() + Environment.NewLine;
                    }
                }
                txtQnAres.Text = txtQnAres.Text.Trim();
            }

        }

        // **********

        private string CalcChangeWork()
        {
            // if in personal period 13,14,16,19,7
            DateTime futuredate = DateTime.Today;

            string pY, pM, pD;

            Calculator.CalcPersonalInfo(mB_Date, futuredate, out pY, out pM, out pD);

            int personalYear = Calculator.GetCorrectNumberFromSplitedString(pY.Split('='));

            if (new int[] { 13, 14, 16, 19, 7 }.Contains(personalYear))
            {
                return "לא מומלץ בתקופה זו להחליף עבודה, קיים קושי למצוא עבודה חדשה.";

            }
            else
            {
                return "אפשרי בתקופה זו להחליף עבודה. מומלץ למצוא עבודה חדשה לפני שאת/ה מתפטר/ת.";

            }

        }

        // **********

        private string CalcCanReturnLoan()
        {
            Dictionary<string, int> chakras = new Dictionary<string, int>();
            Dictionary<string, int> lifePeriods = new Dictionary<string, int>();
            Calc cl = Calculator;

            int curCycle = CalcCurrentCycle();

            FillChakrasDictionary(chakras, lifePeriods);

            if (cl.ChakraContainsValues(chakras, new string[] { "8", "1", "2", "3", "5", "7" }, new int[] { 14, 16 }) ||
                cl.LifePeriodsContainsValues(lifePeriods, new int[] { 14, 16 }, new int[] { 0, 1, 1, 0 }, curCycle))
            {
                return "לא מומלץ לקחת הלוואות. קיים קושי להחזיר הלוואות. פעל בזהירות.";

            }

            if (cl.ChakraContainsValues(chakras, new string[] { "9", "2", "3", "5" }, new int[] { 7 }) ||
                cl.LifePeriodsContainsValues(lifePeriods, new int[] { 7 }, new int[] { 0, 0, 1, 0 }, curCycle))
            {
                return "לא מומלץ לקחת הלוואות. קיים קושי להחזיר הלוואות. פעל בזהירות.";

            }

            if (cl.LifePeriodsContainsValues(lifePeriods, new int[] { 7, 16 }, new int[] { 0, 0, 1, 0 }, 1))
            {
                return "לא מומלץ לקחת הלוואות. קיים קושי להחזיר הלוואות. פעל בזהירות.";

            }

            return "במידה ותיקח היום הלוואה - תצליח להחזירה.";

        }

        // **********

        public bool ConvertDataGridView2CSVfile(DataGridView dgv, string outpath)
        {
            bool finalres = true;
            FileInfo file = new FileInfo(outpath);
            if (file.Exists == true)
            {
                DialogResult dlgres = System.Windows.Forms.DialogResult.OK;

                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                {
                    dlgres = MessageBox.Show("The File:" + Environment.NewLine + outpath + Environment.NewLine + "already exists.... overwite?", "Export Table", MessageBoxButtons.YesNo);
                }
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                {
                    dlgres = MessageBox.Show("הקובץ:" + Environment.NewLine + outpath + Environment.NewLine + "כבר קיים...... האם להחליף את הקובץ?", "ייצוא", MessageBoxButtons.YesNo);
                }

                switch (dlgres)
                {
                    case DialogResult.Abort:
                    case DialogResult.Cancel:
                    case DialogResult.No:
                    case DialogResult.Ignore:
                    case DialogResult.Retry:
                        return false;
                    //break;
                    case DialogResult.Yes:
                    case DialogResult.OK:
                        // just keep going
                        break;
                }
            }

            try
            {
                file.Delete();
                string sEncoding = "";
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                {
                    sEncoding = Encoding.Unicode.ToString();
                }
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                {
                    //sEncoding = "iso-8859-8-i"; // Hebrew Logical
                    sEncoding = "iso-8859-8"; // Hebrew Logical
                                              //sEncoding = "windows-1255"; // Windows Hebrew
                                              //sEncoding = "ibm-862"; // IBM
                }
                FileStream fs = new FileStream(file.FullName, FileMode.OpenOrCreate);
                StreamWriter sw = new StreamWriter(fs, Encoding.GetEncoding(sEncoding));

                #region Header
                string sPrint = "";
                for (int i = 0; i < dgv.Columns.Count; i++)
                {
                    sPrint += dgv.Columns[i].HeaderText.Replace("\n", string.Empty) + ",";

                }
                sPrint = sPrint.Substring(0, sPrint.Length - 1);
                #endregion
                sw.WriteLine(sPrint);

                for (int i = 0; i < dgv.Rows.Count; i++)
                {
                    sPrint = "";
                    for (int c = 0; c < dgv.Rows[i].Cells.Count; c++)
                    {
                        if (i == 0)
                        {
                            sPrint += FlipString(dgv.Rows[i].Cells[c].Value.ToString()) + ",";
                        }
                        else
                        {
                            sPrint += dgv.Rows[i].Cells[c].Value + ",";
                        }
                    }
                    sPrint = sPrint.Substring(0, sPrint.Length - 1);
                    sw.WriteLine(sPrint);
                }


                sw.Close();
                fs.Close();

                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                {
                    MessageBox.Show("Export Complete", "Export", MessageBoxButtons.OK);
                }
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                {
                    MessageBox.Show("הייצוא הושלם", "ייצוא", MessageBoxButtons.OK);
                }
            }
            catch (Exception exp)
            {
                finalres = false;
                string strtxt = "", strcap = "";
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                {
                    strcap = "ייצוא";
                    strtxt = "יצירת הקובץ נכשל";
                }
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                {
                    strcap = "Export";
                    strtxt = "File Creation Failed";
                }
                MessageBox.Show(strtxt + Environment.NewLine + exp.StackTrace, strcap);
            }
            return finalres;
        }

        // **********

        private void btnExportFutureTable_Click(object sender, EventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Filter = "Microsoft Excel (*.csv)|*.csv";
            sfd.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

            string title = "", initialFileName = "";
            if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
            {
                title = "בחר קובץ לייצוא";
                initialFileName = "ייצוא_טבלה.csv";
            }
            if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
            {
                title = "Export Table";
                initialFileName = "Table_Export.csv";
            }
            sfd.Title = title;
            sfd.FileName = initialFileName;

            DialogResult res = sfd.ShowDialog();
            switch (res)
            {
                case System.Windows.Forms.DialogResult.Abort:
                case System.Windows.Forms.DialogResult.Cancel:
                case System.Windows.Forms.DialogResult.Ignore:
                case System.Windows.Forms.DialogResult.No:
                case System.Windows.Forms.DialogResult.None:
                case System.Windows.Forms.DialogResult.Retry:
                    //return;
                    break;
                case System.Windows.Forms.DialogResult.OK:
                case System.Windows.Forms.DialogResult.Yes:
                    ConvertDataGridView2CSVfile(dataFuture, sfd.FileName);
                    break;
            } // switch DialogResult
        }

        // **********

        private void button4_Click_1(object sender, EventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Filter = "Microsoft Excel (*.csv)|*.csv";
            sfd.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

            string title = "", initialFileName = "";
            if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
            {
                title = "בחר קובץ לייצוא";
                initialFileName = "ייצוא_טבלה.csv";
            }
            if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
            {
                title = "Export Table";
                initialFileName = "Table_Export.csv";
            }
            sfd.Title = title;
            sfd.FileName = initialFileName;

            DialogResult res = sfd.ShowDialog();
            switch (res)
            {
                case System.Windows.Forms.DialogResult.Abort:
                case System.Windows.Forms.DialogResult.Cancel:
                case System.Windows.Forms.DialogResult.Ignore:
                case System.Windows.Forms.DialogResult.No:
                case System.Windows.Forms.DialogResult.None:
                case System.Windows.Forms.DialogResult.Retry:
                    //return;
                    break;
                case System.Windows.Forms.DialogResult.OK:
                case System.Windows.Forms.DialogResult.Yes:
                    ConvertDataGridView2CSVfile(dgvQnA, sfd.FileName);
                    break;
            } // switch DialogResult
        }

        // **********

        public UserResult getSpouceResults()
        {
            return SpouceResult;
        }

        // **********

        private void btnCalcSalesMatch_Click(object sender, EventArgs e)
        {
            txtSalesStory.Text = string.Empty;

            CalcSalesMatch();

        }

        // **********

        private void CalcSalesMatch()
        {
            int c = CalcCurrentCycle();
            DateTime dt = DateTimePickerTo.Value;
            string t = "";

            for (int i = 1; i < 5; i++)
            {
                txtSalesStory.Text += Environment.NewLine;

                string[] s = (this.Controls.Find("txt" + i.ToString() + "_1", true)[0] as TextBox).Text.Split("-".ToCharArray());
                DateTimePickerTo.Value = mB_Date.AddYears((Convert.ToInt16(s[0]) + Convert.ToInt16(s[1])) / 2);

                #region intro 4 each cycle

                switch (i)
                {
                    case 1:
                        {
                            if (BLL.AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                            {
                                txtSalesStory.Text += "במחזור ראשון: ";

                            }

                            if (BLL.AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                            {
                                txtSalesStory.Text += "First Cycle: ";

                            }
                            break;

                        }

                    case 2:
                        {
                            if (BLL.AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                            {
                                txtSalesStory.Text += "במחזור שני: ";

                            }

                            if (BLL.AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                            {
                                txtSalesStory.Text += "Second Cycle: ";

                            }
                            break;

                        }

                    case 3:
                        {
                            if (BLL.AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                            {
                                txtSalesStory.Text += "במחזור שלישי: ";

                            }

                            if (BLL.AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                            {
                                txtSalesStory.Text += "Third Cycle: ";

                            }
                            break;

                        }

                    case 4:
                        {
                            if (BLL.AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                            {
                                txtSalesStory.Text += "במחזור רביעי: ";

                            }

                            if (BLL.AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                            {
                                txtSalesStory.Text += "Fourth Cycle: ";

                            }
                            break;

                        }

                }

                #endregion intro 4 each cycle

                CalcSalesMatchPerEachCycle();

                txtSalesStory.Text += Environment.NewLine;

                if (c == i)
                {
                    t = txtFinalSalesMark.Text;

                }

            }

            int sexCreation = Calculator.GetCorrectNumberFromSplitedString(txtNum6.Text.Split(Calculator.Delimiter));
            int solarPlexus = Calculator.GetCorrectNumberFromSplitedString(txtNum5.Text.Split(Calculator.Delimiter));
            int throat = Calculator.GetCorrectNumberFromSplitedString(txtNum3.Text.Split(Calculator.Delimiter));
            int astro = Convert.ToInt16(txtAstroName.Text.Split(" ".ToCharArray()[0])[1].Replace("(", "").Replace(")", ""));
            int meta = Calculator.GetCorrectNumberFromSplitedString(txtNum8.Text.Split(Calculator.Delimiter));
            int[] salesDifficulties = new int[] { 2, 7, 11, 16 };

            if ((Convert.ToSingle(t) >= 8) && (new int[] { 1, 3, 5, 8 }.Contains(sexCreation)))
            {
                txtSalesStory.Text += Environment.NewLine + "הנך בעל כישרון טבעי למכירות.";

            }

            if (salesDifficulties.Contains(solarPlexus) ||
                salesDifficulties.Contains(throat) ||
                salesDifficulties.Contains(astro) ||
                meta == 16)
            {
                txtSalesStory.Text += Environment.NewLine + "קיים קושי בתחום המכירות.";

            }


            DateTimePickerTo.Value = dt;
            txtFinalSalesMark.Text = t;
            txtSalesStory.Text = txtSalesStory.Text.Trim();

        }

        // **********

        private void CalcSalesMatchPerEachCycle()
        {
            List<string> inSalesValues = new List<string>();
            int age = Convert.ToInt16(txtAge.Text.Split("(".ToCharArray()[0])[0]);
            age = Convert.ToInt16(DateTimePickerTo.Value.Subtract(mB_Date).TotalDays / 365);

            #region From Chakra Map

            int upperBndry = Convert.ToInt16(txt1_1.Text.Split("-".ToCharArray()[0])[1].Trim());

            int tmpVal = 0;
            if (age <= upperBndry)
            {
                // Crown
                tmpVal = Calculator.GetCorrectNumberFromSplitedString(txtNum1.Text.Split(Calculator.Delimiter));

            }
            else
            {
                //Third Eye
                tmpVal = Calculator.GetCorrectNumberFromSplitedString(txtNum2.Text.Split(Calculator.Delimiter));

            }
            inSalesValues.Add(tmpVal.ToString());

            tmpVal = Calculator.GetCorrectNumberFromSplitedString(txtNum3.Text.Split(Calculator.Delimiter));
            inSalesValues.Add(tmpVal.ToString());

            tmpVal = Calculator.GetCorrectNumberFromSplitedString(txtNum5.Text.Split(Calculator.Delimiter));
            inSalesValues.Add(tmpVal.ToString());

            tmpVal = Calculator.GetCorrectNumberFromSplitedString(txtNum6.Text.Split(Calculator.Delimiter));
            inSalesValues.Add(tmpVal.ToString());

            tmpVal = Calculator.GetCorrectNumberFromSplitedString(txtPName_Num.Text.Split(Calculator.Delimiter));
            inSalesValues.Add(tmpVal.ToString());

            tmpVal = Convert.ToInt16(txtAstroName.Text.Split("(".ToCharArray()[0])[1].Split(")".ToCharArray()[0])[0]);
            int tmpastronum = tmpVal;
            inSalesValues.Add(tmpVal.ToString());

            #endregion From Chakra Map

            double ChakraMapSalesValue = Calculator.ChakraMap2SalesValue(inSalesValues);

            Control.ControlCollection cntrlCls = this.Controls;

            #region Life Cycles

            int curCycle = 0;
            int[] CycleBoundries = new int[4];
            CycleBoundries[0] = Convert.ToInt16(txt1_1.Text.Split("-".ToCharArray()[0])[1].Trim());
            CycleBoundries[1] = Convert.ToInt16(txt2_1.Text.Split("-".ToCharArray()[0])[1].Trim());
            CycleBoundries[2] = Convert.ToInt16(txt3_1.Text.Split("-".ToCharArray()[0])[1].Trim());
            CycleBoundries[3] = Convert.ToInt16(txt4_1.Text.Split("-".ToCharArray()[0])[1].Trim());

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

            int chlng, clmx, ccl;
            chlng = Calculator.GetCorrectNumberFromSplitedString(cntrlCls.Find("txt" + curCycle.ToString() + "_4", true)[0].Text.Split(Calculator.Delimiter));
            clmx = Calculator.GetCorrectNumberFromSplitedString(cntrlCls.Find("txt" + curCycle.ToString() + "_3", true)[0].Text.Split(Calculator.Delimiter));
            ccl = Calculator.GetCorrectNumberFromSplitedString(cntrlCls.Find("txt" + curCycle.ToString() + "_2", true)[0].Text.Split(Calculator.Delimiter));

            inSalesValues = new List<string>();
            inSalesValues.Add(clmx.ToString());
            inSalesValues.Add(ccl.ToString());

            #endregion Life Cycles

            double LifeCycleBusinessValue = Calculator.LifeCycle2SalesValue(inSalesValues);
            double sum = (ChakraMapSalesValue + LifeCycleBusinessValue) / 2;
            sum = Math.Round(sum, 2);

            if (sum > 10)
            {
                sum = 10;

            }

            string additionalStory = string.Empty;

            if (new int[] { 1, 3, 8, 9, 11, 22 }.Contains(clmx))
            {
                additionalStory = "ההצלחה נובעת ממפת היקום";

            }

            if (sum >= 8 && !new int[] { 1, 3, 5, 8, 9, 11, 22 }.Contains(clmx))
            {
                additionalStory += Environment.NewLine + "ההצלחה נובעת ממפת הצ'אקרות, הנך בעל כישורים טובים מאוד במכירות.";

            }

            txtFinalSalesMark.Text = sum.ToString();

            txtSalesStory.Text += "(" + txtFinalSalesMark.Text + ") " + additionalStory;

        }

        // **********

        private void btnCalcSalesMatch4Group_Click(object sender, EventArgs e)
        {
            List<UserInfo> allGroupSalesPartners = new List<UserInfo>();
            List<DataGridViewRow> RowList = new List<DataGridViewRow>();
            DateTime toCalcDay = DateTimePickerTo.Value;

            #region Gather Data From DataGridView

            UserInfo curSalesPartner;

            for (int i = 0; i < mSalesGroup.Rows.Count - 1; i++)
            {
                DataGridViewRow r = mSalesGroup.Rows[i];
                RowList.Add(r);

                curSalesPartner = new UserInfo();

                curSalesPartner.mFirstName = CellValue2String(r.Cells[0]);
                curSalesPartner.mLastName = CellValue2String(r.Cells[1]);
                curSalesPartner.mFatherName = CellValue2String(r.Cells[3]);
                curSalesPartner.mMotherName = CellValue2String(r.Cells[4]);

                bool resDate = DateTime.TryParse(r.Cells[2].Value.ToString().Trim(), out curSalesPartner.mB_Date);

                if (resDate == false)
                {
                    MessageBox.Show("שגיאה בכתיבת תאריך", "שגיאת נתונים", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                curSalesPartner.mCity = CellValue2String(r.Cells[5]);
                curSalesPartner.mStreet = CellValue2String(r.Cells[6]);

                curSalesPartner.mBuildingNum = CellValue2Double(r.Cells[7]);
                curSalesPartner.mAppNum = CellValue2Double(r.Cells[8]);

                curSalesPartner.mPhones = "";
                curSalesPartner.mEMail = "";

                allGroupSalesPartners.Add(curSalesPartner);
            }

            curSalesPartner = new UserInfo();
            curSalesPartner.mFirstName = txtPrivateName.Text.Trim();
            curSalesPartner.mLastName = txtFamilyName.Text.Trim();
            curSalesPartner.mFatherName = txtFatherName.Text.Trim();
            curSalesPartner.mMotherName = txtMotherName.Text.Trim();
            curSalesPartner.mB_Date = DateTimePickerFrom.Value;
            curSalesPartner.mPhones = txtPhones.Text;
            curSalesPartner.mEMail = txtEMail.Text;

            curSalesPartner.mCity = txtCity.Text.Trim();
            curSalesPartner.mStreet = txtStreet.Text.Trim();
            curSalesPartner.mBuildingNum = Convert.ToDouble(txtBiuldingNum.Text);
            curSalesPartner.mAppNum = Convert.ToDouble(txtAppNum.Text);

            allGroupSalesPartners.Add(curSalesPartner);

            #endregion Gather Data From DataGridView

            double finalMultiResult = 0;
            List<string> personalResults = new List<string>();
            List<double> personalSalesRes = new List<double>();

            #region Calc Business Mark For Each Person

            for (int i = 0; i < allGroupSalesPartners.Count; i++)
            {
                curSalesPartner = allGroupSalesPartners[i];

                txtPrivateName.Text = curSalesPartner.mFirstName;
                txtFamilyName.Text = curSalesPartner.mLastName;
                txtFatherName.Text = curSalesPartner.mFatherName;
                txtMotherName.Text = curSalesPartner.mMotherName;
                DateTimePickerFrom.Value = curSalesPartner.mB_Date;
                txtCity.Text = curSalesPartner.mCity;
                txtStreet.Text = curSalesPartner.mStreet;
                txtBiuldingNum.Text = curSalesPartner.mBuildingNum.ToString();
                txtAppNum.Text = curSalesPartner.mAppNum.ToString();

                DateTimePickerTo.Value = toCalcDay;

                runCalc();

                CalcSalesMatch();

                finalMultiResult += Convert.ToDouble(txtFinalSalesMark.Text.Trim());
                personalSalesRes.Add(Convert.ToDouble(txtFinalSalesMark.Text.Trim()));

                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                {
                    personalResults.Add(curSalesPartner.mFirstName + " " + curSalesPartner.mLastName + ", תוצאה עסקית - " + txtFinalSalesMark.Text.Trim() + " , " + txtGroupSalesStory.Text);
                }
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                {
                    personalResults.Add(curSalesPartner.mFirstName + " " + curSalesPartner.mLastName + ", Bussiness Index - " + txtFinalSalesMark.Text.Trim() + " , " + txtGroupSalesStory.Text);
                }

            }

            #endregion Calc Business Mark For Each Person

            mSalesGroup.Rows.Clear();

            for (int i = 0; i < RowList.Count; i++)
            {
                DataGridViewRow r = RowList[i];
                mSalesGroup.Rows.Add(r);

            }

            mSalesGroup.Update();
            mSalesGroup.Refresh();

            #region Show results

            ReorderBusinessParteners(ref personalSalesRes, ref personalResults);

            for (int i = 0; i < personalResults.Count; i++)
            {
                txtGroupSalesStory.Text += System.Environment.NewLine + personalResults[i];

            }

            #endregion Show results

        }

        private void button6_Click(object sender, EventArgs e)
        {
            txtSecretaryStory.Text = "";
            RunSecretaryCalc4cycles(isSecreatry:true, isAccounting:false);
        }

        private void RunSecretaryCalc4cycles(bool isSecreatry, bool isAccounting)
        {
            int c = CalcCurrentCycle();
            DateTime dt = DateTimePickerTo.Value;
            string t = "";
            for (int i = 1; i < 5; i++)
            {
                if (isSecreatry)
                {
                    txtSecretaryStory.Text += Environment.NewLine;
                }
                else if (isAccounting)
                {
                    txtAccountingStory.Text += Environment.NewLine;
                }

                string[] s = (this.Controls.Find("txt" + i.ToString() + "_1", true)[0] as TextBox).Text.Split("-".ToCharArray());
                DateTimePickerTo.Value = mB_Date.AddYears((Convert.ToInt16(s[0]) + Convert.ToInt16(s[1])) / 2);

                #region intro 4 each cycle
                if (i == 1)
                {
                    if (BLL.AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                    {
                        if (isSecreatry)
                        {
                            txtSecretaryStory.Text += "במחזור ראשון: ";
                        }
                        else if (isAccounting)
                        {
                            txtAccountingStory.Text += "במחזור ראשון: ";
                        }
                    }
                    if (BLL.AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                    {
                        if (isSecreatry)
                        {
                            txtSecretaryStory.Text += "First Cycle: ";
                        }
                        else if (isAccounting)
                        {
                            txtAccountingStory.Text += "First Cycle: ";
                        }
                    }
                }
                if (i == 2)
                {
                    if (BLL.AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                    {
                        if (isSecreatry)
                        {
                            txtSecretaryStory.Text += "במחזור שני: ";
                        }
                        else if (isAccounting)
                        {
                            txtAccountingStory.Text += "במחזור שני: ";
                        }
                    }
                    if (BLL.AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                    {
                        if (isSecreatry)
                        {
                            txtSecretaryStory.Text += "Second Cycle: ";
                        }
                        else if (isAccounting)
                        {
                            txtAccountingStory.Text += "Second Cycle: ";
                        }
                    }
                }
                if (i == 3)
                {
                    if (BLL.AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                    {
                        if (isSecreatry)
                        {
                            txtSecretaryStory.Text += "במחזור שלישי: ";
                        }
                        else if (isAccounting)
                        {
                            txtAccountingStory.Text += "במחזור שלישי: ";
                        }
                    }
                    if (BLL.AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                    {
                        if (isSecreatry)
                        {
                            txtSecretaryStory.Text += "Third Cycle: ";
                        }
                        else if (isAccounting)
                        {
                            txtAccountingStory.Text += "Third Cycle: ";
                        }
                    }
                }
                if (i == 4)
                {
                    if (BLL.AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                    {
                        if (isSecreatry)
                        {
                            txtSecretaryStory.Text += "במחזור רביעי: ";
                        }
                        else if (isAccounting)
                        {
                            txtAccountingStory.Text += "במחזור רביעי: ";
                        }
                    }
                    if (BLL.AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                    {
                        if (isSecreatry)
                        {
                            txtSecretaryStory.Text += "Fourth Cycle: ";
                        }
                        else if (isAccounting)
                        {
                            txtAccountingStory.Text += "Fourth Cycle: ";
                        }
                    }
                }
                #endregion intro 4 each cycle
                SingleSecreteryMatchCalc(isSecreatry:isSecreatry, isAccounting: isAccounting );
                if (isSecreatry)
                {
                    txtSecretaryStory.Text += Environment.NewLine;
                }
                else if (isAccounting)
                {
                    txtAccountingStory.Text += Environment.NewLine;
                }

                if (c == i)
                {
                    if (isSecreatry)
                    {
                        t = txtFinalSeceretyMark.Text;
                    }
                    else if (isAccounting)
                    {
                        t = txtFinalAccountingMark.Text;
                    }
                }
            }

            #region "employee of the month"
            if ((Convert.ToInt16(Calculator.GetCorrectNumberFromSplitedString(txtNum6.Text.Split(Calculator.Delimiter))) == 8) || (Convert.ToInt16(Calculator.GetCorrectNumberFromSplitedString(txtNum3.Text.Split(Calculator.Delimiter))) == 8))
            {
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                {
                    if (isSecreatry)
                    {
                        txtSecretaryStory.Text += Environment.NewLine + "יש לך פוטנציאל עובד מצטיין";
                    }
                    else if (isAccounting)
                    {
                        txtAccountingStory.Text += Environment.NewLine + "יש לך פוטנציאל עובד מצטיין";
                    }
                }
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                {
                    if (isSecreatry)
                    {
                        txtSecretaryStory.Text += Environment.NewLine + "You have the potential of being an excelent employee";
                    }
                    else if (isAccounting)
                    {
                        txtAccountingStory.Text += Environment.NewLine + "You have the potential of being an excelent employee";
                    }
                }
            }
            #endregion

            #region Comment - 7
            int av = Convert.ToInt16(txtAstroName.Text.Split(" ".ToCharArray()[0])[1].Replace("(", "").Replace(")", ""));

            //if ((Calculator.GetCorrectNumberFromSplitedString(txtNum3.Text.Split(Calculator.Delimiter)) == 7) ||
            //     (Calculator.GetCorrectNumberFromSplitedString(txtNum5.Text.Split(Calculator.Delimiter)) == 7) ||
            //     (av == 7))
            if (new int[] { 2, 7, 13, 14, 16, 19 }.Contains(Calculator.GetCorrectNumberFromSplitedString(txtNum3.Text.Split(Calculator.Delimiter))))
            {
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                {
                    if (isSecreatry)
                    {
                        txtSecretaryStory.Text += "יש לך פוטנציאל עובד מצטיין";
                        //"יכולות החשיבה שלך מונעות ממך למצות את הפוטנציאל הגלום בך להצלחה גבוהה יותר. מומלץ לך להוסיף שם על מנת לחדד ולשפר את יכולתך העסקית דבר שיבוא לידי ביטויו בפן העסקי והכלכלי. ";
                    }
                    else if (isAccounting)
                    {
                        txtAccountingStory.Text += "יש לך פוטנציאל עובד מצטיין";
                    }

                }
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                {
                    if (isSecreatry)
                    {
                        txtSecretaryStory.Text += "Though your thinking abilities prevent the accomplishment of your potential to greater success, it is recommended for you to add a name in order to refine your financial strength.";
                    }
                    else if (isAccounting)
                    {
                        txtAccountingStory.Text += "Though your thinking abilities prevent the accomplishment of your potential to greater success, it is recommended for you to add a name in order to refine your financial strength.";
                    }
                }
            }
            #endregion Comment - 7


            DateTimePickerTo.Value = dt;
            
            if (isSecreatry)
            {
                txtFinalSeceretyMark.Text = t;
                txtSecretaryStory.Text = txtSecretaryStory.Text.Trim();
            }
            else if (isAccounting)
            {
                txtFinalAccountingMark.Text = t;
                txtAccountingStory.Text = txtAccountingStory.Text.Trim();
            }
        }

        private void btnMultiPartnersInSecretery_Click(object sender, EventArgs e)
        {
            //run single
            txtSecretaryStory.Text = "";
            RunSecretaryCalc4cycles(isSecreatry: true, isAccounting: false);

            //run multi
            isMultiSecereteryCalc = true;

            List<UserInfo> allGroupSalesPartners = new List<UserInfo>();
            List<DataGridViewRow> RowList = new List<DataGridViewRow>();
            DateTime toCalcDay = DateTimePickerTo.Value;

            #region Gather Data From DataGridView

            UserInfo curSalesPartner;

            for (int i = 0; i < mPartnersSecretery.Rows.Count - 1; i++)
            {
                DataGridViewRow r = mPartnersSecretery.Rows[i];
                RowList.Add(r);

                curSalesPartner = new UserInfo();

                curSalesPartner.mFirstName = CellValue2String(r.Cells[0]);
                curSalesPartner.mLastName = CellValue2String(r.Cells[1]);
                curSalesPartner.mFatherName = CellValue2String(r.Cells[3]);
                curSalesPartner.mMotherName = CellValue2String(r.Cells[4]);

                bool resDate = DateTime.TryParse(r.Cells[2].Value.ToString().Trim(), out curSalesPartner.mB_Date);

                if (resDate == false)
                {
                    MessageBox.Show("שגיאה בכתיבת תאריך", "שגיאת נתונים", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                curSalesPartner.mCity = CellValue2String(r.Cells[5]);
                curSalesPartner.mStreet = CellValue2String(r.Cells[6]);

                curSalesPartner.mBuildingNum = CellValue2Double(r.Cells[7]);
                curSalesPartner.mAppNum = CellValue2Double(r.Cells[8]);

                curSalesPartner.mPhones = "";
                curSalesPartner.mEMail = "";

                allGroupSalesPartners.Add(curSalesPartner);
            }

            curSalesPartner = new UserInfo();
            curSalesPartner.mFirstName = txtPrivateName.Text.Trim();
            curSalesPartner.mLastName = txtFamilyName.Text.Trim();
            curSalesPartner.mFatherName = txtFatherName.Text.Trim();
            curSalesPartner.mMotherName = txtMotherName.Text.Trim();
            curSalesPartner.mB_Date = DateTimePickerFrom.Value;
            curSalesPartner.mPhones = txtPhones.Text;
            curSalesPartner.mEMail = txtEMail.Text;

            curSalesPartner.mCity = txtCity.Text.Trim();
            curSalesPartner.mStreet = txtStreet.Text.Trim();
            curSalesPartner.mBuildingNum = Convert.ToDouble(txtBiuldingNum.Text);
            curSalesPartner.mAppNum = Convert.ToDouble(txtAppNum.Text);

            allGroupSalesPartners.Add(curSalesPartner);

            #endregion Gather Data From DataGridView

            double finalMultiResult = 0;
            List<string> personalResults = new List<string>();
            List<double> personalSalesRes = new List<double>();

            #region Calc Business Mark For Each Person

            for (int i = 0; i < allGroupSalesPartners.Count; i++)
            {
                curSalesPartner = allGroupSalesPartners[i];

                txtPrivateName.Text = curSalesPartner.mFirstName;
                txtFamilyName.Text = curSalesPartner.mLastName;
                txtFatherName.Text = curSalesPartner.mFatherName;
                txtMotherName.Text = curSalesPartner.mMotherName;
                DateTimePickerFrom.Value = curSalesPartner.mB_Date;
                txtCity.Text = curSalesPartner.mCity;
                txtStreet.Text = curSalesPartner.mStreet;
                txtBiuldingNum.Text = curSalesPartner.mBuildingNum.ToString();
                txtAppNum.Text = curSalesPartner.mAppNum.ToString();

                DateTimePickerTo.Value = toCalcDay;

                runCalc();

                SingleSecreteryMatchCalc(isSecreatry: true, isAccounting: false);

                finalMultiResult += Convert.ToDouble(txtFinalSeceretyMark.Text.Trim());
                personalSalesRes.Add(Convert.ToDouble(txtFinalSeceretyMark.Text.Trim()));

                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                {
                    personalResults.Add(curSalesPartner.mFirstName + " " + curSalesPartner.mLastName + ", תוצאה עסקית - " + txtFinalSeceretyMark.Text.Trim() + " , " + txtMultiSecreteryStory.Text);
                }
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                {
                    personalResults.Add(curSalesPartner.mFirstName + " " + curSalesPartner.mLastName + ", Bussiness Index - " + txtFinalSeceretyMark.Text.Trim() + " , " + txtMultiSecreteryStory.Text);
                }

            }

            #endregion Calc Business Mark For Each Person

            mPartnersSecretery.Rows.Clear();

            for (int i = 0; i < RowList.Count; i++)
            {
                DataGridViewRow r = RowList[i];
                mPartnersSecretery.Rows.Add(r);

            }

            mPartnersSecretery.Update();
            mPartnersSecretery.Refresh();

            #region Show results

            ReorderBusinessParteners(ref personalSalesRes, ref personalResults);

            for (int i = 0; i < personalResults.Count; i++)
            {
                txtMultiSecreteryStory.Text += System.Environment.NewLine + personalResults[i];

            }

            #endregion Show results

            isMultiSecereteryCalc = false;
        }

        private void button7_Click(object sender, EventArgs e)
        {
            txtAccountingStory.Text = "";
            RunSecretaryCalc4cycles(isSecreatry: false, isAccounting: true);
        }

        private void button8_Click(object sender, EventArgs e)
        {
            //run single
            txtAccountingStory.Text = "";
            RunSecretaryCalc4cycles(isSecreatry: false, isAccounting: true);

            //run multi
            isMultiSecereteryCalc = true;

            List<UserInfo> allPertners = new List<UserInfo>();
            List<string> yesno = new List<string>();
            List<string> selffix = new List<string>();
            List<DataGridViewRow> RowList = new List<DataGridViewRow>();
            DateTime toCalcDay = DateTimePickerTo.Value;

            #region Gather Data From DataGridView
            UserInfo curPartner;
            for (int i = 0; i < mPartnersAccounting.Rows.Count - 1; i++)
            {
                DataGridViewRow r = mPartnersAccounting.Rows[i];
                RowList.Add(r);

                curPartner = new UserInfo();

                curPartner.mFirstName = CellValue2String(r.Cells[0]);
                curPartner.mLastName = CellValue2String(r.Cells[1]);
                curPartner.mFatherName = CellValue2String(r.Cells[3]);
                curPartner.mMotherName = CellValue2String(r.Cells[4]);
                bool resDate = DateTime.TryParse(r.Cells[2].Value.ToString().Trim(), out curPartner.mB_Date);
                if (resDate == false)
                {
                    MessageBox.Show("שגיאה בכתיבת תאריך", "שגיאת נתונים", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                yesno.Add(CellValue2String(r.Cells[5]));
                selffix.Add(CellValue2String(r.Cells[6]));
                curPartner.mCity = CellValue2String(r.Cells[7]);
                curPartner.mStreet = CellValue2String(r.Cells[8]);

                curPartner.mBuildingNum = CellValue2Double(r.Cells[9]);
                curPartner.mAppNum = CellValue2Double(r.Cells[10]);

                curPartner.mPhones = "";
                curPartner.mEMail = "";

                allPertners.Add(curPartner);
            }

            curPartner = new UserInfo();
            curPartner.mFirstName = txtPrivateName.Text.Trim();
            curPartner.mLastName = txtFamilyName.Text.Trim();
            curPartner.mFatherName = txtFatherName.Text.Trim();
            curPartner.mMotherName = txtMotherName.Text.Trim();
            curPartner.mB_Date = DateTimePickerFrom.Value;
            curPartner.mPhones = txtPhones.Text;
            curPartner.mEMail = txtEMail.Text;

            yesno.Add(cmbBusinessBonus.SelectedItem.ToString());
            selffix.Add(cmbSelfFix.SelectedItem.ToString());

            curPartner.mCity = txtCity.Text.Trim();
            curPartner.mStreet = txtStreet.Text.Trim();
            curPartner.mBuildingNum = Convert.ToDouble(txtBiuldingNum.Text);
            curPartner.mAppNum = Convert.ToDouble(txtAppNum.Text);

            allPertners.Add(curPartner);
            #endregion

            double finalMultiResult = 0;
            List<string> personalResults = new List<string>();
            List<double> persoanlBssnsRes = new List<double>();

            ///
            #region Calc Business Mark For Each Person
            for (int i = 0; i < allPertners.Count; i++)
            {
                curPartner = allPertners[i];
                string bonusyesno = yesno[i];
                string bunosselffix = selffix[i];

                txtPrivateName.Text = curPartner.mFirstName;
                txtFamilyName.Text = curPartner.mLastName;
                txtFatherName.Text = curPartner.mFatherName;
                txtMotherName.Text = curPartner.mMotherName;
                DateTimePickerFrom.Value = curPartner.mB_Date;
                txtCity.Text = curPartner.mCity;
                txtStreet.Text = curPartner.mStreet;
                txtBiuldingNum.Text = curPartner.mBuildingNum.ToString();
                txtAppNum.Text = curPartner.mAppNum.ToString();

                DateTimePickerTo.Value = toCalcDay;

                runCalc();

                SingleSecreteryMatchCalc(isSecreatry: false, isAccounting: true);

                finalMultiResult += Convert.ToDouble(txtFinalAccountingMark.Text.Trim());
                persoanlBssnsRes.Add(Convert.ToDouble(txtFinalAccountingMark.Text.Trim()));

                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                {
                    personalResults.Add(curPartner.mFirstName + " " + curPartner.mLastName + ", תוצאה עסקית - " + txtFinalAccountingMark.Text.Trim() + " , " + txtMultiAccountingStory.Text);
                }
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                {
                    personalResults.Add(curPartner.mFirstName + " " + curPartner.mLastName + ", Bussiness Index - " + txtFinalAccountingMark.Text.Trim() + " , " + txtMultiAccountingStory.Text);
                }
            }
            #endregion Calc Business Mark For Each Person


            mPartnersAccounting.Rows.Clear();

            for (int i = 0; i < RowList.Count; i++)
            {
                DataGridViewRow r = RowList[i];
                mPartnersAccounting.Rows.Add(r);

            }

            mPartnersAccounting.Update();
            mPartnersAccounting.Refresh();

            #region Show results

            ReorderBusinessParteners(ref persoanlBssnsRes, ref personalResults);

            for (int i = 0; i < personalResults.Count; i++)
            {
                txtMultiAccountingStory.Text += System.Environment.NewLine + personalResults[i];
            }

            #endregion Show results

            isMultiSecereteryCalc = false;

        }

        private void button9_Click(object sender, EventArgs e)
        {
            txtSeniorMngStory.Text = "";
            RunManagmentCalc4cycles(isSenior:true);
        }

        private void RunManagmentCalc4cycles(bool isSenior)
        {
            int c = CalcCurrentCycle();
            DateTime dt = DateTimePickerTo.Value;
            string t = "";
            for (int i = 1; i < 5; i++)
            {
                if (isSenior)
                {
                    txtSeniorMngStory.Text += Environment.NewLine;
                }
                else
                {
                    txtJuniorMngStory.Text += Environment.NewLine;
                }

                string[] s = (this.Controls.Find("txt" + i.ToString() + "_1", true)[0] as TextBox).Text.Split("-".ToCharArray());
                DateTimePickerTo.Value = mB_Date.AddYears((Convert.ToInt16(s[0]) + Convert.ToInt16(s[1])) / 2);

                #region intro 4 each cycle
                if (i == 1)
                {

                    if (BLL.AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                    {
                        if (isSenior)
                        {
                            txtSeniorMngStory.Text += "במחזור ראשון: ";
                        }
                        else
                        {
                            txtJuniorMngStory.Text += "במחזור ראשון: ";
                        }
                    }
                    if (BLL.AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                    {
                        if (isSenior)
                        {
                            txtSeniorMngStory.Text += "First Cycle: ";
                        }
                        else
                        {
                            txtJuniorMngStory.Text += "First Cycle: ";
                        }
                    }
                }
                if (i == 2)
                {
                    if (BLL.AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                    {
                        if (isSenior)
                        {
                            txtSeniorMngStory.Text += "במחזור שני: ";
                        }
                        else
                        {
                            txtJuniorMngStory.Text += "במחזור שני: ";
                        }
                    }
                    if (BLL.AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                    {
                        if (isSenior)
                        {
                            txtSeniorMngStory.Text += "Second Cycle: ";
                        }
                        else
                        {
                            txtJuniorMngStory.Text += "Second Cycle: ";
                        }
                    }
                }
                if (i == 3)
                {
                    if (BLL.AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                    {
                        if (isSenior)
                        {
                            txtSeniorMngStory.Text += "במחזור שלישי: ";
                        }
                        else
                        {
                            txtJuniorMngStory.Text += "במחזור שלישי: ";
                        }
                    }
                    if (BLL.AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                    {
                        if (isSenior)
                        {
                            txtSeniorMngStory.Text += "Third Cycle: ";
                        }
                        else
                        {
                            txtJuniorMngStory.Text += "Third Cycle: ";
                        }
                    }
                }
                if (i == 4)
                {
                    if (BLL.AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                    {
                        if (isSenior)
                        {
                            txtSeniorMngStory.Text += "במחזור רביעי: ";
                        }
                        else
                        {
                            txtJuniorMngStory.Text += "במחזור רביעי: ";
                        }
                    }
                    if (BLL.AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                    {
                        if (isSenior)
                        {
                            txtSeniorMngStory.Text += "Fourth Cycle: ";
                        }
                        else
                        {
                            txtJuniorMngStory.Text += "Fourth Cycle: ";
                        }
                    }
                }
                #endregion intro 4 each cycle
                var bonus = isSenior ? cmbSeniorMngBonus.SelectedItem.ToString().Trim() : cmbJuniorMngBonus.SelectedItem.ToString().Trim();
                var selfFix = isSenior ? cmbSeniorMngSelfFix.SelectedItem.ToString().Trim() : cmbJuniorMngSelfFix.SelectedItem.ToString().Trim();
                SingleManagmentMatchCalc(bonus, selfFix, isSenior: isSenior);
                if (isSenior)
                {
                    txtSeniorMngStory.Text += Environment.NewLine;
                }
                else
                {
                    txtJuniorMngStory.Text += Environment.NewLine;
                }

                if (c == i)
                {
                    if (isSenior)
                    {
                        t = txtFinalSeniorMngMark.Text;
                    }
                    else
                    {
                        t = txtFinalJuniorMngMark.Text;
                    }
                }
            }

            #region "employee of the month"
            if ((Convert.ToInt16(Calculator.GetCorrectNumberFromSplitedString(txtNum6.Text.Split(Calculator.Delimiter))) == 8) || (Convert.ToInt16(Calculator.GetCorrectNumberFromSplitedString(txtNum3.Text.Split(Calculator.Delimiter))) == 8))
            {
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                {
                    if (isSenior)
                    {
                        txtSeniorMngStory.Text += Environment.NewLine + "יש לך פוטנציאל עובד מצטיין";
                    }
                    else
                    {
                        txtJuniorMngStory.Text += Environment.NewLine + "יש לך פוטנציאל עובד מצטיין";
                    }
                }
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                {
                    if (isSenior)
                    {
                        txtSeniorMngStory.Text += Environment.NewLine + "You have the potential of being an excelent employee";
                    }
                    else
                    {
                        txtJuniorMngStory.Text += Environment.NewLine + "You have the potential of being an excelent employee";
                    }
                }
            }
            #endregion

            #region Comment - 7
            int av = Convert.ToInt16(txtAstroName.Text.Split(" ".ToCharArray()[0])[1].Replace("(", "").Replace(")", ""));

            //if ((Calculator.GetCorrectNumberFromSplitedString(txtNum3.Text.Split(Calculator.Delimiter)) == 7) ||
            //     (Calculator.GetCorrectNumberFromSplitedString(txtNum5.Text.Split(Calculator.Delimiter)) == 7) ||
            //     (av == 7))
            if (new int[] { 2, 7, 13, 14, 16, 19 }.Contains(Calculator.GetCorrectNumberFromSplitedString(txtNum3.Text.Split(Calculator.Delimiter))))
            {
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                {
                    if (isSenior)
                    {
                        txtSeniorMngStory.Text += "יש לך פוטנציאל עובד מצטיין";
                        //"יכולות החשיבה שלך מונעות ממך למצות את הפוטנציאל הגלום בך להצלחה גבוהה יותר. מומלץ לך להוסיף שם על מנת לחדד ולשפר את יכולתך העסקית דבר שיבוא לידי ביטויו בפן העסקי והכלכלי. ";
                    }
                    else
                    {
                        txtJuniorMngStory.Text += "יש לך פוטנציאל עובד מצטיין";
                    }

                }
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                {
                    if (isSenior)
                    {
                        txtSeniorMngStory.Text += "Though your thinking abilities prevent the accomplishment of your potential to greater success, it is recommended for you to add a name in order to refine your financial strength.";
                    }
                    else
                    {
                        txtJuniorMngStory.Text += "Though your thinking abilities prevent the accomplishment of your potential to greater success, it is recommended for you to add a name in order to refine your financial strength.";
                    }
                }
            }
            #endregion Comment - 7


            DateTimePickerTo.Value = dt;
            if (isSenior)
            {
                txtFinalSeniorMngMark.Text = t;
                txtSeniorMngStory.Text = txtSeniorMngStory.Text.Trim();
            }
            else
            {
                txtFinalJuniorMngMark.Text = t;
                txtJuniorMngStory.Text = txtJuniorMngStory.Text.Trim();
            }
        }

        private void SingleManagmentMatchCalc(string BonusYesOrNo, string SelfFixYesOrNo, bool isSenior)
        {
            List<string> inBusinessValues = new List<string>();
            int age = Convert.ToInt16(txtAge.Text.Split("(".ToCharArray()[0])[0]);
            age = Convert.ToInt16(DateTimePickerTo.Value.Subtract(mB_Date).TotalDays / 365);

            #region From Chakra Map
            int upperBndry = Convert.ToInt16(txt1_1.Text.Split("-".ToCharArray()[0])[1].Trim());

            int tmpVal = 0;
            if (age <= upperBndry)
            {
                // Crown
                tmpVal = Calculator.GetCorrectNumberFromSplitedString(txtNum1.Text.Split(Calculator.Delimiter));
            }
            else
            {
                //Third Eye
                tmpVal = Calculator.GetCorrectNumberFromSplitedString(txtNum2.Text.Split(Calculator.Delimiter));
            }
            inBusinessValues.Add(tmpVal.ToString());


            tmpVal = Calculator.GetCorrectNumberFromSplitedString(txtNum3.Text.Split(Calculator.Delimiter));
            inBusinessValues.Add(tmpVal.ToString());

            tmpVal = Calculator.GetCorrectNumberFromSplitedString(txtNum5.Text.Split(Calculator.Delimiter));
            inBusinessValues.Add(tmpVal.ToString());

            tmpVal = Calculator.GetCorrectNumberFromSplitedString(txtNum6.Text.Split(Calculator.Delimiter));
            inBusinessValues.Add(tmpVal.ToString());

            tmpVal = Calculator.GetCorrectNumberFromSplitedString(txtPName_Num.Text.Split(Calculator.Delimiter));
            inBusinessValues.Add(tmpVal.ToString());

            tmpVal = Convert.ToInt16(txtAstroName.Text.Split("(".ToCharArray()[0])[1].Split(")".ToCharArray()[0])[0]);
            int tmpastronum = tmpVal;
            inBusinessValues.Add(tmpVal.ToString());

            string tmpLvl = txtInfo1.Text.Split(System.Environment.NewLine.ToCharArray()[0])[1].Trim();
            if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
            {
            }
            if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
            {
                if (tmpLvl == "Balanced")
                {
                    tmpLvl = "מאוזן";
                }
                if (tmpLvl == "Not Balanced")
                {
                    tmpLvl = "לא מאוזן";
                }
                if (tmpLvl == "Half Balanced")
                {
                    tmpLvl = "חצי מאוזן";
                }
            }
            inBusinessValues.Add(tmpLvl);
            #endregion
            if (isMultiBusineesCalc == true)
            {
                for (int i = 0; i < inBusinessValues.Count; i++)
                {
                    if (inBusinessValues[i] == "2")
                    {
                        // fix to 7 in the business Chakra chart
                        inBusinessValues.RemoveAt(i);
                        inBusinessValues.Insert(i, "33"); // fir grade 7
                    }
                    if (inBusinessValues[i] == "4")
                    {
                        // fix to 7 in the business Chakra chart
                        inBusinessValues.RemoveAt(i);
                        inBusinessValues.Insert(i, "3");// fir grade 8
                    }
                }
            }

            double ChakraMapBusinessValue = Calculator.ChakraMap2BusinessValue(inBusinessValues);

            Control.ControlCollection cntrlCls = this.Controls;

            #region Life Cycles
            int curCycle = 0;
            int[] CycleBoundries = new int[4];
            CycleBoundries[0] = Convert.ToInt16(txt1_1.Text.Split("-".ToCharArray()[0])[1].Trim());
            CycleBoundries[1] = Convert.ToInt16(txt2_1.Text.Split("-".ToCharArray()[0])[1].Trim());
            CycleBoundries[2] = Convert.ToInt16(txt3_1.Text.Split("-".ToCharArray()[0])[1].Trim());
            CycleBoundries[3] = Convert.ToInt16(txt4_1.Text.Split("-".ToCharArray()[0])[1].Trim());

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

            int chlng, clmx, ccl;
            chlng = Calculator.GetCorrectNumberFromSplitedString(cntrlCls.Find("txt" + curCycle.ToString() + "_4", true)[0].Text.Split(Calculator.Delimiter));
            clmx = Calculator.GetCorrectNumberFromSplitedString(cntrlCls.Find("txt" + curCycle.ToString() + "_3", true)[0].Text.Split(Calculator.Delimiter));
            ccl = Calculator.GetCorrectNumberFromSplitedString(cntrlCls.Find("txt" + curCycle.ToString() + "_2", true)[0].Text.Split(Calculator.Delimiter));

            inBusinessValues = new List<string>();
            //inBusinessValues.Add(chlng.ToString());
            inBusinessValues.Add(clmx.ToString());
            inBusinessValues.Add(ccl.ToString());
            #endregion Life Cycles
            if (isMultiBusineesCalc == true)
            {
                for (int i = 0; i < inBusinessValues.Count; i++)
                {
                    if (inBusinessValues[i] == "2")
                    {
                        // fix to 7 in the business Chakra chart
                        inBusinessValues.RemoveAt(i);
                        inBusinessValues.Insert(i, "33");// for grade 7
                    }
                    if (inBusinessValues[i] == "4")
                    {
                        // fix to 7 in the business Chakra chart
                        inBusinessValues.RemoveAt(i);
                        inBusinessValues.Insert(i, "3");// for grade 8
                    }
                }
            }
            double LifeCycleBusinessValue = Calculator.LifeCycle2BusinessValue(inBusinessValues, "כן" == SelfFixYesOrNo);

            double sum = (ChakraMapBusinessValue + LifeCycleBusinessValue) / 2;

            if (((AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew) && (BonusYesOrNo == "כן")) || ((AppSettings.Instance.ProgramLanguage == AppSettings.Language.English) && (BonusYesOrNo == "Yes")))
            {
                sum += 1;
            }
            else // cmbBusinessBonus.SelectedItem.ToString() == "לא"
            {
                sum += 0;
            }


            if (Calculator.GetCorrectNumberFromSplitedString(txtNum6.Text.Split(Calculator.Delimiter)) == 8)
            {
                if (Calculator.GetCorrectNumberFromSplitedString(txtNum5.Text.Split(Calculator.Delimiter)) != 8)
                {
                    sum += 1;
                }
            }

            if (Calculator.GetCorrectNumberFromSplitedString(txtNum3.Text.Split(Calculator.Delimiter)) == 8)
            {
                sum += 0.5;
            }

            // If sex & creation equals 1, add 0.5 point as bonus
            if (Calculator.GetCorrectNumberFromSplitedString(txtNum6.Text.Split(Calculator.Delimiter)) == 1)
            {
                sum += 0.5;

            }
            #region Bunos of SelfFixing

            if (SelfFixYesOrNo == "כן")
            {
                bool resSelfFix = false;
                #region From LifeCycles
                for (int i = 1; i <= curCycle; i++)
                {
                    for (int c = 2; c < 5; c++)
                    {
                        int tmpNum = Calculator.GetCorrectNumberFromSplitedString(cntrlCls.Find("txt" + i.ToString() + "_" + c.ToString(), true)[0].Text.Split(Calculator.Delimiter));
                        if ((tmpNum == 0) || (Calculator.isMaterNumber(tmpNum)) || ((Calculator.isCarmaticNumber(tmpNum)) && (tmpNum != 16)))
                        {
                            resSelfFix = true;
                        }
                    }
                }

                for (int i = 1; i <= curCycle; i++)
                {
                    int tmpNum = Calculator.GetCorrectNumberFromSplitedString(cntrlCls.Find("txt" + i.ToString() + "_2", true)[0].Text.Split(Calculator.Delimiter));

                    if (tmpNum == 16)
                    {
                        sum = sum + 2;
                        i = 11; // exit loop
                    }
                }
                #endregion From LifeCycles

                if (resSelfFix == true)
                {
                    sum += 0.5;
                }
            }
            else // cmbSelfFix.SelectedItem.ToString() == "לא"
            {
                #region xondt
                //bool isFound = false;
                //#region From MapChakra
                //List<int> nums2check = new List<int>();
                //if (curCycle == 1)
                //{
                //    nums2check.Add(Calculator.GetCorrectNumberFromSplitedString(txtNum1.Text.Split(Calculator.Delimiter)));
                //}
                //else
                //{
                //    nums2check.Add(Calculator.GetCorrectNumberFromSplitedString(txtNum2.Text.Split(Calculator.Delimiter)));
                //}

                //nums2check.Add(Calculator.GetCorrectNumberFromSplitedString(txtNum3.Text.Split(Calculator.Delimiter)));

                //nums2check.Add(Calculator.GetCorrectNumberFromSplitedString(txtNum5.Text.Split(Calculator.Delimiter)));

                //nums2check.Add(Calculator.GetCorrectNumberFromSplitedString(txtNum6.Text.Split(Calculator.Delimiter)));

                //nums2check.Add(Calculator.GetCorrectNumberFromSplitedString(txtPName_Num.Text.Split(Calculator.Delimiter)));

                //for (int i = 0; i < nums2check.Count; i++)
                //{
                //    if (nums2check[i] == 16)
                //    {
                //        isFound = true;
                //    }
                //}
                //#endregion From MapChakra

                //if (isFound == true)
                //{
                //    sum -= 0.5;
                //}
                #endregion xondt
            }

            #endregion Bunos of SelfFixing

            if (sum > 10)
            {
                sum = 10;
            }
            if (isSenior)
            {
                txtFinalSeniorMngMark.Text = Math.Round(sum, 2).ToString();
            }
            else
            {
                txtFinalJuniorMngMark.Text = Math.Round(sum, 2).ToString();
            }

            string story = "";
            if (isSenior)
            {
                if ((sum >= 8.6) & (sum <= 10))
                {
                    if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                    {
                        story = "איכות ניהול הטובה ביותר";
                    }
                    if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                    {
                        story = "Best quality management";
                    }
                }

                if ((sum >= 8.1) & (sum < 8.6))
                {
                    if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                    {
                        story = "איכות ניהול טוב מאוד";
                    }
                    if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                    {
                        story = "Quality management is very good";
                    }
                }

                if ((sum >= 0) & (sum < 8.1))
                {
                    if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                    {
                        story = "איכות ניהול בינוני עד חלש, לא מומלץ";
                    }
                    if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                    {
                        story = "Management quality is weak, not recommended";
                    }
                }
            }
            else
            {
                //מנהל זוטר
                if ((sum >= 9.2) & (sum <= 10))
                {
                    if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                    {
                        story = "איכות ניהול הטובה ביותר";
                    }
                    if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                    {
                        story = "Best quality management";
                    }
                }

                if ((sum >= 8.4) & (sum < 9.21))
                {
                    if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                    {
                        story = "איכות ניהול טוב מאוד";
                    }
                    if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                    {
                        story = "Quality management is very good";
                    }
                }

                if ((sum >= 7.6) & (sum < 8.4))
                {
                    if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                    {
                        story = "איכות הניהול טובה";
                    }
                    if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                    {
                        story = "The quality of management is goodd";
                    }
                }
                if ((sum >= 0) & (sum < 7.6))
                {
                    if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                    {
                        story = "איכות ניהול בינוני עד חלש, לא מומלץ";
                    }
                    if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                    {
                        story = "Management quality is weak, not recommended";
                    }
                }
            }

            if (tmpastronum == 22)
            {
                //if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                //{
                //    story += Environment.NewLine + "היות ומזלך האסטרולוגי עקרב, קל לך להיות עצמאי מצליח או עובד מוערך.";
                //}
                //if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                //{
                //    story += Environment.NewLine + "As astrological sign is Scorpio it is easy for you to be a successful indipendent bussiness or a highly estimated worker.";
                //}
            }
            if (isSenior)
            {
                txtSeniorMngStory.Text += "(" + txtFinalSeniorMngMark.Text + ") " + Environment.NewLine + story;
            }
            else
            {
                txtJuniorMngStory.Text += "(" + txtFinalJuniorMngMark.Text + ") " + Environment.NewLine + story;
            }
        }

        private void button10_Click(object sender, EventArgs e)
        {
            //run single
            txtSeniorMngStory.Text = "";
            RunManagmentCalc4cycles(isSenior: true);

            //run multi
            isMultiBusineesCalc = true;

            isNeedToCalculateWithMsala = true;

            List<UserInfo> allPertners = new List<UserInfo>();
            List<string> yesno = new List<string>();
            List<string> selffix = new List<string>();
            List<DataGridViewRow> RowList = new List<DataGridViewRow>();
            DateTime toCalcDay = DateTimePickerTo.Value;

            #region Gather Data From DataGridView
            UserInfo curPartner;
            for (int i = 0; i < mSeniorMngPartners.Rows.Count - 1; i++)
            {
                DataGridViewRow r = mSeniorMngPartners.Rows[i];
                RowList.Add(r);

                curPartner = new UserInfo();

                curPartner.mFirstName = CellValue2String(r.Cells[0]);
                curPartner.mLastName = CellValue2String(r.Cells[1]);
                curPartner.mFatherName = CellValue2String(r.Cells[3]);
                curPartner.mMotherName = CellValue2String(r.Cells[4]);
                bool resDate = DateTime.TryParse(r.Cells[2].Value.ToString().Trim(), out curPartner.mB_Date);
                if (resDate == false)
                {
                    MessageBox.Show("שגיאה בכתיבת תאריך", "שגיאת נתונים", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                yesno.Add(CellValue2String(r.Cells[5]));
                selffix.Add(CellValue2String(r.Cells[6]));
                curPartner.mCity = CellValue2String(r.Cells[7]);
                curPartner.mStreet = CellValue2String(r.Cells[8]);

                curPartner.mBuildingNum = CellValue2Double(r.Cells[9]);
                curPartner.mAppNum = CellValue2Double(r.Cells[10]);

                curPartner.mPhones = "";
                curPartner.mEMail = "";

                allPertners.Add(curPartner);
            }

            curPartner = new UserInfo();
            curPartner.mFirstName = txtPrivateName.Text.Trim();
            curPartner.mLastName = txtFamilyName.Text.Trim();
            curPartner.mFatherName = txtFatherName.Text.Trim();
            curPartner.mMotherName = txtMotherName.Text.Trim();
            curPartner.mB_Date = DateTimePickerFrom.Value;
            curPartner.mPhones = txtPhones.Text;
            curPartner.mEMail = txtEMail.Text;

            yesno.Add(cmbBusinessBonus.SelectedItem.ToString());
            selffix.Add(cmbSelfFix.SelectedItem.ToString());

            curPartner.mCity = txtCity.Text.Trim();
            curPartner.mStreet = txtStreet.Text.Trim();
            curPartner.mBuildingNum = Convert.ToDouble(txtBiuldingNum.Text);
            curPartner.mAppNum = Convert.ToDouble(txtAppNum.Text);

            allPertners.Add(curPartner);
            #endregion

            double finalMultiResult = 0;
            List<string> personalResults = new List<string>();
            List<double> persoanlBssnsRes = new List<double>();
            #region Calc Business Mark For Each Oertner

            for (int i = 0; i < yesno.Count; i++)
            {
                curPartner = allPertners[i];
                string bonusyesno = yesno[i];
                string bunosselffix = selffix[i];

                txtPrivateName.Text = curPartner.mFirstName;
                txtFamilyName.Text = curPartner.mLastName;
                txtFatherName.Text = curPartner.mFatherName;
                txtMotherName.Text = curPartner.mMotherName;
                DateTimePickerFrom.Value = curPartner.mB_Date;
                txtCity.Text = curPartner.mCity;
                txtStreet.Text = curPartner.mStreet;
                txtBiuldingNum.Text = curPartner.mBuildingNum.ToString();
                txtAppNum.Text = curPartner.mAppNum.ToString();

                DateTimePickerTo.Value = toCalcDay;

                runCalc();

                SingleBusinessMatchCalc(bonusyesno, bunosselffix);

                finalMultiResult += Convert.ToDouble(txtFinalSeniorMngMark.Text.Trim());
                persoanlBssnsRes.Add(Convert.ToDouble(txtFinalSeniorMngMark.Text.Trim()));

                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                {
                    personalResults.Add(curPartner.mFirstName + " " + curPartner.mLastName + ", תוצאה עסקית - " + txtFinalSeniorMngMark.Text.Trim() + " , " + txtFinalMultipleSeniorMngsMartk.Text);
                }
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                {
                    personalResults.Add(curPartner.mFirstName + " " + curPartner.mLastName + ", Bussiness Index - " + txtFinalSeniorMngMark.Text.Trim() + " , " + txtFinalMultipleSeniorMngsMartk.Text);
                }
            }

            finalMultiResult = finalMultiResult / allPertners.Count;
            #endregion


            for (int i = 0; i < RowList.Count; i++)
            {
                DataGridViewRow r = RowList[i];
                mSeniorMngPartners.Rows.Add(r);
            }
            mSeniorMngPartners.Update();
            mSeniorMngPartners.Refresh();

            #region Show results
            txtFinalMultipleSeniorMngsMartk.Text = finalMultiResult.ToString();

            string story = "";
            if ((finalMultiResult >= 8.6) & (finalMultiResult < 10))
            {
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                {
                    story = "איכות ניהול הטובה ביותר";
                }
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                {
                    story = "Best quality management";
                }
            }

            if ((finalMultiResult >= 8.1) & (finalMultiResult < 8.6))
            {
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                {
                    story = "איכות ניהול טוב מאוד";
                }
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                {
                    story = "Quality management is very good";
                }
            }

            if ((finalMultiResult >= 0) & (finalMultiResult < 8.1))
            {
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                {
                    story = "איכות ניהול בינוני עד חלש, לא מומלץ";
                }
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                {
                    story = "Management quality is weak, not recommended";
                }
            }

            txtMultiSeniorMngsStory.Text = story + System.Environment.NewLine + System.Environment.NewLine + "פרוט עבור השותפים:";

            //txtMultiSeniorMngsStory.Text += System.Environment.NewLine +personalResults[personalResults.Count - 1];

            ReorderBusinessParteners(ref persoanlBssnsRes, ref personalResults);

            for (int i = 0; i < personalResults.Count; i++)
            {
                txtMultiSeniorMngsStory.Text += System.Environment.NewLine + personalResults[i];

                //DataGridViewRow r = RowList[i];
                //mSeniorMngPartners.Rows.Add(r);
            }
            //mSeniorMngPartners.Update();
            //mSeniorMngPartners.Refresh();
            #endregion

            isNeedToCalculateWithMsala = false;

            isMultiBusineesCalc = false;
        }

        private void button11_Click(object sender, EventArgs e)
        {
            txtJuniorMngStory.Text = "";
            RunManagmentCalc4cycles(isSenior:false);
        }

        private void button12_Click(object sender, EventArgs e)
        {
            //run single
            txtJuniorMngStory.Text = "";
            RunManagmentCalc4cycles(isSenior: false);

            //run multi
            isMultiBusineesCalc = true;

            isNeedToCalculateWithMsala = true;

            List<UserInfo> allPertners = new List<UserInfo>();
            List<string> yesno = new List<string>();
            List<string> selffix = new List<string>();
            List<DataGridViewRow> RowList = new List<DataGridViewRow>();
            DateTime toCalcDay = DateTimePickerTo.Value;

            #region Gather Data From DataGridView
            UserInfo curPartner;
            for (int i = 0; i < mJuniorMngPartners.Rows.Count - 1; i++)
            {
                DataGridViewRow r = mJuniorMngPartners.Rows[i];
                RowList.Add(r);

                curPartner = new UserInfo();

                curPartner.mFirstName = CellValue2String(r.Cells[0]);
                curPartner.mLastName = CellValue2String(r.Cells[1]);
                curPartner.mFatherName = CellValue2String(r.Cells[3]);
                curPartner.mMotherName = CellValue2String(r.Cells[4]);
                bool resDate = DateTime.TryParse(r.Cells[2].Value.ToString().Trim(), out curPartner.mB_Date);
                if (resDate == false)
                {
                    MessageBox.Show("שגיאה בכתיבת תאריך", "שגיאת נתונים", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                yesno.Add(CellValue2String(r.Cells[5]));
                selffix.Add(CellValue2String(r.Cells[6]));
                curPartner.mCity = CellValue2String(r.Cells[7]);
                curPartner.mStreet = CellValue2String(r.Cells[8]);

                curPartner.mBuildingNum = CellValue2Double(r.Cells[9]);
                curPartner.mAppNum = CellValue2Double(r.Cells[10]);

                curPartner.mPhones = "";
                curPartner.mEMail = "";

                allPertners.Add(curPartner);
            }

            curPartner = new UserInfo();
            curPartner.mFirstName = txtPrivateName.Text.Trim();
            curPartner.mLastName = txtFamilyName.Text.Trim();
            curPartner.mFatherName = txtFatherName.Text.Trim();
            curPartner.mMotherName = txtMotherName.Text.Trim();
            curPartner.mB_Date = DateTimePickerFrom.Value;
            curPartner.mPhones = txtPhones.Text;
            curPartner.mEMail = txtEMail.Text;

            yesno.Add(cmbBusinessBonus.SelectedItem.ToString());
            selffix.Add(cmbSelfFix.SelectedItem.ToString());

            curPartner.mCity = txtCity.Text.Trim();
            curPartner.mStreet = txtStreet.Text.Trim();
            curPartner.mBuildingNum = Convert.ToDouble(txtBiuldingNum.Text);
            curPartner.mAppNum = Convert.ToDouble(txtAppNum.Text);

            allPertners.Add(curPartner);
            #endregion

            double finalMultiResult = 0;
            List<string> personalResults = new List<string>();
            List<double> persoanlBssnsRes = new List<double>();
            #region Calc Business Mark For Each Oertner

            for (int i = 0; i < yesno.Count; i++)
            {
                curPartner = allPertners[i];
                string bonusyesno = yesno[i];
                string bunosselffix = selffix[i];

                txtPrivateName.Text = curPartner.mFirstName;
                txtFamilyName.Text = curPartner.mLastName;
                txtFatherName.Text = curPartner.mFatherName;
                txtMotherName.Text = curPartner.mMotherName;
                DateTimePickerFrom.Value = curPartner.mB_Date;
                txtCity.Text = curPartner.mCity;
                txtStreet.Text = curPartner.mStreet;
                txtBiuldingNum.Text = curPartner.mBuildingNum.ToString();
                txtAppNum.Text = curPartner.mAppNum.ToString();

                DateTimePickerTo.Value = toCalcDay;

                runCalc();

                SingleBusinessMatchCalc(bonusyesno, bunosselffix);

                finalMultiResult += Convert.ToDouble(txtFinalJuniorMngMark.Text.Trim());
                persoanlBssnsRes.Add(Convert.ToDouble(txtFinalJuniorMngMark.Text.Trim()));

                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                {
                    personalResults.Add(curPartner.mFirstName + " " + curPartner.mLastName + ", תוצאה עסקית - " + txtFinalJuniorMngMark.Text.Trim() + " , " + txtFinalMultipleJuniorMngsMartk.Text);
                }
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                {
                    personalResults.Add(curPartner.mFirstName + " " + curPartner.mLastName + ", Bussiness Index - " + txtFinalJuniorMngMark.Text.Trim() + " , " + txtFinalMultipleJuniorMngsMartk.Text);
                }
            }

            finalMultiResult = finalMultiResult / allPertners.Count;
            #endregion


            for (int i = 0; i < RowList.Count; i++)
            {
                DataGridViewRow r = RowList[i];
                mJuniorMngPartners.Rows.Add(r);
            }
            mJuniorMngPartners.Update();
            mJuniorMngPartners.Refresh();

            #region Show results
            txtFinalMultipleJuniorMngsMartk.Text = finalMultiResult.ToString();

            string story = "";
            if ((finalMultiResult >= 9.2) & (finalMultiResult < 10))
            {
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                {
                    story = "איכות ניהול הטובה ביותר";
                }
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                {
                    story = "Best quality management";
                }
            }

            if ((finalMultiResult >=8.4) & (finalMultiResult < 9.26))
            {
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                {
                    story = "איכות ניהול טוב מאוד";
                }
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                {
                    story = "Quality management is very good";
                }
            }

            if ((finalMultiResult >= 7.6) & (finalMultiResult < 8.4))
            {
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                {
                    story = "איכות ניהול טובה";
                }
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                {
                    story = "Quality management is good";
                }
            }
            if ((finalMultiResult >= 0) & (finalMultiResult < 7.6))
            {
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                {
                    story = "איכות ניהול בינוני עד חלש, לא מומלץ";
                }
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                {
                    story = "Management quality is weak, not recommended";
                }
            }

            txtMultiJuniorrMngsStory.Text = story + System.Environment.NewLine + System.Environment.NewLine + "פרוט עבור השותפים:";

            //txtMultiJuniorrMngsStory.Text += System.Environment.NewLine +personalResults[personalResults.Count - 1];

            ReorderBusinessParteners(ref persoanlBssnsRes, ref personalResults);

            for (int i = 0; i < personalResults.Count; i++)
            {
                txtMultiJuniorrMngsStory.Text += System.Environment.NewLine + personalResults[i];

                //DataGridViewRow r = RowList[i];
                //mJuniorMngPartners.Rows.Add(r);
            }
            //mJuniorMngPartners.Update();
            //mJuniorMngPartners.Refresh();
            #endregion

            isNeedToCalculateWithMsala = false;

            isMultiBusineesCalc = false;    
        }

        private void button13_Click(object sender, EventArgs e)
        {
            txtAlternativeStory.Text = "";
            RunAlternativeCalc4cycles();
        }

        private void RunAlternativeCalc4cycles()
        {
            int c = CalcCurrentCycle();
            DateTime dt = DateTimePickerTo.Value;
            string t = "";
            for (int i = 1; i < 5; i++)
            {

                txtAlternativeStory.Text += Environment.NewLine;
               

                string[] s = (this.Controls.Find("txt" + i.ToString() + "_1", true)[0] as TextBox).Text.Split("-".ToCharArray());
                DateTimePickerTo.Value = mB_Date.AddYears((Convert.ToInt16(s[0]) + Convert.ToInt16(s[1])) / 2);

                #region intro 4 each cycle
                if (i == 1)
                {
                    if (BLL.AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                    {

                        txtAlternativeStory.Text += "במחזור ראשון: ";                       
                    }
                    if (BLL.AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                    {

                        txtAlternativeStory.Text += "First Cycle: ";  
                    }
                }
                if (i == 2)
                {
                    if (BLL.AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                    {
                        txtAlternativeStory.Text += "במחזור שני: ";
                    }
                    if (BLL.AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                    {

                        txtAlternativeStory.Text += "Second Cycle: ";
                    }
                }
                if (i == 3)
                {
                    if (BLL.AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                    {

                        txtAlternativeStory.Text += "במחזור שלישי: ";
                    }
                    if (BLL.AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                    {
                        txtAlternativeStory.Text += "Third Cycle: ";
                    }
                }
                if (i == 4)
                {
                    if (BLL.AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                    {
                        txtAlternativeStory.Text += "במחזור רביעי: ";
                    }
                    if (BLL.AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                    {
                        txtAlternativeStory.Text += "Fourth Cycle: ";
                    }
                }
                #endregion intro 4 each cycle
                SingleAlternativeMatchCalc();

                txtAlternativeStory.Text += Environment.NewLine;
                

                if (c == i)
                {
                    t = txtFinalAlternativeMark.Text;
                }
            }

            #region "employee of the month"
            if ((Convert.ToInt16(Calculator.GetCorrectNumberFromSplitedString(txtNum6.Text.Split(Calculator.Delimiter))) == 8) || (Convert.ToInt16(Calculator.GetCorrectNumberFromSplitedString(txtNum3.Text.Split(Calculator.Delimiter))) == 8))
            {
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                {
                    txtAlternativeStory.Text += Environment.NewLine + "יש לך פוטנציאל עובד מצטיין";                   
                }
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                {
                    txtAlternativeStory.Text += Environment.NewLine + "You have the potential of being an excelent employee";
                }
            }
            #endregion

            #region Comment - 7
            int av = Convert.ToInt16(txtAstroName.Text.Split(" ".ToCharArray()[0])[1].Replace("(", "").Replace(")", ""));

            //if ((Calculator.GetCorrectNumberFromSplitedString(txtNum3.Text.Split(Calculator.Delimiter)) == 7) ||
            //     (Calculator.GetCorrectNumberFromSplitedString(txtNum5.Text.Split(Calculator.Delimiter)) == 7) ||
            //     (av == 7))
            if (new int[] { 2, 7, 13, 14, 16, 19 }.Contains(Calculator.GetCorrectNumberFromSplitedString(txtNum3.Text.Split(Calculator.Delimiter))))
            {
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                {
                    txtAlternativeStory.Text += "יש לך פוטנציאל עובד מצטיין";
                    //"יכולות החשיבה שלך מונעות ממך למצות את הפוטנציאל הגלום בך להצלחה גבוהה יותר. מומלץ לך להוסיף שם על מנת לחדד ולשפר את יכולתך העסקית דבר שיבוא לידי ביטויו בפן העסקי והכלכלי. ";
                }
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                {
                    txtAlternativeStory.Text += "Though your thinking abilities prevent the accomplishment of your potential to greater success, it is recommended for you to add a name in order to refine your financial strength.";
                }
            }
            #endregion Comment - 7


            DateTimePickerTo.Value = dt;

            txtFinalAlternativeMark.Text = t;
            txtAlternativeStory.Text = txtAlternativeStory.Text.Trim();
           
        }

        private void SingleAlternativeMatchCalc()
        {
            List<string> inBusinessValues = new List<string>();
            int age = Convert.ToInt16(txtAge.Text.Split("(".ToCharArray()[0])[0]);
            age = Convert.ToInt16(DateTimePickerTo.Value.Subtract(mB_Date).TotalDays / 365);

            #region From Chakra Map
            int upperBndry = Convert.ToInt16(txt1_1.Text.Split("-".ToCharArray()[0])[1].Trim());

            int tmpVal = 0;
            if (age <= upperBndry)
            {
                // Crown
                tmpVal = Calculator.GetCorrectNumberFromSplitedString(txtNum1.Text.Split(Calculator.Delimiter));
            }
            else
            {
                //Third Eye
                tmpVal = Calculator.GetCorrectNumberFromSplitedString(txtNum2.Text.Split(Calculator.Delimiter));
            }
            inBusinessValues.Add(tmpVal.ToString());


            tmpVal = Calculator.GetCorrectNumberFromSplitedString(txtNum3.Text.Split(Calculator.Delimiter));
            inBusinessValues.Add(tmpVal.ToString());

            tmpVal = Calculator.GetCorrectNumberFromSplitedString(txtNum5.Text.Split(Calculator.Delimiter));
            inBusinessValues.Add(tmpVal.ToString());

            tmpVal = Calculator.GetCorrectNumberFromSplitedString(txtNum6.Text.Split(Calculator.Delimiter));
            inBusinessValues.Add(tmpVal.ToString());

            tmpVal = Calculator.GetCorrectNumberFromSplitedString(txtPName_Num.Text.Split(Calculator.Delimiter));
            inBusinessValues.Add(tmpVal.ToString());

            tmpVal = Convert.ToInt16(txtAstroName.Text.Split("(".ToCharArray()[0])[1].Split(")".ToCharArray()[0])[0]);
            int tmpastronum = tmpVal;
            inBusinessValues.Add(tmpVal.ToString());

            string tmpLvl = txtInfo1.Text.Split(System.Environment.NewLine.ToCharArray()[0])[1].Trim();
            if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
            {
            }
            if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
            {
                if (tmpLvl == "Balanced")
                {
                    tmpLvl = "מאוזן";
                }
                if (tmpLvl == "Not Balanced")
                {
                    tmpLvl = "לא מאוזן";
                }
                if (tmpLvl == "Half Balanced")
                {
                    tmpLvl = "חצי מאוזן";
                }
            }
            inBusinessValues.Add(tmpLvl);
            #endregion
            if (isMultiSecereteryCalc == true)
            {
                for (int i = 0; i < inBusinessValues.Count; i++)
                {
                    if (inBusinessValues[i] == "2")
                    {
                        // fix to 7 in the business Chakra chart
                        inBusinessValues.RemoveAt(i);
                        inBusinessValues.Insert(i, "33"); // fir grade 7
                    }
                    if (inBusinessValues[i] == "4")
                    {
                        // fix to 7 in the business Chakra chart
                        inBusinessValues.RemoveAt(i);
                        inBusinessValues.Insert(i, "3");// fir grade 8
                    }
                }
            }

            double ChakraMapBusinessValue = Calculator.ChakraMap2AlternatinvgValue(inBusinessValues);

            Control.ControlCollection cntrlCls = this.Controls;

            #region Life Cycles
            int curCycle = 0;
            int[] CycleBoundries = new int[4];
            CycleBoundries[0] = Convert.ToInt16(txt1_1.Text.Split("-".ToCharArray()[0])[1].Trim());
            CycleBoundries[1] = Convert.ToInt16(txt2_1.Text.Split("-".ToCharArray()[0])[1].Trim());
            CycleBoundries[2] = Convert.ToInt16(txt3_1.Text.Split("-".ToCharArray()[0])[1].Trim());
            CycleBoundries[3] = Convert.ToInt16(txt4_1.Text.Split("-".ToCharArray()[0])[1].Trim());

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

            int chlng, clmx, ccl;
            chlng = Calculator.GetCorrectNumberFromSplitedString(cntrlCls.Find("txt" + curCycle.ToString() + "_4", true)[0].Text.Split(Calculator.Delimiter));
            clmx = Calculator.GetCorrectNumberFromSplitedString(cntrlCls.Find("txt" + curCycle.ToString() + "_3", true)[0].Text.Split(Calculator.Delimiter));
            ccl = Calculator.GetCorrectNumberFromSplitedString(cntrlCls.Find("txt" + curCycle.ToString() + "_2", true)[0].Text.Split(Calculator.Delimiter));

            inBusinessValues = new List<string>();
            //inBusinessValues.Add(chlng.ToString());
            inBusinessValues.Add(clmx.ToString());
            inBusinessValues.Add(ccl.ToString());
            #endregion Life Cycles
            if (isMultiSecereteryCalc == true)
            {
                for (int i = 0; i < inBusinessValues.Count; i++)
                {
                    if (inBusinessValues[i] == "2")
                    {
                        // fix to 7 in the business Chakra chart
                        inBusinessValues.RemoveAt(i);
                        inBusinessValues.Insert(i, "33");// for grade 7
                    }
                    if (inBusinessValues[i] == "4")
                    {
                        // fix to 7 in the business Chakra chart
                        inBusinessValues.RemoveAt(i);
                        inBusinessValues.Insert(i, "3");// for grade 8
                    }
                }
            }
            double LifeCycleBusinessValue = Calculator.LifeCycle2AlternatinvgValue(inBusinessValues);

            double sum = (ChakraMapBusinessValue + LifeCycleBusinessValue) / 2;

            //if (((AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew) && (BonusYesOrNo == "כן")) || ((AppSettings.Instance.ProgramLanguage == AppSettings.Language.English) && (BonusYesOrNo == "Yes")))
            //{
            //    sum += 1;
            //}
            //else // cmbBusinessBonus.SelectedItem.ToString() == "לא"
            //{
            //    sum += 0;
            //}

            //Additional points
            // If sex & creation equals 1 or 8, minus 0.5 point as bonus
            int[] sexCreationNumbers = { 1, 8 };
            if (sexCreationNumbers.Contains(Calculator.GetCorrectNumberFromSplitedString(txtNum6.Text.Split(Calculator.Delimiter)))) 
            {
                sum -= 1;
            }

            // If throat equals0, 2, 4, 6, 7, 9, 11, 33  plus 0.5 point as bonus
            int[] throatNumbers = { 0, 2, 4, 6, 7, 9, 11, 33 };
            if (throatNumbers.Contains(Calculator.GetCorrectNumberFromSplitedString(txtNum3.Text.Split(Calculator.Delimiter)))) 
            {
                sum += 0.5;
            }

            // If sex & creation equals 2, 4, 6, 7, 9, 11, 22, 33, 13, 14, 16, 19, plus 0.5 point as bonus
            int[] sexCreation2Numbers = { 2, 4, 6, 7, 9, 11, 22, 33, 13, 14, 16, 19 };
            if (sexCreation2Numbers.Contains(Calculator.GetCorrectNumberFromSplitedString(txtNum6.Text.Split(Calculator.Delimiter))))
            {
                sum += 0.5;
            }
            // If sex & creation equals  3 or 5
            int[] sexCreation3Numbers = { 3, 5 };
            if (sexCreation3Numbers.Contains(Calculator.GetCorrectNumberFromSplitedString(txtNum6.Text.Split(Calculator.Delimiter))))
            {
                sum -= 0.5;
            }

            // If universe chakra & creation equals  2, 4, 6, 7, 8, 11, 33
            int[] universeNumbers = { 2, 4, 6, 7, 8, 11, 33 };
            int universeNumber = Calculator.GetUniverseNumberFromChakra(txtNum9.Text);
            if (universeNumbers.Contains(universeNumber))
            {
                sum += 0.5;
            }
            
            if (sum > 10)
            {
                sum = 10;
            }

            txtFinalAlternativeMark.Text = Math.Round(sum, 2).ToString();
 
            string story = "";
            if ((sum >= 9) & (sum <= 10))
            {
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                {
                    story = "התאמה טובה מאוד";
                }
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                {
                    story = "Very good match";
                }
            }

            if ((sum >= 8) & (sum < 9))
            {
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                {
                    story = "התאמה טובה";
                }
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                {
                    story = "Good match";
                }
            }

            if ((sum >= 0) & (sum < 8))
            {
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                {
                    story = "התאמה חלשה";
                }
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                {
                    story = "Poor match";
                }
            }

            if (tmpastronum == 22)
            {
                //if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                //{
                //    story += Environment.NewLine + "היות ומזלך האסטרולוגי עקרב, קל לך להיות עובד מוערך.";
                //}
                //if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                //{
                //    story += Environment.NewLine + "As astrological sign is Scorpio it is easy for you to be a successful indipendent bussiness or a highly estimated worker.";
                //}
            }

            txtAlternativeStory.Text += "(" + txtFinalAlternativeMark.Text + ") " + Environment.NewLine + story;           
        }

        private void button14_Click(object sender, EventArgs e)
        {
            //run single
            txtAlternativeStory.Text = "";
            RunAlternativeCalc4cycles();

            //run multi
            isMultiSecereteryCalc = true;

            List<UserInfo> allPertners = new List<UserInfo>();
            List<string> yesno = new List<string>();
            List<string> selffix = new List<string>();
            List<DataGridViewRow> RowList = new List<DataGridViewRow>();
            DateTime toCalcDay = DateTimePickerTo.Value;

            #region Gather Data From DataGridView
            UserInfo curPartner;
            for (int i = 0; i < mPartnersAlternativing.Rows.Count - 1; i++)
            {
                DataGridViewRow r = mPartnersAlternativing.Rows[i];
                RowList.Add(r);

                curPartner = new UserInfo();

                curPartner.mFirstName = CellValue2String(r.Cells[0]);
                curPartner.mLastName = CellValue2String(r.Cells[1]);
                curPartner.mFatherName = CellValue2String(r.Cells[3]);
                curPartner.mMotherName = CellValue2String(r.Cells[4]);
                bool resDate = DateTime.TryParse(r.Cells[2].Value.ToString().Trim(), out curPartner.mB_Date);
                if (resDate == false)
                {
                    MessageBox.Show("שגיאה בכתיבת תאריך", "שגיאת נתונים", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                yesno.Add(CellValue2String(r.Cells[5]));
                selffix.Add(CellValue2String(r.Cells[6]));
                curPartner.mCity = CellValue2String(r.Cells[7]);
                curPartner.mStreet = CellValue2String(r.Cells[8]);

                curPartner.mBuildingNum = CellValue2Double(r.Cells[9]);
                curPartner.mAppNum = CellValue2Double(r.Cells[10]);

                curPartner.mPhones = "";
                curPartner.mEMail = "";

                allPertners.Add(curPartner);
            }

            curPartner = new UserInfo();
            curPartner.mFirstName = txtPrivateName.Text.Trim();
            curPartner.mLastName = txtFamilyName.Text.Trim();
            curPartner.mFatherName = txtFatherName.Text.Trim();
            curPartner.mMotherName = txtMotherName.Text.Trim();
            curPartner.mB_Date = DateTimePickerFrom.Value;
            curPartner.mPhones = txtPhones.Text;
            curPartner.mEMail = txtEMail.Text;

            yesno.Add(cmbBusinessBonus.SelectedItem.ToString());
            selffix.Add(cmbSelfFix.SelectedItem.ToString());

            curPartner.mCity = txtCity.Text.Trim();
            curPartner.mStreet = txtStreet.Text.Trim();
            curPartner.mBuildingNum = Convert.ToDouble(txtBiuldingNum.Text);
            curPartner.mAppNum = Convert.ToDouble(txtAppNum.Text);

            allPertners.Add(curPartner);
            #endregion

            double finalMultiResult = 0;
            List<string> personalResults = new List<string>();
            List<double> persoanlBssnsRes = new List<double>();

            ///
            #region Calc Business Mark For Each Person
            for (int i = 0; i < allPertners.Count; i++)
            {
                curPartner = allPertners[i];
                string bonusyesno = yesno[i];
                string bunosselffix = selffix[i];

                txtPrivateName.Text = curPartner.mFirstName;
                txtFamilyName.Text = curPartner.mLastName;
                txtFatherName.Text = curPartner.mFatherName;
                txtMotherName.Text = curPartner.mMotherName;
                DateTimePickerFrom.Value = curPartner.mB_Date;
                txtCity.Text = curPartner.mCity;
                txtStreet.Text = curPartner.mStreet;
                txtBiuldingNum.Text = curPartner.mBuildingNum.ToString();
                txtAppNum.Text = curPartner.mAppNum.ToString();

                DateTimePickerTo.Value = toCalcDay;

                runCalc();

                SingleAlternativeMatchCalc();

                finalMultiResult += Convert.ToDouble(txtFinalAlternativeMark.Text.Trim());
                persoanlBssnsRes.Add(Convert.ToDouble(txtFinalAlternativeMark.Text.Trim()));

                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                {
                    personalResults.Add(curPartner.mFirstName + " " + curPartner.mLastName + ", תוצאה עסקית - " + txtFinalAlternativeMark.Text.Trim() + " , " + txtMultiAlternativingStory.Text);
                }
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                {
                    personalResults.Add(curPartner.mFirstName + " " + curPartner.mLastName + ", Bussiness Index - " + txtFinalAlternativeMark.Text.Trim() + " , " + txtMultiAlternativingStory.Text);
                }
            }
            #endregion Calc Business Mark For Each Person


            mPartnersAlternativing.Rows.Clear();

            for (int i = 0; i < RowList.Count; i++)
            {
                DataGridViewRow r = RowList[i];
                mPartnersAlternativing.Rows.Add(r);

            }

            mPartnersAlternativing.Update();
            mPartnersAlternativing.Refresh();

            #region Show results

            ReorderBusinessParteners(ref persoanlBssnsRes, ref personalResults);

            for (int i = 0; i < personalResults.Count; i++)
            {
                txtMultiAlternativingStory.Text += System.Environment.NewLine + personalResults[i];
            }

            #endregion Show results

            isMultiSecereteryCalc = false;
        }

        private void button15_Click(object sender, EventArgs e)
        {
            txtBeautyStory.Text = "";
            RunBeautyCalc4cycles();
        }

        private void RunBeautyCalc4cycles()
        {
            int c = CalcCurrentCycle();
            DateTime dt = DateTimePickerTo.Value;
            string t = "";
            for (int i = 1; i < 5; i++)
            {

                txtBeautyStory.Text += Environment.NewLine;


                string[] s = (this.Controls.Find("txt" + i.ToString() + "_1", true)[0] as TextBox).Text.Split("-".ToCharArray());
                DateTimePickerTo.Value = mB_Date.AddYears((Convert.ToInt16(s[0]) + Convert.ToInt16(s[1])) / 2);

                #region intro 4 each cycle
                if (i == 1)
                {
                    if (BLL.AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                    {

                        txtBeautyStory.Text += "במחזור ראשון: ";
                    }
                    if (BLL.AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                    {

                        txtBeautyStory.Text += "First Cycle: ";
                    }
                }
                if (i == 2)
                {
                    if (BLL.AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                    {
                        txtBeautyStory.Text += "במחזור שני: ";
                    }
                    if (BLL.AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                    {

                        txtBeautyStory.Text += "Second Cycle: ";
                    }
                }
                if (i == 3)
                {
                    if (BLL.AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                    {

                        txtBeautyStory.Text += "במחזור שלישי: ";
                    }
                    if (BLL.AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                    {
                        txtBeautyStory.Text += "Third Cycle: ";
                    }
                }
                if (i == 4)
                {
                    if (BLL.AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                    {
                        txtBeautyStory.Text += "במחזור רביעי: ";
                    }
                    if (BLL.AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                    {
                        txtBeautyStory.Text += "Fourth Cycle: ";
                    }
                }
                #endregion intro 4 each cycle
                SingleBeautyMatchCalc();

                txtBeautyStory.Text += Environment.NewLine;


                if (c == i)
                {
                    t = txtFinalBeautyMark.Text;
                }
            }

            #region "employee of the month"
            if ((Convert.ToInt16(Calculator.GetCorrectNumberFromSplitedString(txtNum6.Text.Split(Calculator.Delimiter))) == 8) || (Convert.ToInt16(Calculator.GetCorrectNumberFromSplitedString(txtNum3.Text.Split(Calculator.Delimiter))) == 8))
            {
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                {
                    txtBeautyStory.Text += Environment.NewLine + "יש לך פוטנציאל עובד מצטיין";
                }
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                {
                    txtBeautyStory.Text += Environment.NewLine + "You have the potential of being an excelent employee";
                }
            }
            #endregion

            #region Comment - 7
            int av = Convert.ToInt16(txtAstroName.Text.Split(" ".ToCharArray()[0])[1].Replace("(", "").Replace(")", ""));

            //if ((Calculator.GetCorrectNumberFromSplitedString(txtNum3.Text.Split(Calculator.Delimiter)) == 7) ||
            //     (Calculator.GetCorrectNumberFromSplitedString(txtNum5.Text.Split(Calculator.Delimiter)) == 7) ||
            //     (av == 7))
            if (new int[] { 2, 7, 13, 14, 16, 19 }.Contains(Calculator.GetCorrectNumberFromSplitedString(txtNum3.Text.Split(Calculator.Delimiter))))
            {
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                {
                    txtBeautyStory.Text += "יש לך פוטנציאל עובד מצטיין";
                    //"יכולות החשיבה שלך מונעות ממך למצות את הפוטנציאל הגלום בך להצלחה גבוהה יותר. מומלץ לך להוסיף שם על מנת לחדד ולשפר את יכולתך העסקית דבר שיבוא לידי ביטויו בפן העסקי והכלכלי. ";
                }
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                {
                    txtBeautyStory.Text += "Though your thinking abilities prevent the accomplishment of your potential to greater success, it is recommended for you to add a name in order to refine your financial strength.";
                }
            }
            #endregion Comment - 7


            DateTimePickerTo.Value = dt;

            txtFinalBeautyMark.Text = t;
            txtBeautyStory.Text = txtBeautyStory.Text.Trim();

        }

        private void SingleBeautyMatchCalc()
        {
            List<string> inBusinessValues = new List<string>();
            int age = Convert.ToInt16(txtAge.Text.Split("(".ToCharArray()[0])[0]);
            age = Convert.ToInt16(DateTimePickerTo.Value.Subtract(mB_Date).TotalDays / 365);

            #region From Chakra Map
            int upperBndry = Convert.ToInt16(txt1_1.Text.Split("-".ToCharArray()[0])[1].Trim());

            int tmpVal = 0;
            if (age <= upperBndry)
            {
                // Crown
                tmpVal = Calculator.GetCorrectNumberFromSplitedString(txtNum1.Text.Split(Calculator.Delimiter));
            }
            else
            {
                //Third Eye
                tmpVal = Calculator.GetCorrectNumberFromSplitedString(txtNum2.Text.Split(Calculator.Delimiter));
            }
            inBusinessValues.Add(tmpVal.ToString());


            tmpVal = Calculator.GetCorrectNumberFromSplitedString(txtNum3.Text.Split(Calculator.Delimiter));
            inBusinessValues.Add(tmpVal.ToString());

            tmpVal = Calculator.GetCorrectNumberFromSplitedString(txtNum5.Text.Split(Calculator.Delimiter));
            inBusinessValues.Add(tmpVal.ToString());

            tmpVal = Calculator.GetCorrectNumberFromSplitedString(txtNum6.Text.Split(Calculator.Delimiter));
            inBusinessValues.Add(tmpVal.ToString());

            tmpVal = Calculator.GetCorrectNumberFromSplitedString(txtPName_Num.Text.Split(Calculator.Delimiter));
            inBusinessValues.Add(tmpVal.ToString());

            tmpVal = Convert.ToInt16(txtAstroName.Text.Split("(".ToCharArray()[0])[1].Split(")".ToCharArray()[0])[0]);
            int tmpastronum = tmpVal;
            inBusinessValues.Add(tmpVal.ToString());

            string tmpLvl = txtInfo1.Text.Split(System.Environment.NewLine.ToCharArray()[0])[1].Trim();
            if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
            {
            }
            if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
            {
                if (tmpLvl == "Balanced")
                {
                    tmpLvl = "מאוזן";
                }
                if (tmpLvl == "Not Balanced")
                {
                    tmpLvl = "לא מאוזן";
                }
                if (tmpLvl == "Half Balanced")
                {
                    tmpLvl = "חצי מאוזן";
                }
            }
            inBusinessValues.Add(tmpLvl);
            #endregion
            if (isMultiSecereteryCalc == true)
            {
                for (int i = 0; i < inBusinessValues.Count; i++)
                {
                    if (inBusinessValues[i] == "2")
                    {
                        // fix to 7 in the business Chakra chart
                        inBusinessValues.RemoveAt(i);
                        inBusinessValues.Insert(i, "33"); // fir grade 7
                    }
                    if (inBusinessValues[i] == "4")
                    {
                        // fix to 7 in the business Chakra chart
                        inBusinessValues.RemoveAt(i);
                        inBusinessValues.Insert(i, "3");// fir grade 8
                    }
                }
            }

            double ChakraMapBusinessValue = Calculator.ChakraMap2BeautyValue(inBusinessValues);

            Control.ControlCollection cntrlCls = this.Controls;

            #region Life Cycles
            int curCycle = 0;
            int[] CycleBoundries = new int[4];
            CycleBoundries[0] = Convert.ToInt16(txt1_1.Text.Split("-".ToCharArray()[0])[1].Trim());
            CycleBoundries[1] = Convert.ToInt16(txt2_1.Text.Split("-".ToCharArray()[0])[1].Trim());
            CycleBoundries[2] = Convert.ToInt16(txt3_1.Text.Split("-".ToCharArray()[0])[1].Trim());
            CycleBoundries[3] = Convert.ToInt16(txt4_1.Text.Split("-".ToCharArray()[0])[1].Trim());

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

            int chlng, clmx, ccl;
            chlng = Calculator.GetCorrectNumberFromSplitedString(cntrlCls.Find("txt" + curCycle.ToString() + "_4", true)[0].Text.Split(Calculator.Delimiter));
            clmx = Calculator.GetCorrectNumberFromSplitedString(cntrlCls.Find("txt" + curCycle.ToString() + "_3", true)[0].Text.Split(Calculator.Delimiter));
            ccl = Calculator.GetCorrectNumberFromSplitedString(cntrlCls.Find("txt" + curCycle.ToString() + "_2", true)[0].Text.Split(Calculator.Delimiter));

            inBusinessValues = new List<string>();
            //inBusinessValues.Add(chlng.ToString());
            inBusinessValues.Add(clmx.ToString());
            inBusinessValues.Add(ccl.ToString());
            #endregion Life Cycles
            if (isMultiSecereteryCalc == true)
            {
                for (int i = 0; i < inBusinessValues.Count; i++)
                {
                    if (inBusinessValues[i] == "2")
                    {
                        // fix to 7 in the business Chakra chart
                        inBusinessValues.RemoveAt(i);
                        inBusinessValues.Insert(i, "33");// for grade 7
                    }
                    if (inBusinessValues[i] == "4")
                    {
                        // fix to 7 in the business Chakra chart
                        inBusinessValues.RemoveAt(i);
                        inBusinessValues.Insert(i, "3");// for grade 8
                    }
                }
            }
            double LifeCycleBusinessValue = Calculator.LifeCycle2BeautyValue(inBusinessValues);

            double sum = (ChakraMapBusinessValue + LifeCycleBusinessValue) / 2;

            //if (((AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew) && (BonusYesOrNo == "כן")) || ((AppSettings.Instance.ProgramLanguage == AppSettings.Language.English) && (BonusYesOrNo == "Yes")))
            //{
            //    sum += 1;
            //}
            //else // cmbBusinessBonus.SelectedItem.ToString() == "לא"
            //{
            //    sum += 0;
            //}

            //Additional points
            // If sex & creation equals 1 or 5 or 8, minus 0.5 point as bonus
            int[] sexCreationNumbers = { 1, 5, 8 };
            if (sexCreationNumbers.Contains(Calculator.GetCorrectNumberFromSplitedString(txtNum6.Text.Split(Calculator.Delimiter))))
            {
                sum -= 1;
            }

            // If throat equals0, 2, 4, 6, 7, 9, 11, 33  plus 0.5 point as bonus
            int[] throatNumbers = { 0, 2, 4, 6, 7, 9, 11, 22, 33 };
            if (throatNumbers.Contains(Calculator.GetCorrectNumberFromSplitedString(txtNum3.Text.Split(Calculator.Delimiter))))
            {
                sum += 0.5;
            }

            // If sex & creation equals 2, 4, 6, 7, 9, 11, 22, 33, 13, 14, 16, 19, plus 0.5 point as bonus
            int[] sexCreation2Numbers = { 2, 4, 6, 7, 9, 11, 22, 33, 13, 14, 16, 19 };
            if (sexCreation2Numbers.Contains(Calculator.GetCorrectNumberFromSplitedString(txtNum6.Text.Split(Calculator.Delimiter))))
            {
                sum += 0.5;
            }
           
            // If universe chakra & creation equals  2, 4, 6, 7, 8, 11, 33
            int[] universeNumbers = { 2, 4, 6, 7, 8, 11, 33 };
            int universeNumber = Calculator.GetUniverseNumberFromChakra(txtNum9.Text);
            if (universeNumbers.Contains(universeNumber))
            {
                sum += 0.5;
            }

            if (sum > 10)
            {
                sum = 10;
            }

            txtFinalBeautyMark.Text = Math.Round(sum, 2).ToString();

            string story = "";
            if ((sum >= 9) & (sum <= 10))
            {
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                {
                    story = "התאמה טובה מאוד";
                }
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                {
                    story = "Very good match";
                }
            }

            if ((sum >= 8) & (sum < 9))
            {
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                {
                    story = "התאמה טובה";
                }
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                {
                    story = "Good match";
                }
            }

            if ((sum >= 0) & (sum < 8))
            {
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                {
                    story = "התאמה חלשה";
                }
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                {
                    story = "Poor match";
                }
            }

            if (tmpastronum == 22)
            {
                //if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                //{
                //    story += Environment.NewLine + "היות ומזלך האסטרולוגי עקרב, קל לך להיות עובד מוערך.";
                //}
                //if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                //{
                //    story += Environment.NewLine + "As astrological sign is Scorpio it is easy for you to be a successful indipendent bussiness or a highly estimated worker.";
                //}
            }

            txtBeautyStory.Text += "(" + txtFinalBeautyMark.Text + ") " + Environment.NewLine + story;
        }

        private void button16_Click(object sender, EventArgs e)
        {
            //run single
            txtBeautyStory.Text = "";
            RunBeautyCalc4cycles();

            //run multi
            isMultiSecereteryCalc = true;

            List<UserInfo> allPertners = new List<UserInfo>();
            List<string> yesno = new List<string>();
            List<string> selffix = new List<string>();
            List<DataGridViewRow> RowList = new List<DataGridViewRow>();
            DateTime toCalcDay = DateTimePickerTo.Value;

            #region Gather Data From DataGridView
            UserInfo curPartner;
            for (int i = 0; i < mPartnersBeauty.Rows.Count - 1; i++)
            {
                DataGridViewRow r = mPartnersBeauty.Rows[i];
                RowList.Add(r);

                curPartner = new UserInfo();

                curPartner.mFirstName = CellValue2String(r.Cells[0]);
                curPartner.mLastName = CellValue2String(r.Cells[1]);
                curPartner.mFatherName = CellValue2String(r.Cells[3]);
                curPartner.mMotherName = CellValue2String(r.Cells[4]);
                bool resDate = DateTime.TryParse(r.Cells[2].Value.ToString().Trim(), out curPartner.mB_Date);
                if (resDate == false)
                {
                    MessageBox.Show("שגיאה בכתיבת תאריך", "שגיאת נתונים", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                yesno.Add(CellValue2String(r.Cells[5]));
                selffix.Add(CellValue2String(r.Cells[6]));
                curPartner.mCity = CellValue2String(r.Cells[7]);
                curPartner.mStreet = CellValue2String(r.Cells[8]);

                curPartner.mBuildingNum = CellValue2Double(r.Cells[9]);
                curPartner.mAppNum = CellValue2Double(r.Cells[10]);

                curPartner.mPhones = "";
                curPartner.mEMail = "";

                allPertners.Add(curPartner);
            }

            curPartner = new UserInfo();
            curPartner.mFirstName = txtPrivateName.Text.Trim();
            curPartner.mLastName = txtFamilyName.Text.Trim();
            curPartner.mFatherName = txtFatherName.Text.Trim();
            curPartner.mMotherName = txtMotherName.Text.Trim();
            curPartner.mB_Date = DateTimePickerFrom.Value;
            curPartner.mPhones = txtPhones.Text;
            curPartner.mEMail = txtEMail.Text;

            yesno.Add(cmbBusinessBonus.SelectedItem.ToString());
            selffix.Add(cmbSelfFix.SelectedItem.ToString());

            curPartner.mCity = txtCity.Text.Trim();
            curPartner.mStreet = txtStreet.Text.Trim();
            curPartner.mBuildingNum = Convert.ToDouble(txtBiuldingNum.Text);
            curPartner.mAppNum = Convert.ToDouble(txtAppNum.Text);

            allPertners.Add(curPartner);
            #endregion

            double finalMultiResult = 0;
            List<string> personalResults = new List<string>();
            List<double> persoanlBssnsRes = new List<double>();

            ///
            #region Calc Business Mark For Each Person
            for (int i = 0; i < allPertners.Count; i++)
            {
                curPartner = allPertners[i];
                string bonusyesno = yesno[i];
                string bunosselffix = selffix[i];

                txtPrivateName.Text = curPartner.mFirstName;
                txtFamilyName.Text = curPartner.mLastName;
                txtFatherName.Text = curPartner.mFatherName;
                txtMotherName.Text = curPartner.mMotherName;
                DateTimePickerFrom.Value = curPartner.mB_Date;
                txtCity.Text = curPartner.mCity;
                txtStreet.Text = curPartner.mStreet;
                txtBiuldingNum.Text = curPartner.mBuildingNum.ToString();
                txtAppNum.Text = curPartner.mAppNum.ToString();

                DateTimePickerTo.Value = toCalcDay;

                runCalc();

                SingleBeautyMatchCalc();

                finalMultiResult += Convert.ToDouble(txtFinalBeautyMark.Text.Trim());
                persoanlBssnsRes.Add(Convert.ToDouble(txtFinalBeautyMark.Text.Trim()));

                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                {
                    personalResults.Add(curPartner.mFirstName + " " + curPartner.mLastName + ", תוצאה עסקית - " + txtFinalBeautyMark.Text.Trim() + " , " + txtMultiBeautyStory.Text);
                }
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                {
                    personalResults.Add(curPartner.mFirstName + " " + curPartner.mLastName + ", Bussiness Index - " + txtFinalBeautyMark.Text.Trim() + " , " + txtMultiBeautyStory.Text);
                }
            }
            #endregion Calc Business Mark For Each Person


            mPartnersBeauty.Rows.Clear();

            for (int i = 0; i < RowList.Count; i++)
            {
                DataGridViewRow r = RowList[i];
                mPartnersBeauty.Rows.Add(r);

            }

            mPartnersBeauty.Update();
            mPartnersBeauty.Refresh();

            #region Show results

            ReorderBusinessParteners(ref persoanlBssnsRes, ref personalResults);

            for (int i = 0; i < personalResults.Count; i++)
            {
                txtMultiBeautyStory.Text += System.Environment.NewLine + personalResults[i];
            }

            #endregion Show results

            isMultiSecereteryCalc = false;
        }

        private void button17_Click(object sender, EventArgs e)
        {
            txtManualWorkStory.Text = "";
            RunManualWorkCalc4cycles();
     
        }

        private void RunManualWorkCalc4cycles()
        {
            int c = CalcCurrentCycle();
            DateTime dt = DateTimePickerTo.Value;
            string t = "";
            for (int i = 1; i < 5; i++)
            {

                txtManualWorkStory.Text += Environment.NewLine;


                string[] s = (this.Controls.Find("txt" + i.ToString() + "_1", true)[0] as TextBox).Text.Split("-".ToCharArray());
                DateTimePickerTo.Value = mB_Date.AddYears((Convert.ToInt16(s[0]) + Convert.ToInt16(s[1])) / 2);

                #region intro 4 each cycle
                if (i == 1)
                {
                    if (BLL.AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                    {

                        txtManualWorkStory.Text += "במחזור ראשון: ";
                    }
                    if (BLL.AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                    {

                        txtManualWorkStory.Text += "First Cycle: ";
                    }
                }
                if (i == 2)
                {
                    if (BLL.AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                    {
                        txtManualWorkStory.Text += "במחזור שני: ";
                    }
                    if (BLL.AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                    {

                        txtManualWorkStory.Text += "Second Cycle: ";
                    }
                }
                if (i == 3)
                {
                    if (BLL.AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                    {

                        txtManualWorkStory.Text += "במחזור שלישי: ";
                    }
                    if (BLL.AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                    {
                        txtManualWorkStory.Text += "Third Cycle: ";
                    }
                }
                if (i == 4)
                {
                    if (BLL.AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                    {
                        txtManualWorkStory.Text += "במחזור רביעי: ";
                    }
                    if (BLL.AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                    {
                        txtManualWorkStory.Text += "Fourth Cycle: ";
                    }
                }
                #endregion intro 4 each cycle
                SingleManualWorkMatchCalc();

                txtManualWorkStory.Text += Environment.NewLine;


                if (c == i)
                {
                    t = txtFinalMenualWorkMark.Text;
                }
            }

            #region "employee of the month"
            if ((Convert.ToInt16(Calculator.GetCorrectNumberFromSplitedString(txtNum6.Text.Split(Calculator.Delimiter))) == 8) || (Convert.ToInt16(Calculator.GetCorrectNumberFromSplitedString(txtNum3.Text.Split(Calculator.Delimiter))) == 8))
            {
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                {
                    txtManualWorkStory.Text += Environment.NewLine + "יש לך פוטנציאל עובד מצטיין";
                }
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                {
                    txtManualWorkStory.Text += Environment.NewLine + "You have the potential of being an excelent employee";
                }
            }
            #endregion

            #region Comment - 7
            int av = Convert.ToInt16(txtAstroName.Text.Split(" ".ToCharArray()[0])[1].Replace("(", "").Replace(")", ""));

            //if ((Calculator.GetCorrectNumberFromSplitedString(txtNum3.Text.Split(Calculator.Delimiter)) == 7) ||
            //     (Calculator.GetCorrectNumberFromSplitedString(txtNum5.Text.Split(Calculator.Delimiter)) == 7) ||
            //     (av == 7))
            if (new int[] { 2, 7, 13, 14, 16, 19 }.Contains(Calculator.GetCorrectNumberFromSplitedString(txtNum3.Text.Split(Calculator.Delimiter))))
            {
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                {
                    txtManualWorkStory.Text += "יש לך פוטנציאל עובד מצטיין";
                    //"יכולות החשיבה שלך מונעות ממך למצות את הפוטנציאל הגלום בך להצלחה גבוהה יותר. מומלץ לך להוסיף שם על מנת לחדד ולשפר את יכולתך העסקית דבר שיבוא לידי ביטויו בפן העסקי והכלכלי. ";
                }
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                {
                    txtManualWorkStory.Text += "Though your thinking abilities prevent the accomplishment of your potential to greater success, it is recommended for you to add a name in order to refine your financial strength.";
                }
            }
            #endregion Comment - 7


            DateTimePickerTo.Value = dt;

            txtFinalMenualWorkMark.Text = t;
            txtManualWorkStory.Text = txtManualWorkStory.Text.Trim();

        }

        private void SingleManualWorkMatchCalc()
        {
            List<string> inBusinessValues = new List<string>();
            int age = Convert.ToInt16(txtAge.Text.Split("(".ToCharArray()[0])[0]);
            age = Convert.ToInt16(DateTimePickerTo.Value.Subtract(mB_Date).TotalDays / 365);

            #region From Chakra Map
            int upperBndry = Convert.ToInt16(txt1_1.Text.Split("-".ToCharArray()[0])[1].Trim());

            int tmpVal = 0;
            if (age <= upperBndry)
            {
                // Crown
                tmpVal = Calculator.GetCorrectNumberFromSplitedString(txtNum1.Text.Split(Calculator.Delimiter));
            }
            else
            {
                //Third Eye
                tmpVal = Calculator.GetCorrectNumberFromSplitedString(txtNum2.Text.Split(Calculator.Delimiter));
            }
            inBusinessValues.Add(tmpVal.ToString());


            tmpVal = Calculator.GetCorrectNumberFromSplitedString(txtNum3.Text.Split(Calculator.Delimiter));
            inBusinessValues.Add(tmpVal.ToString());

            tmpVal = Calculator.GetCorrectNumberFromSplitedString(txtNum5.Text.Split(Calculator.Delimiter));
            inBusinessValues.Add(tmpVal.ToString());

            tmpVal = Calculator.GetCorrectNumberFromSplitedString(txtNum6.Text.Split(Calculator.Delimiter));
            inBusinessValues.Add(tmpVal.ToString());

            tmpVal = Calculator.GetCorrectNumberFromSplitedString(txtPName_Num.Text.Split(Calculator.Delimiter));
            inBusinessValues.Add(tmpVal.ToString());

            tmpVal = Convert.ToInt16(txtAstroName.Text.Split("(".ToCharArray()[0])[1].Split(")".ToCharArray()[0])[0]);
            int tmpastronum = tmpVal;
            inBusinessValues.Add(tmpVal.ToString());

            string tmpLvl = txtInfo1.Text.Split(System.Environment.NewLine.ToCharArray()[0])[1].Trim();
            if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
            {
            }
            if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
            {
                if (tmpLvl == "Balanced")
                {
                    tmpLvl = "מאוזן";
                }
                if (tmpLvl == "Not Balanced")
                {
                    tmpLvl = "לא מאוזן";
                }
                if (tmpLvl == "Half Balanced")
                {
                    tmpLvl = "חצי מאוזן";
                }
            }
            inBusinessValues.Add(tmpLvl);
            #endregion
            if (isMultiSecereteryCalc == true)
            {
                for (int i = 0; i < inBusinessValues.Count; i++)
                {
                    if (inBusinessValues[i] == "2")
                    {
                        // fix to 7 in the business Chakra chart
                        inBusinessValues.RemoveAt(i);
                        inBusinessValues.Insert(i, "33"); // fir grade 7
                    }
                    if (inBusinessValues[i] == "4")
                    {
                        // fix to 7 in the business Chakra chart
                        inBusinessValues.RemoveAt(i);
                        inBusinessValues.Insert(i, "3");// fir grade 8
                    }
                }
            }

            double ChakraMapBusinessValue = Calculator.ChakraMap2ManualValue(inBusinessValues);

            Control.ControlCollection cntrlCls = this.Controls;

            #region Life Cycles
            int curCycle = 0;
            int[] CycleBoundries = new int[4];
            CycleBoundries[0] = Convert.ToInt16(txt1_1.Text.Split("-".ToCharArray()[0])[1].Trim());
            CycleBoundries[1] = Convert.ToInt16(txt2_1.Text.Split("-".ToCharArray()[0])[1].Trim());
            CycleBoundries[2] = Convert.ToInt16(txt3_1.Text.Split("-".ToCharArray()[0])[1].Trim());
            CycleBoundries[3] = Convert.ToInt16(txt4_1.Text.Split("-".ToCharArray()[0])[1].Trim());

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

            int chlng, clmx, ccl;
            chlng = Calculator.GetCorrectNumberFromSplitedString(cntrlCls.Find("txt" + curCycle.ToString() + "_4", true)[0].Text.Split(Calculator.Delimiter));
            clmx = Calculator.GetCorrectNumberFromSplitedString(cntrlCls.Find("txt" + curCycle.ToString() + "_3", true)[0].Text.Split(Calculator.Delimiter));
            ccl = Calculator.GetCorrectNumberFromSplitedString(cntrlCls.Find("txt" + curCycle.ToString() + "_2", true)[0].Text.Split(Calculator.Delimiter));

            inBusinessValues = new List<string>();
            //inBusinessValues.Add(chlng.ToString());
            inBusinessValues.Add(clmx.ToString());
            inBusinessValues.Add(ccl.ToString());
            #endregion Life Cycles
            if (isMultiSecereteryCalc == true)
            {
                for (int i = 0; i < inBusinessValues.Count; i++)
                {
                    if (inBusinessValues[i] == "2")
                    {
                        // fix to 7 in the business Chakra chart
                        inBusinessValues.RemoveAt(i);
                        inBusinessValues.Insert(i, "33");// for grade 7
                    }
                    if (inBusinessValues[i] == "4")
                    {
                        // fix to 7 in the business Chakra chart
                        inBusinessValues.RemoveAt(i);
                        inBusinessValues.Insert(i, "3");// for grade 8
                    }
                }
            }
            double LifeCycleBusinessValue = Calculator.LifeCycle2ManualValue(inBusinessValues);

            double sum = (ChakraMapBusinessValue + LifeCycleBusinessValue) / 2;

            //if (((AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew) && (BonusYesOrNo == "כן")) || ((AppSettings.Instance.ProgramLanguage == AppSettings.Language.English) && (BonusYesOrNo == "Yes")))
            //{
            //    sum += 1;
            //}
            //else // cmbBusinessBonus.SelectedItem.ToString() == "לא"
            //{
            //    sum += 0;
            //}

            //Additional points
            // If sex & creation equals 1 or 5 or 8, minus 0.5 point as bonus
            int[] sexCreationNumbers = { 1, 8 };
            if (sexCreationNumbers.Contains(Calculator.GetCorrectNumberFromSplitedString(txtNum6.Text.Split(Calculator.Delimiter))))
            {
                sum -= 1;
            }

            // If throat equals0, 2, 4, 6, 9, 11, 33  plus 0.5 point as bonus
            int[] throatNumbers = { 0, 2, 4, 6, 9, 11, 33 };
            if (throatNumbers.Contains(Calculator.GetCorrectNumberFromSplitedString(txtNum3.Text.Split(Calculator.Delimiter))))
            {
                sum += 0.5;
            }

            // If sex & creation equals 2, 4, 6, 7, 9, 11, 22, 33, 13, 14, 16, 19, plus 0.5 point as bonus
            int[] sexCreation2Numbers = { 2, 4, 6, 7, 9, 11, 22, 33, 13, 14, 16, 19 };
            if (sexCreation2Numbers.Contains(Calculator.GetCorrectNumberFromSplitedString(txtNum6.Text.Split(Calculator.Delimiter))))
            {
                sum += 0.5;
            }

            int[] sexCreation3Numbers = { 3, 5 };
            if (sexCreation3Numbers.Contains(Calculator.GetCorrectNumberFromSplitedString(txtNum6.Text.Split(Calculator.Delimiter))))
            {
                sum -= 0.5;
            }

            // If universe chakra & creation equals  2, 4, 6, 7, 8, 11, 33
            int[] universeNumbers = { 2, 4, 6, 7, 8, 11, 33 };
            int universeNumber = Calculator.GetUniverseNumberFromChakra(txtNum9.Text);
            if (universeNumbers.Contains(universeNumber))
            {
                sum += 0.5;
            }

            if (sum > 10)
            {
                sum = 10;
            }

            txtFinalMenualWorkMark.Text = Math.Round(sum, 2).ToString();

            string story = "";
            if ((sum >= 9) & (sum <= 10))
            {
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                {
                    story = "התאמה טובה מאוד";
                }
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                {
                    story = "Very good match";
                }
            }

            if ((sum >= 8) & (sum < 9))
            {
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                {
                    story = "התאמה טובה";
                }
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                {
                    story = "Good match";
                }
            }

            if ((sum >= 0) & (sum < 8))
            {
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                {
                    story = "התאמה חלשה";
                }
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                {
                    story = "Poor match";
                }
            }

            if (tmpastronum == 22)
            {
                //if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                //{
                //    story += Environment.NewLine + "היות ומזלך האסטרולוגי עקרב, קל לך להיות עובד מוערך.";
                //}
                //if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                //{
                //    story += Environment.NewLine + "As astrological sign is Scorpio it is easy for you to be a successful indipendent bussiness or a highly estimated worker.";
                //}
            }

            txtManualWorkStory.Text += "(" + txtFinalMenualWorkMark.Text + ") " + Environment.NewLine + story;
        }

        private void button18_Click(object sender, EventArgs e)
        {
            //Run single
            txtManualWorkStory.Text = "";
            RunManualWorkCalc4cycles();

            //Run multi
            isMultiSecereteryCalc = true;

            List<UserInfo> allPertners = new List<UserInfo>();
            List<string> yesno = new List<string>();
            List<string> selffix = new List<string>();
            List<DataGridViewRow> RowList = new List<DataGridViewRow>();
            DateTime toCalcDay = DateTimePickerTo.Value;

            #region Gather Data From DataGridView
            UserInfo curPartner;
            for (int i = 0; i < mPartnersManual.Rows.Count - 1; i++)
            {
                DataGridViewRow r = mPartnersManual.Rows[i];
                RowList.Add(r);

                curPartner = new UserInfo();

                curPartner.mFirstName = CellValue2String(r.Cells[0]);
                curPartner.mLastName = CellValue2String(r.Cells[1]);
                curPartner.mFatherName = CellValue2String(r.Cells[3]);
                curPartner.mMotherName = CellValue2String(r.Cells[4]);
                bool resDate = DateTime.TryParse(r.Cells[2].Value.ToString().Trim(), out curPartner.mB_Date);
                if (resDate == false)
                {
                    MessageBox.Show("שגיאה בכתיבת תאריך", "שגיאת נתונים", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                yesno.Add(CellValue2String(r.Cells[5]));
                selffix.Add(CellValue2String(r.Cells[6]));
                curPartner.mCity = CellValue2String(r.Cells[7]);
                curPartner.mStreet = CellValue2String(r.Cells[8]);

                curPartner.mBuildingNum = CellValue2Double(r.Cells[9]);
                curPartner.mAppNum = CellValue2Double(r.Cells[10]);

                curPartner.mPhones = "";
                curPartner.mEMail = "";

                allPertners.Add(curPartner);
            }

            curPartner = new UserInfo();
            curPartner.mFirstName = txtPrivateName.Text.Trim();
            curPartner.mLastName = txtFamilyName.Text.Trim();
            curPartner.mFatherName = txtFatherName.Text.Trim();
            curPartner.mMotherName = txtMotherName.Text.Trim();
            curPartner.mB_Date = DateTimePickerFrom.Value;
            curPartner.mPhones = txtPhones.Text;
            curPartner.mEMail = txtEMail.Text;

            yesno.Add(cmbBusinessBonus.SelectedItem.ToString());
            selffix.Add(cmbSelfFix.SelectedItem.ToString());

            curPartner.mCity = txtCity.Text.Trim();
            curPartner.mStreet = txtStreet.Text.Trim();
            curPartner.mBuildingNum = Convert.ToDouble(txtBiuldingNum.Text);
            curPartner.mAppNum = Convert.ToDouble(txtAppNum.Text);

            allPertners.Add(curPartner);
            #endregion

            double finalMultiResult = 0;
            List<string> personalResults = new List<string>();
            List<double> persoanlBssnsRes = new List<double>();

            ///
            #region Calc Business Mark For Each Person
            for (int i = 0; i < allPertners.Count; i++)
            {
                curPartner = allPertners[i];
                string bonusyesno = yesno[i];
                string bunosselffix = selffix[i];

                txtPrivateName.Text = curPartner.mFirstName;
                txtFamilyName.Text = curPartner.mLastName;
                txtFatherName.Text = curPartner.mFatherName;
                txtMotherName.Text = curPartner.mMotherName;
                DateTimePickerFrom.Value = curPartner.mB_Date;
                txtCity.Text = curPartner.mCity;
                txtStreet.Text = curPartner.mStreet;
                txtBiuldingNum.Text = curPartner.mBuildingNum.ToString();
                txtAppNum.Text = curPartner.mAppNum.ToString();

                DateTimePickerTo.Value = toCalcDay;

                runCalc();

                SingleManualWorkMatchCalc();

                finalMultiResult += Convert.ToDouble(txtFinalMenualWorkMark.Text.Trim());
                persoanlBssnsRes.Add(Convert.ToDouble(txtFinalMenualWorkMark.Text.Trim()));

                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                {
                    personalResults.Add(curPartner.mFirstName + " " + curPartner.mLastName + ", תוצאה עסקית - " + txtFinalMenualWorkMark.Text.Trim() + " , " + txtMultiManualStory.Text);
                }
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                {
                    personalResults.Add(curPartner.mFirstName + " " + curPartner.mLastName + ", Bussiness Index - " + txtFinalMenualWorkMark.Text.Trim() + " , " + txtMultiManualStory.Text);
                }
            }
            #endregion Calc Business Mark For Each Person


            mPartnersManual.Rows.Clear();

            for (int i = 0; i < RowList.Count; i++)
            {
                DataGridViewRow r = RowList[i];
                mPartnersManual.Rows.Add(r);

            }

            mPartnersManual.Update();
            mPartnersManual.Refresh();

            #region Show results

            ReorderBusinessParteners(ref persoanlBssnsRes, ref personalResults);

            for (int i = 0; i < personalResults.Count; i++)
            {
                txtMultiManualStory.Text += System.Environment.NewLine + personalResults[i];
            }

            #endregion Show results

            isMultiSecereteryCalc = false;
        }

        private void button19_Click(object sender, EventArgs e)
        {
            txtHiTechWorkStory.Text = "";
            RunHiTechCalc4cycles();
        }

        private void RunHiTechCalc4cycles()
        {
            int c = CalcCurrentCycle();
            DateTime dt = DateTimePickerTo.Value;
            string t = "";
            for (int i = 1; i < 5; i++)
            {

                txtHiTechWorkStory.Text += Environment.NewLine;


                string[] s = (this.Controls.Find("txt" + i.ToString() + "_1", true)[0] as TextBox).Text.Split("-".ToCharArray());
                DateTimePickerTo.Value = mB_Date.AddYears((Convert.ToInt16(s[0]) + Convert.ToInt16(s[1])) / 2);

                #region intro 4 each cycle
                if (i == 1)
                {
                    if (BLL.AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                    {

                        txtHiTechWorkStory.Text += "במחזור ראשון: ";
                    }
                    if (BLL.AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                    {

                        txtHiTechWorkStory.Text += "First Cycle: ";
                    }
                }
                if (i == 2)
                {
                    if (BLL.AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                    {
                        txtHiTechWorkStory.Text += "במחזור שני: ";
                    }
                    if (BLL.AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                    {

                        txtHiTechWorkStory.Text += "Second Cycle: ";
                    }
                }
                if (i == 3)
                {
                    if (BLL.AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                    {

                        txtHiTechWorkStory.Text += "במחזור שלישי: ";
                    }
                    if (BLL.AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                    {
                        txtHiTechWorkStory.Text += "Third Cycle: ";
                    }
                }
                if (i == 4)
                {
                    if (BLL.AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                    {
                        txtHiTechWorkStory.Text += "במחזור רביעי: ";
                    }
                    if (BLL.AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                    {
                        txtHiTechWorkStory.Text += "Fourth Cycle: ";
                    }
                }
                #endregion intro 4 each cycle
                SingleHiTechMatchCalc();

                txtHiTechWorkStory.Text += Environment.NewLine;


                if (c == i)
                {
                    t = txtFinalHiTechMark.Text;
                }
            }

            #region "employee of the month"
            if ((Convert.ToInt16(Calculator.GetCorrectNumberFromSplitedString(txtNum6.Text.Split(Calculator.Delimiter))) == 8) || (Convert.ToInt16(Calculator.GetCorrectNumberFromSplitedString(txtNum3.Text.Split(Calculator.Delimiter))) == 8))
            {
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                {
                    txtHiTechWorkStory.Text += Environment.NewLine + "יש לך פוטנציאל עובד מצטיין";
                }
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                {
                    txtHiTechWorkStory.Text += Environment.NewLine + "You have the potential of being an excelent employee";
                }
            }
            #endregion

            #region Comment - 7
            int av = Convert.ToInt16(txtAstroName.Text.Split(" ".ToCharArray()[0])[1].Replace("(", "").Replace(")", ""));

            //if ((Calculator.GetCorrectNumberFromSplitedString(txtNum3.Text.Split(Calculator.Delimiter)) == 7) ||
            //     (Calculator.GetCorrectNumberFromSplitedString(txtNum5.Text.Split(Calculator.Delimiter)) == 7) ||
            //     (av == 7))
            if (new int[] { 2, 7, 13, 14, 16, 19 }.Contains(Calculator.GetCorrectNumberFromSplitedString(txtNum3.Text.Split(Calculator.Delimiter))))
            {
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                {
                    txtHiTechWorkStory.Text += "יש לך פוטנציאל עובד מצטיין";
                    //"יכולות החשיבה שלך מונעות ממך למצות את הפוטנציאל הגלום בך להצלחה גבוהה יותר. מומלץ לך להוסיף שם על מנת לחדד ולשפר את יכולתך העסקית דבר שיבוא לידי ביטויו בפן העסקי והכלכלי. ";
                }
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                {
                    txtHiTechWorkStory.Text += "Though your thinking abilities prevent the accomplishment of your potential to greater success, it is recommended for you to add a name in order to refine your financial strength.";
                }
            }
            #endregion Comment - 7


            DateTimePickerTo.Value = dt;

            txtFinalHiTechMark.Text = t;
            txtHiTechWorkStory.Text = txtHiTechWorkStory.Text.Trim();

        }

        private void SingleHiTechMatchCalc()
        {
            List<string> inBusinessValues = new List<string>();
            int age = Convert.ToInt16(txtAge.Text.Split("(".ToCharArray()[0])[0]);
            age = Convert.ToInt16(DateTimePickerTo.Value.Subtract(mB_Date).TotalDays / 365);

            #region From Chakra Map
            int upperBndry = Convert.ToInt16(txt1_1.Text.Split("-".ToCharArray()[0])[1].Trim());

            int tmpVal = 0;
            if (age <= upperBndry)
            {
                // Crown
                tmpVal = Calculator.GetCorrectNumberFromSplitedString(txtNum1.Text.Split(Calculator.Delimiter));
            }
            else
            {
                //Third Eye
                tmpVal = Calculator.GetCorrectNumberFromSplitedString(txtNum2.Text.Split(Calculator.Delimiter));
            }
            inBusinessValues.Add(tmpVal.ToString());


            tmpVal = Calculator.GetCorrectNumberFromSplitedString(txtNum3.Text.Split(Calculator.Delimiter));
            inBusinessValues.Add(tmpVal.ToString());

            tmpVal = Calculator.GetCorrectNumberFromSplitedString(txtNum5.Text.Split(Calculator.Delimiter));
            inBusinessValues.Add(tmpVal.ToString());

            tmpVal = Calculator.GetCorrectNumberFromSplitedString(txtNum6.Text.Split(Calculator.Delimiter));
            inBusinessValues.Add(tmpVal.ToString());

            tmpVal = Calculator.GetCorrectNumberFromSplitedString(txtPName_Num.Text.Split(Calculator.Delimiter));
            inBusinessValues.Add(tmpVal.ToString());

            tmpVal = Convert.ToInt16(txtAstroName.Text.Split("(".ToCharArray()[0])[1].Split(")".ToCharArray()[0])[0]);
            int tmpastronum = tmpVal;
            inBusinessValues.Add(tmpVal.ToString());

            string tmpLvl = txtInfo1.Text.Split(System.Environment.NewLine.ToCharArray()[0])[1].Trim();
            if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
            {
            }
            if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
            {
                if (tmpLvl == "Balanced")
                {
                    tmpLvl = "מאוזן";
                }
                if (tmpLvl == "Not Balanced")
                {
                    tmpLvl = "לא מאוזן";
                }
                if (tmpLvl == "Half Balanced")
                {
                    tmpLvl = "חצי מאוזן";
                }
            }
            inBusinessValues.Add(tmpLvl);
            #endregion
            if (isMultiSecereteryCalc == true)
            {
                for (int i = 0; i < inBusinessValues.Count; i++)
                {
                    if (inBusinessValues[i] == "2")
                    {
                        // fix to 7 in the business Chakra chart
                        inBusinessValues.RemoveAt(i);
                        inBusinessValues.Insert(i, "33"); // fir grade 7
                    }
                    if (inBusinessValues[i] == "4")
                    {
                        // fix to 7 in the business Chakra chart
                        inBusinessValues.RemoveAt(i);
                        inBusinessValues.Insert(i, "3");// fir grade 8
                    }
                }
            }

            double ChakraMapBusinessValue = Calculator.ChakraMap2HiTechValue(inBusinessValues);

            Control.ControlCollection cntrlCls = this.Controls;

            #region Life Cycles
            int curCycle = 0;
            int[] CycleBoundries = new int[4];
            CycleBoundries[0] = Convert.ToInt16(txt1_1.Text.Split("-".ToCharArray()[0])[1].Trim());
            CycleBoundries[1] = Convert.ToInt16(txt2_1.Text.Split("-".ToCharArray()[0])[1].Trim());
            CycleBoundries[2] = Convert.ToInt16(txt3_1.Text.Split("-".ToCharArray()[0])[1].Trim());
            CycleBoundries[3] = Convert.ToInt16(txt4_1.Text.Split("-".ToCharArray()[0])[1].Trim());

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

            int chlng, clmx, ccl;
            chlng = Calculator.GetCorrectNumberFromSplitedString(cntrlCls.Find("txt" + curCycle.ToString() + "_4", true)[0].Text.Split(Calculator.Delimiter));
            clmx = Calculator.GetCorrectNumberFromSplitedString(cntrlCls.Find("txt" + curCycle.ToString() + "_3", true)[0].Text.Split(Calculator.Delimiter));
            ccl = Calculator.GetCorrectNumberFromSplitedString(cntrlCls.Find("txt" + curCycle.ToString() + "_2", true)[0].Text.Split(Calculator.Delimiter));

            inBusinessValues = new List<string>();
            //inBusinessValues.Add(chlng.ToString());
            inBusinessValues.Add(clmx.ToString());
            inBusinessValues.Add(ccl.ToString());
            #endregion Life Cycles
            if (isMultiSecereteryCalc == true)
            {
                for (int i = 0; i < inBusinessValues.Count; i++)
                {
                    if (inBusinessValues[i] == "2")
                    {
                        // fix to 7 in the business Chakra chart
                        inBusinessValues.RemoveAt(i);
                        inBusinessValues.Insert(i, "33");// for grade 7
                    }
                    if (inBusinessValues[i] == "4")
                    {
                        // fix to 7 in the business Chakra chart
                        inBusinessValues.RemoveAt(i);
                        inBusinessValues.Insert(i, "3");// for grade 8
                    }
                }
            }
            //לפי הטבלאות
            //inBusinessValues = ערך נומורלוגי
            //LifeCycleBusinessValue =ערך תעסוקתי
            double LifeCycleBusinessValue = Calculator.LifeCycle2HiTechValue(inBusinessValues);

            double sum = (ChakraMapBusinessValue + LifeCycleBusinessValue) / 2;

            //if (((AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew) && (BonusYesOrNo == "כן")) || ((AppSettings.Instance.ProgramLanguage == AppSettings.Language.English) && (BonusYesOrNo == "Yes")))
            //{
            //    sum += 1;
            //}
            //else // cmbBusinessBonus.SelectedItem.ToString() == "לא"
            //{
            //    sum += 0;
            //}

            //Additional points
            // If sex & creation equals  1, 3, 5, 8, plus 1 point as bonus
            int[] sexCreationNumbers = { 1, 3, 5, 8 };
            if (sexCreationNumbers.Contains(Calculator.GetCorrectNumberFromSplitedString(txtNum6.Text.Split(Calculator.Delimiter))))
            {
                sum += 1;
            }

            // If throat equals 1, 3, 5, 7, 8, 9, 16  plus 0.5 point as bonus
            int[] throatNumbers = { 1, 3, 5, 7, 8, 9, 16 };
            if (throatNumbers.Contains(Calculator.GetCorrectNumberFromSplitedString(txtNum3.Text.Split(Calculator.Delimiter))))
            {
                sum += 0.5;
            }

            // If throat equals 2, 4, 6, 13, 14, 19, 33  minus 0.3 point as bonus
            int[] throat2Numbers = { 2, 4, 6, 13, 14, 19, 33 };
            if (throat2Numbers.Contains(Calculator.GetCorrectNumberFromSplitedString(txtNum3.Text.Split(Calculator.Delimiter))))
            {
                sum -= 0.3;
            }

            // If sex & creation equals 2, 4, 6, 11, 22, 33 minus 1 point as bonus
            int[] sexCreation2Numbers = { 2, 4, 6, 11, 22, 33 };
            if (sexCreation2Numbers.Contains(Calculator.GetCorrectNumberFromSplitedString(txtNum6.Text.Split(Calculator.Delimiter))))
            {
                sum -= 1;
            }

            //צאקרת העל או עין שלישית
            int[] eyeCreationNumbers = { 1, 3, 5, 8, 9, 11, 22, 14, 19 };
            if (eyeCreationNumbers.Contains(Calculator.GetCorrectNumberFromSplitedString(txtNum8.Text.Split(Calculator.Delimiter))) ||
                eyeCreationNumbers.Contains(Calculator.GetCorrectNumberFromSplitedString(txtNum2.Text.Split(Calculator.Delimiter))))
            {
                sum -= 0.4;
            }

            // If universe chakra & creation equals   1, 3, 5, 7, 8, 9, 22
            int[] universeNumbers = { 1, 3, 5, 7, 8, 9, 22 };
            int universeNumber = Calculator.GetUniverseNumberFromChakra(txtNum9.Text);
            if (universeNumbers.Contains(universeNumber))
            {
                sum += 0.5;
            }

            #region Additional Msala
            if (isNeedToCalculateWithMsala)
            {
                var isFound = false;
                for (var i = 1; i < 5; i++)
                {
                    if (Calculator.GetCorrectNumberFromSplitedString(cntrlCls.Find("txtX" + i + "_2", true)[0].Text.Split(Calculator.Delimiter)) == 11 ||
                        Calculator.GetCorrectNumberFromSplitedString(cntrlCls.Find("txtX" + i + "_2", true)[0].Text.Split(Calculator.Delimiter)) == 22)
                    {
                        sum += 0.5;
                        isFound = true;
                        break;
                    }
                }
                if (!isFound)
                {
                    for (var i = 1; i < 5; i++)
                    {
                        if (Calculator.GetCorrectNumberFromSplitedString(cntrlCls.Find("txtX" + i + "_3", true)[0].Text.Split(Calculator.Delimiter)) == 11 ||
                        Calculator.GetCorrectNumberFromSplitedString(cntrlCls.Find("txtX" + i + "_3", true)[0].Text.Split(Calculator.Delimiter)) == 22)
                        {
                            sum += 0.5;
                            isFound = true;
                            break;
                        }
                    }
                }
            }
            #endregion

            if (sum > 10)
            {
                sum = 10;
            }

            txtFinalHiTechMark.Text = Math.Round(sum, 2).ToString();

            string story = "";
            if ((sum >= 9) & (sum <= 10))
            {
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                {
                    story = "התאמה טובה מאוד";
                }
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                {
                    story = "Very good match";
                }
            }

            if ((sum >= 8) & (sum < 9))
            {
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                {
                    story = "התאמה טובה";
                }
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                {
                    story = "Good match";
                }
            }

            if ((sum >= 0) & (sum < 8))
            {
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                {
                    story = "התאמה חלשה";
                }
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                {
                    story = "Poor match";
                }
            }

            if (tmpastronum == 22)
            {
                //if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                //{
                //    story += Environment.NewLine + "היות ומזלך האסטרולוגי עקרב, קל לך להיות עובד מוערך.";
                //}
                //if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                //{
                //    story += Environment.NewLine + "As astrological sign is Scorpio it is easy for you to be a successful indipendent bussiness or a highly estimated worker.";
                //}
            }

            txtHiTechWorkStory.Text += "(" + txtFinalHiTechMark.Text + ") " + Environment.NewLine + story;
        }

        private void button20_Click(object sender, EventArgs e)
        {
            //Run single
            txtHiTechWorkStory.Text = "";
            RunHiTechCalc4cycles();

            //Run multi
            isMultiSecereteryCalc = true;

            List<UserInfo> allPertners = new List<UserInfo>();
            List<string> yesno = new List<string>();
            List<string> selffix = new List<string>();
            List<DataGridViewRow> RowList = new List<DataGridViewRow>();
            DateTime toCalcDay = DateTimePickerTo.Value;

            #region Gather Data From DataGridView
            UserInfo curPartner;
            for (int i = 0; i < mPartnersHiTech.Rows.Count - 1; i++)
            {
                DataGridViewRow r = mPartnersHiTech.Rows[i];
                RowList.Add(r);

                curPartner = new UserInfo();

                curPartner.mFirstName = CellValue2String(r.Cells[0]);
                curPartner.mLastName = CellValue2String(r.Cells[1]);
                curPartner.mFatherName = CellValue2String(r.Cells[3]);
                curPartner.mMotherName = CellValue2String(r.Cells[4]);
                bool resDate = DateTime.TryParse(r.Cells[2].Value.ToString().Trim(), out curPartner.mB_Date);
                if (resDate == false)
                {
                    MessageBox.Show("שגיאה בכתיבת תאריך", "שגיאת נתונים", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                yesno.Add(CellValue2String(r.Cells[5]));
                selffix.Add(CellValue2String(r.Cells[6]));
                curPartner.mCity = CellValue2String(r.Cells[7]);
                curPartner.mStreet = CellValue2String(r.Cells[8]);

                curPartner.mBuildingNum = CellValue2Double(r.Cells[9]);
                curPartner.mAppNum = CellValue2Double(r.Cells[10]);

                curPartner.mPhones = "";
                curPartner.mEMail = "";

                allPertners.Add(curPartner);
            }

            curPartner = new UserInfo();
            curPartner.mFirstName = txtPrivateName.Text.Trim();
            curPartner.mLastName = txtFamilyName.Text.Trim();
            curPartner.mFatherName = txtFatherName.Text.Trim();
            curPartner.mMotherName = txtMotherName.Text.Trim();
            curPartner.mB_Date = DateTimePickerFrom.Value;
            curPartner.mPhones = txtPhones.Text;
            curPartner.mEMail = txtEMail.Text;

            yesno.Add(cmbBusinessBonus.SelectedItem.ToString());
            selffix.Add(cmbSelfFix.SelectedItem.ToString());

            curPartner.mCity = txtCity.Text.Trim();
            curPartner.mStreet = txtStreet.Text.Trim();
            curPartner.mBuildingNum = Convert.ToDouble(txtBiuldingNum.Text);
            curPartner.mAppNum = Convert.ToDouble(txtAppNum.Text);

            allPertners.Add(curPartner);
            #endregion

            double finalMultiResult = 0;
            List<string> personalResults = new List<string>();
            List<double> persoanlBssnsRes = new List<double>();

            ///
            #region Calc Business Mark For Each Person
            for (int i = 0; i < allPertners.Count; i++)
            {
                curPartner = allPertners[i];
                string bonusyesno = yesno[i];
                string bunosselffix = selffix[i];

                txtPrivateName.Text = curPartner.mFirstName;
                txtFamilyName.Text = curPartner.mLastName;
                txtFatherName.Text = curPartner.mFatherName;
                txtMotherName.Text = curPartner.mMotherName;
                DateTimePickerFrom.Value = curPartner.mB_Date;
                txtCity.Text = curPartner.mCity;
                txtStreet.Text = curPartner.mStreet;
                txtBiuldingNum.Text = curPartner.mBuildingNum.ToString();
                txtAppNum.Text = curPartner.mAppNum.ToString();

                DateTimePickerTo.Value = toCalcDay;

                runCalc();

                SingleHiTechMatchCalc();

                finalMultiResult += Convert.ToDouble(txtFinalHiTechMark.Text.Trim());
                persoanlBssnsRes.Add(Convert.ToDouble(txtFinalHiTechMark.Text.Trim()));

                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                {
                    personalResults.Add(curPartner.mFirstName + " " + curPartner.mLastName + ", תוצאה עסקית - " + txtFinalHiTechMark.Text.Trim() + " , " + txtMultiHiTechStory.Text);
                }
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                {
                    personalResults.Add(curPartner.mFirstName + " " + curPartner.mLastName + ", Bussiness Index - " + txtFinalHiTechMark.Text.Trim() + " , " + txtMultiHiTechStory.Text);
                }
            }
            #endregion Calc Business Mark For Each Person


            mPartnersHiTech.Rows.Clear();

            for (int i = 0; i < RowList.Count; i++)
            {
                DataGridViewRow r = RowList[i];
                mPartnersHiTech.Rows.Add(r);

            }

            mPartnersHiTech.Update();
            mPartnersHiTech.Refresh();

            #region Show results

            ReorderBusinessParteners(ref persoanlBssnsRes, ref personalResults);

            for (int i = 0; i < personalResults.Count; i++)
            {
                txtMultiHiTechStory.Text += System.Environment.NewLine + personalResults[i];
            }

            #endregion Show results

            isMultiSecereteryCalc = false;
        }

        private void button21_Click(object sender, EventArgs e)
        {
            txtLowTechWorkStory.Text = "";
            RunLowTechCalc4cycles();
        }

        private void RunLowTechCalc4cycles()
        {
            int c = CalcCurrentCycle();
            DateTime dt = DateTimePickerTo.Value;
            string t = "";
            for (int i = 1; i < 5; i++)
            {

                txtLowTechWorkStory.Text += Environment.NewLine;


                string[] s = (this.Controls.Find("txt" + i.ToString() + "_1", true)[0] as TextBox).Text.Split("-".ToCharArray());
                DateTimePickerTo.Value = mB_Date.AddYears((Convert.ToInt16(s[0]) + Convert.ToInt16(s[1])) / 2);

                #region intro 4 each cycle
                if (i == 1)
                {
                    if (BLL.AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                    {

                        txtLowTechWorkStory.Text += "במחזור ראשון: ";
                    }
                    if (BLL.AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                    {

                        txtLowTechWorkStory.Text += "First Cycle: ";
                    }
                }
                if (i == 2)
                {
                    if (BLL.AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                    {
                        txtLowTechWorkStory.Text += "במחזור שני: ";
                    }
                    if (BLL.AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                    {

                        txtLowTechWorkStory.Text += "Second Cycle: ";
                    }
                }
                if (i == 3)
                {
                    if (BLL.AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                    {

                        txtLowTechWorkStory.Text += "במחזור שלישי: ";
                    }
                    if (BLL.AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                    {
                        txtLowTechWorkStory.Text += "Third Cycle: ";
                    }
                }
                if (i == 4)
                {
                    if (BLL.AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                    {
                        txtLowTechWorkStory.Text += "במחזור רביעי: ";
                    }
                    if (BLL.AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                    {
                        txtLowTechWorkStory.Text += "Fourth Cycle: ";
                    }
                }
                #endregion intro 4 each cycle
                SingleLowTechMatchCalc();

                txtLowTechWorkStory.Text += Environment.NewLine;


                if (c == i)
                {
                    t = txtFinalLowTechMark.Text;
                }
            }

            #region "employee of the month"
            if ((Convert.ToInt16(Calculator.GetCorrectNumberFromSplitedString(txtNum6.Text.Split(Calculator.Delimiter))) == 8) || (Convert.ToInt16(Calculator.GetCorrectNumberFromSplitedString(txtNum3.Text.Split(Calculator.Delimiter))) == 8))
            {
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                {
                    txtLowTechWorkStory.Text += Environment.NewLine + "יש לך פוטנציאל עובד מצטיין";
                }
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                {
                    txtLowTechWorkStory.Text += Environment.NewLine + "You have the potential of being an excelent employee";
                }
            }
            #endregion

            #region Comment - 7
            int av = Convert.ToInt16(txtAstroName.Text.Split(" ".ToCharArray()[0])[1].Replace("(", "").Replace(")", ""));

            //if ((Calculator.GetCorrectNumberFromSplitedString(txtNum3.Text.Split(Calculator.Delimiter)) == 7) ||
            //     (Calculator.GetCorrectNumberFromSplitedString(txtNum5.Text.Split(Calculator.Delimiter)) == 7) ||
            //     (av == 7))
            if (new int[] { 2, 7, 13, 14, 16, 19 }.Contains(Calculator.GetCorrectNumberFromSplitedString(txtNum3.Text.Split(Calculator.Delimiter))))
            {
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                {
                    txtLowTechWorkStory.Text += "יש לך פוטנציאל עובד מצטיין";
                    //"יכולות החשיבה שלך מונעות ממך למצות את הפוטנציאל הגלום בך להצלחה גבוהה יותר. מומלץ לך להוסיף שם על מנת לחדד ולשפר את יכולתך העסקית דבר שיבוא לידי ביטויו בפן העסקי והכלכלי. ";
                }
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                {
                    txtLowTechWorkStory.Text += "Though your thinking abilities prevent the accomplishment of your potential to greater success, it is recommended for you to add a name in order to refine your financial strength.";
                }
            }
            #endregion Comment - 7


            DateTimePickerTo.Value = dt;

            txtFinalLowTechMark.Text = t;
            txtLowTechWorkStory.Text = txtLowTechWorkStory.Text.Trim();

        }

        private void SingleLowTechMatchCalc()
        {
            List<string> inBusinessValues = new List<string>();
            int age = Convert.ToInt16(txtAge.Text.Split("(".ToCharArray()[0])[0]);
            age = Convert.ToInt16(DateTimePickerTo.Value.Subtract(mB_Date).TotalDays / 365);

            #region From Chakra Map
            int upperBndry = Convert.ToInt16(txt1_1.Text.Split("-".ToCharArray()[0])[1].Trim());

            int tmpVal = 0;
            if (age <= upperBndry)
            {
                // Crown
                tmpVal = Calculator.GetCorrectNumberFromSplitedString(txtNum1.Text.Split(Calculator.Delimiter));
            }
            else
            {
                //Third Eye
                tmpVal = Calculator.GetCorrectNumberFromSplitedString(txtNum2.Text.Split(Calculator.Delimiter));
            }
            inBusinessValues.Add(tmpVal.ToString());


            tmpVal = Calculator.GetCorrectNumberFromSplitedString(txtNum3.Text.Split(Calculator.Delimiter));
            inBusinessValues.Add(tmpVal.ToString());

            tmpVal = Calculator.GetCorrectNumberFromSplitedString(txtNum5.Text.Split(Calculator.Delimiter));
            inBusinessValues.Add(tmpVal.ToString());

            tmpVal = Calculator.GetCorrectNumberFromSplitedString(txtNum6.Text.Split(Calculator.Delimiter));
            inBusinessValues.Add(tmpVal.ToString());

            tmpVal = Calculator.GetCorrectNumberFromSplitedString(txtPName_Num.Text.Split(Calculator.Delimiter));
            inBusinessValues.Add(tmpVal.ToString());

            tmpVal = Convert.ToInt16(txtAstroName.Text.Split("(".ToCharArray()[0])[1].Split(")".ToCharArray()[0])[0]);
            int tmpastronum = tmpVal;
            inBusinessValues.Add(tmpVal.ToString());

            string tmpLvl = txtInfo1.Text.Split(System.Environment.NewLine.ToCharArray()[0])[1].Trim();
            if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
            {
            }
            if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
            {
                if (tmpLvl == "Balanced")
                {
                    tmpLvl = "מאוזן";
                }
                if (tmpLvl == "Not Balanced")
                {
                    tmpLvl = "לא מאוזן";
                }
                if (tmpLvl == "Half Balanced")
                {
                    tmpLvl = "חצי מאוזן";
                }
            }
            inBusinessValues.Add(tmpLvl);
            #endregion
            if (isMultiSecereteryCalc == true)
            {
                for (int i = 0; i < inBusinessValues.Count; i++)
                {
                    if (inBusinessValues[i] == "2")
                    {
                        // fix to 7 in the business Chakra chart
                        inBusinessValues.RemoveAt(i);
                        inBusinessValues.Insert(i, "33"); // fir grade 7
                    }
                    if (inBusinessValues[i] == "4")
                    {
                        // fix to 7 in the business Chakra chart
                        inBusinessValues.RemoveAt(i);
                        inBusinessValues.Insert(i, "3");// fir grade 8
                    }
                }
            }

            double ChakraMapBusinessValue = Calculator.ChakraMap2LowTechValue(inBusinessValues);

            Control.ControlCollection cntrlCls = this.Controls;

            #region Life Cycles
            int curCycle = 0;
            int[] CycleBoundries = new int[4];
            CycleBoundries[0] = Convert.ToInt16(txt1_1.Text.Split("-".ToCharArray()[0])[1].Trim());
            CycleBoundries[1] = Convert.ToInt16(txt2_1.Text.Split("-".ToCharArray()[0])[1].Trim());
            CycleBoundries[2] = Convert.ToInt16(txt3_1.Text.Split("-".ToCharArray()[0])[1].Trim());
            CycleBoundries[3] = Convert.ToInt16(txt4_1.Text.Split("-".ToCharArray()[0])[1].Trim());

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

            int chlng, clmx, ccl;
            chlng = Calculator.GetCorrectNumberFromSplitedString(cntrlCls.Find("txt" + curCycle.ToString() + "_4", true)[0].Text.Split(Calculator.Delimiter));
            clmx = Calculator.GetCorrectNumberFromSplitedString(cntrlCls.Find("txt" + curCycle.ToString() + "_3", true)[0].Text.Split(Calculator.Delimiter));
            ccl = Calculator.GetCorrectNumberFromSplitedString(cntrlCls.Find("txt" + curCycle.ToString() + "_2", true)[0].Text.Split(Calculator.Delimiter));

            inBusinessValues = new List<string>();
            //inBusinessValues.Add(chlng.ToString());
            inBusinessValues.Add(clmx.ToString());
            inBusinessValues.Add(ccl.ToString());
            #endregion Life Cycles
            if (isMultiSecereteryCalc == true)
            {
                for (int i = 0; i < inBusinessValues.Count; i++)
                {
                    if (inBusinessValues[i] == "2")
                    {
                        // fix to 7 in the business Chakra chart
                        inBusinessValues.RemoveAt(i);
                        inBusinessValues.Insert(i, "33");// for grade 7
                    }
                    if (inBusinessValues[i] == "4")
                    {
                        // fix to 7 in the business Chakra chart
                        inBusinessValues.RemoveAt(i);
                        inBusinessValues.Insert(i, "3");// for grade 8
                    }
                }
            }
            //לפי הטבלאות
            //inBusinessValues = ערך נומורלוגי
            //LifeCycleBusinessValue =ערך תעסוקתי
            double LifeCycleBusinessValue = Calculator.LifeCycle2LowTechValue(inBusinessValues);

            double sum = (ChakraMapBusinessValue + LifeCycleBusinessValue) / 2;

            //if (((AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew) && (BonusYesOrNo == "כן")) || ((AppSettings.Instance.ProgramLanguage == AppSettings.Language.English) && (BonusYesOrNo == "Yes")))
            //{
            //    sum += 1;
            //}
            //else // cmbBusinessBonus.SelectedItem.ToString() == "לא"
            //{
            //    sum += 0;
            //}

            //Additional points
            // If sex & creation equals  1, 8, minus 1 point as bonus
            int[] sexCreationNumbers = { 1, 8 };
            if (sexCreationNumbers.Contains(Calculator.GetCorrectNumberFromSplitedString(txtNum6.Text.Split(Calculator.Delimiter))))
            {
                sum -= 1;
            }

            // If throat equals 0, 2, 4, 6, 11, 22, 33  plus 0.5 point as bonus
            int[] throatNumbers = { 0, 2, 4, 6, 11, 22, 33 };
            if (throatNumbers.Contains(Calculator.GetCorrectNumberFromSplitedString(txtNum3.Text.Split(Calculator.Delimiter))))
            {
                sum += 0.5;
            }

            // If throat equals 7, 13, 14, 16, 19  minus 0.5 point as bonus
            int[] throat2Numbers = { 7, 13, 14, 16, 19 };
            if (throat2Numbers.Contains(Calculator.GetCorrectNumberFromSplitedString(txtNum3.Text.Split(Calculator.Delimiter))))
            {
                sum -= 0.5;
            }

            // If sex & creation equals 3, 5 minus 0.5 point as bonus
            int[] sexCreation2Numbers = { 3, 5 };
            if (sexCreation2Numbers.Contains(Calculator.GetCorrectNumberFromSplitedString(txtNum6.Text.Split(Calculator.Delimiter))))
            {
                sum -= 0.5;
            }

            // If sex & creation equals 2, 4, 6, 7, 9, 11, 22, 33, 13, 14, 16, 19 plus 0.5 point as bonus
            int[] sexCreation3Numbers = { 2, 4, 6, 7, 9, 11, 22, 33, 13, 14, 16, 19 };
            if (sexCreation3Numbers.Contains(Calculator.GetCorrectNumberFromSplitedString(txtNum6.Text.Split(Calculator.Delimiter))))
            {
                sum += 0.5;
            }

            // If universe chakra & creation equals   1, 3, 5, 8, 9, 22
            int[] universeNumbers = { 1, 3, 5, 8, 9, 22 };
            int universeNumber = Calculator.GetUniverseNumberFromChakra(txtNum9.Text);
            if (universeNumbers.Contains(universeNumber))
            {
                sum -= 0.5;
            }

            // If sun chakra & creation equals   7, 13, 14, 16, 19
            int[] sunNumbers = { 7, 13, 14, 16, 19 };
            if (sunNumbers.Contains(Calculator.GetCorrectNumberFromSplitedString(txtNum5.Text.Split(Calculator.Delimiter))))
            {
                sum -= 0.5;
            }

            #region Additional Msala
            if (isNeedToCalculateWithMsala)
            {
                var isFound = false;
                for (var i = 1; i < 5; i++)
                {
                    if (Calculator.GetCorrectNumberFromSplitedString(cntrlCls.Find("txtX" + i + "_2", true)[0].Text.Split(Calculator.Delimiter)) == 11 ||
                        Calculator.GetCorrectNumberFromSplitedString(cntrlCls.Find("txtX" + i + "_2", true)[0].Text.Split(Calculator.Delimiter)) == 22)
                    {
                        sum += 0.5;
                        isFound = true;
                        break;
                    }
                }
                if (!isFound)
                {
                    for (var i = 1; i < 5; i++)
                    {
                        if (Calculator.GetCorrectNumberFromSplitedString(cntrlCls.Find("txtX" + i + "_3", true)[0].Text.Split(Calculator.Delimiter)) == 11 ||
                        Calculator.GetCorrectNumberFromSplitedString(cntrlCls.Find("txtX" + i + "_3", true)[0].Text.Split(Calculator.Delimiter)) == 22)
                        {
                            sum += 0.5;
                            isFound = true;
                            break;
                        }
                    }
                }
            }
            #endregion

            if (sum > 10)
            {
                sum = 10;
            }

            txtFinalLowTechMark.Text = Math.Round(sum, 2).ToString();

            string story = "";
            if ((sum >= 9) & (sum <= 10))
            {
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                {
                    story = "התאמה טובה מאוד";
                }
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                {
                    story = "Very good match";
                }
            }

            if ((sum >= 8) & (sum < 9))
            {
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                {
                    story = "התאמה טובה";
                }
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                {
                    story = "Good match";
                }
            }

            if ((sum >= 0) & (sum < 8))
            {
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                {
                    story = "התאמה חלשה";
                }
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                {
                    story = "Poor match";
                }
            }

            if (tmpastronum == 22)
            {
                //if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                //{
                //    story += Environment.NewLine + "היות ומזלך האסטרולוגי עקרב, קל לך להיות עובד מוערך.";
                //}
                //if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                //{
                //    story += Environment.NewLine + "As astrological sign is Scorpio it is easy for you to be a successful indipendent bussiness or a highly estimated worker.";
                //}
            }

            txtLowTechWorkStory.Text += "(" + txtFinalLowTechMark.Text + ") " + Environment.NewLine + story;
        }

        private void button22_Click(object sender, EventArgs e)
        {
            //Run single
            txtHiTechWorkStory.Text = "";
            RunLowTechCalc4cycles();

            //Run multi
            isMultiSecereteryCalc = true;

            List<UserInfo> allPertners = new List<UserInfo>();
            List<string> yesno = new List<string>();
            List<string> selffix = new List<string>();
            List<DataGridViewRow> RowList = new List<DataGridViewRow>();
            DateTime toCalcDay = DateTimePickerTo.Value;

            #region Gather Data From DataGridView
            UserInfo curPartner;
            for (int i = 0; i < mPartnersLowTech.Rows.Count - 1; i++)
            {
                DataGridViewRow r = mPartnersLowTech.Rows[i];
                RowList.Add(r);

                curPartner = new UserInfo();

                curPartner.mFirstName = CellValue2String(r.Cells[0]);
                curPartner.mLastName = CellValue2String(r.Cells[1]);
                curPartner.mFatherName = CellValue2String(r.Cells[3]);
                curPartner.mMotherName = CellValue2String(r.Cells[4]);
                bool resDate = DateTime.TryParse(r.Cells[2].Value.ToString().Trim(), out curPartner.mB_Date);
                if (resDate == false)
                {
                    MessageBox.Show("שגיאה בכתיבת תאריך", "שגיאת נתונים", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                yesno.Add(CellValue2String(r.Cells[5]));
                selffix.Add(CellValue2String(r.Cells[6]));
                curPartner.mCity = CellValue2String(r.Cells[7]);
                curPartner.mStreet = CellValue2String(r.Cells[8]);

                curPartner.mBuildingNum = CellValue2Double(r.Cells[9]);
                curPartner.mAppNum = CellValue2Double(r.Cells[10]);

                curPartner.mPhones = "";
                curPartner.mEMail = "";

                allPertners.Add(curPartner);
            }

            curPartner = new UserInfo();
            curPartner.mFirstName = txtPrivateName.Text.Trim();
            curPartner.mLastName = txtFamilyName.Text.Trim();
            curPartner.mFatherName = txtFatherName.Text.Trim();
            curPartner.mMotherName = txtMotherName.Text.Trim();
            curPartner.mB_Date = DateTimePickerFrom.Value;
            curPartner.mPhones = txtPhones.Text;
            curPartner.mEMail = txtEMail.Text;

            yesno.Add(cmbBusinessBonus.SelectedItem.ToString());
            selffix.Add(cmbSelfFix.SelectedItem.ToString());

            curPartner.mCity = txtCity.Text.Trim();
            curPartner.mStreet = txtStreet.Text.Trim();
            curPartner.mBuildingNum = Convert.ToDouble(txtBiuldingNum.Text);
            curPartner.mAppNum = Convert.ToDouble(txtAppNum.Text);

            allPertners.Add(curPartner);
            #endregion

            double finalMultiResult = 0;
            List<string> personalResults = new List<string>();
            List<double> persoanlBssnsRes = new List<double>();

            ///
            #region Calc Business Mark For Each Person
            for (int i = 0; i < allPertners.Count; i++)
            {
                curPartner = allPertners[i];
                string bonusyesno = yesno[i];
                string bunosselffix = selffix[i];

                txtPrivateName.Text = curPartner.mFirstName;
                txtFamilyName.Text = curPartner.mLastName;
                txtFatherName.Text = curPartner.mFatherName;
                txtMotherName.Text = curPartner.mMotherName;
                DateTimePickerFrom.Value = curPartner.mB_Date;
                txtCity.Text = curPartner.mCity;
                txtStreet.Text = curPartner.mStreet;
                txtBiuldingNum.Text = curPartner.mBuildingNum.ToString();
                txtAppNum.Text = curPartner.mAppNum.ToString();

                DateTimePickerTo.Value = toCalcDay;

                runCalc();

                SingleHiTechMatchCalc();

                finalMultiResult += Convert.ToDouble(txtFinalLowTechMark.Text.Trim());
                persoanlBssnsRes.Add(Convert.ToDouble(txtFinalLowTechMark.Text.Trim()));

                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                {
                    personalResults.Add(curPartner.mFirstName + " " + curPartner.mLastName + ", תוצאה עסקית - " + txtFinalLowTechMark.Text.Trim() + " , " + txtMultiLowTechStory.Text);
                }
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                {
                    personalResults.Add(curPartner.mFirstName + " " + curPartner.mLastName + ", Bussiness Index - " + txtFinalLowTechMark.Text.Trim() + " , " + txtMultiLowTechStory.Text);
                }
            }
            #endregion Calc Business Mark For Each Person


            mPartnersLowTech.Rows.Clear();

            for (int i = 0; i < RowList.Count; i++)
            {
                DataGridViewRow r = RowList[i];
                mPartnersLowTech.Rows.Add(r);

            }

            mPartnersLowTech.Update();
            mPartnersLowTech.Refresh();

            #region Show results

            ReorderBusinessParteners(ref persoanlBssnsRes, ref personalResults);

            for (int i = 0; i < personalResults.Count; i++)
            {
                txtMultiLowTechStory.Text += System.Environment.NewLine + personalResults[i];
            }

            #endregion Show results

            isMultiSecereteryCalc = false;
        }


        // **********

        public Omega.Objects.UserInfo getSpouceInfo()
        {
            return curPartner;
        }

    }

    /*
    internal class WordWriter
    {
        #region Data Members
        private object mMissing = System.Reflection.Missing.Value;
        private object mSavePath;
        private object mTemplet; //= (object)Path.Combine(mApplicationMainDir, @"Templets\ThisTemplet.dot"); // need to find way 2 insert a templet here...
        private object oTrue = true as object;
        private object oFalse = false as object;


        private Microsoft.Office.Interop.Word.ApplicationClass mWordApp = new Microsoft.Office.Interop.Word.ApplicationClass();
        private Microsoft.Office.Interop.Word.Document mCurDoc;

        private Microsoft.Office.Interop.Word.Font mHeaderStyle = new Microsoft.Office.Interop.Word.FontClass();
        private Microsoft.Office.Interop.Word.Font mBodyStyle = new Microsoft.Office.Interop.Word.FontClass();

        private string mTempDirPath;// = System.IO.Path.Combine(System.Environment.CurrentDirectory, "TempDir");
        private string mFullName;

        private int picCount = 0;
        private Range mCurRange;
        private object mCurLocation;
        #endregion 

        #region Constructor
            public WordWriter(string ApplicationMainDir, string sDocPath, string firstname, string lastname)
            {
                mTemplet = (object)Path.Combine(ApplicationMainDir, @"Templets\ThisTemplet.dot");
                mTempDirPath = System.IO.Path.Combine(ApplicationMainDir, "TempDir");

                mFullName = firstname + " " + lastname;

                mSavePath = sDocPath as object;

                DirectoryInfo dir = new DirectoryInfo(System.IO.Path.GetDirectoryName(mSavePath.ToString()));
                foreach (FileInfo fi in dir.GetFiles(".doc"))
                {
                    if (fi.FullName == mSavePath.ToString())
                    {
                        fi.Delete();
                    }
                }

                #region Font Styles
                mHeaderStyle.Name = "David";
                mHeaderStyle.Size = 24;
                mHeaderStyle.Bold = 1;
                mHeaderStyle.Color = WdColor.wdColorIndigo;

                mBodyStyle.Name = "Times New Roman";
                mBodyStyle.Size = 16;
                mBodyStyle.Bold = 0;
                mBodyStyle.Color = WdColor.wdColorAutomatic;
                #endregion Font Styles

                #region TempDir_Create
                DirectoryInfo tmpDir = new DirectoryInfo(mTempDirPath);
                if (tmpDir.Exists == true)
                {
                    tmpDir.Delete(true);
                }
                tmpDir.Create();
                #endregion
            }

            public WordWriter()
            {
            }
        #endregion

        #region Puclic Methods
            public void CreateDocument()
            {
                mCurDoc = mWordApp.Documents.Add(ref mTemplet, ref mMissing, ref mMissing, ref mMissing);
                                
                mCurLocation = 0;
                mCurRange = mCurDoc.Range(ref mCurLocation, ref mMissing);
            }

            public void TypeHeader()
            {
                Range HeaderRange;
                object curStart = 0;
                object curEnd = 0;
                HeaderRange = mCurDoc.Range(ref curStart, ref mMissing);
                SetFontToRagne(ref HeaderRange, mHeaderStyle);
                HeaderRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter; // CENTERED
                HeaderRange.ParagraphFormat.LeftIndent = 0;

                string sHeader = "נומרולוגיית הצ'אקרות עבור:" + System.Environment.NewLine + mFullName;

                HeaderRange.InsertAfter(sHeader + System.Environment.NewLine);
                HeaderRange.InsertAfter(System.Environment.NewLine);
                mCurLocation = HeaderRange.StoryLength - 1;
            }

            public void TypeBodyWithPicture(Control cntrl, string sText)
            {

                mCurRange = mCurDoc.Range(ref mCurLocation, ref mMissing);
                SetFontToRagne(ref mCurRange, mBodyStyle);
                mCurRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;

                mCurRange.InsertAfter(System.Environment.NewLine); mCurRange.InsertAfter(sText); mCurRange.InsertAfter(System.Environment.NewLine);

                mCurLocation = mCurRange.StoryLength - 1;
                mCurRange = mCurDoc.Range(ref mCurLocation, ref mMissing);

                picCount++;
                mCurDoc.Tables.Add(mCurRange, 1, 1, ref mMissing, ref mMissing);
                Range rngPic = mCurDoc.Tables[mCurDoc.Tables.Count].Range;

                mCurRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                string sPicPath = System.IO.Path.Combine(mTempDirPath, "Pic" + picCount.ToString() + ".jpg");
                CreateImageFromApp(sPicPath, cntrl);
                rngPic.InlineShapes.AddPicture(sPicPath, ref oFalse, ref oTrue, ref mMissing);

                mCurRange.InsertAfter(System.Environment.NewLine);

                mCurLocation = mCurRange.StoryLength - 1;
                mCurRange = mCurDoc.Range(ref mCurLocation, ref mMissing);

                object BreakType = Microsoft.Office.Interop.Word.WdBreakType.wdPageBreak;
                mCurRange.InsertBreak(ref BreakType);

                mCurLocation = mCurRange.StoryLength -1 ;//mCurDoc.Tables[mCurDoc.Tables.Count].Range.End - mCurDoc.Tables[mCurDoc.Tables.Count].Range.Start + mCurRange.StoryLength - 1;
            }

            public void InsertPageBreak()
            {
                object BreakType = Microsoft.Office.Interop.Word.WdBreakType.wdPageBreak;
                mCurRange.InsertBreak(ref BreakType);
            }

            public void FinishDoc()
            {
                DirectoryInfo tmpDir = new DirectoryInfo(mTempDirPath);
                tmpDir.Delete(true);

                mCurDoc.SaveAs(ref mSavePath, ref mMissing, ref mMissing, ref mMissing, ref mMissing, ref mMissing, ref mMissing, ref mMissing, ref mMissing, ref mMissing, ref mMissing, ref mMissing, ref mMissing, ref mMissing, ref mMissing, ref mMissing);
                //mWordApp.Visible = true;
                mWordApp.Quit(ref mMissing, ref mMissing, ref mMissing);
                System.Diagnostics.Process.Start(Path.GetDirectoryName(mSavePath.ToString()));
            }

            public void OpenWordFile()
            {
                mWordApp.Activate();
                mWordApp.Visible = true;
                
                //mCurDoc.Close(ref oTrue, ref mMissing, ref mMissing);
                mCurDoc.ActiveWindow.Close(ref oTrue, ref mMissing);
                
                mWordApp.Application.Visible = true;
                mWordApp.Documents.Open(ref mSavePath, ref mMissing, ref mMissing, ref mMissing, ref mMissing, ref mMissing, ref mMissing, ref mMissing, ref mMissing, ref mMissing, ref mMissing, ref mMissing, ref mMissing, ref mMissing,ref mMissing,ref mMissing);
            }

            public void ReplcaeTempletFile(string sNewFileLocation, string sTempletLocation)
            {
                object sNFL = sNewFileLocation;
                mCurDoc = mWordApp.Documents.Open(ref sNFL, ref oTrue, ref mMissing, ref mMissing, ref mMissing, ref mMissing, ref mMissing, ref mMissing, ref mMissing, ref mMissing, ref mMissing, ref mMissing, ref mMissing, ref mMissing, ref mMissing, ref mMissing);

                object sTmplt = sTempletLocation;
                mCurDoc.SaveAs(ref sTmplt, ref mMissing, ref mMissing, ref mMissing, ref mMissing, ref mMissing, ref mMissing, ref mMissing, ref mMissing, ref mMissing, ref mMissing, ref mMissing, ref mMissing, ref mMissing, ref mMissing, ref mMissing);
                mCurDoc.Close(ref oTrue, ref mMissing, ref mMissing); ;
                mWordApp.Quit(ref mMissing, ref oFalse, ref mMissing);
                
                GC.Collect();
                GC.WaitForFullGCComplete();
            }
        #endregion

        #region Private Methods
            private void SetFontToRagne(ref Range rng, Microsoft.Office.Interop.Word.Font font)
            {
                rng.Font.Bold = font.Bold;
                rng.Font.Color = font.Color;
                rng.Font.Name = font.Name;
                rng.Font.Size = font.Size;
            }

            private void CreateImageFromApp(string FullName, Control Cntrl)
            {
                Graphics g = Cntrl.CreateGraphics();
                Bitmap picBitMap = new Bitmap(Cntrl.Width, Cntrl.Height, g);

                Cntrl.DrawToBitmap(picBitMap, System.Drawing.Rectangle.FromLTRB(0, 0, Cntrl.Width, Cntrl.Height ));

                picBitMap.Save(FullName, System.Drawing.Imaging.ImageFormat.Jpeg);
            }
        #endregion
    }
    */


}