using System;
using System.Collections.Generic;
using System.Windows.Forms;
using BLL;
using System.IO;
using ErrorLogger;


namespace Omega
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]

        #region Private Members
        #endregion

        static void Main()
        {
            //BLL.AppSettings OmegaAppSettings = new AppSettings(System.IO.Path.GetDirectoryName(System.Windows.Forms.Application.ExecutablePath));

            try
            {
                BLL.AppSettings.Instance.ChangeSecuritySettingsForApplicationDirectory();

                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.WaitForFullGCComplete();

                FileInfo dbFile = new FileInfo(AppSettings.Instance.GetXmlDbPath());
                if (dbFile.Exists == true)
                {
                    File.SetAttributes(dbFile.FullName, FileAttributes.Normal);
                    File.SetAttributes(dbFile.FullName, ~FileAttributes.Hidden);
                    dbFile.IsReadOnly = false;
                }
                else
                {
                    bool test = false;
                    
                    #region Seek for older db Positions
                    List<string> lst = new List<string> { "C:\\Program Files\\OmersSimpleSolutions\\Dynamic Chakra Numerology\\",
                                                          "C:\\Program Files(x86)\\OmersSimpleSolutions\\Dynamic Chakra Numerology\\"};

                        foreach (string str in lst)
                        {

                            DirectoryInfo dir = new DirectoryInfo(str);
                            if (dir.Exists == true)
                            {
                                foreach (FileInfo file in dir.GetFiles())
                                {
                                    if (file.Name.ToLower().Contains("db.xml") == true)
                                    {
                                        test = true;
                                        file.CopyTo(BLL.AppSettings.Instance.GetXmlDbPath());
                                        break;
                                    }
                                }
                            }
                            if (test == true)
                            {
                                break;
                            }
                        }
                    #endregion

                    if (test == false)
                    {
                        bool isDBcreated = AppSettings.Instance.CreateNewDB();
                        //Omega.Properties.Resources.CopyTo(System.IO.Path.Combine(AppSettings.Instance.AppmMainDir,"Db.xml"));
                    }
                    dbFile.IsReadOnly = false;
                }
                //dbFile = new FileInfo(AppSettings.Instance.AppmMainDir + "\\DB.mdb");
                //dbFile.IsReadOnly = false;


                //if (AppSettings.Instance.ProgramType != AppSettings.ProgType.Trial)
                //{
                //  if (AppSettings.Instance.isFileInRegistry("Microsoft.Office.Interop.Word.dll") == false)
                //  {
                bool res = true;
                res = AppSettings.Instance.RegisterDLLs(AppSettings.Instance.AppmMainDir + "\\Microsoft.Office.Interop.Word.dll");
                res = AppSettings.Instance.RegisterDLLs(AppSettings.Instance.AppmMainDir + "\\Microsoft.Office.Interop.Word.xml");
                res = AppSettings.Instance.RegisterDLLs(AppSettings.Instance.AppmMainDir + "\\EnvDTE.dll");
                res = AppSettings.Instance.RegisterDLLs(AppSettings.Instance.AppmMainDir + "\\VsWebSite.Interop.dll");
                res = AppSettings.Instance.RegisterDLLs(AppSettings.Instance.AppmMainDir + "\\VsWebSite.Interop90.dll");

                //  }
                //}

                if ((AppSettings.Instance.ProgramType == AppSettings.ProgType.Expert ) || (AppSettings.Instance.ProgramType == AppSettings.ProgType.Normal ))
                {
                    bool resFileNameChane = AppSettings.Instance.ChangeReportFileNames();
                }


                DateTime MasterENDDate = dbFile.CreationTime;
                if ((AppSettings.Instance.ProgramType == AppSettings.ProgType.Trial) && (MasterENDDate.Subtract(DateTime.Now).TotalDays > AppSettings.Instance.TrialTimeDuration))
                {
                    FileInfo dllFile = new FileInfo(AppSettings.Instance.AppmMainDir + "\\BLL.dll");
                    dllFile.Delete();
                    MessageBox.Show("Trial Vesrion Expired" + System.Environment.NewLine + "תכנת הנסיון פגה", "נומרולוגיית הצ'אקרות", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    return;
                }
                else
                {
                    //DateTime tExp = new DateTime(2013, 01, 01);
                    //FileInfo hiddenSecurityFile = new FileInfo(AppSettings.Instance.AppmMainDir + "\\BLLp.manifest.xml");
                    //if (hiddenSecurityFile.Exists == false)
                    //{
                    //    hiddenSecurityFile.Create();
                    //}
                    //if ((DateTime.Now.Subtract(tExp).TotalDays < 0) && (hiddenSecurityFile.Exists == true))
                    //{
                        Application.EnableVisualStyles();
                        Application.SetCompatibleTextRenderingDefault(false);
                        Application.Run(new MainForm());
                    //}
                }
            }
            catch (Exception ex)
            {
                ErrorLogger.Logger.Instance.WriteErrorLog(ex);

                DialogResult dlgRes = MessageBox.Show("Project Error" + Environment.NewLine + "it is recommanded to view error and send it to the application distributer..." + Environment.NewLine + "would you like to open the log file and view the error?", "Application Error", MessageBoxButtons.YesNo, MessageBoxIcon.Error);
                if (dlgRes == DialogResult.Yes)
                {
                    System.Diagnostics.Process.Start(ErrorLogger.Logger.Instance.Path);
                }
            }

                
        }
    }
}
