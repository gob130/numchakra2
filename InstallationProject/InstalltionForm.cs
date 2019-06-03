using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.IO;
using System.Windows.Forms;

using IWshRuntimeLibrary;
using System.Runtime.InteropServices;
using Microsoft.Win32;
using System.Diagnostics;


namespace InstallationProject
{
    public partial class InstalltionForm : Form
    {
        #region Private Members
        private string mInstallAppFolder = System.Environment.CurrentDirectory;
        private string mInternalFilesDir = System.Environment.CurrentDirectory + "\\InternalFiles";
        private string mMSInstallersDir = System.Environment.CurrentDirectory + "\\MS_Install";
        private string mDestinationFolder = @"C:\ChakraNumerology";
        #endregion

        public InstalltionForm()
        {
            InitializeComponent();
        }

        private void cmdInstall_Click(object sender, EventArgs e)
        {
            bool resInsApp = true;
            try
            {

                bool IsDBCopyNessecery = false;
                //mDestinationFolder = @"D:\_Progs\ChakraNumerology";

                DirectoryInfo mDestFolder = new DirectoryInfo(mDestinationFolder);

                if (mDestFolder.Exists == false)
                {
                    IsDBCopyNessecery = true;
                    mDestFolder.Create();
                }
                else
                {
                    IsDBCopyNessecery = false;
                }

                DirectoryInfo mInternalFiles = new DirectoryInfo(mInternalFilesDir);
                FileInfo[] InternalFiles = mInternalFiles.GetFiles();
                DirectoryInfo[] InternalDris = mInternalFiles.GetDirectories();

                mProgressBar.Step = 1;
                mProgressBar.Minimum = 0;
                mProgressBar.Maximum = InternalFiles.Length;
                foreach (DirectoryInfo dir in InternalDris)
                {
                    mProgressBar.Maximum += dir.GetFiles().Length;
                }


                foreach (FileInfo file in InternalFiles)
                {
                    if ((file.Name == "DB.mdb") && (IsDBCopyNessecery == true))
                    {
                        file.CopyTo(mDestinationFolder + "\\" + file.Name);
                    }

                    if (file.Name != "DB.mdb")
                    {
                        file.CopyTo(mDestinationFolder + "\\" + file.Name, true);
                    }

                    if (System.IO.Path.GetExtension(file.Name).ToLower() == "dll")
                    {
                        RegisterDLLs(mDestinationFolder + "\\" + file.Name);
                    }

                    mProgressBar.Value++;
                }
                foreach (DirectoryInfo dir in InternalDris)
                {
                    /*
                    #region Copy Directories
                    DirectoryInfo newDir = new DirectoryInfo(mDestinationFolder + @"\" + dir.Name);
                    if (newDir.Exists == false)
                    {
                        newDir.Create();
                    }
                    foreach (FileInfo file in dir.GetFiles())
                    {
                        file.CopyTo(newDir.FullName + @"\" + file.Name, true);
                        mProgressBar.Value++;
                    }
                    #endregion Copy Directories
                    */
                    CopyAll(dir, new DirectoryInfo(mDestinationFolder + @"\" + dir.Name));
                }

                CreateShortCutAtDestkop();
            }
            catch
            {
                resInsApp = false;
                //DirectoryInfo dir = new DirectoryInfo(mDestinationFolder);
                //dir.Delete(true);
            }

            if (resInsApp == true)
            {
                MessageBox.Show("התכנה הותקנה בהצלחה על המחשב", "סיום ההתקנה", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show("התכנה לא הותקנה", "סיום ההתקנה", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            try
            {
                InstallDotNet2();
            }
            catch
            {

            }
            this.Close();
        }

        private bool InstallDotNet2()
        {
            bool res = true;
            try
            {
                RegistryKey rkDotNetV20 = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\.NETFramework\policy\v2.0");
                RegistryKey rkDotNetV30 = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\.NETFramework\policy\v3.0");
                RegistryKey rkDotNetV35 = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\.NETFramework\policy\v3.5");
                RegistryKey rkDotNetV40 = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\.NETFramework\policy\v4.0");
                if ((rkDotNetV20 == null) && (rkDotNetV30 == null) && (rkDotNetV35 == null) && (rkDotNetV40 == null))
                {
                    MessageBox.Show("יש צורך בהתקנת תוכנה נוספת מבית מיקרוסופט" + System.Environment.NewLine + "...ההתקנה תתחיל כעת","המשך התקנה",MessageBoxButtons.OK,MessageBoxIcon.Information);
                    System.Diagnostics.Process.Start(mMSInstallersDir + "\\MS_Install\\dotnetfxv2.0.exe");
                }
                else
                {
                } 
            }
            catch
            {
                res = false;
            }

            return res;
        }

        private bool RegisterDLLs(string dllPath)
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

        private bool UnRegisterDLLs(string dllPath)
        {
            bool res = true;
            try
            {
                //'/s' : indicates regsvr32.exe to run silently.
                string fileinfo = "/u" + " " + "\"" + dllPath + "\"";

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

        private void CreateShortCutAtDestkop()
        {
            string targetdir = System.Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
            string mShortcutPath = targetdir + "\\נומרולוגיית הצאקרות.lnk";
            string mAppPath = mDestinationFolder + "\\NumChakra.exe";

            // Create a new instance of WshShellClass
            WshShellClass WshShell = new WshShellClass();

            // Create the shortcut
            IWshRuntimeLibrary.IWshShortcut MyShortcut;

            // Choose the path for the shortcut
            MyShortcut = (IWshRuntimeLibrary.IWshShortcut)WshShell.CreateShortcut(mShortcutPath);

            // Where the shortcut should point to
            MyShortcut.TargetPath = mAppPath;

            // Description for the shortcut
            MyShortcut.Description = "Launch Numerology Chakra Application";

            // Location for the shortcut's icon
            MyShortcut.IconLocation = mDestinationFolder + @"\CHKR_NUM.ico";

            // Create the shortcut at the given path
            MyShortcut.Save();

            //FileInfo fiSC = mDestFolder.GetFiles("*.lnk")[0];
            //fiSC.CopyTo(targetdir);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog BrowseFolder = new FolderBrowserDialog();
            BrowseFolder.Description = "בחר תיקיית התקנה חדשה - יש לשים לב להרשאות";
            BrowseFolder.ShowNewFolderButton = true;

            DialogResult res = BrowseFolder.ShowDialog();
            if (res == DialogResult.OK)
            {
                mDestinationFolder = BrowseFolder.SelectedPath;
            }
            else
            {
                return;
            }
        }

        private void cmdUnInstall_Click(object sender, EventArgs e)
        {
            DialogResult res = MessageBox.Show("האם להסיר את התכנה?","בדיקת משתמש",MessageBoxButtons.YesNo,MessageBoxIcon.Warning);
            if ((res == DialogResult.No) || (res == DialogResult.Cancel))
            {
                return;
            }
            else // Yes
            {
                DirectoryInfo mAppDir = new DirectoryInfo(mDestinationFolder);
                if (mAppDir.Exists == true)
                {
                    mAppDir.Delete(true);
                }
                else
                {
                    MessageBox.Show("התכנה לא הותקנה במקומה המקורי - אנא נווט אל נתכנה עצמה?", "שגיאת הסרה", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    FolderBrowserDialog fbd = new FolderBrowserDialog();
                    fbd.Description = "נווט אל התיקייה בה מותקנת התכנה כעת...";
                    res = fbd.ShowDialog();
                    if (res == DialogResult.Cancel)
                        return;
                    else
                    {
                        bool ans = UnRegisterDLLs(fbd.SelectedPath + "\\BLL.dll");

                        mAppDir = new DirectoryInfo(fbd.SelectedPath);
                        mAppDir.Delete(true);
                    }
                }
            }
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            System.Diagnostics.Process.Start("mailto:yaakobi999@gmail.com");
        }

        private static void CopyAll(DirectoryInfo source, DirectoryInfo target)
        {
            if (Directory.Exists(target.FullName) == false)
            {
                Directory.CreateDirectory(target.FullName);
            }
            
            //// Copy each file into it's new directory.
            foreach (FileInfo fi in source.GetFiles())
            {
                if (fi.Name == "DB.mdb")
                {
                    FileInfo fDB = new FileInfo(System.IO.Path.Combine(target.ToString(), fi.Name));
                    if (fDB.Exists != true) // db does not exsits
                    {
                        fi.CopyTo(System.IO.Path.Combine(target.ToString(), fi.Name), true);
                    }
                }
                else
                {
                    fi.CopyTo(System.IO.Path.Combine(target.ToString(), fi.Name), true);
                }
            }

            // Copy each subdirectory using recursion.
            foreach (DirectoryInfo diSourceSubDir in source.GetDirectories())
            {
                DirectoryInfo nextTargetSubDir =
                    target.CreateSubdirectory(diSourceSubDir.Name);
                CopyAll(diSourceSubDir, nextTargetSubDir);
            }
        }
    }
}
