using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO;

namespace FixAndUpdate
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                string curpath = System.IO.Path.GetDirectoryName(System.Windows.Forms.Application.ExecutablePath);
                string InternalFilesPath = curpath + "\\files";

                DirectoryInfo FilesDir = new DirectoryInfo(InternalFilesPath);

                //DirectoryInfo AppDirXP      = new DirectoryInfo(@"C:\Program Files\OmersSimpleSolutions\Dynamic Chakra Numerology\");
                DirectoryInfo AppDirWin7x32 = new DirectoryInfo(@"C:\Program Files\OmersSimpleSolutions\Dynamic Chakra Numerology\");
                DirectoryInfo AppDirWin7x64 = new DirectoryInfo(@"C:\Program Files (x86)\OmersSimpleSolutions\Dynamic Chakra Numerology\");

                DirectoryInfo AppDir = new DirectoryInfo(@"C:\");
                if (AppDirWin7x32.Exists == true)
                {
                    AppDir = AppDirWin7x32;
                }
                if (AppDirWin7x64.Exists == true)
                {
                    AppDir = AppDirWin7x64;
                }

                foreach (FileInfo File in FilesDir.GetFiles())
                {
                    string newFilePath = System.IO.Path.Combine(AppDir.FullName, File.Name);
                    File.CopyTo(newFilePath, true);
                }

                MessageBox.Show("התיקון הושלם בהצלחה", "Process Ended....", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception exp)
            {
                MessageBox.Show(exp.Message);
            }

            this.Close();
        }
    }
}
