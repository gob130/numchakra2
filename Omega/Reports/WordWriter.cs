#region Namespaces

using log4net;
using Microsoft.Office.Interop.Word;
using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Windows.Forms;
using System.Xml;

#endregion Namespaces

// **********

namespace Omega.Reports
{
    public class WordWriter
    {
        #region Data Members

        private object mMissing = System.Reflection.Missing.Value;
        private object mSavePath;
        private object mTemplet; //= (object)Path.Combine(mApplicationMainDir, @"Templets\ThisTemplet.dot"); // need to find way 2 insert a templet here...
        private object oTrue = true as object;
        private object oFalse = false as object;

        private static Form mMainForm;
        private Microsoft.Office.Interop.Word.ApplicationClass mWordApp = new Microsoft.Office.Interop.Word.ApplicationClass();
        private Microsoft.Office.Interop.Word.Document mCurDoc;

        private Microsoft.Office.Interop.Word.Font mHeaderStyle = new Microsoft.Office.Interop.Word.FontClass();
        private Microsoft.Office.Interop.Word.Font mBodyStyle = new Microsoft.Office.Interop.Word.FontClass();

        private string mTempDirPath;// = System.IO.Path.Combine(System.Environment.CurrentDirectory, "TempDir");
        private string mFullName;

        private int picCount = 0;
        private Range mCurRange;
        private object mCurLocation;
        private static readonly ILog mlog = LogManager.GetLogger("WordWriter");

        #endregion

        #region Constructor

        public WordWriter(string ApplicationMainDir, string sDocPath, string firstname, string lastname)
        {
            mlog.Info("WordWriter object created...");
            picCount = 0;
            mMainForm = ReportDataProvider.Instance.mMainForm;

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
            try
            {
                mlog.Info("Try to create directory: " + mTempDirPath);
                DirectoryInfo tmpDir = new DirectoryInfo(mTempDirPath);
                if (tmpDir.Exists == true)
                {
                    tmpDir.Delete(true);
                }

                tmpDir.Create();
                mlog.Info("Created");

            }
            catch (Exception ex)
            {
                mlog.Error(ex);

            }

            bool prem = BLL.AppSettings.Instance.AddPremissions2SpecificFolder(mTempDirPath);

        }
        #endregion

        public WordWriter()
        {
            mlog.Info("WordWriter object created...");
            mMainForm = MainForm.ActiveForm;

        }

        #endregion

        public void CloseWordWriter()
        {
            mWordApp.Quit();
        }

        #region Puclic Methods

        #region Type Styles

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

            mCurLocation = mCurRange.StoryLength - 1;//mCurDoc.Tables[mCurDoc.Tables.Count].Range.End - mCurDoc.Tables[mCurDoc.Tables.Count].Range.Start + mCurRange.StoryLength - 1;
        }
        /*
        public void TypeText(FontClass thisFont, string text, string alignment)
        {
            Range textRange;
            object curStart = 0;
            object curEnd = 0;
            textRange = mCurDoc.Range(ref mCurLocation, ref mMissing);
            SetFontToRagne(ref textRange, thisFont);

            textRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight; // Right
            switch (alignment.ToLower())
            {
                case "center":
                    textRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter; // CENTERED
                    break;
                case "left":
                    textRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft; // Right
                    break;
                case "right":
                    textRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight; // Right
                    break;
            }

            textRange.ParagraphFormat.LeftIndent = 0;

            string sHeader = text;

            textRange.InsertAfter(sHeader + " ");
            //textRange.InsertAfter(System.Environment.NewLine);

            mCurLocation = mCurRange.StoryLength - 1;
            mCurRange = mCurDoc.Range(ref mCurLocation, ref mMissing);
            InsertBreak();
        }
        */
        public void TypeText(Microsoft.Office.Interop.Word.Font thisFont, string text, string alignment)
        {
            Microsoft.Office.Interop.Word.Range textRange;
            object curStart = 0;
            object curEnd = 0;
            textRange = mCurDoc.Range(ref mCurLocation, ref mMissing);

            textRange.InsertBefore(text.ToString());
            textRange.Select();


            Microsoft.Office.Interop.Word.Font curFont = thisFont;
            SetFontToRagne(ref textRange, curFont);

            textRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight; // Right
            switch (alignment.ToLower())
            {
                case "center":
                    textRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter; // CENTERED
                    break;
                case "left":
                    textRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft; // Right
                    break;
                case "right":
                    textRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight; // Right
                    break;
            }

            textRange.ParagraphFormat.LeftIndent = 0;


            textRange.InsertParagraphAfter();
            textRange.Select();
            //textRange.InsertAfter(System.Environment.NewLine);
            //InsertBreak();

            mCurLocation = mCurRange.StoryLength - 1;
            mCurRange = mCurDoc.Range(ref mCurLocation, ref mMissing);

        }

        public void TypeImage(string path2image)
        {
            mCurRange = mCurDoc.Range(ref mCurLocation, ref mMissing);

            mCurRange.ParagraphFormat.LeftIndent = 0;

            mCurRange.InsertAfter(System.Environment.NewLine);
            mCurLocation = mCurRange.StoryLength - 1;
            mCurRange = mCurDoc.Range(ref mCurLocation, ref mMissing);

            //mCurRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

            picCount++;
            mCurDoc.Tables.Add(mCurRange, 1, 1, ref mMissing, ref mMissing);
            Range rngPic = mCurDoc.Tables[mCurDoc.Tables.Count].Range;

            if (ReportDataProvider.Instance.isImageReserved(path2image.Trim()) == true)
            {
                path2image = path2image.Trim();
                string sPicPath = System.IO.Path.Combine(mTempDirPath, "Pic" + picCount.ToString() + "." + System.Drawing.Imaging.ImageFormat.Png.ToString());
                CreateImageFromApp(sPicPath, ReportDataProvider.Instance.Reserved2ControlName(path2image));
                rngPic.InlineShapes.AddPicture(sPicPath, ref oFalse, ref oTrue, ref mMissing);
            }
            else // image is a file on FileSystem path
            {
                // chack if image is internal (as file)...

                if (path2image.Contains("InternalImages") == true)
                {
                    path2image = path2image.Replace("InternalImages", System.IO.Path.Combine(BLL.AppSettings.Instance.AppmMainDir, @"Templets\ReportStyles\InternalImages"));
                }

                rngPic.InlineShapes.AddPicture(path2image, ref oFalse, ref oTrue, ref mMissing);


            }
            mCurRange.InsertAfter(System.Environment.NewLine);

            mCurLocation = mCurRange.StoryLength - 1;
            mCurRange = mCurDoc.Range(ref mCurLocation, ref mMissing);

            /*
            mCurRange.InsertAfter(System.Environment.NewLine);

            mCurLocation = mCurRange.StoryLength - 1;
            mCurRange = mCurDoc.Range(ref mCurLocation, ref mMissing);
            */
        }

        public void InsertPageBreak()
        {
            object BreakType = Microsoft.Office.Interop.Word.WdBreakType.wdPageBreak;
            mCurRange.InsertBreak(ref BreakType);

            //mCurRange = mCurDoc.Range(ref mCurLocation, ref mMissing);
            mCurLocation = mCurRange.StoryLength - 1;
        }
        public void InsertBreak()
        {
            object BreakType = Microsoft.Office.Interop.Word.WdBreakType.wdLineBreakClearLeft;
            mCurRange.InsertBreak(ref BreakType);

            //mCurRange = mCurDoc.Range(ref mCurLocation, ref mMissing);
            mCurLocation = mCurRange.StoryLength - 1;
        }
        #endregion Type Styles

        #region XML
        /// <summary>
        /// Prints the XML Report Syle to .Doc file
        /// </summary>
        /// <param name="xmlDoc">XML Report format</param>
        public void TypeXml2Doc(XmlDocument xmlDoc)
        {
            XmlNodeList root = xmlDoc.ChildNodes;

            int PageCounter = 0;

            try
            {
                mlog.Info("Filling document with data from XML file");

                foreach (XmlNode curPageNode in root[1].ChildNodes)
                {
                    PageCounter++;

                    TypeXmlPage(curPageNode);

                    // Add Page Break between pages
                    if (PageCounter < root[1].ChildNodes.Count)
                    {
                        InsertPageBreak();
                    }
                }

                mlog.Info("Document filled");

            }
            catch (Exception ex)
            {
                mlog.Error(ex);

            }

        }

        private void TypeXmlPage(XmlNode curPage)
        {
            foreach (XmlNode node in curPage.ChildNodes)
            {
                switch (node.Name.ToLower())
                {
                    case "text":
                        {
                            wordText wrdTxt = new wordText();
                            bool res = wrdTxt.XmlNode2Text(node);
                            TypeText(wrdTxt.Font, wrdTxt.Text.Trim(), wrdTxt.Alignment.ToString());

                            break;

                        }

                    case "image":
                        {
                            //TypeImage(string path2image)
                            TypeImage(node.InnerText.Trim());

                            break;

                        }

                }

            }

        }

        #endregion XML

        #region General
        public void CreateDocument()
        {
            mCurDoc = mWordApp.Documents.Add(ref mTemplet, ref mMissing, ref mMissing, ref mMissing);

            mCurLocation = 0;
            mCurRange = mCurDoc.Range(ref mCurLocation, ref mMissing);
        }

        public void FinishDoc()
        {
            mCurDoc.SaveAs(ref mSavePath, ref mMissing, ref mMissing, ref mMissing, ref mMissing, ref mMissing, ref mMissing, ref mMissing, ref mMissing, ref mMissing, ref mMissing, ref mMissing, ref mMissing, ref mMissing, ref mMissing, ref mMissing);
            mWordApp.Quit(ref mMissing, ref mMissing, ref mMissing);

            GC.Collect();
            GC.WaitForPendingFinalizers();

            System.Threading.Thread.Sleep(500);
            System.Diagnostics.Process.Start(Path.GetDirectoryName(mSavePath.ToString()));


            DirectoryInfo tmpDir = new DirectoryInfo(mTempDirPath);
            tmpDir.Delete(true);

        }


        public void OpenWordFile()
        {
            mWordApp.Activate();
            mWordApp.Visible = true;

            //mCurDoc.Close(ref oTrue, ref mMissing, ref mMissing);
            mCurDoc.ActiveWindow.Close(ref oTrue, ref mMissing);

            mWordApp.Application.Visible = true;
            mWordApp.Documents.Open(ref mSavePath, ref mMissing, ref mMissing, ref mMissing, ref mMissing, ref mMissing, ref mMissing, ref mMissing, ref mMissing, ref mMissing, ref mMissing, ref mMissing, ref mMissing, ref mMissing, ref mMissing, ref mMissing);

        }

        public void ReplcaeTempletFile(string sNewFileLocation, string sTempletLocation)
        {
            object sNFL = sNewFileLocation;
            mCurDoc = mWordApp.Documents.Open(ref sNFL, ref oTrue, ref mMissing, ref mMissing, ref mMissing, ref mMissing, ref mMissing, ref mMissing, ref mMissing, ref mMissing, ref mMissing, true, ref mMissing, ref mMissing, ref mMissing, ref mMissing);

            object sTmplt = sTempletLocation;

            mCurDoc.SaveAs(ref sTmplt, ref mMissing, ref mMissing, ref mMissing, ref mMissing, ref mMissing, ref mMissing, ref mMissing, ref mMissing, ref mMissing, ref mMissing, ref mMissing, ref mMissing, ref mMissing, ref mMissing, ref mMissing);
            mCurDoc.Close(ref oTrue, ref mMissing, ref mMissing); ;
            mWordApp.Quit(ref mMissing, ref oFalse, ref mMissing);

            GC.Collect();
            GC.WaitForFullGCComplete();
        }
        #endregion General
        #endregion

        #region Private Methods

        private void SetFontToRagne(ref Range rng, Microsoft.Office.Interop.Word.Font font)
        {
            //rng.Font = new Microsoft.Office.Interop.Word.FontClass();
            rng.Font.Bold = font.Bold;
            rng.Font.BoldBi = font.Bold;
            rng.Font.Color = font.Color;
            rng.Font.Underline = font.Underline;
            rng.Font.Italic = font.Italic;
            rng.Font.Name = font.Name;
            rng.Font.NameAscii = font.NameAscii;
            rng.Font.NameBi = font.NameBi;
            rng.Font.NameOther = font.NameOther;
            rng.Font.Size = font.Size;
            rng.Font.SizeBi = font.Size;


            //rng.Font = font as FontClass;            
        }

        private void CreateImageFromApp(string FullPath2SaveImage, Control Cntrl)
        {
            mMainForm.Invoke(new MethodInvoker(delegate ()
            {
                //mMainForm.Show();
                //mMainForm.Focus();
                //mMainForm.Refresh();
                mlog.Info("Creating image file from application");
                if (Cntrl.Name == "grobxCycles")
                {
                    mlog.Info("Creating image file from life cycles");
                    Control.ControlCollection cntrlCls = mMainForm.Controls;

                    TabControl ttt = cntrlCls.Find("mTabs", true)[0] as TabControl;
                    ttt.TabPages["tabPage5"].Show();
                    ttt.TabPages["tabPage5"].Refresh();

                }

            }));            

            //Cntrl.Invoke(new MethodInvoker(delegate ()
            //{
            //    Cntrl.Show();
            //    Cntrl.Refresh();
            //    g = Cntrl.CreateGraphics();

            //}));


            Cntrl.Invoke(new MethodInvoker(delegate ()
            {
                mlog.Info("Creating bitmap object");
                Bitmap picBitMap = new Bitmap(Cntrl.Width, Cntrl.Height);

                Cntrl.DrawToBitmap(picBitMap, System.Drawing.Rectangle.FromLTRB(0, 0, Cntrl.Width, Cntrl.Height));
                picBitMap.Save(FullPath2SaveImage, ImageFormat.Png);
                mlog.Info("Bitmap object saved successfuly");

            }));

        }

        private void CreateImageFromApp(string FullPath2SaveImage, string CntrlName)
        {
            Control.ControlCollection cntrlCls = mMainForm.Controls;
            CreateImageFromApp(FullPath2SaveImage, cntrlCls.Find(CntrlName, true)[0]);
        }

        #endregion

    }

}
