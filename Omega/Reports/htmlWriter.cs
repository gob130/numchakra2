using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Xml;
using Omega;
using BLL;
using System.Windows.Forms;
using System.Drawing;

namespace Omega.Reports
{
    class htmlWriter
    {
        string quotationmark = (Convert.ToChar(34)).ToString(); 

        private string mDirPath = "";
        private string mPath2HTMfile = "";
        private string mPath2ImageDir = "";
        private FileStream fileHandler;
        private StreamWriter Writer;
        private int picCount;

        private string mName = "";
        private string mRepName = "";

        //private string mTempDirPath;
        private Form mMainForm;



        public htmlWriter(string path)
        {
            DateTime curTime = DateTime.Now;
            //יעקבי נשר-סולן - דוח ראשי.html
            string allname = (System.IO.Path.GetFileNameWithoutExtension(path)).Replace(" - ".ToString(), "*".ToString());
            string[] splt = allname.Split("*".ToCharArray());
            mName = splt[0].Trim().Replace("-","_");
            mRepName = splt[1].Trim();

            mDirPath = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(path), mName.Replace(" ".ToString(), "_".ToString()) + "_" + mRepName.Replace(" ".ToString(),"_".ToString()) + "_" + curTime.Date.Year.ToString() + "_" + curTime.Date.Month.ToString() + "_" + curTime.Date.Day.ToString());
            mPath2HTMfile = System.IO.Path.Combine(mDirPath, mName.Replace(" ".ToString(), "_".ToString()) + "_" + curTime.Date.Year.ToString() + "_" + curTime.Date.Month.ToString() + "_" + curTime.Date.Day.ToString() + ".html");
            mPath2ImageDir = System.IO.Path.Combine(mDirPath,"Images");

            DirectoryInfo dir = new DirectoryInfo(mDirPath);
            if (dir.Exists == false)
            {
                dir.Create();
            }

            dir = new DirectoryInfo(mPath2ImageDir);
            if (dir.Exists == false)
            {
                dir.Create();
            }
           
            fileHandler = new FileStream(mPath2HTMfile, FileMode.OpenOrCreate);
            Writer = new StreamWriter(fileHandler);

            mMainForm = ReportDataProvider.Instance.mMainForm;
            picCount = 0;
        }

        public bool TypeXml2HTML(XmlDocument xmlDoc)
        {
            bool res = true ;
            try
            {
                Writer.WriteLine("<!DOCTYPE html>");
                Writer.WriteLine("<html>");
                #region Print Head
                Writer.WriteLine("<head>");
                Writer.WriteLine("<title>" + mName.Replace("_","-") + " : " + mRepName + "</title>");
                if (BLL.AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                {//{direction:rtl;}
                    Writer.WriteLine("<style>" + Environment.NewLine + "div.container {direction:rtl;text-align:justify;}" + Environment.NewLine +
                        "p {direction:rtl;text-align:right;}" + Environment.NewLine + "</style>");
                }
                if (BLL.AppSettings.Instance.ProgramLanguage != AppSettings.Language.Hebrew)
                {//{direction:rtl;}
                    Writer.WriteLine("<style>" + Environment.NewLine + "div.container {direction:ltr;text-align:justify;}" + Environment.NewLine +
                        "p {direction:ltr;text-align:justify;}" + Environment.NewLine + "</style>");
                }
                Writer.WriteLine("</head>");
                #endregion Print Head

                Writer.WriteLine("<body>");
                Writer.WriteLine("<div class=" + quotationmark + "container" + quotationmark + " style= " + quotationmark + "text-align:right;"+ quotationmark +">");

                XmlNodeList root = xmlDoc.ChildNodes;

                int PageCounter = 0;
                foreach (XmlNode curPageNode in root[1].ChildNodes)
                {
                    TypeXmlPage(curPageNode);
                    PageCounter++;
                    // Add Page Break between pages
                    if (PageCounter < root[1].ChildNodes.Count)
                    {
                        Writer.WriteLine("<hr>");
                    }
                }
                Writer.WriteLine("</div>");
                Writer.WriteLine("</body>");
                Writer.WriteLine("</html>");

            }
            catch
            {
                res = false;
            }
            return res;
        }

        public void OpenReoprtOnDefaultBrowser()
        {
            Writer.Close();
            fileHandler.Close();

            System.Diagnostics.Process.Start(fileHandler.Name);
        }

        private void TypeXmlPage(XmlNode curPage)
        {
            foreach (XmlNode node in curPage.ChildNodes)
            {
                switch (node.Name.ToLower())
                {
                    case "text":
                        TypeText(node);
                        break;
                    case "image":
                        //TypeImage(string path2image)
                        TypeImage(node.InnerText.Trim());
                        break;
                }
            }
        }

        public void TypeText(XmlNode node)
        {
            string fontName = "", fontSize= "", textalign= "", text = "";
            bool fontBold = false, fontUnderline = false, fontItalique = false;
            bool res = Node2TextAndFontStyle(node,out text, out  fontName, out  fontSize, out  textalign, out  fontBold, out  fontUnderline, out fontItalique);
            
            string style = "font-family:" + fontName + ";color:balck;font-size:" + fontSize + ";text-align:" + textalign + ";";
            string fText = AddBold(AddItalic(AddUnderline(text, fontUnderline), fontItalique), fontBold);
            //fText = AddRTL(fText,BLL.AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew);
            //Writer.WriteLine("<p style=" + quotationmark + style + quotationmark + ">" + fText + "</p>");
            Writer.WriteLine("<pre style=" + quotationmark + style + quotationmark + ">" + fText + "</pre>");
            Writer.WriteLine("<br>");
        }

        public void TypeImage(string path2image)
        {
            picCount++;
            string sPicPath = "";

            if (ReportDataProvider.Instance.isImageReserved(path2image.Trim()) == true)
            {
                path2image = path2image.Trim();
                sPicPath = System.IO.Path.Combine(mPath2ImageDir, "Pic" + picCount.ToString() + "." + System.Drawing.Imaging.ImageFormat.Png.ToString());
                CreateImageFromApp(sPicPath, ReportDataProvider.Instance.Reserved2ControlName(path2image));
            }
            else // image is a file on FileSystem path
            {
                // chack if image is internal (as file)...

                if (path2image.Contains("InternalImages") == true)
                {
                    sPicPath = path2image.Replace("InternalImages", System.IO.Path.Combine(BLL.AppSettings.Instance.AppmMainDir, @"Templets\ReportStyles\InternalImages"));
                }
            }
            Writer.WriteLine("<div id=" + quotationmark + "image_" + picCount.ToString() + quotationmark + " style="+ quotationmark + "text-align:center;"+quotationmark +">"); 
            Writer.WriteLine("<img src=" + quotationmark + @"Images\" + System.IO.Path.GetFileName(sPicPath) + quotationmark + ">");
            Writer.WriteLine("</div>");
            Writer.WriteLine("<br>");
            //mPath2ImageDir
        }

        private void CreateImageFromApp(string FullPath2SaveImage, Control Cntrl)
        {
            mMainForm.Show();
            mMainForm.Focus();
            mMainForm.Refresh();

            if (Cntrl.Name == "grobxCycles")
            {
                Control.ControlCollection cntrlCls = mMainForm.Controls;

                TabControl ttt = cntrlCls.Find("mTabs", true)[0] as TabControl;
                ttt.TabPages["tabPage5"].Show();
                ttt.TabPages["tabPage5"].Refresh();
            }

            Cntrl.Show();
            Cntrl.Refresh();

            Graphics g = Cntrl.CreateGraphics();
            Bitmap picBitMap = new Bitmap(Cntrl.Width, Cntrl.Height, g);

            Cntrl.DrawToBitmap(picBitMap, System.Drawing.Rectangle.FromLTRB(0, 0, Cntrl.Width, Cntrl.Height));

            picBitMap.Save(FullPath2SaveImage, System.Drawing.Imaging.ImageFormat.Png);
        }

        private void CreateImageFromApp(string FullPath2SaveImage, string CntrlName)
        {
            Control.ControlCollection cntrlCls = mMainForm.Controls;
            CreateImageFromApp(FullPath2SaveImage, cntrlCls.Find(CntrlName, true)[0]);
        }

        private bool Node2TextAndFontStyle(XmlNode node, out string text, out string fontName, out string fontSize, out string textalign, out bool fontBold, out bool fontUnderline, out bool fontItalique)
        {
            bool res = true;

            try
            {
                #region Font Info
                string[] splt = node.OuterXml.Split(" ".ToCharArray()[0]);

                fontName = node.Attributes.GetNamedItem("font").InnerText.Trim();// splt[1].Split((char)34)[1].Trim();
                fontSize = node.Attributes.GetNamedItem("size").InnerText.Trim();//splt[2].Split((char)34)[1].Trim();
                fontBold = (node.Attributes.GetNamedItem("bold").InnerText.Trim() == "true");
                fontUnderline = (node.Attributes.GetNamedItem("underline").InnerText.Trim() == "true");
                fontItalique = (node.Attributes.GetNamedItem("italic").InnerText.Trim() == "true");
                textalign = node.Attributes.GetNamedItem("alignment").InnerText.Trim();
                #endregion

                //mAlignment = (EnumProvider.Alignment)EnumProvider.Instance.GetAlignenmentEnumFromDescription(textalign);

                #region TextParser
                string text2WordDoc;
                bool resText = ReportDataProvider.Instance.TextParser(node.InnerText, out text2WordDoc);

                text = text2WordDoc;
                #endregion
            }
            catch
            {
                res = false;
                fontName = ""; fontSize = ""; textalign = ""; text = "";
                fontBold = false; fontUnderline = false; fontItalique = false;
            }
            return res;
        }

        private string AddBold(string txt, bool b)
        {
            if (b == true)
            {
                return "<b>" + txt + "</b>";
            }
            else
            {
                return txt;
            }
        }
        private string AddUnderline(string txt, bool u)
        {
            if (u == true)
            {
                return "<u>" + txt + "</u>";
            }
            else
            {
                return txt;
            }
        }
        private string AddItalic(string txt, bool i)
        {
            if (i == true)
            {
                return "<i>" + txt + "</i>";
            }
            else
            {
                return txt;
            }
        }
        private string AddRTL(string txt, bool rtl)
        {
            string text = "";
            if (rtl == true)
            {
                text = "<bdo dir=" + quotationmark + "rtl" + quotationmark + ">" + txt + "</bdo>";
            }
            else
            {text = txt;
            }
            return text;
        }
    }
}
