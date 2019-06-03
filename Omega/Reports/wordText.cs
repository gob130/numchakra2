using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using System.IO;
using Microsoft.Office.Interop;
using Microsoft.Office.Interop.Word;
using Omega.Enums;

namespace Omega.Reports
{
    public class wordText
    {
        private Font mFont;
        private string mText;
        private EnumProvider.Alignment mAlignment;

        #region Properties
        public string Text
        {
            get
            {
                return mText;
            }
        }
        public EnumProvider.Alignment Alignment
        {
            get
            {
                return mAlignment;
            }
        }
        public Font Font
        {
            get
            {
                return mFont;
            }
        }
        #endregion Properties

        #region Constructor
        public wordText()
        {
            mFont = new FontClass();
            mText = "";
            mAlignment = EnumProvider.Alignment.Right;
        }

        public wordText(FontClass font, string text, EnumProvider.Alignment alignment)
        {
            mFont = font;
            mText = text;
            mAlignment = alignment;
        }
        #endregion Constructor

        public bool XmlNode2Text(XmlNode node)
        {
            bool res = true;
            try
            {
                #region Font Info
                string[] splt = node.OuterXml.Split(" ".ToCharArray()[0]);

                string fontName = node.Attributes.GetNamedItem("font").InnerText.Trim();// splt[1].Split((char)34)[1].Trim();
                string fontSize = node.Attributes.GetNamedItem("size").InnerText.Trim();//splt[2].Split((char)34)[1].Trim();
                bool fontBold = (node.Attributes.GetNamedItem("bold").InnerText.Trim() == "true");
                bool fontUnderline = (node.Attributes.GetNamedItem("underline").InnerText.Trim() == "true");
                bool fontItalique = (node.Attributes.GetNamedItem("italic").InnerText.Trim() == "true");
                string align = node.Attributes.GetNamedItem("alignment").InnerText.Trim();

                mFont = new Microsoft.Office.Interop.Word.FontClass();
                mFont.Name = fontName;
                mFont.NameAscii = fontName;
                mFont.NameBi = fontName;
                mFont.NameOther = fontName;
                mFont.Size = (float)Convert.ToDouble(fontSize);
                mFont.Bold = Convert.ToInt16(fontBold);
                mFont.Italic = Convert.ToInt16(fontItalique);
                if (fontUnderline)
                {
                    mFont.Underline = WdUnderline.wdUnderlineSingle;
                }
                else
                {
                    mFont.Underline = WdUnderline.wdUnderlineNone;
                }
                #endregion

                mAlignment = (EnumProvider.Alignment)EnumProvider.Instance.GetAlignenmentEnumFromDescription(align);

                #region TextParser
                string text2WordDoc;
                bool resText = ReportDataProvider.Instance.TextParser(node.InnerText, out text2WordDoc);
                #endregion

                mText = text2WordDoc;
            }
            catch
            {
                mFont = new FontClass();
                mText = "";

                res = false;
            }

            return res;
        }
    }
}
