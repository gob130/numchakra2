using System;
using System.Collections.Generic;
using System.Text;
using Omega;
using BLL;

namespace Omega.Objects
{
    public class FullResults
    {
        #region Chakra Data
        public int Crown;
        public int Throught;
        public int ThirdEye;
        public int Heart;
        public int SexAndCreation;
        public int Sun;
        public int Base;
        public int Universal;
        public int Superior;
        
        public int Astro;
        public string AstroName;
        
        public Enums.EnumProvider.Balance Balanced;
        #endregion

        #region LifeCycles
        public int lcAges1;
        public int lcAges2;
        public int lcAges3;
        public int lcAges4;

        public int lcClimax1;
        public int lcClimax2;
        public int lcClimax3;
        public int lcClimax4;

        public int lcCycle1;
        public int lcCycle2;
        public int lcCycle3;
        public int lcCycle4;

        public int lcChalange1;
        public int lcChalange2;
        public int lcChalange3;
        public int lcChalange4;

        public int currentCycle;
        public int lcCCAges;
        public int lcCCclimax;
        public int lcCCcycle;
        public int lcCCchalange;

        #endregion LifeCycles

        #region Others
        public int BssnssSingle;
        public int BssnssMulti;
        public int Marriage;
        public int Learn;
        public int ADHD;
        public int Health;
        #endregion

        private int iNoData = -999;
        private string sNoData = "-999";

        public FullResults()
        {
            BLL.Calc c = new BLL.Calc();
            try
            {
                //c.GetCorrectNumberFromSplitedString(Reports.ReportDataProvider.Instance.mMainForm.Controls.Find("txtNum1", true)[0].Text.Trim().Split(c.Delimiter));
                #region Chakra
                Crown          = c.GetCorrectNumberFromSplitedString(Reports.ReportDataProvider.Instance.mMainForm.Controls.Find("txtNum1", true)[0].Text.Trim().Split(c.Delimiter));
                ThirdEye       = c.GetCorrectNumberFromSplitedString(Reports.ReportDataProvider.Instance.mMainForm.Controls.Find("txtNum2", true)[0].Text.Trim().Split(c.Delimiter));
                Throught       = c.GetCorrectNumberFromSplitedString(Reports.ReportDataProvider.Instance.mMainForm.Controls.Find("txtNum3", true)[0].Text.Trim().Split(c.Delimiter));
                Heart          = c.GetCorrectNumberFromSplitedString(Reports.ReportDataProvider.Instance.mMainForm.Controls.Find("txtNum4", true)[0].Text.Trim().Split(c.Delimiter));
                Sun            = c.GetCorrectNumberFromSplitedString(Reports.ReportDataProvider.Instance.mMainForm.Controls.Find("txtNum5", true)[0].Text.Trim().Split(c.Delimiter));
                SexAndCreation = c.GetCorrectNumberFromSplitedString(Reports.ReportDataProvider.Instance.mMainForm.Controls.Find("txtNum6", true)[0].Text.Trim().Split(c.Delimiter));
                Base           = c.GetCorrectNumberFromSplitedString(Reports.ReportDataProvider.Instance.mMainForm.Controls.Find("txtNum7", true)[0].Text.Trim().Split(c.Delimiter));
                Superior       = c.GetCorrectNumberFromSplitedString(Reports.ReportDataProvider.Instance.mMainForm.Controls.Find("txtNum8", true)[0].Text.Trim().Split(c.Delimiter));
                Universal      = c.GetCorrectNumberFromSplitedString(Reports.ReportDataProvider.Instance.mMainForm.Controls.Find("txtNum9", true)[0].Text.Trim().Split(c.Delimiter));

                string txt = Reports.ReportDataProvider.Instance.mMainForm.Controls.Find("txtAstroName", true)[0].Text.Trim();

                Astro = Convert.ToInt16(txt.Split(" ".ToCharArray()[0])[1].Substring(1, 2).Replace(")", ""));
                AstroName = c.GetAstroLuckNameByNumber(Astro);

                txt = (Reports.ReportDataProvider.Instance.mMainForm.Controls.Find("txtInfo1", true)[0].Text.Trim()).Split(Environment.NewLine.ToCharArray()[0])[0].ToString();
                    Balanced = Enums.EnumProvider.Balance._לא_מאוזן_;
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                {
                    if (txt == "מאוזן")
                    {
                        Balanced = Enums.EnumProvider.Balance._מאוזן_;
                    }
                    if (txt == "לא מאוזן")
                    {
                        Balanced = Enums.EnumProvider.Balance._לא_מאוזן_;
                    }
                    if (txt == "מאוזן חלקית")
                    {
                        Balanced = Enums.EnumProvider.Balance._מאוזן_חלקית_;
                    }
                }
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                {
                    if (txt.ToLower() == "balanced")
                    {
                        Balanced = Enums.EnumProvider.Balance._מאוזן_;
                    }
                    if (txt.ToLower() == "no balanced")
                    {
                        Balanced = Enums.EnumProvider.Balance._לא_מאוזן_;
                    }
                    if (txt.ToLower() == "half balanced")
                    {
                        Balanced = Enums.EnumProvider.Balance._מאוזן_חלקית_;
                    }
                }
                #endregion Chakra

                #region Life Cycles
                lcAges1 = Convert.ToInt16(Reports.ReportDataProvider.Instance.mMainForm.Controls.Find("txt1_1", true)[0].Text.Trim().Split("-".ToCharArray()[0])[1]);
                lcAges2 = Convert.ToInt16(Reports.ReportDataProvider.Instance.mMainForm.Controls.Find("txt2_1", true)[0].Text.Trim().Split("-".ToCharArray()[0])[1]);
                lcAges3 = Convert.ToInt16(Reports.ReportDataProvider.Instance.mMainForm.Controls.Find("txt3_1", true)[0].Text.Trim().Split("-".ToCharArray()[0])[1]);
                lcAges4 = Convert.ToInt16(Reports.ReportDataProvider.Instance.mMainForm.Controls.Find("txt4_1", true)[0].Text.Trim().Split("-".ToCharArray()[0])[1]);

                lcCycle1 = c.GetCorrectNumberFromSplitedString(Reports.ReportDataProvider.Instance.mMainForm.Controls.Find("txt1_2", true)[0].Text.Trim().Split(c.Delimiter));
                lcCycle2 = c.GetCorrectNumberFromSplitedString(Reports.ReportDataProvider.Instance.mMainForm.Controls.Find("txt2_2", true)[0].Text.Trim().Split(c.Delimiter));
                lcCycle3 = c.GetCorrectNumberFromSplitedString(Reports.ReportDataProvider.Instance.mMainForm.Controls.Find("txt3_2", true)[0].Text.Trim().Split(c.Delimiter));
                lcCycle4 = c.GetCorrectNumberFromSplitedString(Reports.ReportDataProvider.Instance.mMainForm.Controls.Find("txt4_2", true)[0].Text.Trim().Split(c.Delimiter));
                
                lcClimax1 = c.GetCorrectNumberFromSplitedString(Reports.ReportDataProvider.Instance.mMainForm.Controls.Find("txt1_3", true)[0].Text.Trim().Split(c.Delimiter));
                lcClimax2 = c.GetCorrectNumberFromSplitedString(Reports.ReportDataProvider.Instance.mMainForm.Controls.Find("txt2_3", true)[0].Text.Trim().Split(c.Delimiter));
                lcClimax3 = c.GetCorrectNumberFromSplitedString(Reports.ReportDataProvider.Instance.mMainForm.Controls.Find("txt3_3", true)[0].Text.Trim().Split(c.Delimiter));
                lcClimax4 = c.GetCorrectNumberFromSplitedString(Reports.ReportDataProvider.Instance.mMainForm.Controls.Find("txt4_3", true)[0].Text.Trim().Split(c.Delimiter));

                lcChalange1 = c.GetCorrectNumberFromSplitedString(Reports.ReportDataProvider.Instance.mMainForm.Controls.Find("txt1_4", true)[0].Text.Trim().Split(c.Delimiter));
                lcChalange2 = c.GetCorrectNumberFromSplitedString(Reports.ReportDataProvider.Instance.mMainForm.Controls.Find("txt2_4", true)[0].Text.Trim().Split(c.Delimiter));
                lcChalange3 = c.GetCorrectNumberFromSplitedString(Reports.ReportDataProvider.Instance.mMainForm.Controls.Find("txt3_4", true)[0].Text.Trim().Split(c.Delimiter));
                lcChalange4 = c.GetCorrectNumberFromSplitedString(Reports.ReportDataProvider.Instance.mMainForm.Controls.Find("txt4_4", true)[0].Text.Trim().Split(c.Delimiter));

                currentCycle = CalcCurrentCycle();
                lcCCAges     = c.GetCorrectNumberFromSplitedString(Reports.ReportDataProvider.Instance.mMainForm.Controls.Find("txt" + currentCycle.ToString() + "_1", true)[0].Text.Trim().Split(c.Delimiter));
                lcCCclimax   = c.GetCorrectNumberFromSplitedString(Reports.ReportDataProvider.Instance.mMainForm.Controls.Find("txt" + currentCycle.ToString() + "_2", true)[0].Text.Trim().Split(c.Delimiter));
                lcCCcycle    = c.GetCorrectNumberFromSplitedString(Reports.ReportDataProvider.Instance.mMainForm.Controls.Find("txt" + currentCycle.ToString() + "_3", true)[0].Text.Trim().Split(c.Delimiter));
                lcCCchalange = c.GetCorrectNumberFromSplitedString(Reports.ReportDataProvider.Instance.mMainForm.Controls.Find("txt" + currentCycle.ToString() + "_4", true)[0].Text.Trim().Split(c.Delimiter));

                #endregion Life Cycles

                #region Others
                BssnssSingle = c.GetCorrectNumberFromSplitedString(Reports.ReportDataProvider.Instance.mMainForm.Controls.Find("txtFinalBusinessMark", true)[0].Text.Trim().Split(c.Delimiter));
                if (Reports.ReportDataProvider.Instance.mMainForm.Controls.Find("txtFinalMultipleBusineesMartk", true)[0].Text.Trim().Length > 0)
                {
                    BssnssMulti = c.GetCorrectNumberFromSplitedString(Reports.ReportDataProvider.Instance.mMainForm.Controls.Find("txtFinalMultipleBusineesMartk", true)[0].Text.Trim().Split(c.Delimiter)); 
                }
                if (Reports.ReportDataProvider.Instance.mMainForm.Controls.Find("txtMrrgMark", true)[0].Text.Trim().Length > 0)
                {
                    Marriage = c.GetCorrectNumberFromSplitedString(Reports.ReportDataProvider.Instance.mMainForm.Controls.Find("txtMrrgMark", true)[0].Text.Trim().Split(c.Delimiter));
                }
                Learn = c.GetCorrectNumberFromSplitedString(Reports.ReportDataProvider.Instance.mMainForm.Controls.Find("txtLeanSccss", true)[0].Text.Trim().Split(c.Delimiter));
                ADHD = c.GetCorrectNumberFromSplitedString(Reports.ReportDataProvider.Instance.mMainForm.Controls.Find("txtAtt", true)[0].Text.Trim().Split(c.Delimiter));
                Health = c.GetCorrectNumberFromSplitedString(Reports.ReportDataProvider.Instance.mMainForm.Controls.Find("txtHealthValue", true)[0].Text.Trim().Split(c.Delimiter));
                #endregion
            }
            catch //(Exception excp)
            {
                #region Chakra
                Crown = iNoData;
                Throught = iNoData;
                ThirdEye = iNoData;
                Heart = iNoData;
                SexAndCreation = iNoData;
                Sun = iNoData;
                Base = iNoData;
                Universal = iNoData;
                Superior = iNoData;

                Astro = iNoData;
                AstroName = sNoData;

                Balanced = Enums.EnumProvider.Balance._לא_מאוזן_;
                #endregion Chakra

                #region Life Cycles
                lcAges1 = iNoData;
                lcAges2 = iNoData;
                lcAges3 = iNoData;
                lcAges4 = iNoData;

                lcClimax1 = iNoData;
                lcClimax2 = iNoData;
                lcClimax3 = iNoData;
                lcClimax4 = iNoData;

                lcCycle1 = iNoData;
                lcCycle2 = iNoData;
                lcCycle3 = iNoData;
                lcCycle4 = iNoData;

                lcChalange1 = iNoData;
                lcChalange2 = iNoData;
                lcChalange3 = iNoData;
                lcChalange4 = iNoData;

                currentCycle = iNoData;
                lcCCAges = iNoData;
                lcCCclimax = iNoData;
                lcCCcycle = iNoData;
                lcCCchalange = iNoData;
                #endregion Life Cycles

                #region Others
                BssnssSingle = iNoData;
                BssnssMulti = iNoData;
                Marriage = iNoData;
                Learn = iNoData;
                ADHD = iNoData;
                Health = iNoData;
                #endregion
            }
        }

        private int CalcCurrentCycle()
        {
            int age = Convert.ToInt16(Reports.ReportDataProvider.Instance.mMainForm.Controls.Find("txtAge", true)[0].Text.Split("(".ToCharArray()[0])[0]);

            int curCycle = 0;
            int[] CycleBoundries = new int[4];
            CycleBoundries[0] = Convert.ToInt16(Reports.ReportDataProvider.Instance.mMainForm.Controls.Find("txt1_1", true)[0].Text.Split("-".ToCharArray()[0])[1].Trim());
            CycleBoundries[1] = Convert.ToInt16(Reports.ReportDataProvider.Instance.mMainForm.Controls.Find("txt2_1", true)[0].Text.Split("-".ToCharArray()[0])[1].Trim());
            CycleBoundries[2] = Convert.ToInt16(Reports.ReportDataProvider.Instance.mMainForm.Controls.Find("txt3_1", true)[0].Text.Split("-".ToCharArray()[0])[1].Trim());
            CycleBoundries[3] = Convert.ToInt16(Reports.ReportDataProvider.Instance.mMainForm.Controls.Find("txt4_1", true)[0].Text.Split("-".ToCharArray()[0])[1].Trim());

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
    }
}

