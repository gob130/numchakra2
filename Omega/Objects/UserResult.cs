using System;
using System.Collections.Generic;
using System.Text;

namespace Omega.Objects
{
    public class UserResult
    {
        public string Crown;
        public string Throught;
        public string ThirdEye;
        public string Heart;
        public string SexAndCreation;
        public string Sun;
        public string Base;
        public string PrivateNameNum;
        public string Astro;
        
        public string LC_cycle;
        public string LC_Climax;
        public string LC_Chalange;

        public string LC_cycle1, LC_cycle2, LC_cycle3, LC_cycle4;
        public string LC_chlng1, LC_chlng2, LC_chlng3, LC_chlng4;
        public string LC_climx1, LC_climx2, LC_climx3, LC_climx4;

        public Omega.Enums.EnumProvider.Balance MapBalanceStatus;
        public string stryMarr, stryBssnss,stryLearn,stryADHD,stryHealth;
        public double mrkMarr, mrkBssnss,mrkLearn,mrkADHD,mrkHealth;

        public bool isMapStrong;

        public DateTime birthday;

        public UserResult()
        {
            Crown = "";
            Throught = "";
            ThirdEye = "";
            Heart = "";
            SexAndCreation = "";
            Sun = "";
            Base = "";

            Astro = "";
            PrivateNameNum = "";

            LC_cycle = "";
            LC_Climax = "";
            LC_Chalange = "";

            birthday = new DateTime(2000, 1, 1);

        }

        public UserResult(string crown, string throught, string thirdeye, string heart, string sexcreation, string sun, string basis)
        {
            Crown = crown;
            Throught = throught;
            ThirdEye = thirdeye;
            Heart = heart;
            SexAndCreation = sexcreation ;
            Sun = sun;
            Base = basis;
        }

        public bool GatherDataFromGUI(System.Windows.Forms.Control.ControlCollection Controls,int currentcycle)
        {
            bool res = true;
            string str = "";
            try
            {
                Crown = Controls.Find("txtNum1", true)[0].Text;
                ThirdEye = Controls.Find("txtNum2", true)[0].Text;
                Throught = Controls.Find("txtNum3", true)[0].Text;
                Heart = Controls.Find("txtNum4", true)[0].Text;
                Sun = Controls.Find("txtNum5", true)[0].Text;
                SexAndCreation = Controls.Find("txtNum6", true)[0].Text;
                Base = Controls.Find("txtNum7", true)[0].Text;
                PrivateNameNum = Controls.Find("txtPName_Num", true)[0].Text;
                Astro = Controls.Find("txtAstroName", true)[0].Text;

                #region Balance
                str = Controls.Find("txtInfo1", true)[0].Text.Split(Environment.NewLine.ToCharArray()[0])[1];
                if (BLL.AppSettings.Instance.ProgramLanguage == BLL.AppSettings.Language.English)
                {
                    if (str == Enums.EnumProvider.Instance.GetBalanceEnumFromDescription(Enums.EnumProvider.Balance._מאוזן_.ToString()))
                    {
                        MapBalanceStatus = Enums.EnumProvider.Balance._מאוזן_;
                    }
                    if (str == Enums.EnumProvider.Instance.GetBalanceEnumFromDescription(Enums.EnumProvider.Balance._לא_מאוזן_.ToString()))
                    {
                        MapBalanceStatus = Enums.EnumProvider.Balance._לא_מאוזן_;
                    }
                    if (str == Enums.EnumProvider.Instance.GetBalanceEnumFromDescription(Enums.EnumProvider.Balance._מאוזן_חלקית_.ToString()))
                    {
                        MapBalanceStatus = Enums.EnumProvider.Balance._מאוזן_חלקית_ ;
                    }
                }
                if (BLL.AppSettings.Instance.ProgramLanguage == BLL.AppSettings.Language.English)
                {
                    if (str == "מאוזן")
                    {
                        MapBalanceStatus = Enums.EnumProvider.Balance._מאוזן_;
                    }
                    if (str == "לא מאוזן")
                    {
                        MapBalanceStatus = Enums.EnumProvider.Balance._לא_מאוזן_;
                    }
                    if (str == "חצי מאוזן")
                    {
                        MapBalanceStatus = Enums.EnumProvider.Balance._מאוזן_חלקית_ ;
                    }
                }
                #endregion Balance

                if (bool.TryParse(Controls.Find("txtMapStrong", true)[0].Text,out isMapStrong))
                {
                }
                else
                {
                }

                LC_chlng1 = Controls.Find("txt1_4", true)[0].Text;
                LC_chlng2 = Controls.Find("txt2_4", true)[0].Text;
                LC_chlng3 = Controls.Find("txt3_4", true)[0].Text;
                LC_chlng4 = Controls.Find("txt4_4", true)[0].Text;

                LC_cycle1 = Controls.Find("txt1_2", true)[0].Text;
                LC_cycle2 = Controls.Find("txt2_2", true)[0].Text;
                LC_cycle3 = Controls.Find("txt3_2", true)[0].Text;
                LC_cycle4 = Controls.Find("txt4_2", true)[0].Text;

                LC_climx1 = Controls.Find("txt1_3", true)[0].Text;
                LC_climx2 = Controls.Find("txt2_3", true)[0].Text;
                LC_climx3 = Controls.Find("txt3_3", true)[0].Text;
                LC_climx4 = Controls.Find("txt4_3", true)[0].Text;

                LC_Climax = Controls.Find("txt" + currentcycle.ToString() + "_3", true)[0].Text;
                LC_Chalange  = Controls.Find("txt" + currentcycle.ToString() + "_4", true)[0].Text;
                LC_cycle = Controls.Find("txt" + currentcycle.ToString() + "_2", true)[0].Text;



                mrkBssnss = Convert.ToDouble(Controls.Find("txtFinalBusinessMark", true)[0].Text);
                stryBssnss = Controls.Find("txtBusinessStory", true)[0].Text;

                mrkMarr = Convert.ToDouble( Controls.Find("txtMrrgMark", true)[0].Text);
                stryMarr = Controls.Find("txtMrrgStory", true)[0].Text;

                mrkLearn = Convert.ToDouble(Controls.Find("txtLeanSccss", true)[0].Text);
                stryLearn = Controls.Find("txtLearnStory", true)[0].Text;

                mrkADHD = Convert.ToDouble(Controls.Find("txtAtt", true)[0].Text);
                stryADHD = Controls.Find("txtAttStory", true)[0].Text;

                mrkHealth = Convert.ToDouble(Controls.Find("txtHealthValue", true)[0].Text);
                stryHealth = Controls.Find("txtHealthStory", true)[0].Text;
                
            }
            catch //(Exception exp)
            {
                res = false;
            }
            return res;
        }
    }
}
