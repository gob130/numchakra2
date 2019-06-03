using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Management;
using System.Text;

namespace BLL
{
    public class Calc
    {
        #region Private Members
        //private System.Data.OleDb.OleDbConnection mDBConnection;
        //private ADODB.Connection mDBConnection;
        //private string mDBLocation;
        //private string mDBName;

        bool isPersonMaster;

        private char mDelimiter = "=".ToCharArray()[0];
        private int[] mDelayingNumbers = new int[] { 2, 7 }; // Self destruction
        private int[] mMasterNumbers = new int[3] { 9, 11, 22 };
        private int[] mCarmaticNumbers = new int[4] { 13, 14, 16, 19 };
        private int[] mHalfCarmaticNumbers = new int[1] { 15 };
        //private int[] mDynamicChakraOpenValues = new int[9] { 0, 2, 6, 7, 8, 9, 11, 22, 33 };
        private int[][] mDynamicChakraOpenValues = new int[][]
        {
            new int[] { 1, 3, 4, 5, 6, 8, 9, 11, 22, 33 }, // Crown
            new int[] { 1, 2, 3, 4, 5, 6, 8, 9, 11, 22, 33 }, // Third Eye
            new int[] { 0, 1, 3, 4, 5, 6, 8, 9, 22, 33 }, // Throat
            new int[] { 1, 2, 3, 4, 5, 6, 7, 8 }, // Heart
            new int[] { 1, 3, 4, 5, 6, 8, 9, 22, 33 }, // Solar Plexus
            new int[] { 1, 3, 5, 8 }, // Sex and Creation
            new int[] { 1, 2, 3, 4, 5, 6, 7, 8, 9, 11, 22, 33 }, // Root/Base
            new int[] { 1, 2, 3, 4, 5, 6, 7, 8, 9, 11, 22, 33 }, // Meta/firstname
            new int[] { 1, 3, 4, 5, 6, 8, 9, 22, 33 } // Universe / Astrologic luck

        };

        //private string mDBPassword;

        //private string mClientTableName = "Clients";
        //private string mPSWRDTableName = "PSS";
        //private string mNumsTableName = "Num_Tables";
        //private string mLucksTableName = "Luck_Table";

        private int[] Letters = new int[27]; // Letters Array by ASCII Code
        private string[] Letters_Master = new string[2];
        private string[] Letters_Crown = new string[3];
        private string[] Letters_ThirdEye = new string[2];
        private string[] Letters_Throught = new string[5];
        private string[] Letters_Heart = new string[2];
        private string[] Letters_Sun = new string[2];
        private string[] Letters_Sex_Creation = new string[5];
        private string[] Letters_Root = new string[3];

        private AstroLuck AstroData = new AstroLuck();
        private Proffing mProffer = new Proffing();

        private int[,] mHealth = new int[1, 1];

        private int[,] mCombHramonic = new int[39, 2];
        private int[,] mCombHalfHramonic = new int[20, 2];
        private int[,] mCombDissHramonic = new int[19, 2];

        private int[,] mStreetCombHramonic = new int[38, 2];
        private int[,] mStreetCombHalfHramonic = new int[21, 2];
        private int[,] mStreetCombDissHramonic = new int[19, 2];

        private string[,] mBusinessChakraChart = new string[23, 2];
        private string[,] mBusinessLifeCycleChart = new string[23, 2];

        private Dictionary<string, int> mSalesChakraChart = new Dictionary<string, int>();
        private Dictionary<string, int> mSecretaryAndAccountingChakraChart = new Dictionary<string, int>();
        private Dictionary<string, int> mAlternativingChakraChart = new Dictionary<string, int>();
        private Dictionary<string, int> mBeautyChakraChart = new Dictionary<string, int>();
        private Dictionary<string, int> mManualWorkChakraChart = new Dictionary<string, int>();
        private Dictionary<string, double> mHiTechChakraChart = new Dictionary<string, double>();
        private Dictionary<string, double> mLowTechChakraChart = new Dictionary<string, double>();

        private string[,] mRelationCompetability_TermLong = new string[138, 3];
        private string[,] mRelationCompetability_TermShort = new string[138, 3];
        private Tuple<int, int, int>[] mSexualCompatability = new Tuple<int, int, int>[12];

        private string[,] mLearnSccssAttPrbl_Chkra = new string[26, 2];
        private string[,] mLearnSccssAttPrbl_LC = new string[17, 2];

        private int[,] mHealthTable;


        private int[] mChakra1_OpeningValues; // Master & Crown
        private int[] mChakra2_OpeningValues; // Third Eye
        private int[] mChakra3_OpeningValues; // Throat
        private int[] mChakra4_OpeningValues; // Heart
        private int[] mChakra5_OpeningValues; // Solar Plexus
        private int[] mChakra6_OpeningValues; // Sex Creation
        private int[] mChakra7_OpeningValues; // Root
        #endregion //Private Members

        #region Constructor
        public Calc(AppSettings.Language lang, bool isMaster)
        {
            start(lang);
            isPersonMaster = isMaster;
        }

        public Calc()
        {
            start(AppSettings.Language.Hebrew);
            isPersonMaster = true;
        }

        private void start(AppSettings.Language lang)
        {
            //mDBLocation = System.IO.Directory.GetCurrentDirectory();
            //mDBName = "DB.mdb";
            //mDBPassword = "Idan_Neta_Omer_Yaakobi_31415_1935_dyi";

            switch (lang)
            {
                case AppSettings.Language.Hebrew:
                    #region Letters Init.
                    Letters[0] = 1;// א
                    Letters[1] = 2;//ב
                    Letters[2] = 3; // ג
                    Letters[3] = 4; // ד
                    Letters[4] = 5; // ה
                    Letters[5] = 6;// ו
                    Letters[6] = 7; //ז
                    Letters[7] = 8; //ח
                    Letters[8] = 9;//ט

                    Letters[9] = 1;// י
                    Letters[10] = 2;// ך
                    Letters[11] = 2;// כ
                    Letters[12] = 3;//ל
                    Letters[13] = 4;// ם
                    Letters[14] = 4;// n
                    Letters[15] = 5; //ן
                    Letters[16] = 5; //נ
                    Letters[17] = 6;//ס
                    Letters[18] = 7;//ע
                    Letters[19] = 8;//ף
                    Letters[20] = 8;//פ
                    Letters[21] = 9;// ץ
                    Letters[22] = 9;//צ

                    Letters[23] = 1;// ק
                    Letters[24] = 2;//ר
                    Letters[25] = 3;//ש
                    Letters[26] = 4;//ת
                    #endregion //Letters Init.

                    #region Chakra_Letters
                    Letters_Master[0] = "ט";// -(int)"א";
                    Letters_Master[1] = "צ";// - (int)"א";

                    Letters_Crown[0] = "ב";// - (int)"א";
                    Letters_Crown[1] = "כ";// - (int)"א";
                    Letters_Crown[2] = "ר";// - (int)"א";

                    Letters_ThirdEye[0] = "ז";// - (int)"א";
                    Letters_ThirdEye[1] = "ע";// - (int)"א";

                    Letters_Throught[0] = "ג";// - (int)"א";
                    Letters_Throught[1] = "ל";// - (int)"א";
                    Letters_Throught[2] = "ש";// - (int)"א";
                    Letters_Throught[3] = "ה";// - (int)"א";
                    Letters_Throught[4] = "נ";// - (int)"א";

                    Letters_Heart[0] = "ו";// - (int)"א";
                    Letters_Heart[1] = "ס";// - (int)"א";

                    Letters_Sun[0] = "ט";// - (int)"א";
                    Letters_Sun[1] = "צ";// - (int)"א";

                    Letters_Sex_Creation[0] = "א";// - (int)"א";
                    Letters_Sex_Creation[1] = "י";// - (int)"א";
                    Letters_Sex_Creation[2] = "ק";// - (int)"א";        THOL
                    Letters_Sex_Creation[3] = "ח";// - (int)"א";        THOL
                    Letters_Sex_Creation[4] = "פ";// - (int)"א";        THOL

                    Letters_Root[0] = "ד";// - (int)"א";
                    Letters_Root[1] = "מ";// - (int)"א";
                    Letters_Root[2] = "ת";// - (int)"א";
                    #endregion

                    break;
                case AppSettings.Language.English:
                    Letters = new int[26];
                    #region Letters Init.
                    Letters[0] = 1;// a
                    Letters[1] = 2;//b
                    Letters[2] = 3; //c 
                    Letters[3] = 4; // d
                    Letters[4] = 5; // e
                    Letters[5] = 6;// f
                    Letters[6] = 7; //g
                    Letters[7] = 8; //h
                    Letters[8] = 9;//I

                    Letters[9] = 1;// J
                    Letters[10] = 2;// K
                    Letters[11] = 3;// l
                    Letters[12] = 4;//M
                    Letters[13] = 5;// N
                    Letters[14] = 6;// O
                    Letters[15] = 7; //P
                    Letters[16] = 8; //Q
                    Letters[17] = 9;//R

                    Letters[18] = 1;//S
                    Letters[19] = 2;//T
                    Letters[20] = 3;//U
                    Letters[21] = 4;// V
                    Letters[22] = 5;//W
                    Letters[23] = 6;// X
                    Letters[24] = 7;//Y
                    Letters[25] = 8;//Z
                    #endregion //Letters Init.

                    #region Chaktra_Letters
                    Letters_Master = new string[2] { "", "" };
                    Letters_Crown = new string[3] { "b", "k", "t" };
                    Letters_ThirdEye = new string[3] { "g", "p", "y" };
                    Letters_Throught = new string[6] { "c", "l", "u", "e", "n", "w" };
                    Letters_Heart = new string[3] { "f", "o", "x" };
                    Letters_Sun = new string[2] { "i", "r" };
                    Letters_Sex_Creation = new string[6] { "a", "j", "s", "h", "q", "z" };
                    Letters_Root = new string[3] { "d", "m", "v" };
                    #endregion Chaktra Opening

                    break;
            }

            #region Hramonic_Diss_Half_Combinations
            #region Hramonic
            mCombHramonic[0, 0] = 1; mCombHramonic[0, 1] = 3;

            mCombHramonic[1, 0] = 4; mCombHramonic[1, 1] = 1;
            mCombHramonic[2, 0] = 4; mCombHramonic[2, 1] = 2;

            mCombHramonic[3, 0] = 5; mCombHramonic[3, 1] = 1;
            mCombHramonic[4, 0] = 5; mCombHramonic[4, 1] = 3;

            mCombHramonic[5, 0] = 6; mCombHramonic[5, 1] = 2;
            mCombHramonic[6, 0] = 6; mCombHramonic[6, 1] = 3;
            mCombHramonic[7, 0] = 6; mCombHramonic[7, 1] = 4;

            mCombHramonic[8, 0] = 8; mCombHramonic[8, 1] = 1;
            mCombHramonic[9, 0] = 8; mCombHramonic[9, 1] = 3;
            mCombHramonic[10, 0] = 8; mCombHramonic[10, 1] = 4;
            mCombHramonic[11, 0] = 8; mCombHramonic[11, 1] = 5;
            mCombHramonic[12, 0] = 8; mCombHramonic[12, 1] = 6;

            mCombHramonic[13, 0] = 9; mCombHramonic[13, 1] = 1;
            mCombHramonic[14, 0] = 9; mCombHramonic[14, 1] = 3;
            mCombHramonic[15, 0] = 9; mCombHramonic[15, 1] = 5;
            mCombHramonic[16, 0] = 9; mCombHramonic[16, 1] = 6;
            mCombHramonic[17, 0] = 9; mCombHramonic[17, 1] = 8;

            mCombHramonic[18, 0] = 11; mCombHramonic[18, 1] = 1;
            mCombHramonic[19, 0] = 11; mCombHramonic[19, 1] = 4;
            mCombHramonic[20, 0] = 11; mCombHramonic[20, 1] = 3;
            mCombHramonic[21, 0] = 11; mCombHramonic[21, 1] = 5;
            mCombHramonic[22, 0] = 11; mCombHramonic[22, 1] = 6;
            mCombHramonic[23, 0] = 11; mCombHramonic[23, 1] = 8;
            mCombHramonic[24, 0] = 11; mCombHramonic[24, 1] = 9;

            mCombHramonic[25, 0] = 22; mCombHramonic[25, 1] = 1;
            mCombHramonic[26, 0] = 22; mCombHramonic[26, 1] = 3;
            mCombHramonic[27, 0] = 22; mCombHramonic[27, 1] = 5;
            mCombHramonic[28, 0] = 22; mCombHramonic[28, 1] = 6;
            mCombHramonic[29, 0] = 22; mCombHramonic[29, 1] = 8;
            mCombHramonic[30, 0] = 22; mCombHramonic[30, 1] = 9;
            mCombHramonic[31, 0] = 22; mCombHramonic[31, 1] = 11;

            mCombHramonic[32, 0] = 33; mCombHramonic[32, 1] = 2;
            mCombHramonic[33, 0] = 33; mCombHramonic[33, 1] = 3;
            mCombHramonic[34, 0] = 33; mCombHramonic[34, 1] = 4;
            mCombHramonic[35, 0] = 33; mCombHramonic[35, 1] = 8;
            mCombHramonic[36, 0] = 33; mCombHramonic[36, 1] = 9;
            mCombHramonic[37, 0] = 33; mCombHramonic[37, 1] = 11;
            mCombHramonic[38, 0] = 33; mCombHramonic[38, 1] = 22;
            #endregion Hramonic

            #region HalfHarmonic
            mCombHalfHramonic[0, 0] = 3; mCombHalfHramonic[0, 1] = 3;

            mCombHalfHramonic[1, 0] = 5; mCombHalfHramonic[1, 1] = 5;

            mCombHalfHramonic[2, 0] = 6; mCombHalfHramonic[2, 1] = 1;

            mCombHalfHramonic[3, 0] = 7; mCombHalfHramonic[3, 1] = 1;
            mCombHalfHramonic[4, 0] = 7; mCombHalfHramonic[4, 1] = 2;
            mCombHalfHramonic[5, 0] = 7; mCombHalfHramonic[5, 1] = 4;

            mCombHalfHramonic[6, 0] = 8; mCombHalfHramonic[6, 1] = 2;
            mCombHalfHramonic[7, 0] = 8; mCombHalfHramonic[7, 1] = 7;

            mCombHalfHramonic[08, 0] = 9; mCombHalfHramonic[08, 1] = 2;
            mCombHalfHramonic[09, 0] = 9; mCombHalfHramonic[09, 1] = 4;
            mCombHalfHramonic[10, 0] = 9; mCombHalfHramonic[10, 1] = 7;
            mCombHalfHramonic[11, 0] = 9; mCombHalfHramonic[11, 1] = 9;

            mCombHalfHramonic[12, 0] = 11; mCombHalfHramonic[12, 1] = 2;
            mCombHalfHramonic[13, 0] = 11; mCombHalfHramonic[13, 1] = 7;
            mCombHalfHramonic[14, 0] = 11; mCombHalfHramonic[14, 1] = 11;

            mCombHalfHramonic[15, 0] = 22; mCombHalfHramonic[15, 1] = 2;
            mCombHalfHramonic[16, 0] = 22; mCombHalfHramonic[16, 1] = 4;
            mCombHalfHramonic[17, 0] = 22; mCombHalfHramonic[17, 1] = 7;
            mCombHalfHramonic[18, 0] = 22; mCombHalfHramonic[18, 1] = 22;

            mCombHalfHramonic[19, 0] = 33; mCombHalfHramonic[19, 1] = 1;
            #endregion

            #region Diss_Harmonic
            mCombDissHramonic[0, 0] = 1; mCombDissHramonic[0, 1] = 1;

            mCombDissHramonic[1, 0] = 2; mCombDissHramonic[1, 1] = 1;
            mCombDissHramonic[2, 0] = 2; mCombDissHramonic[2, 1] = 2;

            mCombDissHramonic[3, 0] = 3; mCombDissHramonic[3, 1] = 2;

            mCombDissHramonic[4, 0] = 4; mCombDissHramonic[4, 1] = 3;
            mCombDissHramonic[5, 0] = 4; mCombDissHramonic[5, 1] = 4;

            mCombDissHramonic[6, 0] = 5; mCombDissHramonic[6, 1] = 2;
            mCombDissHramonic[7, 0] = 5; mCombDissHramonic[7, 1] = 4;

            mCombDissHramonic[8, 0] = 6; mCombDissHramonic[8, 1] = 5;
            mCombDissHramonic[9, 0] = 6; mCombDissHramonic[9, 1] = 6;

            mCombDissHramonic[10, 0] = 7; mCombDissHramonic[10, 1] = 3;
            mCombDissHramonic[11, 0] = 7; mCombDissHramonic[11, 1] = 5;
            mCombDissHramonic[12, 0] = 7; mCombDissHramonic[12, 1] = 6;
            mCombDissHramonic[13, 0] = 7; mCombDissHramonic[13, 1] = 7;

            mCombDissHramonic[14, 0] = 8; mCombDissHramonic[14, 1] = 8;

            mCombDissHramonic[15, 0] = 33; mCombDissHramonic[15, 1] = 5;
            mCombDissHramonic[16, 0] = 33; mCombDissHramonic[16, 1] = 6;
            mCombDissHramonic[17, 0] = 33; mCombDissHramonic[17, 1] = 7;
            mCombDissHramonic[18, 0] = 33; mCombDissHramonic[18, 1] = 33;
            #endregion Diss_Harmonic

            mStreetCombHramonic = mCombHramonic;
            mStreetCombHalfHramonic = mCombHalfHramonic;
            mStreetCombDissHramonic = mCombDissHramonic;
            #endregion

            #region Business

            mBusinessChakraChart[0, 0] = "1"; mBusinessChakraChart[0, 1] = "10";
            mBusinessChakraChart[1, 0] = "2"; mBusinessChakraChart[1, 1] = "5";
            mBusinessChakraChart[2, 0] = "3"; mBusinessChakraChart[2, 1] = "8";
            mBusinessChakraChart[3, 0] = "4"; mBusinessChakraChart[3, 1] = "7";
            mBusinessChakraChart[4, 0] = "5"; mBusinessChakraChart[4, 1] = "9";
            mBusinessChakraChart[5, 0] = "6"; mBusinessChakraChart[5, 1] = "6";
            mBusinessChakraChart[6, 0] = "7"; mBusinessChakraChart[6, 1] = "5";
            mBusinessChakraChart[7, 0] = "8"; mBusinessChakraChart[7, 1] = "10";
            mBusinessChakraChart[8, 0] = "9"; mBusinessChakraChart[8, 1] = "10";
            mBusinessChakraChart[9, 0] = "11"; mBusinessChakraChart[9, 1] = "10";
            mBusinessChakraChart[10, 0] = "22"; mBusinessChakraChart[10, 1] = "10";
            mBusinessChakraChart[11, 0] = "33"; mBusinessChakraChart[11, 1] = "7";
            mBusinessChakraChart[12, 0] = "13"; mBusinessChakraChart[12, 1] = "7";
            mBusinessChakraChart[13, 0] = "14"; mBusinessChakraChart[13, 1] = "9";
            mBusinessChakraChart[14, 0] = "15"; mBusinessChakraChart[14, 1] = "7";
            mBusinessChakraChart[15, 0] = "16"; mBusinessChakraChart[15, 1] = "1";
            mBusinessChakraChart[16, 0] = "19"; mBusinessChakraChart[16, 1] = "7";
            mBusinessChakraChart[17, 0] = "20"; mBusinessChakraChart[17, 1] = "8";
            mBusinessChakraChart[18, 0] = "30"; mBusinessChakraChart[18, 1] = "8";
            mBusinessChakraChart[19, 0] = "0"; mBusinessChakraChart[19, 1] = "0"; //6,6,1,10
            mBusinessChakraChart[20, 0] = "מאוזן"; mBusinessChakraChart[20, 1] = "10";
            mBusinessChakraChart[21, 0] = "חצי מאוזן"; mBusinessChakraChart[21, 1] = "8";
            mBusinessChakraChart[22, 0] = "לא מאוזן"; mBusinessChakraChart[22, 1] = "4";

            mBusinessLifeCycleChart[0, 0] = "1"; mBusinessLifeCycleChart[0, 1] = "10";
            mBusinessLifeCycleChart[1, 0] = "2"; mBusinessLifeCycleChart[1, 1] = "5";
            mBusinessLifeCycleChart[2, 0] = "3"; mBusinessLifeCycleChart[2, 1] = "8";
            mBusinessLifeCycleChart[3, 0] = "4"; mBusinessLifeCycleChart[3, 1] = "7";
            mBusinessLifeCycleChart[4, 0] = "5"; mBusinessLifeCycleChart[4, 1] = "9";
            mBusinessLifeCycleChart[5, 0] = "6"; mBusinessLifeCycleChart[5, 1] = "6";
            mBusinessLifeCycleChart[6, 0] = "7"; mBusinessLifeCycleChart[6, 1] = "5";
            mBusinessLifeCycleChart[7, 0] = "8"; mBusinessLifeCycleChart[7, 1] = "10";
            mBusinessLifeCycleChart[8, 0] = "9"; mBusinessLifeCycleChart[8, 1] = "10";
            mBusinessLifeCycleChart[9, 0] = "11"; mBusinessLifeCycleChart[9, 1] = "10";
            mBusinessLifeCycleChart[10, 0] = "22"; mBusinessLifeCycleChart[10, 1] = "10";
            mBusinessLifeCycleChart[11, 0] = "33"; mBusinessLifeCycleChart[11, 1] = "7";
            mBusinessLifeCycleChart[12, 0] = "13"; mBusinessLifeCycleChart[12, 1] = "5";
            mBusinessLifeCycleChart[13, 0] = "14"; mBusinessLifeCycleChart[13, 1] = "5";
            mBusinessLifeCycleChart[14, 0] = "15"; mBusinessLifeCycleChart[14, 1] = "7";
            mBusinessLifeCycleChart[15, 0] = "16"; mBusinessLifeCycleChart[15, 1] = "-10";
            mBusinessLifeCycleChart[16, 0] = "19"; mBusinessLifeCycleChart[16, 1] = "5";
            mBusinessLifeCycleChart[17, 0] = "20"; mBusinessLifeCycleChart[17, 1] = "8";
            mBusinessLifeCycleChart[18, 0] = "30"; mBusinessLifeCycleChart[18, 1] = "8";
            mBusinessLifeCycleChart[19, 0] = "0"; mBusinessLifeCycleChart[19, 1] = "0"; //6,6,1,10
            mBusinessLifeCycleChart[20, 0] = "מאוזן"; mBusinessLifeCycleChart[20, 1] = "10";
            mBusinessLifeCycleChart[21, 0] = "חצי מאוזן"; mBusinessLifeCycleChart[21, 1] = "8";
            mBusinessLifeCycleChart[22, 0] = "לא מאוזן"; mBusinessLifeCycleChart[22, 1] = "4";

            // Fill sales chakra chart values
            mSalesChakraChart.Add("0", 6);
            mSalesChakraChart.Add("1", 10);
            mSalesChakraChart.Add("2", 6);
            mSalesChakraChart.Add("3", 9);
            mSalesChakraChart.Add("4", 7);
            mSalesChakraChart.Add("5", 9);
            mSalesChakraChart.Add("6", 6);
            mSalesChakraChart.Add("7", 3);
            mSalesChakraChart.Add("8", 10);
            mSalesChakraChart.Add("9", 10);
            mSalesChakraChart.Add("11", 10);
            mSalesChakraChart.Add("22", 10);
            mSalesChakraChart.Add("33", 6);
            mSalesChakraChart.Add("13", 8);
            mSalesChakraChart.Add("14", 8);
            mSalesChakraChart.Add("16", 3);
            mSalesChakraChart.Add("19", 7);

            //Fill secretary chakra chart values
            mSecretaryAndAccountingChakraChart.Add("0", 8);
            mSecretaryAndAccountingChakraChart.Add("1", 6);
            mSecretaryAndAccountingChakraChart.Add("2", 9);
            mSecretaryAndAccountingChakraChart.Add("3", 8);
            mSecretaryAndAccountingChakraChart.Add("4", 10);
            mSecretaryAndAccountingChakraChart.Add("5", 6);
            mSecretaryAndAccountingChakraChart.Add("6", 10);
            mSecretaryAndAccountingChakraChart.Add("7", 5);
            mSecretaryAndAccountingChakraChart.Add("8", 7);
            mSecretaryAndAccountingChakraChart.Add("9", 8);
            mSecretaryAndAccountingChakraChart.Add("11", 10);
            mSecretaryAndAccountingChakraChart.Add("22", 10);
            mSecretaryAndAccountingChakraChart.Add("33", 10);
            mSecretaryAndAccountingChakraChart.Add("13", 9);
            mSecretaryAndAccountingChakraChart.Add("14", 9);
            mSecretaryAndAccountingChakraChart.Add("16", 9);
            mSecretaryAndAccountingChakraChart.Add("19", 7);
            mSecretaryAndAccountingChakraChart.Add("מאוזן", 10);
            mSecretaryAndAccountingChakraChart.Add("חצי מאוזן",8);
            mSecretaryAndAccountingChakraChart.Add("לא מאוזן", 4);

            //Fill alternative chakra chart values
            mAlternativingChakraChart.Add("0", 9);
            mAlternativingChakraChart.Add("1", 6);
            mAlternativingChakraChart.Add("2", 10);
            mAlternativingChakraChart.Add("3", 8);
            mAlternativingChakraChart.Add("4", 10);
            mAlternativingChakraChart.Add("5", 6);
            mAlternativingChakraChart.Add("6", 10);
            mAlternativingChakraChart.Add("7", 8);
            mAlternativingChakraChart.Add("8", 7);
            mAlternativingChakraChart.Add("9", 8);
            mAlternativingChakraChart.Add("11", 10);
            mAlternativingChakraChart.Add("22", 10);
            mAlternativingChakraChart.Add("33", 10);
            mAlternativingChakraChart.Add("13", 6);
            mAlternativingChakraChart.Add("14", 5);
            mAlternativingChakraChart.Add("16", 9);
            mAlternativingChakraChart.Add("19", 7);
            mAlternativingChakraChart.Add("מאוזן", 7);
            mAlternativingChakraChart.Add("חצי מאוזן", 8);
            mAlternativingChakraChart.Add("לא מאוזן", 9);

           
            //Fill beauty chakra chart values
            mBeautyChakraChart.Add("0", 9);
            mBeautyChakraChart.Add("1", 6);
            mBeautyChakraChart.Add("2", 10);
            mBeautyChakraChart.Add("3", 10);
            mBeautyChakraChart.Add("4", 10);
            mBeautyChakraChart.Add("5", 6);
            mBeautyChakraChart.Add("6", 10);
            mBeautyChakraChart.Add("7", 10);
            mBeautyChakraChart.Add("8", 6);
            mBeautyChakraChart.Add("9", 8);
            mBeautyChakraChart.Add("11", 10);
            mBeautyChakraChart.Add("22", 10);
            mBeautyChakraChart.Add("33", 10);
            mBeautyChakraChart.Add("13", 6);
            mBeautyChakraChart.Add("14", 8);
            mBeautyChakraChart.Add("16", 9);
            mBeautyChakraChart.Add("19", 8);
            mBeautyChakraChart.Add("מאוזן", 7);
            mBeautyChakraChart.Add("חצי מאוזן", 8);
            mBeautyChakraChart.Add("לא מאוזן", 9);

            //Fill manual work chakra chart values
            mManualWorkChakraChart.Add("0", 9);
            mManualWorkChakraChart.Add("1", 6);
            mManualWorkChakraChart.Add("2", 10);
            mManualWorkChakraChart.Add("3", 8);
            mManualWorkChakraChart.Add("4", 10);
            mManualWorkChakraChart.Add("5", 6);
            mManualWorkChakraChart.Add("6", 10);
            mManualWorkChakraChart.Add("7", 8);
            mManualWorkChakraChart.Add("8", 8);
            mManualWorkChakraChart.Add("9", 8);
            mManualWorkChakraChart.Add("11", 10);
            mManualWorkChakraChart.Add("22", 10);
            mManualWorkChakraChart.Add("33", 10);
            mManualWorkChakraChart.Add("13", 9);
            mManualWorkChakraChart.Add("14", 8);
            mManualWorkChakraChart.Add("16", 9);
            mManualWorkChakraChart.Add("19", 8);
            mManualWorkChakraChart.Add("מאוזן", 7);
            mManualWorkChakraChart.Add("חצי מאוזן", 8);
            mManualWorkChakraChart.Add("לא מאוזן", 9);


            //Fill hi-tech chakra chart values
            mHiTechChakraChart.Add("0", 8);
            mHiTechChakraChart.Add("1", 9);
            mHiTechChakraChart.Add("2", 6);
            mHiTechChakraChart.Add("3", 10.5);
            mHiTechChakraChart.Add("4", 6);
            mHiTechChakraChart.Add("5", 10.5);
            mHiTechChakraChart.Add("6", 6);
            mHiTechChakraChart.Add("7", 10);
            mHiTechChakraChart.Add("8", 8);
            mHiTechChakraChart.Add("9", 9);
            mHiTechChakraChart.Add("11", 9);
            mHiTechChakraChart.Add("22", 9);
            mHiTechChakraChart.Add("33", 7);
            mHiTechChakraChart.Add("13", 6);
            mHiTechChakraChart.Add("14", 8);
            mHiTechChakraChart.Add("16", 9);
            mHiTechChakraChart.Add("19", 6);
            mHiTechChakraChart.Add("מאוזן", 10);
            mHiTechChakraChart.Add("חצי מאוזן", 8);
            mHiTechChakraChart.Add("לא מאוזן", 4);

            //Fill low-tech chakra chart values
            mLowTechChakraChart.Add("0", 8);
            mLowTechChakraChart.Add("1", 6);
            mLowTechChakraChart.Add("2", 10);
            mLowTechChakraChart.Add("3", 8);
            mLowTechChakraChart.Add("4", 10);
            mLowTechChakraChart.Add("5", 8);
            mLowTechChakraChart.Add("6", 10);
            mLowTechChakraChart.Add("7", 8);
            mLowTechChakraChart.Add("8", 8);
            mLowTechChakraChart.Add("9", 7);
            mLowTechChakraChart.Add("11", 10);
            mLowTechChakraChart.Add("22", 10);
            mLowTechChakraChart.Add("33", 10);
            mLowTechChakraChart.Add("13", 9);
            mLowTechChakraChart.Add("14", 7);
            mLowTechChakraChart.Add("16", 9);
            mLowTechChakraChart.Add("19", 7);
            mLowTechChakraChart.Add("מאוזן", 7);
            mLowTechChakraChart.Add("חצי מאוזן", 8);
            mLowTechChakraChart.Add("לא מאוזן", 9);
            #endregion

            #region Marriage
            mRelationCompetability_TermLong[0, 0] = "1"; mRelationCompetability_TermLong[0, 1] = "1"; mRelationCompetability_TermLong[0, 2] = "5";
            mRelationCompetability_TermLong[1, 0] = "1"; mRelationCompetability_TermLong[1, 1] = "2"; mRelationCompetability_TermLong[1, 2] = "7";
            mRelationCompetability_TermLong[2, 0] = "1"; mRelationCompetability_TermLong[2, 1] = "3"; mRelationCompetability_TermLong[2, 2] = "5";
            mRelationCompetability_TermLong[3, 0] = "1"; mRelationCompetability_TermLong[3, 1] = "4"; mRelationCompetability_TermLong[3, 2] = "8";
            mRelationCompetability_TermLong[4, 0] = "1"; mRelationCompetability_TermLong[4, 1] = "5"; mRelationCompetability_TermLong[4, 2] = "6";
            mRelationCompetability_TermLong[5, 0] = "1"; mRelationCompetability_TermLong[5, 1] = "6"; mRelationCompetability_TermLong[5, 2] = "8";
            mRelationCompetability_TermLong[6, 0] = "1"; mRelationCompetability_TermLong[6, 1] = "7"; mRelationCompetability_TermLong[6, 2] = "4";
            mRelationCompetability_TermLong[7, 0] = "1"; mRelationCompetability_TermLong[7, 1] = "8"; mRelationCompetability_TermLong[7, 2] = "8";
            mRelationCompetability_TermLong[8, 0] = "1"; mRelationCompetability_TermLong[8, 1] = "9"; mRelationCompetability_TermLong[8, 2] = "10";
            mRelationCompetability_TermLong[9, 0] = "1"; mRelationCompetability_TermLong[9, 1] = "11"; mRelationCompetability_TermLong[9, 2] = "8";
            mRelationCompetability_TermLong[10, 0] = "1"; mRelationCompetability_TermLong[10, 1] = "22"; mRelationCompetability_TermLong[10, 2] = "10";
            mRelationCompetability_TermLong[11, 0] = "1"; mRelationCompetability_TermLong[11, 1] = "33"; mRelationCompetability_TermLong[11, 2] = "8";
            mRelationCompetability_TermLong[12, 0] = "2"; mRelationCompetability_TermLong[12, 1] = "2"; mRelationCompetability_TermLong[12, 2] = "10";
            mRelationCompetability_TermLong[13, 0] = "2"; mRelationCompetability_TermLong[13, 1] = "3"; mRelationCompetability_TermLong[13, 2] = "8";
            mRelationCompetability_TermLong[14, 0] = "2"; mRelationCompetability_TermLong[14, 1] = "4"; mRelationCompetability_TermLong[14, 2] = "10";
            mRelationCompetability_TermLong[15, 0] = "2"; mRelationCompetability_TermLong[15, 1] = "5"; mRelationCompetability_TermLong[15, 2] = "6";
            mRelationCompetability_TermLong[16, 0] = "2"; mRelationCompetability_TermLong[16, 1] = "6"; mRelationCompetability_TermLong[16, 2] = "10";
            mRelationCompetability_TermLong[17, 0] = "2"; mRelationCompetability_TermLong[17, 1] = "7"; mRelationCompetability_TermLong[17, 2] = "3";
            mRelationCompetability_TermLong[18, 0] = "2"; mRelationCompetability_TermLong[18, 1] = "8"; mRelationCompetability_TermLong[18, 2] = "10";
            mRelationCompetability_TermLong[19, 0] = "2"; mRelationCompetability_TermLong[19, 1] = "9"; mRelationCompetability_TermLong[19, 2] = "10";
            mRelationCompetability_TermLong[20, 0] = "2"; mRelationCompetability_TermLong[20, 1] = "11"; mRelationCompetability_TermLong[20, 2] = "10";
            mRelationCompetability_TermLong[21, 0] = "2"; mRelationCompetability_TermLong[21, 1] = "22"; mRelationCompetability_TermLong[21, 2] = "10";
            mRelationCompetability_TermLong[22, 0] = "2"; mRelationCompetability_TermLong[22, 1] = "33"; mRelationCompetability_TermLong[22, 2] = "10";
            mRelationCompetability_TermLong[23, 0] = "3"; mRelationCompetability_TermLong[23, 1] = "3"; mRelationCompetability_TermLong[23, 2] = "8";
            mRelationCompetability_TermLong[24, 0] = "3"; mRelationCompetability_TermLong[24, 1] = "4"; mRelationCompetability_TermLong[24, 2] = "7";
            mRelationCompetability_TermLong[25, 0] = "3"; mRelationCompetability_TermLong[25, 1] = "5"; mRelationCompetability_TermLong[25, 2] = "7";
            mRelationCompetability_TermLong[26, 0] = "3"; mRelationCompetability_TermLong[26, 1] = "6"; mRelationCompetability_TermLong[26, 2] = "8";
            mRelationCompetability_TermLong[27, 0] = "3"; mRelationCompetability_TermLong[27, 1] = "7"; mRelationCompetability_TermLong[27, 2] = "4";
            mRelationCompetability_TermLong[28, 0] = "3"; mRelationCompetability_TermLong[28, 1] = "8"; mRelationCompetability_TermLong[28, 2] = "8";
            mRelationCompetability_TermLong[29, 0] = "3"; mRelationCompetability_TermLong[29, 1] = "9"; mRelationCompetability_TermLong[29, 2] = "10";
            mRelationCompetability_TermLong[30, 0] = "3"; mRelationCompetability_TermLong[30, 1] = "11"; mRelationCompetability_TermLong[30, 2] = "8";
            mRelationCompetability_TermLong[31, 0] = "3"; mRelationCompetability_TermLong[31, 1] = "22"; mRelationCompetability_TermLong[31, 2] = "10";
            mRelationCompetability_TermLong[32, 0] = "3"; mRelationCompetability_TermLong[32, 1] = "33"; mRelationCompetability_TermLong[32, 2] = "8";
            mRelationCompetability_TermLong[33, 0] = "4"; mRelationCompetability_TermLong[33, 1] = "4"; mRelationCompetability_TermLong[33, 2] = "9";
            mRelationCompetability_TermLong[34, 0] = "4"; mRelationCompetability_TermLong[34, 1] = "5"; mRelationCompetability_TermLong[34, 2] = "6";
            mRelationCompetability_TermLong[35, 0] = "4"; mRelationCompetability_TermLong[35, 1] = "6"; mRelationCompetability_TermLong[35, 2] = "10";
            mRelationCompetability_TermLong[36, 0] = "4"; mRelationCompetability_TermLong[36, 1] = "7"; mRelationCompetability_TermLong[36, 2] = "5";
            mRelationCompetability_TermLong[37, 0] = "4"; mRelationCompetability_TermLong[37, 1] = "8"; mRelationCompetability_TermLong[37, 2] = "10";
            mRelationCompetability_TermLong[38, 0] = "4"; mRelationCompetability_TermLong[38, 1] = "9"; mRelationCompetability_TermLong[38, 2] = "10";
            mRelationCompetability_TermLong[39, 0] = "4"; mRelationCompetability_TermLong[39, 1] = "11"; mRelationCompetability_TermLong[39, 2] = "10";
            mRelationCompetability_TermLong[40, 0] = "4"; mRelationCompetability_TermLong[40, 1] = "22"; mRelationCompetability_TermLong[40, 2] = "10";
            mRelationCompetability_TermLong[41, 0] = "4"; mRelationCompetability_TermLong[41, 1] = "33"; mRelationCompetability_TermLong[41, 2] = "10";
            mRelationCompetability_TermLong[42, 0] = "5"; mRelationCompetability_TermLong[42, 1] = "5"; mRelationCompetability_TermLong[42, 2] = "6";
            mRelationCompetability_TermLong[43, 0] = "5"; mRelationCompetability_TermLong[43, 1] = "6"; mRelationCompetability_TermLong[43, 2] = "6";
            mRelationCompetability_TermLong[44, 0] = "5"; mRelationCompetability_TermLong[44, 1] = "7"; mRelationCompetability_TermLong[44, 2] = "3";
            mRelationCompetability_TermLong[45, 0] = "5"; mRelationCompetability_TermLong[45, 1] = "8"; mRelationCompetability_TermLong[45, 2] = "6";
            mRelationCompetability_TermLong[46, 0] = "5"; mRelationCompetability_TermLong[46, 1] = "9"; mRelationCompetability_TermLong[46, 2] = "10";
            mRelationCompetability_TermLong[47, 0] = "5"; mRelationCompetability_TermLong[47, 1] = "11"; mRelationCompetability_TermLong[47, 2] = "8";
            mRelationCompetability_TermLong[48, 0] = "5"; mRelationCompetability_TermLong[48, 1] = "22"; mRelationCompetability_TermLong[48, 2] = "10";
            mRelationCompetability_TermLong[49, 0] = "5"; mRelationCompetability_TermLong[49, 1] = "33"; mRelationCompetability_TermLong[49, 2] = "6";
            mRelationCompetability_TermLong[50, 0] = "6"; mRelationCompetability_TermLong[50, 1] = "6"; mRelationCompetability_TermLong[50, 2] = "10";
            mRelationCompetability_TermLong[51, 0] = "6"; mRelationCompetability_TermLong[51, 1] = "7"; mRelationCompetability_TermLong[51, 2] = "3";
            mRelationCompetability_TermLong[52, 0] = "6"; mRelationCompetability_TermLong[52, 1] = "8"; mRelationCompetability_TermLong[52, 2] = "10";
            mRelationCompetability_TermLong[53, 0] = "6"; mRelationCompetability_TermLong[53, 1] = "9"; mRelationCompetability_TermLong[53, 2] = "10";
            mRelationCompetability_TermLong[54, 0] = "6"; mRelationCompetability_TermLong[54, 1] = "11"; mRelationCompetability_TermLong[54, 2] = "10";
            mRelationCompetability_TermLong[55, 0] = "6"; mRelationCompetability_TermLong[55, 1] = "22"; mRelationCompetability_TermLong[55, 2] = "10";
            mRelationCompetability_TermLong[56, 0] = "6"; mRelationCompetability_TermLong[56, 1] = "33"; mRelationCompetability_TermLong[56, 2] = "10";
            mRelationCompetability_TermLong[57, 0] = "7"; mRelationCompetability_TermLong[57, 1] = "7"; mRelationCompetability_TermLong[57, 2] = "2";
            mRelationCompetability_TermLong[58, 0] = "7"; mRelationCompetability_TermLong[58, 1] = "8"; mRelationCompetability_TermLong[58, 2] = "5";
            mRelationCompetability_TermLong[59, 0] = "7"; mRelationCompetability_TermLong[59, 1] = "9"; mRelationCompetability_TermLong[59, 2] = "6";
            mRelationCompetability_TermLong[60, 0] = "7"; mRelationCompetability_TermLong[60, 1] = "11"; mRelationCompetability_TermLong[60, 2] = "6";
            mRelationCompetability_TermLong[61, 0] = "7"; mRelationCompetability_TermLong[61, 1] = "22"; mRelationCompetability_TermLong[61, 2] = "6";
            mRelationCompetability_TermLong[62, 0] = "7"; mRelationCompetability_TermLong[62, 1] = "33"; mRelationCompetability_TermLong[62, 2] = "3";
            mRelationCompetability_TermLong[63, 0] = "8"; mRelationCompetability_TermLong[63, 1] = "8"; mRelationCompetability_TermLong[63, 2] = "10";
            mRelationCompetability_TermLong[64, 0] = "8"; mRelationCompetability_TermLong[64, 1] = "9"; mRelationCompetability_TermLong[64, 2] = "10";
            mRelationCompetability_TermLong[65, 0] = "8"; mRelationCompetability_TermLong[65, 1] = "11"; mRelationCompetability_TermLong[65, 2] = "10";
            mRelationCompetability_TermLong[66, 0] = "8"; mRelationCompetability_TermLong[66, 1] = "22"; mRelationCompetability_TermLong[66, 2] = "10";
            mRelationCompetability_TermLong[67, 0] = "8"; mRelationCompetability_TermLong[67, 1] = "33"; mRelationCompetability_TermLong[67, 2] = "10";
            mRelationCompetability_TermLong[68, 0] = "9"; mRelationCompetability_TermLong[68, 1] = "9"; mRelationCompetability_TermLong[68, 2] = "10";
            mRelationCompetability_TermLong[69, 0] = "9"; mRelationCompetability_TermLong[69, 1] = "11"; mRelationCompetability_TermLong[69, 2] = "10";
            mRelationCompetability_TermLong[70, 0] = "9"; mRelationCompetability_TermLong[70, 1] = "22"; mRelationCompetability_TermLong[70, 2] = "10";
            mRelationCompetability_TermLong[71, 0] = "9"; mRelationCompetability_TermLong[71, 1] = "33"; mRelationCompetability_TermLong[71, 2] = "10";
            mRelationCompetability_TermLong[72, 0] = "11"; mRelationCompetability_TermLong[72, 1] = "11"; mRelationCompetability_TermLong[72, 2] = "10";
            mRelationCompetability_TermLong[73, 0] = "11"; mRelationCompetability_TermLong[73, 1] = "22"; mRelationCompetability_TermLong[73, 2] = "10";
            mRelationCompetability_TermLong[74, 0] = "11"; mRelationCompetability_TermLong[74, 1] = "33"; mRelationCompetability_TermLong[74, 2] = "10";
            mRelationCompetability_TermLong[75, 0] = "22"; mRelationCompetability_TermLong[75, 1] = "22"; mRelationCompetability_TermLong[75, 2] = "10";
            mRelationCompetability_TermLong[76, 0] = "22"; mRelationCompetability_TermLong[76, 1] = "33"; mRelationCompetability_TermLong[76, 2] = "10";
            mRelationCompetability_TermLong[77, 0] = "33"; mRelationCompetability_TermLong[77, 1] = "33"; mRelationCompetability_TermLong[77, 2] = "10";
            mRelationCompetability_TermLong[78, 0] = "13"; mRelationCompetability_TermLong[78, 1] = "1"; mRelationCompetability_TermLong[78, 2] = "3";
            mRelationCompetability_TermLong[79, 0] = "13"; mRelationCompetability_TermLong[79, 1] = "2"; mRelationCompetability_TermLong[79, 2] = "7";
            mRelationCompetability_TermLong[80, 0] = "13"; mRelationCompetability_TermLong[80, 1] = "3"; mRelationCompetability_TermLong[80, 2] = "5";
            mRelationCompetability_TermLong[81, 0] = "13"; mRelationCompetability_TermLong[81, 1] = "4"; mRelationCompetability_TermLong[81, 2] = "7";
            mRelationCompetability_TermLong[82, 0] = "13"; mRelationCompetability_TermLong[82, 1] = "5"; mRelationCompetability_TermLong[82, 2] = "4";
            mRelationCompetability_TermLong[83, 0] = "13"; mRelationCompetability_TermLong[83, 1] = "6"; mRelationCompetability_TermLong[83, 2] = "8";
            mRelationCompetability_TermLong[84, 0] = "13"; mRelationCompetability_TermLong[84, 1] = "7"; mRelationCompetability_TermLong[84, 2] = "3";
            mRelationCompetability_TermLong[85, 0] = "13"; mRelationCompetability_TermLong[85, 1] = "8"; mRelationCompetability_TermLong[85, 2] = "7";
            mRelationCompetability_TermLong[86, 0] = "13"; mRelationCompetability_TermLong[86, 1] = "9"; mRelationCompetability_TermLong[86, 2] = "7";
            mRelationCompetability_TermLong[87, 0] = "13"; mRelationCompetability_TermLong[87, 1] = "11"; mRelationCompetability_TermLong[87, 2] = "7";
            mRelationCompetability_TermLong[88, 0] = "13"; mRelationCompetability_TermLong[88, 1] = "22"; mRelationCompetability_TermLong[88, 2] = "7";
            mRelationCompetability_TermLong[89, 0] = "13"; mRelationCompetability_TermLong[89, 1] = "33"; mRelationCompetability_TermLong[89, 2] = "8";
            mRelationCompetability_TermLong[90, 0] = "14"; mRelationCompetability_TermLong[90, 1] = "1"; mRelationCompetability_TermLong[90, 2] = "4";
            mRelationCompetability_TermLong[91, 0] = "14"; mRelationCompetability_TermLong[91, 1] = "2"; mRelationCompetability_TermLong[91, 2] = "4";
            mRelationCompetability_TermLong[92, 0] = "14"; mRelationCompetability_TermLong[92, 1] = "3"; mRelationCompetability_TermLong[92, 2] = "6";
            mRelationCompetability_TermLong[93, 0] = "14"; mRelationCompetability_TermLong[93, 1] = "4"; mRelationCompetability_TermLong[93, 2] = "4";
            mRelationCompetability_TermLong[94, 0] = "14"; mRelationCompetability_TermLong[94, 1] = "5"; mRelationCompetability_TermLong[94, 2] = "4";
            mRelationCompetability_TermLong[95, 0] = "14"; mRelationCompetability_TermLong[95, 1] = "6"; mRelationCompetability_TermLong[95, 2] = "4";
            mRelationCompetability_TermLong[96, 0] = "14"; mRelationCompetability_TermLong[96, 1] = "7"; mRelationCompetability_TermLong[96, 2] = "1";
            mRelationCompetability_TermLong[97, 0] = "14"; mRelationCompetability_TermLong[97, 1] = "8"; mRelationCompetability_TermLong[97, 2] = "6";
            mRelationCompetability_TermLong[98, 0] = "14"; mRelationCompetability_TermLong[98, 1] = "9"; mRelationCompetability_TermLong[98, 2] = "6";
            mRelationCompetability_TermLong[99, 0] = "14"; mRelationCompetability_TermLong[99, 1] = "11"; mRelationCompetability_TermLong[99, 2] = "6";
            mRelationCompetability_TermLong[100, 0] = "14"; mRelationCompetability_TermLong[100, 1] = "22"; mRelationCompetability_TermLong[100, 2] = "6";
            mRelationCompetability_TermLong[101, 0] = "14"; mRelationCompetability_TermLong[101, 1] = "33"; mRelationCompetability_TermLong[101, 2] = "4";
            mRelationCompetability_TermLong[102, 0] = "15"; mRelationCompetability_TermLong[102, 1] = "1"; mRelationCompetability_TermLong[102, 2] = "7";
            mRelationCompetability_TermLong[103, 0] = "15"; mRelationCompetability_TermLong[103, 1] = "2"; mRelationCompetability_TermLong[103, 2] = "8";
            mRelationCompetability_TermLong[104, 0] = "15"; mRelationCompetability_TermLong[104, 1] = "3"; mRelationCompetability_TermLong[104, 2] = "7";
            mRelationCompetability_TermLong[105, 0] = "15"; mRelationCompetability_TermLong[105, 1] = "4"; mRelationCompetability_TermLong[105, 2] = "8";
            mRelationCompetability_TermLong[106, 0] = "15"; mRelationCompetability_TermLong[106, 1] = "5"; mRelationCompetability_TermLong[106, 2] = "7";
            mRelationCompetability_TermLong[107, 0] = "15"; mRelationCompetability_TermLong[107, 1] = "6"; mRelationCompetability_TermLong[107, 2] = "8";
            mRelationCompetability_TermLong[108, 0] = "15"; mRelationCompetability_TermLong[108, 1] = "7"; mRelationCompetability_TermLong[108, 2] = "3";
            mRelationCompetability_TermLong[109, 0] = "15"; mRelationCompetability_TermLong[109, 1] = "8"; mRelationCompetability_TermLong[109, 2] = "8";
            mRelationCompetability_TermLong[110, 0] = "15"; mRelationCompetability_TermLong[110, 1] = "9"; mRelationCompetability_TermLong[110, 2] = "8";
            mRelationCompetability_TermLong[111, 0] = "15"; mRelationCompetability_TermLong[111, 1] = "11"; mRelationCompetability_TermLong[111, 2] = "8";
            mRelationCompetability_TermLong[112, 0] = "15"; mRelationCompetability_TermLong[112, 1] = "22"; mRelationCompetability_TermLong[112, 2] = "8";
            mRelationCompetability_TermLong[113, 0] = "15"; mRelationCompetability_TermLong[113, 1] = "33"; mRelationCompetability_TermLong[113, 2] = "8";
            mRelationCompetability_TermLong[114, 0] = "16"; mRelationCompetability_TermLong[114, 1] = "1"; mRelationCompetability_TermLong[114, 2] = "2";
            mRelationCompetability_TermLong[115, 0] = "16"; mRelationCompetability_TermLong[115, 1] = "2"; mRelationCompetability_TermLong[115, 2] = "1";
            mRelationCompetability_TermLong[116, 0] = "16"; mRelationCompetability_TermLong[116, 1] = "3"; mRelationCompetability_TermLong[116, 2] = "2";
            mRelationCompetability_TermLong[117, 0] = "16"; mRelationCompetability_TermLong[117, 1] = "4"; mRelationCompetability_TermLong[117, 2] = "3";
            mRelationCompetability_TermLong[118, 0] = "16"; mRelationCompetability_TermLong[118, 1] = "5"; mRelationCompetability_TermLong[118, 2] = "1";
            mRelationCompetability_TermLong[119, 0] = "16"; mRelationCompetability_TermLong[119, 1] = "6"; mRelationCompetability_TermLong[119, 2] = "1";
            mRelationCompetability_TermLong[120, 0] = "16"; mRelationCompetability_TermLong[120, 1] = "7"; mRelationCompetability_TermLong[120, 2] = "4";
            mRelationCompetability_TermLong[121, 0] = "16"; mRelationCompetability_TermLong[121, 1] = "8"; mRelationCompetability_TermLong[121, 2] = "3";
            mRelationCompetability_TermLong[122, 0] = "16"; mRelationCompetability_TermLong[122, 1] = "9"; mRelationCompetability_TermLong[122, 2] = "4";
            mRelationCompetability_TermLong[123, 0] = "16"; mRelationCompetability_TermLong[123, 1] = "11"; mRelationCompetability_TermLong[123, 2] = "4";
            mRelationCompetability_TermLong[124, 0] = "16"; mRelationCompetability_TermLong[124, 1] = "22"; mRelationCompetability_TermLong[124, 2] = "4";
            mRelationCompetability_TermLong[125, 0] = "16"; mRelationCompetability_TermLong[125, 1] = "33"; mRelationCompetability_TermLong[125, 2] = "1";
            mRelationCompetability_TermLong[126, 0] = "19"; mRelationCompetability_TermLong[126, 1] = "1"; mRelationCompetability_TermLong[126, 2] = "5";
            mRelationCompetability_TermLong[127, 0] = "19"; mRelationCompetability_TermLong[127, 1] = "2"; mRelationCompetability_TermLong[127, 2] = "7";
            mRelationCompetability_TermLong[128, 0] = "19"; mRelationCompetability_TermLong[128, 1] = "3"; mRelationCompetability_TermLong[128, 2] = "5";
            mRelationCompetability_TermLong[129, 0] = "19"; mRelationCompetability_TermLong[129, 1] = "4"; mRelationCompetability_TermLong[129, 2] = "7";
            mRelationCompetability_TermLong[130, 0] = "19"; mRelationCompetability_TermLong[130, 1] = "5"; mRelationCompetability_TermLong[130, 2] = "5";
            mRelationCompetability_TermLong[131, 0] = "19"; mRelationCompetability_TermLong[131, 1] = "6"; mRelationCompetability_TermLong[131, 2] = "7";
            mRelationCompetability_TermLong[132, 0] = "19"; mRelationCompetability_TermLong[132, 1] = "7"; mRelationCompetability_TermLong[132, 2] = "3";
            mRelationCompetability_TermLong[133, 0] = "19"; mRelationCompetability_TermLong[133, 1] = "8"; mRelationCompetability_TermLong[133, 2] = "7";
            mRelationCompetability_TermLong[134, 0] = "19"; mRelationCompetability_TermLong[134, 1] = "9"; mRelationCompetability_TermLong[134, 2] = "7";
            mRelationCompetability_TermLong[135, 0] = "19"; mRelationCompetability_TermLong[135, 1] = "11"; mRelationCompetability_TermLong[135, 2] = "6";
            mRelationCompetability_TermLong[136, 0] = "19"; mRelationCompetability_TermLong[136, 1] = "22"; mRelationCompetability_TermLong[136, 2] = "7";
            mRelationCompetability_TermLong[137, 0] = "19"; mRelationCompetability_TermLong[137, 1] = "33"; mRelationCompetability_TermLong[137, 2] = "7";


            mRelationCompetability_TermShort[0, 0] = "1"; mRelationCompetability_TermShort[0, 1] = "1"; mRelationCompetability_TermShort[0, 2] = "9";
            mRelationCompetability_TermShort[1, 0] = "1"; mRelationCompetability_TermShort[1, 1] = "2"; mRelationCompetability_TermShort[1, 2] = "4";
            mRelationCompetability_TermShort[2, 0] = "1"; mRelationCompetability_TermShort[2, 1] = "3"; mRelationCompetability_TermShort[2, 2] = "9";
            mRelationCompetability_TermShort[3, 0] = "1"; mRelationCompetability_TermShort[3, 1] = "4"; mRelationCompetability_TermShort[3, 2] = "5";
            mRelationCompetability_TermShort[4, 0] = "1"; mRelationCompetability_TermShort[4, 1] = "5"; mRelationCompetability_TermShort[4, 2] = "9";
            mRelationCompetability_TermShort[5, 0] = "1"; mRelationCompetability_TermShort[5, 1] = "6"; mRelationCompetability_TermShort[5, 2] = "5";
            mRelationCompetability_TermShort[6, 0] = "1"; mRelationCompetability_TermShort[6, 1] = "7"; mRelationCompetability_TermShort[6, 2] = "2";
            mRelationCompetability_TermShort[7, 0] = "1"; mRelationCompetability_TermShort[7, 1] = "8"; mRelationCompetability_TermShort[7, 2] = "6";
            mRelationCompetability_TermShort[8, 0] = "1"; mRelationCompetability_TermShort[8, 1] = "9"; mRelationCompetability_TermShort[8, 2] = "10";
            mRelationCompetability_TermShort[9, 0] = "1"; mRelationCompetability_TermShort[9, 1] = "11"; mRelationCompetability_TermShort[9, 2] = "10";
            mRelationCompetability_TermShort[10, 0] = "1"; mRelationCompetability_TermShort[10, 1] = "22"; mRelationCompetability_TermShort[10, 2] = "10";
            mRelationCompetability_TermShort[11, 0] = "1"; mRelationCompetability_TermShort[11, 1] = "33"; mRelationCompetability_TermShort[11, 2] = "5";
            mRelationCompetability_TermShort[12, 0] = "2"; mRelationCompetability_TermShort[12, 1] = "2"; mRelationCompetability_TermShort[12, 2] = "4";
            mRelationCompetability_TermShort[13, 0] = "2"; mRelationCompetability_TermShort[13, 1] = "3"; mRelationCompetability_TermShort[13, 2] = "4";
            mRelationCompetability_TermShort[14, 0] = "2"; mRelationCompetability_TermShort[14, 1] = "4"; mRelationCompetability_TermShort[14, 2] = "4";
            mRelationCompetability_TermShort[15, 0] = "2"; mRelationCompetability_TermShort[15, 1] = "5"; mRelationCompetability_TermShort[15, 2] = "6";
            mRelationCompetability_TermShort[16, 0] = "2"; mRelationCompetability_TermShort[16, 1] = "6"; mRelationCompetability_TermShort[16, 2] = "4";
            mRelationCompetability_TermShort[17, 0] = "2"; mRelationCompetability_TermShort[17, 1] = "7"; mRelationCompetability_TermShort[17, 2] = "1";
            mRelationCompetability_TermShort[18, 0] = "2"; mRelationCompetability_TermShort[18, 1] = "8"; mRelationCompetability_TermShort[18, 2] = "4";
            mRelationCompetability_TermShort[19, 0] = "2"; mRelationCompetability_TermShort[19, 1] = "9"; mRelationCompetability_TermShort[19, 2] = "4";
            mRelationCompetability_TermShort[20, 0] = "2"; mRelationCompetability_TermShort[20, 1] = "11"; mRelationCompetability_TermShort[20, 2] = "4";
            mRelationCompetability_TermShort[21, 0] = "2"; mRelationCompetability_TermShort[21, 1] = "22"; mRelationCompetability_TermShort[21, 2] = "4";
            mRelationCompetability_TermShort[22, 0] = "2"; mRelationCompetability_TermShort[22, 1] = "33"; mRelationCompetability_TermShort[22, 2] = "4";
            mRelationCompetability_TermShort[23, 0] = "3"; mRelationCompetability_TermShort[23, 1] = "3"; mRelationCompetability_TermShort[23, 2] = "10";
            mRelationCompetability_TermShort[24, 0] = "3"; mRelationCompetability_TermShort[24, 1] = "4"; mRelationCompetability_TermShort[24, 2] = "4";
            mRelationCompetability_TermShort[25, 0] = "3"; mRelationCompetability_TermShort[25, 1] = "5"; mRelationCompetability_TermShort[25, 2] = "9";
            mRelationCompetability_TermShort[26, 0] = "3"; mRelationCompetability_TermShort[26, 1] = "6"; mRelationCompetability_TermShort[26, 2] = "4";
            mRelationCompetability_TermShort[27, 0] = "3"; mRelationCompetability_TermShort[27, 1] = "7"; mRelationCompetability_TermShort[27, 2] = "2";
            mRelationCompetability_TermShort[28, 0] = "3"; mRelationCompetability_TermShort[28, 1] = "8"; mRelationCompetability_TermShort[28, 2] = "4";
            mRelationCompetability_TermShort[29, 0] = "3"; mRelationCompetability_TermShort[29, 1] = "9"; mRelationCompetability_TermShort[29, 2] = "10";
            mRelationCompetability_TermShort[30, 0] = "3"; mRelationCompetability_TermShort[30, 1] = "11"; mRelationCompetability_TermShort[30, 2] = "10";
            mRelationCompetability_TermShort[31, 0] = "3"; mRelationCompetability_TermShort[31, 1] = "22"; mRelationCompetability_TermShort[31, 2] = "10";
            mRelationCompetability_TermShort[32, 0] = "3"; mRelationCompetability_TermShort[32, 1] = "33"; mRelationCompetability_TermShort[32, 2] = "4";
            mRelationCompetability_TermShort[33, 0] = "4"; mRelationCompetability_TermShort[33, 1] = "4"; mRelationCompetability_TermShort[33, 2] = "6";
            mRelationCompetability_TermShort[34, 0] = "4"; mRelationCompetability_TermShort[34, 1] = "5"; mRelationCompetability_TermShort[34, 2] = "4";
            mRelationCompetability_TermShort[35, 0] = "4"; mRelationCompetability_TermShort[35, 1] = "6"; mRelationCompetability_TermShort[35, 2] = "4";
            mRelationCompetability_TermShort[36, 0] = "4"; mRelationCompetability_TermShort[36, 1] = "7"; mRelationCompetability_TermShort[36, 2] = "6";
            mRelationCompetability_TermShort[37, 0] = "4"; mRelationCompetability_TermShort[37, 1] = "8"; mRelationCompetability_TermShort[37, 2] = "6";
            mRelationCompetability_TermShort[38, 0] = "4"; mRelationCompetability_TermShort[38, 1] = "9"; mRelationCompetability_TermShort[38, 2] = "6";
            mRelationCompetability_TermShort[39, 0] = "4"; mRelationCompetability_TermShort[39, 1] = "11"; mRelationCompetability_TermShort[39, 2] = "6";
            mRelationCompetability_TermShort[40, 0] = "4"; mRelationCompetability_TermShort[40, 1] = "22"; mRelationCompetability_TermShort[40, 2] = "6";
            mRelationCompetability_TermShort[41, 0] = "4"; mRelationCompetability_TermShort[41, 1] = "33"; mRelationCompetability_TermShort[41, 2] = "4";
            mRelationCompetability_TermShort[42, 0] = "5"; mRelationCompetability_TermShort[42, 1] = "5"; mRelationCompetability_TermShort[42, 2] = "10";
            mRelationCompetability_TermShort[43, 0] = "5"; mRelationCompetability_TermShort[43, 1] = "6"; mRelationCompetability_TermShort[43, 2] = "4";
            mRelationCompetability_TermShort[44, 0] = "5"; mRelationCompetability_TermShort[44, 1] = "7"; mRelationCompetability_TermShort[44, 2] = "2";
            mRelationCompetability_TermShort[45, 0] = "5"; mRelationCompetability_TermShort[45, 1] = "8"; mRelationCompetability_TermShort[45, 2] = "6";
            mRelationCompetability_TermShort[46, 0] = "5"; mRelationCompetability_TermShort[46, 1] = "9"; mRelationCompetability_TermShort[46, 2] = "10";
            mRelationCompetability_TermShort[47, 0] = "5"; mRelationCompetability_TermShort[47, 1] = "11"; mRelationCompetability_TermShort[47, 2] = "10";
            mRelationCompetability_TermShort[48, 0] = "5"; mRelationCompetability_TermShort[48, 1] = "22"; mRelationCompetability_TermShort[48, 2] = "10";
            mRelationCompetability_TermShort[49, 0] = "5"; mRelationCompetability_TermShort[49, 1] = "33"; mRelationCompetability_TermShort[49, 2] = "4";
            mRelationCompetability_TermShort[50, 0] = "6"; mRelationCompetability_TermShort[50, 1] = "6"; mRelationCompetability_TermShort[50, 2] = "6";
            mRelationCompetability_TermShort[51, 0] = "6"; mRelationCompetability_TermShort[51, 1] = "7"; mRelationCompetability_TermShort[51, 2] = "3";
            mRelationCompetability_TermShort[52, 0] = "6"; mRelationCompetability_TermShort[52, 1] = "8"; mRelationCompetability_TermShort[52, 2] = "4";
            mRelationCompetability_TermShort[53, 0] = "6"; mRelationCompetability_TermShort[53, 1] = "9"; mRelationCompetability_TermShort[53, 2] = "4";
            mRelationCompetability_TermShort[54, 0] = "6"; mRelationCompetability_TermShort[54, 1] = "11"; mRelationCompetability_TermShort[54, 2] = "4";
            mRelationCompetability_TermShort[55, 0] = "6"; mRelationCompetability_TermShort[55, 1] = "22"; mRelationCompetability_TermShort[55, 2] = "4";
            mRelationCompetability_TermShort[56, 0] = "6"; mRelationCompetability_TermShort[56, 1] = "33"; mRelationCompetability_TermShort[56, 2] = "4";
            mRelationCompetability_TermShort[57, 0] = "7"; mRelationCompetability_TermShort[57, 1] = "7"; mRelationCompetability_TermShort[57, 2] = "6";
            mRelationCompetability_TermShort[58, 0] = "7"; mRelationCompetability_TermShort[58, 1] = "8"; mRelationCompetability_TermShort[58, 2] = "3";
            mRelationCompetability_TermShort[59, 0] = "7"; mRelationCompetability_TermShort[59, 1] = "9"; mRelationCompetability_TermShort[59, 2] = "3";
            mRelationCompetability_TermShort[60, 0] = "7"; mRelationCompetability_TermShort[60, 1] = "11"; mRelationCompetability_TermShort[60, 2] = "3";
            mRelationCompetability_TermShort[61, 0] = "7"; mRelationCompetability_TermShort[61, 1] = "22"; mRelationCompetability_TermShort[61, 2] = "3";
            mRelationCompetability_TermShort[62, 0] = "7"; mRelationCompetability_TermShort[62, 1] = "33"; mRelationCompetability_TermShort[62, 2] = "3";
            mRelationCompetability_TermShort[63, 0] = "8"; mRelationCompetability_TermShort[63, 1] = "8"; mRelationCompetability_TermShort[63, 2] = "8";
            mRelationCompetability_TermShort[64, 0] = "8"; mRelationCompetability_TermShort[64, 1] = "9"; mRelationCompetability_TermShort[64, 2] = "4";
            mRelationCompetability_TermShort[65, 0] = "8"; mRelationCompetability_TermShort[65, 1] = "11"; mRelationCompetability_TermShort[65, 2] = "4";
            mRelationCompetability_TermShort[66, 0] = "8"; mRelationCompetability_TermShort[66, 1] = "22"; mRelationCompetability_TermShort[66, 2] = "4";
            mRelationCompetability_TermShort[67, 0] = "8"; mRelationCompetability_TermShort[67, 1] = "33"; mRelationCompetability_TermShort[67, 2] = "4";
            mRelationCompetability_TermShort[68, 0] = "9"; mRelationCompetability_TermShort[68, 1] = "9"; mRelationCompetability_TermShort[68, 2] = "10";
            mRelationCompetability_TermShort[69, 0] = "9"; mRelationCompetability_TermShort[69, 1] = "11"; mRelationCompetability_TermShort[69, 2] = "10";
            mRelationCompetability_TermShort[70, 0] = "9"; mRelationCompetability_TermShort[70, 1] = "22"; mRelationCompetability_TermShort[70, 2] = "10";
            mRelationCompetability_TermShort[71, 0] = "9"; mRelationCompetability_TermShort[71, 1] = "33"; mRelationCompetability_TermShort[71, 2] = "4";
            mRelationCompetability_TermShort[72, 0] = "11"; mRelationCompetability_TermShort[72, 1] = "11"; mRelationCompetability_TermShort[72, 2] = "10";
            mRelationCompetability_TermShort[73, 0] = "11"; mRelationCompetability_TermShort[73, 1] = "22"; mRelationCompetability_TermShort[73, 2] = "10";
            mRelationCompetability_TermShort[74, 0] = "11"; mRelationCompetability_TermShort[74, 1] = "33"; mRelationCompetability_TermShort[74, 2] = "4";
            mRelationCompetability_TermShort[75, 0] = "22"; mRelationCompetability_TermShort[75, 1] = "22"; mRelationCompetability_TermShort[75, 2] = "10";
            mRelationCompetability_TermShort[76, 0] = "22"; mRelationCompetability_TermShort[76, 1] = "33"; mRelationCompetability_TermShort[76, 2] = "4";
            mRelationCompetability_TermShort[77, 0] = "33"; mRelationCompetability_TermShort[77, 1] = "33"; mRelationCompetability_TermShort[77, 2] = "4";
            mRelationCompetability_TermShort[78, 0] = "13"; mRelationCompetability_TermShort[78, 1] = "1"; mRelationCompetability_TermShort[78, 2] = "6";
            mRelationCompetability_TermShort[79, 0] = "13"; mRelationCompetability_TermShort[79, 1] = "2"; mRelationCompetability_TermShort[79, 2] = "4";
            mRelationCompetability_TermShort[80, 0] = "13"; mRelationCompetability_TermShort[80, 1] = "3"; mRelationCompetability_TermShort[80, 2] = "5";
            mRelationCompetability_TermShort[81, 0] = "13"; mRelationCompetability_TermShort[81, 1] = "4"; mRelationCompetability_TermShort[81, 2] = "4";
            mRelationCompetability_TermShort[82, 0] = "13"; mRelationCompetability_TermShort[82, 1] = "5"; mRelationCompetability_TermShort[82, 2] = "6";
            mRelationCompetability_TermShort[83, 0] = "13"; mRelationCompetability_TermShort[83, 1] = "6"; mRelationCompetability_TermShort[83, 2] = "4";
            mRelationCompetability_TermShort[84, 0] = "13"; mRelationCompetability_TermShort[84, 1] = "7"; mRelationCompetability_TermShort[84, 2] = "3";
            mRelationCompetability_TermShort[85, 0] = "13"; mRelationCompetability_TermShort[85, 1] = "8"; mRelationCompetability_TermShort[85, 2] = "4";
            mRelationCompetability_TermShort[86, 0] = "13"; mRelationCompetability_TermShort[86, 1] = "9"; mRelationCompetability_TermShort[86, 2] = "6";
            mRelationCompetability_TermShort[87, 0] = "13"; mRelationCompetability_TermShort[87, 1] = "11"; mRelationCompetability_TermShort[87, 2] = "6";
            mRelationCompetability_TermShort[88, 0] = "13"; mRelationCompetability_TermShort[88, 1] = "22"; mRelationCompetability_TermShort[88, 2] = "6";
            mRelationCompetability_TermShort[89, 0] = "13"; mRelationCompetability_TermShort[89, 1] = "33"; mRelationCompetability_TermShort[89, 2] = "4";
            mRelationCompetability_TermShort[90, 0] = "14"; mRelationCompetability_TermShort[90, 1] = "1"; mRelationCompetability_TermShort[90, 2] = "6";
            mRelationCompetability_TermShort[91, 0] = "14"; mRelationCompetability_TermShort[91, 1] = "2"; mRelationCompetability_TermShort[91, 2] = "4";
            mRelationCompetability_TermShort[92, 0] = "14"; mRelationCompetability_TermShort[92, 1] = "3"; mRelationCompetability_TermShort[92, 2] = "8";
            mRelationCompetability_TermShort[93, 0] = "14"; mRelationCompetability_TermShort[93, 1] = "4"; mRelationCompetability_TermShort[93, 2] = "4";
            mRelationCompetability_TermShort[94, 0] = "14"; mRelationCompetability_TermShort[94, 1] = "5"; mRelationCompetability_TermShort[94, 2] = "8";
            mRelationCompetability_TermShort[95, 0] = "14"; mRelationCompetability_TermShort[95, 1] = "6"; mRelationCompetability_TermShort[95, 2] = "4";
            mRelationCompetability_TermShort[96, 0] = "14"; mRelationCompetability_TermShort[96, 1] = "7"; mRelationCompetability_TermShort[96, 2] = "3";
            mRelationCompetability_TermShort[97, 0] = "14"; mRelationCompetability_TermShort[97, 1] = "8"; mRelationCompetability_TermShort[97, 2] = "4";
            mRelationCompetability_TermShort[98, 0] = "14"; mRelationCompetability_TermShort[98, 1] = "9"; mRelationCompetability_TermShort[98, 2] = "8";
            mRelationCompetability_TermShort[99, 0] = "14"; mRelationCompetability_TermShort[99, 1] = "11"; mRelationCompetability_TermShort[99, 2] = "8";
            mRelationCompetability_TermShort[100, 0] = "14"; mRelationCompetability_TermShort[100, 1] = "22"; mRelationCompetability_TermShort[100, 2] = "8";
            mRelationCompetability_TermShort[101, 0] = "14"; mRelationCompetability_TermShort[101, 1] = "33"; mRelationCompetability_TermShort[101, 2] = "4";
            mRelationCompetability_TermShort[102, 0] = "15"; mRelationCompetability_TermShort[102, 1] = "1"; mRelationCompetability_TermShort[102, 2] = "6";
            mRelationCompetability_TermShort[103, 0] = "15"; mRelationCompetability_TermShort[103, 1] = "2"; mRelationCompetability_TermShort[103, 2] = "4";
            mRelationCompetability_TermShort[104, 0] = "15"; mRelationCompetability_TermShort[104, 1] = "3"; mRelationCompetability_TermShort[104, 2] = "7";
            mRelationCompetability_TermShort[105, 0] = "15"; mRelationCompetability_TermShort[105, 1] = "4"; mRelationCompetability_TermShort[105, 2] = "4";
            mRelationCompetability_TermShort[106, 0] = "15"; mRelationCompetability_TermShort[106, 1] = "5"; mRelationCompetability_TermShort[106, 2] = "7";
            mRelationCompetability_TermShort[107, 0] = "15"; mRelationCompetability_TermShort[107, 1] = "6"; mRelationCompetability_TermShort[107, 2] = "4";
            mRelationCompetability_TermShort[108, 0] = "15"; mRelationCompetability_TermShort[108, 1] = "7"; mRelationCompetability_TermShort[108, 2] = "3";
            mRelationCompetability_TermShort[109, 0] = "15"; mRelationCompetability_TermShort[109, 1] = "8"; mRelationCompetability_TermShort[109, 2] = "4";
            mRelationCompetability_TermShort[110, 0] = "15"; mRelationCompetability_TermShort[110, 1] = "9"; mRelationCompetability_TermShort[110, 2] = "7";
            mRelationCompetability_TermShort[111, 0] = "15"; mRelationCompetability_TermShort[111, 1] = "11"; mRelationCompetability_TermShort[111, 2] = "7";
            mRelationCompetability_TermShort[112, 0] = "15"; mRelationCompetability_TermShort[112, 1] = "22"; mRelationCompetability_TermShort[112, 2] = "7";
            mRelationCompetability_TermShort[113, 0] = "15"; mRelationCompetability_TermShort[113, 1] = "33"; mRelationCompetability_TermShort[113, 2] = "4";
            mRelationCompetability_TermShort[114, 0] = "16"; mRelationCompetability_TermShort[114, 1] = "1"; mRelationCompetability_TermShort[114, 2] = "2";
            mRelationCompetability_TermShort[115, 0] = "16"; mRelationCompetability_TermShort[115, 1] = "2"; mRelationCompetability_TermShort[115, 2] = "1";
            mRelationCompetability_TermShort[116, 0] = "16"; mRelationCompetability_TermShort[116, 1] = "3"; mRelationCompetability_TermShort[116, 2] = "2";
            mRelationCompetability_TermShort[117, 0] = "16"; mRelationCompetability_TermShort[117, 1] = "4"; mRelationCompetability_TermShort[117, 2] = "3";
            mRelationCompetability_TermShort[118, 0] = "16"; mRelationCompetability_TermShort[118, 1] = "5"; mRelationCompetability_TermShort[118, 2] = "1";
            mRelationCompetability_TermShort[119, 0] = "16"; mRelationCompetability_TermShort[119, 1] = "6"; mRelationCompetability_TermShort[119, 2] = "1";
            mRelationCompetability_TermShort[120, 0] = "16"; mRelationCompetability_TermShort[120, 1] = "7"; mRelationCompetability_TermShort[120, 2] = "6";
            mRelationCompetability_TermShort[121, 0] = "16"; mRelationCompetability_TermShort[121, 1] = "8"; mRelationCompetability_TermShort[121, 2] = "3";
            mRelationCompetability_TermShort[122, 0] = "16"; mRelationCompetability_TermShort[122, 1] = "9"; mRelationCompetability_TermShort[122, 2] = "3";
            mRelationCompetability_TermShort[123, 0] = "16"; mRelationCompetability_TermShort[123, 1] = "11"; mRelationCompetability_TermShort[123, 2] = "3";
            mRelationCompetability_TermShort[124, 0] = "16"; mRelationCompetability_TermShort[124, 1] = "22"; mRelationCompetability_TermShort[124, 2] = "3";
            mRelationCompetability_TermShort[125, 0] = "16"; mRelationCompetability_TermShort[125, 1] = "33"; mRelationCompetability_TermShort[125, 2] = "1";
            mRelationCompetability_TermShort[126, 0] = "19"; mRelationCompetability_TermShort[126, 1] = "1"; mRelationCompetability_TermShort[126, 2] = "8";
            mRelationCompetability_TermShort[127, 0] = "19"; mRelationCompetability_TermShort[127, 1] = "2"; mRelationCompetability_TermShort[127, 2] = "4";
            mRelationCompetability_TermShort[128, 0] = "19"; mRelationCompetability_TermShort[128, 1] = "3"; mRelationCompetability_TermShort[128, 2] = "8";
            mRelationCompetability_TermShort[129, 0] = "19"; mRelationCompetability_TermShort[129, 1] = "4"; mRelationCompetability_TermShort[129, 2] = "5";
            mRelationCompetability_TermShort[130, 0] = "19"; mRelationCompetability_TermShort[130, 1] = "5"; mRelationCompetability_TermShort[130, 2] = "8";
            mRelationCompetability_TermShort[131, 0] = "19"; mRelationCompetability_TermShort[131, 1] = "6"; mRelationCompetability_TermShort[131, 2] = "4";
            mRelationCompetability_TermShort[132, 0] = "19"; mRelationCompetability_TermShort[132, 1] = "7"; mRelationCompetability_TermShort[132, 2] = "3";
            mRelationCompetability_TermShort[133, 0] = "19"; mRelationCompetability_TermShort[133, 1] = "8"; mRelationCompetability_TermShort[133, 2] = "4";
            mRelationCompetability_TermShort[134, 0] = "19"; mRelationCompetability_TermShort[134, 1] = "9"; mRelationCompetability_TermShort[134, 2] = "8";
            mRelationCompetability_TermShort[135, 0] = "19"; mRelationCompetability_TermShort[135, 1] = "11"; mRelationCompetability_TermShort[135, 2] = "8";
            mRelationCompetability_TermShort[136, 0] = "19"; mRelationCompetability_TermShort[136, 1] = "22"; mRelationCompetability_TermShort[136, 2] = "8";
            mRelationCompetability_TermShort[137, 0] = "19"; mRelationCompetability_TermShort[137, 1] = "33"; mRelationCompetability_TermShort[137, 2] = "4";

            mSexualCompatability = new Tuple<int, int, int>[]
            {
                Tuple.Create(0, 1, 8), Tuple.Create(0, 4, 7), Tuple.Create(0, 7, 3), Tuple.Create(0, 11, 10),
                Tuple.Create(0, 2, 10), Tuple.Create(0, 5, 10), Tuple.Create(0, 8, 7), Tuple.Create(0, 22, 10),
                Tuple.Create(0, 3, 8), Tuple.Create(0, 6, 8), Tuple.Create(0, 9, 10), Tuple.Create(0, 33, 6),

                Tuple.Create(1, 1, 9), Tuple.Create(1, 4, 8), Tuple.Create(1, 7, 4), Tuple.Create(1, 11, 9),
                Tuple.Create(1, 2, 8), Tuple.Create(1, 5, 7), Tuple.Create(1, 8, 8), Tuple.Create(1, 22, 10),
                Tuple.Create(1, 3, 7), Tuple.Create(1, 6, 7), Tuple.Create(1, 9, 10), Tuple.Create(1, 33, 7),

                Tuple.Create(2, 1, 8), Tuple.Create(2, 4, 7), Tuple.Create(2, 7, 3), Tuple.Create(2, 11, 10),
                Tuple.Create(2, 2, 10), Tuple.Create(2, 5, 10), Tuple.Create(2, 8, 7), Tuple.Create(2, 22, 10),
                Tuple.Create(2, 3, 8), Tuple.Create(2, 6, 8), Tuple.Create(2, 9, 10), Tuple.Create(2, 33, 6),

                Tuple.Create(3, 1, 7), Tuple.Create(3, 4, 7), Tuple.Create(3, 7, 4), Tuple.Create(3, 11, 7),
                Tuple.Create(3, 2, 8), Tuple.Create(3, 5, 9), Tuple.Create(3, 8, 8), Tuple.Create(3, 22, 10),
                Tuple.Create(3, 3, 9), Tuple.Create(3, 6, 7), Tuple.Create(3, 9, 9), Tuple.Create(3, 33, 6),

                Tuple.Create(4, 1, 8), Tuple.Create(4, 4, 9), Tuple.Create(4, 7, 6), Tuple.Create(4, 11, 9),
                Tuple.Create(4, 2, 7), Tuple.Create(4, 5, 9), Tuple.Create(4, 8, 10), Tuple.Create(4, 22, 10),
                Tuple.Create(4, 3, 7), Tuple.Create(4, 6, 8), Tuple.Create(4, 9, 9), Tuple.Create(4, 33, 8),

                Tuple.Create(5, 1, 7), Tuple.Create(5, 4, 9), Tuple.Create(5, 7, 3), Tuple.Create(5, 11, 9),
                Tuple.Create(5, 2, 10), Tuple.Create(5, 5, 10), Tuple.Create(5, 8, 9), Tuple.Create(5, 22, 10),
                Tuple.Create(5, 3, 9), Tuple.Create(5, 6, 7), Tuple.Create(5, 9, 10), Tuple.Create(5, 33, 6),

                Tuple.Create(6, 1, 7), Tuple.Create(6, 4, 8), Tuple.Create(6, 7, 4), Tuple.Create(6, 11, 9),
                Tuple.Create(6, 2, 8), Tuple.Create(6, 5, 7), Tuple.Create(6, 8, 8), Tuple.Create(6, 22, 9),
                Tuple.Create(6, 3, 7), Tuple.Create(6, 6, 8), Tuple.Create(6, 9, 9), Tuple.Create(6, 33, 7),

                Tuple.Create(7, 1, 4), Tuple.Create(7, 4, 6), Tuple.Create(7, 7, 7), Tuple.Create(7, 11, 4),
                Tuple.Create(7, 2, 3), Tuple.Create(7, 5, 3), Tuple.Create(7, 8, 5), Tuple.Create(7, 22, 6),
                Tuple.Create(7, 3, 4), Tuple.Create(7, 6, 4), Tuple.Create(7, 9, 5), Tuple.Create(7, 33, 6),

                Tuple.Create(8, 1, 8), Tuple.Create(8, 4, 10), Tuple.Create(8, 7, 5), Tuple.Create(8, 11, 10),
                Tuple.Create(8, 2, 7), Tuple.Create(8, 5, 9), Tuple.Create(8, 8, 10), Tuple.Create(8, 22, 10),
                Tuple.Create(8, 3, 8), Tuple.Create(8, 6, 8), Tuple.Create(8, 9, 10), Tuple.Create(8, 33, 7),

                Tuple.Create(9, 1, 10), Tuple.Create(9, 4, 9), Tuple.Create(9, 7, 5), Tuple.Create(9, 11, 10),
                Tuple.Create(9, 2, 10), Tuple.Create(9, 5, 10), Tuple.Create(9, 8, 10), Tuple.Create(9, 22, 10),
                Tuple.Create(9, 3, 9), Tuple.Create(9, 6, 9), Tuple.Create(9, 9, 10), Tuple.Create(9, 33, 7),

                Tuple.Create(11, 1, 9), Tuple.Create(11, 4, 9), Tuple.Create(11, 7, 4), Tuple.Create(11, 11, 10),
                Tuple.Create(11, 2, 10), Tuple.Create(11, 5, 9), Tuple.Create(11, 8, 10), Tuple.Create(11, 22, 10),
                Tuple.Create(11, 3, 7), Tuple.Create(11, 6, 9), Tuple.Create(11, 9, 10), Tuple.Create(11, 33, 6),

                Tuple.Create(22, 1, 10), Tuple.Create(22, 4, 10), Tuple.Create(22, 7, 6), Tuple.Create(22, 11, 10),
                Tuple.Create(22, 2, 10), Tuple.Create(22, 5, 10), Tuple.Create(22, 8, 10), Tuple.Create(22, 22, 10),
                Tuple.Create(22, 3, 10), Tuple.Create(22, 6, 9), Tuple.Create(22, 9, 10), Tuple.Create(22, 33, 7),

                Tuple.Create(33, 1, 7), Tuple.Create(33, 4, 8), Tuple.Create(33, 7, 6), Tuple.Create(33, 11, 6),
                Tuple.Create(33, 2, 6), Tuple.Create(33, 5, 6), Tuple.Create(33, 8, 7), Tuple.Create(33, 22, 7),
                Tuple.Create(33, 3, 6), Tuple.Create(33, 6, 7), Tuple.Create(33, 9, 7), Tuple.Create(33, 33, 8),

                Tuple.Create(13, 1, 6), Tuple.Create(13, 4, 7), Tuple.Create(13, 7, 2), Tuple.Create(13, 11, 5),
                Tuple.Create(13, 2, 5), Tuple.Create(13, 5, 6), Tuple.Create(13, 8, 6), Tuple.Create(13, 22, 6),
                Tuple.Create(13, 3, 6), Tuple.Create(13, 6, 6), Tuple.Create(13, 9, 6), Tuple.Create(13, 33, 5),

                Tuple.Create(14, 1, 7), Tuple.Create(14, 4, 6), Tuple.Create(14, 7, 3), Tuple.Create(14, 11, 6),
                Tuple.Create(14, 2, 6), Tuple.Create(14, 5, 7), Tuple.Create(14, 8, 6), Tuple.Create(14, 22, 7),
                Tuple.Create(14, 3, 7), Tuple.Create(14, 6, 6), Tuple.Create(14, 9, 7), Tuple.Create(14, 33, 6),

                Tuple.Create(16, 1, 4), Tuple.Create(16, 4, 6), Tuple.Create(16, 7, 7), Tuple.Create(16, 11, 4),
                Tuple.Create(16, 2, 3), Tuple.Create(16, 5, 3), Tuple.Create(16, 8, 5), Tuple.Create(16, 22, 6),
                Tuple.Create(16, 3, 4), Tuple.Create(16, 6, 4), Tuple.Create(16, 9, 5), Tuple.Create(16, 33, 6),

                Tuple.Create(19, 1, 7), Tuple.Create(19, 4, 6), Tuple.Create(19, 7, 3), Tuple.Create(19, 11, 5),
                Tuple.Create(19, 2, 5), Tuple.Create(19, 5, 7), Tuple.Create(19, 8, 7), Tuple.Create(19, 22, 7),
                Tuple.Create(19, 3, 7), Tuple.Create(19, 6, 6), Tuple.Create(19, 9, 7), Tuple.Create(19, 33, 5)
            };

            #endregion // Marriage

            #region Learning Success And Attantion Problems
            mLearnSccssAttPrbl_Chkra[0, 0] = "0"; mLearnSccssAttPrbl_Chkra[0, 1] = "10";
            mLearnSccssAttPrbl_Chkra[1, 0] = "1"; mLearnSccssAttPrbl_Chkra[1, 1] = "7";
            mLearnSccssAttPrbl_Chkra[2, 0] = "2"; mLearnSccssAttPrbl_Chkra[2, 1] = "8";
            mLearnSccssAttPrbl_Chkra[3, 0] = "3"; mLearnSccssAttPrbl_Chkra[3, 1] = "9";
            mLearnSccssAttPrbl_Chkra[4, 0] = "4"; mLearnSccssAttPrbl_Chkra[4, 1] = "9";
            mLearnSccssAttPrbl_Chkra[5, 0] = "5"; mLearnSccssAttPrbl_Chkra[5, 1] = "9";
            mLearnSccssAttPrbl_Chkra[6, 0] = "6"; mLearnSccssAttPrbl_Chkra[6, 1] = "8";
            mLearnSccssAttPrbl_Chkra[7, 0] = "7"; mLearnSccssAttPrbl_Chkra[7, 1] = "10";
            mLearnSccssAttPrbl_Chkra[8, 0] = "8"; mLearnSccssAttPrbl_Chkra[8, 1] = "10";
            mLearnSccssAttPrbl_Chkra[9, 0] = "9"; mLearnSccssAttPrbl_Chkra[9, 1] = "10";
            mLearnSccssAttPrbl_Chkra[10, 0] = "11"; mLearnSccssAttPrbl_Chkra[10, 1] = "10";
            mLearnSccssAttPrbl_Chkra[11, 0] = "22"; mLearnSccssAttPrbl_Chkra[11, 1] = "10";
            mLearnSccssAttPrbl_Chkra[12, 0] = "33"; mLearnSccssAttPrbl_Chkra[12, 1] = "7";
            mLearnSccssAttPrbl_Chkra[13, 0] = "13"; mLearnSccssAttPrbl_Chkra[13, 1] = "5";
            mLearnSccssAttPrbl_Chkra[14, 0] = "14"; mLearnSccssAttPrbl_Chkra[14, 1] = "6";
            mLearnSccssAttPrbl_Chkra[15, 0] = "16"; mLearnSccssAttPrbl_Chkra[15, 1] = "7";
            mLearnSccssAttPrbl_Chkra[16, 0] = "19"; mLearnSccssAttPrbl_Chkra[16, 1] = "6";
            mLearnSccssAttPrbl_Chkra[17, 0] = "20"; mLearnSccssAttPrbl_Chkra[17, 1] = "9";
            mLearnSccssAttPrbl_Chkra[18, 0] = "21"; mLearnSccssAttPrbl_Chkra[18, 1] = "8";
            mLearnSccssAttPrbl_Chkra[19, 0] = "23"; mLearnSccssAttPrbl_Chkra[19, 1] = "8";
            mLearnSccssAttPrbl_Chkra[20, 0] = "25"; mLearnSccssAttPrbl_Chkra[20, 1] = "6";
            mLearnSccssAttPrbl_Chkra[21, 0] = "30"; mLearnSccssAttPrbl_Chkra[21, 1] = "10";
            mLearnSccssAttPrbl_Chkra[22, 0] = "31"; mLearnSccssAttPrbl_Chkra[22, 1] = "8";
            mLearnSccssAttPrbl_Chkra[23, 0] = "מאוזן"; mLearnSccssAttPrbl_Chkra[23, 1] = "10";
            mLearnSccssAttPrbl_Chkra[24, 0] = "חצי מאוזן"; mLearnSccssAttPrbl_Chkra[24, 1] = "8";
            mLearnSccssAttPrbl_Chkra[25, 0] = "לא מאוזן"; mLearnSccssAttPrbl_Chkra[25, 1] = "4";

            mLearnSccssAttPrbl_LC[0, 0] = "1"; mLearnSccssAttPrbl_LC[0, 1] = "7";
            mLearnSccssAttPrbl_LC[1, 0] = "2"; mLearnSccssAttPrbl_LC[1, 1] = "7";
            mLearnSccssAttPrbl_LC[2, 0] = "4"; mLearnSccssAttPrbl_LC[2, 1] = "7";
            mLearnSccssAttPrbl_LC[3, 0] = "6"; mLearnSccssAttPrbl_LC[3, 1] = "7";
            mLearnSccssAttPrbl_LC[4, 0] = "33"; mLearnSccssAttPrbl_LC[4, 1] = "7";
            mLearnSccssAttPrbl_LC[5, 0] = "3"; mLearnSccssAttPrbl_LC[5, 1] = "9";
            mLearnSccssAttPrbl_LC[6, 0] = "5"; mLearnSccssAttPrbl_LC[6, 1] = "9";
            mLearnSccssAttPrbl_LC[7, 0] = "7"; mLearnSccssAttPrbl_LC[7, 1] = "9";
            mLearnSccssAttPrbl_LC[8, 0] = "8"; mLearnSccssAttPrbl_LC[8, 1] = "9";
            mLearnSccssAttPrbl_LC[9, 0] = "16"; mLearnSccssAttPrbl_LC[9, 1] = "9";
            mLearnSccssAttPrbl_LC[10, 0] = "0"; mLearnSccssAttPrbl_LC[10, 1] = "10";
            mLearnSccssAttPrbl_LC[11, 0] = "9"; mLearnSccssAttPrbl_LC[11, 1] = "10";
            mLearnSccssAttPrbl_LC[12, 0] = "11"; mLearnSccssAttPrbl_LC[12, 1] = "10";
            mLearnSccssAttPrbl_LC[13, 0] = "22"; mLearnSccssAttPrbl_LC[13, 1] = "10";
            mLearnSccssAttPrbl_LC[14, 0] = "13"; mLearnSccssAttPrbl_LC[14, 1] = "5";
            mLearnSccssAttPrbl_LC[15, 0] = "14"; mLearnSccssAttPrbl_LC[15, 1] = "5";
            mLearnSccssAttPrbl_LC[16, 0] = "19"; mLearnSccssAttPrbl_LC[16, 1] = "5";
            #endregion

            #region Chkara_Opening_Nums         
            mChakra1_OpeningValues = new int[] { 1, 2, 3, 4, 5, 6, 8, 9, 11, 22, 33 }; // Crown
            mChakra2_OpeningValues = new int[] { 1, 2, 3, 4, 5, 6, 8, 9, 11, 22, 33 }; // Third Eye
            mChakra3_OpeningValues = new int[] { 1, 3, 5, 8, 9, 11, 22 }; // Throat
            mChakra4_OpeningValues = new int[] { 1, 2, 3, 4, 5, 6, 7, 8 }; // Heart
            mChakra5_OpeningValues = new int[] { 1, 3, 5, 8, 9, 11, 22 }; // Solar Plexus
            mChakra6_OpeningValues = new int[] { 1, 3, 5, 8, 9, 11, 22 }; // Sex Creation
            mChakra7_OpeningValues = new int[] { 1, 2, 3, 4, 5, 6, 7, 8, 9, 11, 22, 33 }; // Root  
            #endregion

            #region Health Tables
            mHealthTable = new int[16, 4] { {1,10,8,5},
                                                {2,6,6,10},
                                                {3,7,7,5},
                                                {4,9,7,5},
                                                {5,7,6,5},
                                                {6,7,7,6},
                                                {7,7,5,8},
                                                {8,10,8,10},
                                                {9,10,8,10},
                                                {11,10,8,10},
                                                {22,10,8,10},
                                                {33,9,7,10},
                                                {13,8,7,5},
                                                {14,7,6,5},
                                                {16,5,5,8},
                                                {19,6,6,7}   };
            //mHealthTable = new int[16, 4] { {1,10,8,5},
            //                                    {2,6,6,10},
            //                                    {3,7,7,5},
            //                                    {4,9,7,5},
            //                                    {5,7,6,5},
            //                                    {6,7,7,10},
            //                                    {7,7,5,8},
            //                                    {8,10,8,10},
            //                                    {9,10,8,10},
            //                                    {11,10,8,10},
            //                                    {22,10,8,10},
            //                                    {33,9,7,10},
            //                                    {13,8,7,5},
            //                                    {14,7,6,5},
            //                                    {16,5,5,8},
            //                                    {19,6,6,7}   };
            #endregion
        }
        #endregion

        #region Terminator
        /// <summary> this sub will close all the proccesses open before closing the DLL
        /// 
        /// </summary>
        public void Terminate()
        {
            //try
            //{
            //    if (mDBConnection.State == System.Data.ConnectionState.Open)
            //    {
            //        mDBConnection.Close();
            //    }
            //}
            //catch
            //{
            //    // db wal closed
            //}
        }
        #endregion

        #region Properties
        public char Delimiter
        {
            get
            {
                return mDelimiter;
            }
        }

        public bool isPasswordOK
        {
            get
            {
                return mProffer.Proofed;
            }
        }
        #endregion

        #region Public Methods

        //#region DB
        ///// <summary> this will open a new connection to the DB
        ///// 
        ///// </summary>
        ///// <param name="InPath">path of the db file location</param>
        //public void OpenDB()
        //{
        //    //string InPath = mDBLocation + "\\" + mDBName;

        //    string InPath = System.IO.Path.Combine(System.Environment.CurrentDirectory,mDBName);

        //    mDBConnection = new OleDbConnection();
        //    //mDBConnection.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" + "Data Source=" + InPath + ";Jet OLEDB:Database Password=" + mDBPassword + ";";
        //    mDBConnection.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" + "Data Source=" + InPath.ToString() + ";User id=; Password=" + mDBPassword + ";"; // User Id=Admin
        //    mDBConnection.Open();
        //}
        //#endregion

        public bool isDynamicChakraOpen(int chakraIdx, int[] vals, out string found)
        {
            bool res = false;
            found = "";

            for (int i = 0; i < vals.Length; i++)
            {
                #region Old Code


                //if (chakraIdx > 0 && isNumInArray(mDynamicChakraOpenValues[chakraIdx - 1], vals[i]))
                //{
                //    res = true;
                //    found += vals[i].ToString() + ",";
                //}

                #endregion Old Code

                if (chakraIdx > 0 && vals.Length >= 2 && (mMasterNumbers.Any(m => m == vals[0]) || (mCarmaticNumbers.Any(c => c == vals[0]))) && mDynamicChakraOpenValues[chakraIdx - 1].Any(n => n == vals[0]))
                {                    
                    res = true;
                    found += vals[i].ToString() + ",";
                    break;

                }
                else if (chakraIdx > 0 && vals.Length >= 2 && (!mMasterNumbers.Any(m => m == vals[0]) || (!mCarmaticNumbers.Any(c => c == vals[0]))) && mDynamicChakraOpenValues[chakraIdx - 1].Any(n => n == vals[i]))
                {
                    res = true;
                    found += vals[i].ToString() + ",";

                }
                else if (chakraIdx > 0 && vals.Length == 1 && mDynamicChakraOpenValues[chakraIdx - 1].Any(n => n == vals[i]))
                {
                    res = true;
                    found += vals[i].ToString() + ",";

                }

            }

            if (found.Length > 0)
            {
                found = found.Substring(0, found.Length - 1);
            }

            return res;
        }

        public void SetMainMasterValue(bool masterValue)
        {
            isPersonMaster = masterValue;
        }

        public bool isMaterNumber(int num)
        {
            if (isPersonMaster == true)
            {
                for (int i = 0; i < mMasterNumbers.Length; i++)
                {
                    if (num == mMasterNumbers[i])
                    {
                        return true;
                    }
                }
                return false;
            }
            else
            {
                for (int i = 0; i < mMasterNumbers.Length; i++)
                {
                    if ((num == mMasterNumbers[i]) && ((num != 11) && (num != 22)))
                    {
                        return true;
                    }
                }
                return false;
            }

        }

        public bool isCarmaticNumber(int num)
        {
            for (int i = 0; i < mCarmaticNumbers.Length; i++)
            {
                if (num == mCarmaticNumbers[i])
                {
                    return true;
                }
            }
            return false;
        }

        public bool isHalfCarmaticNumber(int num)
        {
            for (int i = 0; i < mHalfCarmaticNumbers.Length; i++)
            {
                if (num == mHalfCarmaticNumbers[i])
                {
                    return true;
                }
            }
            return false;
        }

        public bool isDelayingNumber(int num)
        {
            return mDelayingNumbers.Any(n => n == num);

        }

        /// <summary> replaces the final letters of the language by the normal letters
        /// 
        /// </summary>
        /// <param name="name">given name (string)</param>
        /// <returns>fixed name (string)</returns>
        public string ChangeFinalChars(string name)
        {
            string FixedName = "";

            char[] sp = name.ToCharArray();
            char OutC;
            for (int i = 0; i < sp.Length; i++)
            {
                if ((sp[i] != "-".ToCharArray()[0]) && (sp[i] != " ".ToCharArray()[0]) && (sp[i] != "'".ToCharArray()[0]))  // ignoring "-" & " " & "'"
                {
                    bool res = isFinalLetter(sp[i], out OutC);

                    if (res == true)
                    {
                        FixedName = FixedName + OutC;
                    }
                    else
                    {
                        FixedName = FixedName + sp[i];
                    }
                }
            }

            return FixedName.ToString();
        }

        /// <summary> Calculates Sum
        /// 
        /// </summary>
        /// <param name="num">in number</param>
        /// <returns>out string for textbox</returns>
        public string CalcSum(int num)
        {
            return Nums_2_String(num);
        }

        /// <summary>Returns the Head Calc.
        /// HEAD
        /// </summary>
        /// <param name="FirstName">string of thje fist name</param>
        /// <returns>nums string</returns>
        public string VouleCalc(string FirstName, string LastName)
        {
            char[] FName = FirstName.ToCharArray();
            int sum = 0;
            for (int i = 0; i < FirstName.Length; i++)
            {
                if (isVoule(FName[i]) == true)
                {
                    sum = sum + char2int(FName[i]);
                }
            }

            string sumF = Nums_2_String(sum);

            char[] LName = LastName.ToCharArray();
            sum = 0;
            for (int i = 0; i < LastName.Length; i++)
            {
                if (isVoule(LName[i]) == true)
                {
                    sum = sum + char2int(LName[i]);
                }
            }

            string sumL = Nums_2_String(sum);

            string[] spFN = sumF.Split(mDelimiter);
            string[] spLN = sumL.Split(mDelimiter);

            sum = 0;
            if (is_Special(Convert.ToInt16(spFN[0])))
            {
                sum = sum + Convert.ToInt16(spFN[0]);
            }
            else
            {
                sum = sum + Convert.ToInt16(spFN[spFN.Length - 1]);
            }

            if (is_Special(Convert.ToInt16(spLN[0])))
            {
                sum = sum + Convert.ToInt16(spLN[0]);
            }
            else
            {
                sum = sum + Convert.ToInt16(spLN[spLN.Length - 1]);
            }

            if ((is_Special(sum) == false) && (sum > 9))
            {
                return (sum.ToString() + mDelimiter.ToString() + Nums_2_String(sum)).ToString();
            }
            else
            {
                return Nums_2_String(sum);
            }
        }

        /// <summary>Returns the Legs Calc.
        /// LEGS
        /// </summary>
        /// <param name="FirstName">string of thje fist name</param>
        /// <returns>nums string</returns>
        public string ConsonantsCalc(string FirstName, string LastName)
        {
            char[] FName = FirstName.ToCharArray();
            int sum = 0;
            for (int i = 0; i < FirstName.Length; i++)
            {
                if (isVoule(FName[i]) == false)
                {
                    sum = sum + char2int(FName[i]);
                }
            }

            string sumF = Nums_2_String(sum);

            char[] LName = LastName.ToCharArray();
            sum = 0;
            for (int i = 0; i < LastName.Length; i++)
            {
                if (isVoule(LName[i]) == false)
                {
                    sum = sum + char2int(LName[i]);
                }
            }

            string sumL = Nums_2_String(sum);

            string[] spFN = sumF.Split(mDelimiter);
            string[] spLN = sumL.Split(mDelimiter);

            sum = 0;
            if (is_Special(Convert.ToInt16(spFN[0])))
            {
                sum = sum + Convert.ToInt16(spFN[0]);
            }
            else
            {
                sum = sum + Convert.ToInt16(spFN[spFN.Length - 1]);
            }

            if (is_Special(Convert.ToInt16(spLN[0])))
            {
                sum = sum + Convert.ToInt16(spLN[0]);
            }
            else
            {
                sum = sum + Convert.ToInt16(spLN[spLN.Length - 1]);
            }

            if ((is_Special(sum) == false) && (sum > 9))
            {
                return (sum.ToString() + mDelimiter.ToString() + Nums_2_String(sum)).ToString();
            }
            else
            {
                return Nums_2_String(sum);
            }
        }

        /// <summary> full name calculation
        /// HANDS
        /// </summary>
        /// <param name="FistName">string of the first name</param>
        /// <param name="LastName">string of the last name</param>
        /// <returns>nums string</returns>
        public string FullNameCalc(string FirstName, string LastName)
        {
            char[] FName = FirstName.ToCharArray();
            char[] LName = LastName.ToCharArray();
            int sum = 0;

            for (int i = 0; i < FirstName.Length; i++)
            {
                sum = sum + char2int(FName[i]);
            }

            string sumFname = Nums_2_String(sum);

            sum = 0;
            for (int i = 0; i < LastName.Length; i++)
            {
                sum = sum + char2int(LName[i]);
            }

            string sumLname = Nums_2_String(sum);

            string[] spFN = sumFname.Split(mDelimiter);
            string[] spLN = sumLname.Split(mDelimiter);

            sum = 0;
            if (is_Special(Convert.ToInt16(spFN[0])))
            {
                sum = sum + Convert.ToInt16(spFN[0]);
            }
            else
            {
                sum = sum + Convert.ToInt16(spFN[spFN.Length - 1]);
            }

            if (is_Special(Convert.ToInt16(spLN[0])))
            {
                sum = sum + Convert.ToInt16(spLN[0]);
            }
            else
            {
                sum = sum + Convert.ToInt16(spLN[spLN.Length - 1]);
            }

            if ((is_Special(sum) == false) && (sum > 9))
            {
                return (sum.ToString() + mDelimiter.ToString() + Nums_2_String(sum)).ToString();
            }
            else
            {
                return Nums_2_String(sum);
            }
        }

        /// <summary> calculates the 
        /// מקלעת השמש
        /// </summary>
        /// <param name="BirthDay">the birthday in a DateTime format</param>
        /// <returns>nums string</returns>
        public string FateCalc(DateTime BirthDay)
        {
            int d = BirthDay.Day;
            int m = BirthDay.Month;
            int y = BirthDay.Year;

            string[] sD = Nums_2_String(d).Split(mDelimiter);
            string[] sM = Nums_2_String(m).Split(mDelimiter);
            string[] sY = Nums_2_String(y).Split(mDelimiter);

            int cD = Convert.ToInt16(sD[0]);
            int cM = Convert.ToInt16(sM[0]);
            int cY = Convert.ToInt16(sY[0]);

            string scD = Nums_2_String(cD);
            string scM = Nums_2_String(cM);
            string scY = Nums_2_String(cY);

            d = Convert.ToInt16(scD.Split(mDelimiter)[0]);
            m = Convert.ToInt16(scM.Split(mDelimiter)[0]);
            y = Convert.ToInt16(scY.Split(mDelimiter)[0]);

            int sum = d + m + y;

            if ((is_Special(sum) == false) && (sum > 9))
            {
                return (sum.ToString() + mDelimiter.ToString() + Nums_2_String(sum)).ToString();
            }
            else
            {
                return Nums_2_String(sum);
            }

        }

        /// <summary> calculates the wound
        /// HEART
        /// </summary>
        /// <param name="fullNameCalc">the full name calculation</param>
        /// <param name="faithCalc">faith calculation</param>
        /// <returns>nums string</returns>
        public string WoundCalc(string fullNameCalc, string FateCalc)
        {
            string[] sFN = fullNameCalc.Split(mDelimiter);
            string[] sF = FateCalc.Split(mDelimiter);

            int cName = Convert.ToInt16(sFN[sFN.Length - 1]);
            int cFate = Convert.ToInt16(sF[sF.Length - 1]);

            return Nums_2_String(Math.Abs(cName - cFate));
        }

        /// <summary> Maturity Calculation
        /// 
        /// </summary>
        /// <param name="fullNameCalc">the full name calculation</param>
        /// <param name="faithCalc">faith calculation</param>
        /// <returns>nums string</returns>
        public string MaturityCalc(string fullNameCalc, string FateCalc)
        {
            string[] spFullN = fullNameCalc.Split(mDelimiter);
            string[] spFate = FateCalc.Split(mDelimiter);

            int num1 = GetCorrectNumberFromSplitedString(spFullN);
            int num2 = GetCorrectNumberFromSplitedString(spFate);

            string outs = Nums_2_String(num1 + num2);

            if ((num1 + num2 >= 10) && (is_Special(GetCorrectNumberFromSplitedString(outs.Split(mDelimiter))) == false))
            {
                outs = (num1 + num2).ToString() + mDelimiter + outs;
            }

            return outs;
        }

        /// <summary> Day of the month
        /// Calculates the day of the month
        /// </summary>
        /// <param name="BirthDay">the persons birthday</param>
        /// <returns>nums string</returns>
        public string DayOfTheMonth(DateTime BirthDay)
        {
            if ((is_Special(BirthDay.Day) == false) && (BirthDay.Day > 9))
            {
                return (BirthDay.Day.ToString() + mDelimiter.ToString() + Nums_2_String(BirthDay.Day)).ToString();

            }
            else
            {
                return Nums_2_String(BirthDay.Day);
            }
        }

        public string PrivateNameCalc(string FirstName)
        {
            char[] FName = FirstName.ToCharArray();

            int sum = 0;

            for (int i = 0; i < FirstName.Length; i++)
            {
                sum = sum + char2int(FName[i]);
            }


            if ((is_Special(sum) == false) && (sum > 9))
            {
                return (sum.ToString() + mDelimiter.ToString() + Nums_2_String(sum)).ToString();
            }
            else
            {
                return Nums_2_String(sum);
            }
        }

        public int GetAstroData(DateTime date, out string LuckName)
        {
            return AstroData.Date2AstroLuck(date, out LuckName);
        }

        public string GetAstroLuckNameByNumber(int num)
        {
            return AstroData.GetAstroNameByNumber(num);
        }

        #region Chakra Calcs.
        public string CalcMasterChkra(string Name, string outputVal)
        {
            return CalcChakra(Name, outputVal, Letters_Master, mChakra1_OpeningValues);
        }

        public string CalcCrownChakra(string Name, string outputVal)
        {
            return CalcChakra(Name, outputVal, Letters_Crown, mChakra1_OpeningValues);
        }

        public string CalcHeartChakra(string Name, string outputVal)
        {
            return CalcChakra(Name, outputVal, Letters_Heart, mChakra4_OpeningValues);
        }

        public string CalcRootChakra(string Name, string outputVal)
        {
            return CalcChakra(Name, outputVal, Letters_Root, mChakra7_OpeningValues);
        }

        public string CalcSex_CreationChakra(string Name, string outputVal)
        {
            return CalcChakra(Name, outputVal, Letters_Sex_Creation, mChakra6_OpeningValues);
        }

        public string CalcSunChakra(string Name, string outputVal)
        {
            return CalcChakra(Name, outputVal, Letters_Sun, mChakra5_OpeningValues);
        }

        public string CalcThirdEyeChakra(string Name, string outputVal)
        {
            return CalcChakra(Name, outputVal, Letters_ThirdEye, mChakra2_OpeningValues);
        }

        public string CalcThroughtChakra(string Name, string outputVal)
        {
            return CalcChakra(Name, outputVal, Letters_Throught, mChakra3_OpeningValues);
        }
        #endregion

        /// <summary> Constructing Strings to show on the Intansive map
        /// מפה אינטנסיבית
        /// </summary>
        /// <param name="FirstName">Persons First Name</param>
        /// <param name="LastName">Persons Last Name</param>
        /// <returns>9 strings (in a string array) to be placed into the Map</returns>
        public string[] CreateIntensiveMap(string FirstName, string LastName)
        {
            string innrDelimiter = " , ";
            string[] outLS = new string[9];

            string FullName = ChangeFinalChars(FirstName) + ChangeFinalChars(LastName);

            for (int i = 0; i < 9; i++)
            {
                outLS[i] = "";
            }

            for (long i = 0; i < FullName.Length; i++)
            {
                int num = char2int(FullName.Substring((int)i, 1).ToCharArray()[0]);

                if (num > 0)
                {
                    if (outLS[num - 1] == "")
                    {
                        outLS[num - 1] = num.ToString();
                    }
                    else
                    {
                        outLS[num - 1] = outLS[num - 1] + innrDelimiter + num.ToString();
                    }

                }

            }
            return outLS;
        }

        /// <summary> Constructing Strings to show on the Pitagoras Squares
        /// ריבועי פיתגורס
        /// </summary>
        /// <param name="FirstName">Persons First Name</param>
        /// <param name="LastName">Persons Last Name</param>
        /// <returns>9 strings (in a string array) to be placed into the Map</returns>
        public string[] CreatePitagorasSquares(DateTime BirthDay)
        {
            string innrDelimiter = " , ";
            string[] outLS = new string[9];

            string NUMs = BirthDay.Day.ToString() + BirthDay.Month.ToString() + BirthDay.Year.ToString();

            for (int i = 0; i < 9; i++)
            {
                outLS[i] = "";
            }

            for (long i = 0; i < NUMs.Length; i++)
            {
                int num = Convert.ToInt16(NUMs.Substring((int)i, 1));
                if (num != 0)
                {
                    if (outLS[num - 1] == "")
                    {
                        outLS[num - 1] = num.ToString();
                    }
                    else
                    {
                        outLS[num - 1] = outLS[num - 1] + innrDelimiter + num.ToString();
                    }
                }
            }
            return outLS;
        }

        /// <summary> Calcs. the Life Cycles by the birthday only
        /// מחזורי החיים
        /// </summary>
        /// <param name="BDay">Birth Day - DateTime object</param>
        /// <returns>16 strings (string array) by order of appearance</returns>
        public int[] CareteLifeCycles(DateTime BDay)
        {
            int[] outLS = new int[16];
            for (int i = 0; i < 16; i++)
            {
                outLS[i] = 0;
            }

            string sMiklaat = FateCalc(BDay);
            string[] sMikSplit = sMiklaat.Split(mDelimiter);

            #region Ages
            outLS[0] = 36 - Convert.ToInt16(sMikSplit[sMikSplit.Length - 1]);
            outLS[4] = outLS[0] + 9;
            outLS[8] = outLS[4] + 9;
            outLS[12] = 120;
            #endregion

            #region Cycles
            outLS[1] = GetCorrectNumberFromSplitedString(Nums_2_String(BDay.Month).Split(mDelimiter)); //Nums_2_String
            outLS[5] = BDay.Day; //Nums_2_String
            outLS[9] = outLS[5];//Nums_2_String

            char[] cy = BDay.Year.ToString().ToCharArray();
            int sum = 0;
            for (int i = 0; i < cy.Length; i++)
            {
                sum = sum + Convert.ToInt16(cy[i]) - 48;
            }
            outLS[13] = sum; //Nums_2_String

            #endregion Cycles


            string[] y, m, d;
            y = Nums_2_String(sum).Split(mDelimiter);
            m = Nums_2_String(BDay.Month).Split(mDelimiter);
            d = Nums_2_String(BDay.Day).Split(mDelimiter);

            int sy, sd, sm;
            sy = GetCorrectNumberFromSplitedString(y); //Convert.ToInt16(y[y.Length - 1]);
            sm = GetCorrectNumberFromSplitedString(m); //Convert.ToInt16(m[m.Length - 1]);
            sd = GetCorrectNumberFromSplitedString(d); //Convert.ToInt16(d[d.Length - 1]);

            #region Climax
            outLS[2] = GetCorrectNumberFromSplitedString(Nums_2_String(BDay.Month).Split(mDelimiter)) + GetCorrectNumberFromSplitedString(Nums_2_String(BDay.Day).Split(mDelimiter));
            outLS[6] = sy + GetCorrectNumberFromSplitedString(Nums_2_String(BDay.Day).Split(mDelimiter));
            outLS[10] = outLS[6] + outLS[2];
            outLS[14] = sy + sm;
            #endregion

            #region Chalange
            outLS[3] = Math.Abs(Convert.ToInt16(m[m.Length - 1]) - Convert.ToInt16(d[d.Length - 1]));
            outLS[7] = Math.Abs(Convert.ToInt16(y[y.Length - 1]) - Convert.ToInt16(d[d.Length - 1]));
            outLS[11] = Math.Abs(outLS[3] - outLS[7]);
            outLS[15] = Math.Abs(Convert.ToInt16(y[y.Length - 1]) - Convert.ToInt16(m[m.Length - 1]));
            #endregion

            return outLS;
        }

        public void CalcPersonalInfo(DateTime BDay, DateTime Time2Calc, out string sPersonalYear, out string sPersonalMonth, out string sPersonalDay)
        {
            int PersonalYear, PersonalMonth, PersonalDay;

            int sum = 0;
            string[] sNum = Nums_2_String(BDay.Day).Split(mDelimiter);
            sum += GetCorrectNumberFromSplitedString(sNum);//Convert.ToInt16(sNum[sNum.Length - 1]);

            sNum = Nums_2_String(BDay.Month).Split(mDelimiter);
            sum += GetCorrectNumberFromSplitedString(sNum);//Convert.ToInt16(sNum[sNum.Length - 1]);

            DateTime today = Time2Calc;
            DateTime MovedBDay;
            //bool isDatePossible = DateTime.TryParse(today.Year.ToString() + "-" + BDay.Month.ToString() + "-" + BDay.Day.ToString(), out MovedBDay);
            //if (isDatePossible == false)
            //{
            //    MovedBDay = new DateTime(today.Year, BDay.Month, BDay.Day-1);
            //}

            MovedBDay = BDay.AddYears(today.Year - BDay.Year);
            TimeSpan diff = MovedBDay.Subtract(today);
            int y;
            if (diff.Days <= 0)
            {
                y = today.Year;
            }
            else
            {
                y = today.Year - 1;
            }

            char[] cy = y.ToString().ToCharArray();

            int sum2 = 0;
            for (int i = 0; i < cy.Length; i++)
            {
                sum2 += Convert.ToInt16(cy[i]) - 48;
            }

            sNum = Nums_2_String(sum2).Split(mDelimiter);
            //sum += Convert.ToInt16(sNum[sNum.Length - 1]);
            sum += GetCorrectNumberFromSplitedString(sNum);

            sPersonalYear = Nums_2_String(sum);
            PersonalYear = GetCorrectNumberFromSplitedString(Nums_2_String(sum).Split(mDelimiter));

            // month
            sum = PersonalYear + GetCorrectNumberFromSplitedString(Nums_2_String(today.Month).Split(mDelimiter));
            sNum = Nums_2_String(sum).Split(mDelimiter);
            PersonalMonth = GetCorrectNumberFromSplitedString(Nums_2_String(sum).Split(mDelimiter));//Convert.ToInt16(sNum[sNum.Length-1]);
            sPersonalMonth = Nums_2_String(sum);

            // day
            PersonalMonth = GetCorrectNumberFromSplitedString(sPersonalMonth.Split(mDelimiter));
            sum = PersonalMonth + GetCorrectNumberFromSplitedString(Nums_2_String(today.Day).Split(mDelimiter));
            sNum = Nums_2_String(sum).Split(mDelimiter);
            PersonalDay = GetCorrectNumberFromSplitedString(Nums_2_String(sum).Split(mDelimiter)); //Convert.ToInt16(sNum[sNum.Length - 1]);
            sPersonalDay = Nums_2_String(sum);
        }

        /// <summary> Calculates the Combined Map
        /// using Intensive Map + Pitagoras Map
        /// </summary>
        /// <param name="inIntensiveMapData">Intensive Map [9 values]</param>
        /// <param name="inPiatgorasMapData">Pitagoras Map [9 values]</param>
        /// <returns>Combined Map [9 values]</returns>
        public string[] CalcCombinedMap(string[] inIntensiveMapData, string[] inPiatgorasMapData)
        {
            string[] outCM = new string[9];

            for (int i = 0; i < inIntensiveMapData.Length; i++)
            {
                string sTemp = "";
                if (inIntensiveMapData[i] != "")
                {
                    sTemp = inIntensiveMapData[i];
                    if (inPiatgorasMapData[i] != "")
                    {
                        sTemp += " , " + inPiatgorasMapData[i];
                    }
                }
                else  // inIntensiveMapData[i] == ""
                {
                    sTemp = inPiatgorasMapData[i];
                }

                //= inIntensiveMapData[i] + " , " + inPiatgorasMapData[i];
                //string[] sSplit = sTemp.Split(",".ToCharArray()[0]);
                outCM[i] = sTemp;// +"  [" + sSplit.Length.ToString() + "]";
            }

            return outCM;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sPrivateName"></param>
        /// <param name="sLastName"></param>
        /// <param name="BDay"></param>
        /// <param name="MotherName"></param>
        /// <param name="FatherName"></param>
        /// <returns></returns>
        public int CalcParentsPresent(string sFirstName, string sLastName, DateTime BDay, string sMotherName, string sFatherName)
        {
            string chkraSexCreat = FullNameCalc(sFirstName, sLastName); ;
            string sFate = FateCalc(BDay);
            string sMomCalc = CalcName(sMotherName);
            string sDadCalc = CalcName(sFatherName);

            string[] sALL = new string[4] { chkraSexCreat, sFate, sMomCalc, sDadCalc };
            string[] sSplit;
            double sum = 0;

            for (int i = 0; i < sALL.Length; i++)
            {
                int tmpNum;
                sSplit = sALL[i].Split(mDelimiter);
                bool tmpIsSpecial = false;
                for (int j = 0; j < sSplit.Length; j++)
                {
                    tmpNum = Convert.ToInt16(sSplit[j]);
                    if (is_Special(tmpNum))
                    {
                        if (isCarmaticNumber(tmpNum))
                            sum -= tmpNum;
                        else sum += tmpNum;

                        tmpIsSpecial = true;
                    }
                }
                if (tmpIsSpecial == false)
                {
                    sum += Convert.ToInt16(sSplit[sSplit.Length - 1]);
                }
            }

            return (int)Math.Round(sum / 4.0);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sFirstName"></param>
        /// <param name="sLastName"></param>
        /// <param name="BDay"></param>
        /// <param name="sMotherName"></param>
        /// <param name="sFatherName"></param>
        /// <returns></returns>
        public int CalcUnique(string sFirstName, string sLastName, DateTime BDay)
        {
            string chkraSexCreat = FullNameCalc(sFirstName, sLastName);
            string sCrown = DayOfTheMonth(BDay);
            string sFate = FateCalc(BDay);
            string sNameCalc = CalcName(sFirstName);
            string sAstroLuck; int nAstroNum = AstroData.Date2AstroLuck(BDay, out sAstroLuck);

            string[] s = sFate.Split(mDelimiter);
            int testNUm = Convert.ToInt16(s[s.Length - 1]);
            int TestedValue = 36 - testNUm;

            DateTime today = DateTime.Now;
            int age = DateTime.Now.Year - BDay.Year;

            string[] sALL;
            string[] sSplit;
            double sum = 0;

            if (age <= TestedValue)
            {
                sALL = new string[5] { chkraSexCreat, sCrown, sFate, sNameCalc, nAstroNum.ToString() };
            }
            else
            {
                string sThirdEye = MaturityCalc(sNameCalc, sFate);
                string sHead = VouleCalc(sFirstName, sLastName);
                sALL = new string[6] { sThirdEye, sFate, chkraSexCreat, sHead, sNameCalc, nAstroNum.ToString() };
            }


            for (int i = 0; i < sALL.Length; i++)
            {
                int tmpNum;
                sSplit = sALL[i].Split(mDelimiter);
                bool tmpIsSpecial = false;
                for (int j = 0; j < sSplit.Length; j++)
                {
                    tmpNum = Convert.ToInt16(sSplit[j]);
                    if (is_Special(tmpNum))
                    {
                        if (isCarmaticNumber(tmpNum))
                            sum -= tmpNum;
                        else sum += tmpNum;

                        tmpIsSpecial = true;
                    }
                }
                if (tmpIsSpecial == false)
                {
                    sum += Convert.ToInt16(sSplit[sSplit.Length - 1]);
                }
            }
            return (int)Math.Round(sum / sALL.Length);




        }

        public string TestHramonicInfo(int num1, int num2)
        {
            string outHarmonic = "מאוזן";
            string outDissHarmonic = "לא מאוזן";
            string outHalfHarmonic = "חצי מאוזן";

            int[,] tmpArr = mCombHramonic;
            for (int i = 0; i < tmpArr.Length / 2; i++)
            {
                if (((num1 == tmpArr[i, 0]) && (num2 == tmpArr[i, 1])) || ((num2 == tmpArr[i, 0]) && (num1 == tmpArr[i, 1])))
                {
                    return outHarmonic;
                }
            }

            tmpArr = mCombHalfHramonic;
            for (int i = 0; i < tmpArr.Length / 2; i++)
            {
                if (((num1 == tmpArr[i, 0]) && (num2 == tmpArr[i, 1])) || ((num2 == tmpArr[i, 0]) && (num1 == tmpArr[i, 1])))
                {
                    return outHalfHarmonic;
                }
            }

            tmpArr = mCombDissHramonic;
            for (int i = 0; i < tmpArr.Length / 2; i++)
            {
                if (((num1 == tmpArr[i, 0]) && (num2 == tmpArr[i, 1])) || ((num2 == tmpArr[i, 0]) && (num1 == tmpArr[i, 1])))
                {
                    return outDissHarmonic;
                }
            }

            return "";
        }

        public int GetUniverseNumberFromChakra(string val)
        {
            var from = val.IndexOf('(') + 1;
            var to = val.IndexOf(')');
            var length = to - from;
            return int.Parse(val.Substring(from, length));
        }

        public int GetCorrectNumberFromSplitedString(string[] nums)
        {
            for (int i = 0; i < nums.Length; i++)
            {
                int tmpnum = Convert.ToInt16(nums[i]);
                if (is_Special(tmpnum) == true)
                    return tmpnum;
            }

            return Convert.ToInt16(nums[nums.Length - 1]);
        }

        public string CityCompetability(string sMikSun, string sCityName)
        {
            if (sCityName.Length > 0)
            {
                int iSunMik = GetCorrectNumberFromSplitedString(sMikSun.Split(mDelimiter));

                int sum = 0;
                for (int i = 0; i < sCityName.Length; i++)
                {
                    sum = sum + char2int(sCityName.Substring(i, 1).ToCharArray()[0]);
                }

                int iCity = GetCorrectNumberFromSplitedString(Nums_2_String(sum).Split(mDelimiter));

                return TestCompatebilityCityApp(iSunMik, iCity);
            }
            else
            {
                return "-";
            }
        }

        public string AppCompetability(string sMikSun, int iAppNum)
        {
            int iSunMik = GetCorrectNumberFromSplitedString(sMikSun.Split(mDelimiter));

            int iApp = GetCorrectNumberFromSplitedString(Nums_2_String(iAppNum).Split(mDelimiter));
            return TestCompatebilityCityApp(iSunMik, iApp);
        }

        public string CheckDepresedMap(string[] sChakras, string sPrivatName, int iAstroNum)
        {
            string sOutStr = "";
            bool found2 = false, found7 = false;
            int[] iArr = new int[sChakras.Length + 1];

            for (int i = 0; i < sChakras.Length; i++)
            {
                iArr[i] = GetCorrectNumberFromSplitedString(sChakras[i].Split(mDelimiter));

            }

            //iArr[6] = GetCorrectNumberFromSplitedString(sPrivatName.Split(mDelimiter));
            iArr[iArr.Length - 1] = iAstroNum;

            for (int i = 0; i < iArr.Length; i++)
            {
                if ((iArr[i] == 2) && (i != 2)) // Skip third eye index
                {
                    found2 = true;
                }

                if ((iArr[i] == 7) || (iArr[i] == 16))
                {
                    found7 = true;
                }

            }

            if ((found2 == true) && (found7 == true))
            {
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                {
                    sOutStr = "נטייה למפה דכאונית";
                }
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                {
                    sOutStr = "tendency to a depressive map";
                }
                return sOutStr;
            }

            //if ((found11 == true) && (found7 == true))
            //{
            //    if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
            //    {
            //        sOutStr = "בדוק נטייה לדכאון";
            //    }
            //    if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
            //    {
            //        sOutStr = "Check Tendency for Depression";
            //    }
            //}

            return sOutStr;
        }

        public string CheckStress(string[] sChakras, string sPrivatName, int iAstroNum, int firstClimax)
        {
            string sOutStr = "";
            bool found7 = false;
            int[] iArr = new int[sChakras.Length + 1];

            for (int i = 0; i < sChakras.Length; i++)
            {
                iArr[i] = GetCorrectNumberFromSplitedString(sChakras[i].Split(mDelimiter));

            }

            //iArr[6] = GetCorrectNumberFromSplitedString(sPrivatName.Split(mDelimiter));
            iArr[iArr.Length - 1] = iAstroNum;

            for (int i = 0; i < iArr.Length; i++)
            {
                if ((iArr[i] == 7) || (iArr[i] == 16))
                {
                    found7 = true;
                }
            }

            if ((firstClimax == 16) || (firstClimax == 7))
            {
                found7 = true;
            }

            if (found7 == true)
            {
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                {
                    sOutStr = "נטייה לחרדות";
                }
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                {
                    sOutStr = "Tendency towards anxiety";
                }
            }

            return sOutStr;
        }

        public string CheckAttention(string[] sChakras, string sPrivatName, int iAstroNum)
        {
            string sOutStr = "";
            bool foundC = false, foundN = false;
            List<int> Critical = new List<int> { 13, 14, 16, 19 };
            List<int> nonCritical = new List<int> { 1, 3, 5 };

            int[] iArr = new int[7];

            for (int i = 0; i < sChakras.Length; i++)
            {
                iArr[i] = GetCorrectNumberFromSplitedString(sChakras[i].Split(mDelimiter));
            }

            iArr[5] = GetCorrectNumberFromSplitedString(sPrivatName.Split(mDelimiter));
            iArr[6] = iAstroNum;

            for (int i = 0; i < iArr.Length; i++)
            {
                if (isNumInList(iArr[i], Critical))
                {
                    foundC = true;
                }
                else
                {
                    if (isNumInList(iArr[i], nonCritical))
                    {
                        foundN = true;
                    }
                }
            }

            if (foundC == true)
            {
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                {
                    sOutStr = "הפרעות קשב וריכוז";
                }
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                {
                    sOutStr = "ADHD appears";
                }
                return sOutStr;
            }

            if (foundN == true)
            {
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                {
                    sOutStr = "נטייה להפרעות קשב וריכוז";
                }
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                {
                    sOutStr = "Tendency to ADHD";
                }
                return sOutStr;
            }

            return sOutStr;
        }

        public string CheckFearFromUnkown(string[] sChakras, string sPrivateName, int iAstroNum)
        {
            string outStr = "";

            bool res = false;
            for (int i = 0; i < sChakras.Length; i++)
            {
                if (GetCorrectNumberFromSplitedString(sChakras[i].Split(mDelimiter)) == 2)
                {
                    res = true;
                }
            }

            if (iAstroNum == 2)
            {
                res = true;
            }

            //if (GetCorrectNumberFromSplitedString(sPrivateName.Split(mDelimiter)) == 2)
            //{
            //    res = true;
            //}

            if (res == true)
            {
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
                {
                    outStr = "נטייה לחשש ופחד מאירועים לא מוכרים.";
                }
                if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
                {
                    outStr = "Tendency towards unkown events fright.";
                }
            }

            return outStr;
        }

        #region Business
        private bool found16 = false;

        public double ChakraMap2BusinessValue(List<string> Values)
        {
            List<int> outBusinessBalues = new List<int>();

            for (int i = 0; i < Values.Count; i++)
            {
                outBusinessBalues.Add(ConvertChakraMapValues2Numbers(Values[i]));
            }

            if (outBusinessBalues[1] == 0) // Throught Chakra
            {
                bool resMasterInList = isHaveMasterInList(Values, Values.Count - 1);
                bool resCarmaticInList = isHaveCarmaticInList(Values, Values.Count - 1);

                if ((resMasterInList == false) && (resCarmaticInList == false))
                {
                    outBusinessBalues[1] = 6;
                }

                if ((resMasterInList == true) && (resCarmaticInList == true))
                {
                    outBusinessBalues[1] = 6;
                }

                if ((resMasterInList == true) && (resCarmaticInList == false))
                {
                    outBusinessBalues[1] = 10;
                }

                if ((resMasterInList == false) && (resCarmaticInList == true))
                {
                    outBusinessBalues[1] = 1;
                }
            }

            double FinalRes = 0;
            for (int i = 0; i < outBusinessBalues.Count; i++)
            {
                FinalRes += outBusinessBalues[i];
            }

            return FinalRes / outBusinessBalues.Count;
        }

        public double LifeCycle2BusinessValue(List<string> Values, bool isSelfFix)
        {
            bool found1122 = false;
            // cycle data will always be last
            found16 = false;
            List<int> outBusinessBalues = new List<int>();

            for (int i = 0; i < Values.Count; i++)
            {
                if (Values[i] == "16") found16 = true;

                if (new string[] { "11", "22" }.Contains(Values[i])) found1122 = true;

                if (i == (Values.Count - 1)) // cycle data will always be last
                {
                    #region Cycle Special Treatment
                    if (GetCorrectNumberFromSplitedString(Values[i].Split(mDelimiter)) == 16)
                    {
                        outBusinessBalues.Add(4);
                    }
                    else
                    {
                        outBusinessBalues.Add(ConvertLifeCycleValues2Numbers(Values[i]));

                    }
                    #endregion Cycle Special Treatment
                }
                else
                {
                    outBusinessBalues.Add(ConvertLifeCycleValues2Numbers(Values[i]));
                }
            }

            bool resMasterInList = isHaveMasterInList(Values, Values.Count);
            bool resCarmaticInList = isHaveCarmaticInList(Values, Values.Count);

            if (outBusinessBalues[0] == 0) // challange is zero
            {
                if ((resMasterInList == false) && (resCarmaticInList == false))
                {
                    outBusinessBalues[0] = 6;
                }

                if ((resMasterInList == true) && (resCarmaticInList == true))
                {
                    outBusinessBalues[0] = 6;
                }

                if ((resMasterInList == true) && (resCarmaticInList == false))
                {
                    outBusinessBalues[0] = 10;
                }

                if ((resMasterInList == false) && (resCarmaticInList == true))
                {
                    outBusinessBalues[0] = 6;
                }
            }

            if ((found16 == true) && (resMasterInList == false) && (resCarmaticInList == true))
            {
                for (int i = 0; i < Values.Count - 1; i++) // ignoring Cycle for testing
                {
                    if ((Values[i] == "16") && (isSelfFix == false))
                    {
                        outBusinessBalues[i] = -10;
                    }
                    if ((Values[i] == "16") && (isSelfFix == true))
                    {
                        outBusinessBalues[i] = 6;
                    }
                }
            }

            double FinalRes = 0;
            for (int i = 0; i < outBusinessBalues.Count; i++)
            {
                FinalRes += outBusinessBalues[i];
            }

            if (found1122)
            {
                FinalRes += 0.5;

            }

            return FinalRes / outBusinessBalues.Count;
        }

        // **********

        public double ChakraMap2SalesValue(List<string> Values)
        {
            List<int> outSalesValues = new List<int>();

            for (int i = 0; i < Values.Count; i++)
            {
                outSalesValues.Add(mSalesChakraChart[Values[i]]);

            }

            if (outSalesValues[1] == 0) // Throught Chakra
            {
                bool resMasterInList = isHaveMasterInList(Values, Values.Count - 1);
                bool resCarmaticInList = isHaveCarmaticInList(Values, Values.Count - 1);

                if ((resMasterInList == false) && (resCarmaticInList == false))
                {
                    outSalesValues[1] = 6;
                }

                if ((resMasterInList == true) && (resCarmaticInList == true))
                {
                    outSalesValues[1] = 6;
                }

                if ((resMasterInList == true) && (resCarmaticInList == false))
                {
                    outSalesValues[1] = 10;
                }

                if ((resMasterInList == false) && (resCarmaticInList == true))
                {
                    outSalesValues[1] = 1;
                }
            }

            double FinalRes = outSalesValues.Sum();

            return FinalRes / outSalesValues.Count;
        }

        // **********

        public double LifeCycle2SalesValue(List<string> Values)
        {
            List<int> outSalesValues = new List<int>();

            for (int i = 0; i < Values.Count; i++)
            {
                if (i == (Values.Count - 1)) // cycle data will always be last
                {
                    #region Cycle Special Treatment

                    if (GetCorrectNumberFromSplitedString(Values[i].Split(mDelimiter)) == 16)
                    {
                        outSalesValues.Add(4);

                    }
                    else
                    {
                        outSalesValues.Add(mSalesChakraChart[Values[i]]);

                    }

                    #endregion Cycle Special Treatment
                }
                else
                {
                    outSalesValues.Add(mSalesChakraChart[Values[i]]);

                }

            }

            bool resMasterInList = isHaveMasterInList(Values, Values.Count);
            bool resCarmaticInList = isHaveCarmaticInList(Values, Values.Count);

            if (outSalesValues[0] == 0) // challange is zero
            {
                if ((resMasterInList == false) && (resCarmaticInList == false))
                {
                    outSalesValues[0] = 6;

                }

                if ((resMasterInList == true) && (resCarmaticInList == true))
                {
                    outSalesValues[0] = 6;

                }

                if ((resMasterInList == true) && (resCarmaticInList == false))
                {
                    outSalesValues[0] = 10;

                }

                if ((resMasterInList == false) && (resCarmaticInList == true))
                {
                    outSalesValues[0] = 6;

                }
            }

            double FinalRes = outSalesValues.Sum();

            return FinalRes / outSalesValues.Count;
        }

        /********/
        //Secretery calculate value
        public double ChakraMap2SecretaryAndAccountingValue(List<string> Values)
        {
            List<int> outSecretarValues = new List<int>();

            for (int i = 0; i < Values.Count; i++)
            {
                outSecretarValues.Add(mSecretaryAndAccountingChakraChart[Values[i]]);

            }

            double FinalRes = outSecretarValues.Sum();

            return FinalRes / outSecretarValues.Count;
        }


        public double LifeCycle2SecretaryAndAccountingValue(List<string> Values)
        {
            List<int> outSecretarValues = new List<int>();

            for (int i = 0; i < Values.Count; i++)
            {
                if (i == (Values.Count - 1)) // cycle data will always be last
                {
                    #region Cycle Special Treatment

                    if (GetCorrectNumberFromSplitedString(Values[i].Split(mDelimiter)) == 16)
                    {
                        outSecretarValues.Add(4);

                    }
                    else
                    {
                        outSecretarValues.Add(mSecretaryAndAccountingChakraChart[Values[i]]);

                    }

                    #endregion Cycle Special Treatment
                }
                else
                {
                    outSecretarValues.Add(mSecretaryAndAccountingChakraChart[Values[i]]);

                }

            }

            double FinalRes = outSecretarValues.Sum();

            return FinalRes / outSecretarValues.Count;
        }

        // **********

        public double ChakraMap2AlternatinvgValue(List<string> Values)
        {
            List<int> outValues = new List<int>();

            for (int i = 0; i < Values.Count; i++)
            {
                outValues.Add(mAlternativingChakraChart[Values[i]]);

            }

            double FinalRes = outValues.Sum();

            return FinalRes / outValues.Count;
        }

        public double LifeCycle2AlternatinvgValue(List<string> Values)
        {
            List<int> outValues = new List<int>();

            for (int i = 0; i < Values.Count; i++)
            {
                if (i == (Values.Count - 1)) // cycle data will always be last
                {
                    #region Cycle Special Treatment

                    if (GetCorrectNumberFromSplitedString(Values[i].Split(mDelimiter)) == 16)
                    {
                        outValues.Add(4);

                    }
                    else
                    {
                        outValues.Add(mAlternativingChakraChart[Values[i]]);

                    }

                    #endregion Cycle Special Treatment
                }
                else
                {
                    outValues.Add(mAlternativingChakraChart[Values[i]]);

                }

            }

            double FinalRes = outValues.Sum();

            return FinalRes / outValues.Count;
        }

        public double ChakraMap2BeautyValue(List<string> Values)
        {
            List<int> outValues = new List<int>();

            for (int i = 0; i < Values.Count; i++)
            {
                outValues.Add(mBeautyChakraChart[Values[i]]);

            }

            double FinalRes = outValues.Sum();

            return FinalRes / outValues.Count;
        }

        public double LifeCycle2BeautyValue(List<string> Values)
        {
            List<int> outValues = new List<int>();

            for (int i = 0; i < Values.Count; i++)
            {
                if (i == (Values.Count - 1)) // cycle data will always be last
                {
                    #region Cycle Special Treatment

                    if (GetCorrectNumberFromSplitedString(Values[i].Split(mDelimiter)) == 16)
                    {
                        outValues.Add(4);

                    }
                    else
                    {
                        outValues.Add(mBeautyChakraChart[Values[i]]);

                    }

                    #endregion Cycle Special Treatment
                }
                else
                {
                    outValues.Add(mBeautyChakraChart[Values[i]]);

                }

            }

            double FinalRes = outValues.Sum();

            return FinalRes / outValues.Count;
        }

        public double ChakraMap2ManualValue(List<string> Values)
        {
            List<int> outValues = new List<int>();

            for (int i = 0; i < Values.Count; i++)
            {
                outValues.Add(mManualWorkChakraChart[Values[i]]);

            }

            double FinalRes = outValues.Sum();

            return FinalRes / outValues.Count;
        }

        public double LifeCycle2ManualValue(List<string> Values)
        {
            List<int> outValues = new List<int>();

            for (int i = 0; i < Values.Count; i++)
            {
                if (i == (Values.Count - 1)) // cycle data will always be last
                {
                    #region Cycle Special Treatment

                    if (GetCorrectNumberFromSplitedString(Values[i].Split(mDelimiter)) == 16)
                    {
                        outValues.Add(4);

                    }
                    else
                    {
                        outValues.Add(mManualWorkChakraChart[Values[i]]);

                    }

                    #endregion Cycle Special Treatment
                }
                else
                {
                    outValues.Add(mManualWorkChakraChart[Values[i]]);

                }

            }

            double FinalRes = outValues.Sum();

            return FinalRes / outValues.Count;
        }

         public double ChakraMap2HiTechValue(List<string> Values)
        {
            List<double> outValues = new List<double>();

            for (int i = 0; i < Values.Count; i++)
            {
                outValues.Add(mHiTechChakraChart[Values[i]]);

            }

            double FinalRes = outValues.Sum();

            return FinalRes / outValues.Count;
        }

        public double LifeCycle2HiTechValue(List<string> Values)
        {
            List<double> outValues = new List<double>();

            for (int i = 0; i < Values.Count; i++)
            {
                if (i == (Values.Count - 1)) // cycle data will always be last
                {
                    #region Cycle Special Treatment

                    if (GetCorrectNumberFromSplitedString(Values[i].Split(mDelimiter)) == 16)
                    {
                        outValues.Add(4);

                    }
                    else
                    {
                        outValues.Add(mHiTechChakraChart[Values[i]]);

                    }

                    #endregion Cycle Special Treatment
                }
                else
                {
                    outValues.Add(mHiTechChakraChart[Values[i]]);

                }

            }

            double FinalRes = outValues.Sum();

            return FinalRes / outValues.Count;
        }

        public double ChakraMap2LowTechValue(List<string> Values)
        {
            List<double> outValues = new List<double>();

            for (int i = 0; i < Values.Count; i++)
            {
                outValues.Add(mLowTechChakraChart[Values[i]]);

            }

            double FinalRes = outValues.Sum();

            return FinalRes / outValues.Count;
        }

        public double LifeCycle2LowTechValue(List<string> Values)
        {
            List<double> outValues = new List<double>();

            for (int i = 0; i < Values.Count; i++)
            {
                if (i == (Values.Count - 1)) // cycle data will always be last
                {
                    #region Cycle Special Treatment

                    if (GetCorrectNumberFromSplitedString(Values[i].Split(mDelimiter)) == 16)
                    {
                        outValues.Add(4);

                    }
                    else
                    {
                        outValues.Add(mLowTechChakraChart[Values[i]]);

                    }

                    #endregion Cycle Special Treatment
                }
                else
                {
                    outValues.Add(mLowTechChakraChart[Values[i]]);

                }

            }

            double FinalRes = outValues.Sum();

            return FinalRes / outValues.Count;
        }

        private int ConvertChakraMapValues2Numbers(string inValue)
        {
            for (int i = 0; i < mBusinessChakraChart.Length; i++)
            {
                if (inValue == mBusinessChakraChart[i, 0])
                {
                    return Convert.ToInt16(mBusinessChakraChart[i, 1]);
                }
            }

            return -1;
        }

        private int ConvertLifeCycleValues2Numbers(string inValue)
        {
            for (int i = 0; i < mBusinessLifeCycleChart.Length; i++)
            {
                if (inValue == mBusinessLifeCycleChart[i, 0])
                {
                    return Convert.ToInt16(mBusinessLifeCycleChart[i, 1]);

                }

            }

            return -1;
        }

        private bool isHaveMasterInList(List<string> inlist, int count)
        {
            for (int i = 0; i < count; i++)
            {
                if (isMaterNumber(GetCorrectNumberFromSplitedString(inlist[i].Split(mDelimiter))) == true)
                {
                    return true;
                }
            }

            return false;
        }

        private bool isHaveCarmaticInList(List<string> inlist, int count)
        {
            for (int i = 0; i < count; i++)
            {
                if (isCarmaticNumber(GetCorrectNumberFromSplitedString(inlist[i].Split(mDelimiter))) == true)
                {
                    return true;

                    //if (GetCorrectNumberFromSplitedString(inlist[i].Split(mDelimiter)) == 16)
                    //{
                    //    found16 = true;
                    //}
                }
            }

            return false;
        }
        #endregion Business

        #region date 2 heb
        public string GeorgianDate2HebrewJewishDateString(DateTime anyDate)
        {

            System.Text.StringBuilder hebrewFormatedString = new System.Text.StringBuilder();
            // Create the hebrew culture to use hebrew (Jewish) calendar 
            CultureInfo jewishCulture = CultureInfo.CreateSpecificCulture("he-IL");
            jewishCulture.DateTimeFormat.Calendar = new HebrewCalendar();

            // Day of the week in the format " " 
            hebrewFormatedString.Append(anyDate.ToString("dddd", jewishCulture) + " ");
            // Day of the month in the format "'" 
            hebrewFormatedString.Append(anyDate.ToString("dd", jewishCulture) + " ");
            // Month and year in the format " " 
            hebrewFormatedString.Append("" + anyDate.ToString("y", jewishCulture));

            return hebrewFormatedString.ToString();

        }

        public DateTime HebrewJewishDateString2GeorgianDate(string strHebrew)
        {
            CultureInfo cultureJewish = CultureInfo.CreateSpecificCulture("he-IL");

            cultureJewish.DateTimeFormat.Calendar = new HebrewCalendar();

            // create a hebrew date
            //DateTime dtAnyDateHebrew = new DateTime(5767, 1, 1, new System.Globalization.HebrewCalendar());

            // BTW: You can compare DateTimes like integers
            //DateTime dtHebrewBegin = new DateTime(5767, 1, 1, new System.Globalization.HebrewCalendar());
            //DateTime dtHebrewMiddle = new DateTime(5768, 1, 1, new System.Globalization.HebrewCalendar());
            //DateTime dtHebrewEnd = new DateTime(5769, 1, 1, new System.Globalization.HebrewCalendar());
            // ...

            DateTime dtResult = DateTime.Parse(strHebrew, cultureJewish);
            return dtResult;
        }
        #endregion

        public bool isNumInArray(int[] arr, int num)
        {
            for (int i = 0; i < arr.Length; i++)
            {
                if (num == arr[i])
                {
                    return true;
                }
            }

            return false;
        }

        public bool isNumInArray(string[] arr, int num)
        {
            for (int i = 0; i < arr.Length; i++)
            {
                if (num == Convert.ToInt16(arr[i]))
                {
                    return true;
                }
            }

            return false;
        }

        public bool isNumInList(int num, List<int> list)
        {
            bool res = false;
            foreach (int tmpNum in list)
            {
                if (num == tmpNum)
                {
                    res = true;
                }
            }
            return res;
        }

        public int MinArr(int[] nums)
        {

            if (nums.Length > 1)
            {
                int res = nums[0];
                for (int i = 1; i < nums.Length; i++)
                {
                    if (res > nums[i])
                    {
                        res = nums[i];
                    }
                }
                return res;
            }
            else
            {
                return nums[0];
            }
        }
        #endregion //Public Methods

        #region Private Methods

        /// <summary> sums the digits in a given number
        /// 
        /// </summary>
        /// <param name="num">integer number</param>
        /// <returns>string to set to a textbox</returns>
        private string Nums_2_String(int num)
        {
            int InNum = num;
            int res, sum = 0;
            string strOut = "";
            //strOut = num.ToString();
            if (is_Special(num) == false)
            {
                while (num != 0)
                {
                    Math.DivRem(num, 10, out res);

                    sum += res;

                    num = (int)(num / 10);
                }
            }
            else // true
            {
                sum = num;
            }

            strOut += sum.ToString();
            //strOut = strOut + mDelimiter + sum.ToString();

            if (sum > 9)
            {
                num = sum;
                sum = 0;
                while (num != 0)
                {
                    Math.DivRem(num, 10, out res);

                    sum += res;

                    num = (int)(num / 10);
                }
                strOut = strOut + mDelimiter + sum.ToString();

                if (sum > 9)
                {
                    num = sum;
                    sum = 0;
                    while (num != 0)
                    {
                        Math.DivRem(num, 10, out res);

                        sum += res;

                        num = (int)(num / 10);
                    }
                    strOut = strOut + mDelimiter + sum.ToString();
                }
            }


            string[] Splitted = strOut.Split(mDelimiter);

            if (is_Special(Convert.ToInt16(Splitted[0])) == true)
            {
                return strOut;
            }

            if (is_Special(InNum) == true)
            {
                return Splitted[Splitted.Length - 2] + mDelimiter + Splitted[Splitted.Length - 1];
            }
            else
            {
                return Splitted[Splitted.Length - 1];
            }

        }

        /// <summary> tests if a given number is a special number - master or carmatic
        /// 
        /// </summary>
        /// <param name="num">input number</param>
        /// <returns>boolean - True / False</returns>
        private bool is_Special(int num)
        {
            bool ans = false;

            // Master numbers
            if (isMaterNumber(num))//((num == 11) || (num == 22)|| (num == 33))
            {
                ans = true;
            }
            // Carmatic Numbers
            if (isCarmaticNumber(num))//((num == 13) || (num == 14) || (num == 16) || (num == 19))
            {
                ans = true;
            }

            return ans;
        }

        /// <summary> Final Letters Recognition
        /// 
        /// </summary>
        /// <param name="c">in char</param>
        /// <param name="OutC">out char</param>
        /// <returns>boolean value</returns>
        private bool isFinalLetter(char c, out char OutC)
        {
            bool test1 = (c == "ך".ToCharArray()[0]);
            bool test2 = (c == "ם".ToCharArray()[0]);
            bool test3 = (c == "ן".ToCharArray()[0]);
            bool test4 = (c == "ץ".ToCharArray()[0]);
            bool test5 = (c == "ף".ToCharArray()[0]);

            bool res;
            OutC = Convert.ToChar("9");

            if (test1)
            {
                OutC = "כ".ToCharArray()[0];
            }
            if (test2)
            {
                OutC = "מ".ToCharArray()[0];
            }
            if (test3)
            {
                OutC = "נ".ToCharArray()[0];
            }
            if (test4)
            {
                OutC = "צ".ToCharArray()[0];
            }
            if (test5)
            {
                OutC = "פ".ToCharArray()[0];
            }

            if (test1 || test2 || test3 || test4 || test5)
            {
                res = true;
            }
            else
            {
                res = false;
            }

            return res;
        }

        /// <summary> Converts From Letters to numbers
        /// GIMATRIA
        /// </summary>
        /// <param name="c">in char</param>
        /// <returns>out number</returns>
        private int char2int(char c)
        {
            int index = 0;
            if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
            {
                index = (int)c - (int)"א".ToCharArray()[0];
            }

            if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
            {
                index = (int)c - (int)"a".ToCharArray()[0];
            }

            // Eliminate from invalid letter for current program langauge
            if (index < 0)
            {
                index = 0;

            }
            return Letters[index];

        }

        /// <summary> Calculates a name value
        /// 
        /// </summary>
        /// <param name="name">string of the name</param>
        /// <returns>string of the num</returns>
        private string CalcName(string name)
        {
            char[] cName = name.ToCharArray();
            int[] iName = new int[cName.Length];
            for (int i = 0; i < cName.Length; i++)
            {
                iName[i] = char2int(cName[i]);
            }

            int num = 0;
            for (int i = 0; i < iName.Length; i++)
            {
                num += iName[i];
            }

            return Nums_2_String(num);
        }

        /// <summary> tests rather a char ia a voule or not
        /// 
        /// </summary>
        /// <param name="c">in char</param>
        /// <returns>true OR false</returns>
        private bool isVoule(char c)
        {
            bool res = false;

            int currindex;

            if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.Hebrew)
            {
                currindex = (int)c - (int)"א".ToCharArray()[0];
                res = ((currindex == 0) || (currindex == 4) || (currindex == 5) || (currindex == 9));
            }

            if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
            {
                currindex = (int)c - (int)"a".ToCharArray()[0];
                res = ((currindex == 0) || (currindex == 4) || (currindex == 8) || (currindex == 14) || (currindex == 20) || (currindex == 24));
            }

            return res;
        }

        /// <summary> ChakraCalc
        /// 
        /// </summary>
        /// <param name="Name">Name of the Person</param>
        /// <param name="Letters_of_Chakra">Chakra Special Letters</param>
        /// <returns>out String</returns>
        private string CalcChakra(string Name, string ChakraOutput, string[] Letters_of_Chakra, int[] Nums_of_Chakra)
        {
            string outputString = "";

            //int sum = 0;
            for (int i = 0; i < Name.Length; i++)
            {
                for (int index = 0; index < Letters_of_Chakra.Length; index++)
                {
                    string l = Name.Substring(i, 1);
                    if (l == Letters_of_Chakra[index])
                    {
                        if (outputString == "")
                        {
                            outputString = Letters_of_Chakra[index];
                        }
                        else
                        {
                            if (isCharInString(Letters_of_Chakra[index].ToCharArray()[0], outputString) == false)
                            {
                                outputString = outputString + Letters_of_Chakra[index];
                            }
                        }

                    }
                }
            }

            if (outputString != "")
            {
                string tmp = outputString;
                outputString = "";
                for (int i = 0; i < tmp.Length; i++)
                {
                    if (outputString == "")
                    {
                        outputString = tmp.Substring(0, 1);
                    }
                    else
                    {
                        outputString = outputString + "," + tmp.Substring(i, 1);
                    }
                }
            }

            if (outputString == "")
            {
                string[] splt = ChakraOutput.Split(mDelimiter);

                for (int i = 0; i < splt.Length; i++)
                {
                    for (int j = 0; j < Nums_of_Chakra.Length; j++)
                    {
                        if (Convert.ToInt16(splt[i]) == Nums_of_Chakra[j])
                        {
                            outputString = Nums_of_Chakra[j].ToString();
                            return outputString;
                        }
                    }
                }
            }

            return outputString;
        }

        private bool isCharInString(char c, string s)
        {
            bool res = false;

            for (int i = 0; i < s.Length; i++)
            {
                if (s.Substring(i, 1) == c.ToString())
                    res = true;
            }

            return res;
        }

        private string TestCompatebilityCityApp(int iSunMik, int iNum)
        {
            string outHarmonic = "מאוזן";
            string outDissHarmonic = "לא מאוזן";
            string outHalfHarmonic = "חצי מאוזן";

            if (iSunMik == iNum)
                return outHarmonic;

            if (iSunMik == 16 && iNum == 7)
                return outHarmonic;
            if (iSunMik == 19 && iNum == 1)
                return outHarmonic;
            if (iSunMik == 13 && iNum == 4)
                return outHarmonic;
            if (iSunMik == 14 && iNum == 5)
                return outHarmonic;

            int num1, num2;
            num1 = iSunMik;
            num2 = iNum;
            int[,] tmpArr = mStreetCombHramonic;
            for (int i = 0; i < tmpArr.Length / 2; i++)
            {
                if (((num1 == tmpArr[i, 0]) && (num2 == tmpArr[i, 1])) || ((num2 == tmpArr[i, 0]) && (num1 == tmpArr[i, 1])))
                {
                    return outHarmonic;
                }
            }

            tmpArr = mStreetCombHalfHramonic;
            for (int i = 0; i < tmpArr.Length / 2; i++)
            {
                if (((num1 == tmpArr[i, 0]) && (num2 == tmpArr[i, 1])) || ((num2 == tmpArr[i, 0]) && (num1 == tmpArr[i, 1])))
                {
                    return outHalfHarmonic;
                }
            }

            tmpArr = mStreetCombDissHramonic;
            for (int i = 0; i < tmpArr.Length / 2; i++)
            {
                if (((num1 == tmpArr[i, 0]) && (num2 == tmpArr[i, 1])) || ((num2 == tmpArr[i, 0]) && (num1 == tmpArr[i, 1])))
                {
                    return outDissHarmonic;
                }
            }

            return outHalfHarmonic;
        }



        private List<int> GetValuesFromTable_Simple(List<string> personValues, string[,] arrValues)
        {
            List<int> outVals = new List<int>();

            foreach (string str in personValues)
            {
                for (int i = 0; i < arrValues.GetLength(0); i++)
                {
                    if (arrValues[i, 0] == GetCorrectNumberFromSplitedString(str.Split(mDelimiter)).ToString())
                    {
                        outVals.Add(Convert.ToInt16(arrValues[i, 1]));
                    }
                }
            }

            return outVals;
        }


        #endregion //Private Methods

        #region Proofiong
        public string Proof_CreateFirstKey()
        {
            return mProffer.CreateFirstKey();
        }

        public bool Proof_TestKeys(string FirstKey, string SecondKey)
        {
            return mProffer.TestPassword(FirstKey, SecondKey);
        }

        public bool Proof_Test_First_Key_With_PC(string sFirstKey)
        {
            return mProffer.IsFirstKeyMatchesPC(sFirstKey);
        }
        #endregion

        #region Marriage
        public double MarriageCompetabilityCalc(bool LongTerm, List<string> partner1, bool isBounus1, List<string> partner2, bool isBounus2)
        {
            List<int> values = new List<int>();
            double resAns = -1;

            for (int i = 0; i < (partner1.Count - 4); i++)
            {
                int num1 = Convert.ToInt16(partner1[i]);
                int num2 = Convert.ToInt16(partner2[i]);

                values.Add(GetMarriageValue(num1, num2, LongTerm));
            }

            //if (Convert.ToInt16(partner1[partner1.Count - 4]) < Convert.ToInt16(partner2[partner1.Count - 4]))
            //{
            //    values.Add(Convert.ToInt16(partner1[partner1.Count - 4]));
            //}
            //else
            //{
            //    values.Add(Convert.ToInt16(partner2[partner1.Count - 4]));
            //}

            #region MEUZAN
            int tmp1, tmp2;
            tmp1 = -1;
            tmp2 = -1;
            if (partner1[partner1.Count - 1] == "מאוזן")
            {
                tmp1 = 10;
            }
            if (partner1[partner1.Count - 1] == "חצי מאוזן")
            {
                tmp1 = 8;
            }
            if (partner1[partner1.Count - 1] == "לא מאוזן")
            {
                tmp1 = 4;
            }

            if (partner2[partner1.Count - 1] == "מאוזן")
            {
                tmp2 = 10;
            }
            if (partner2[partner1.Count - 1] == "חצי מאוזן")
            {
                tmp2 = 8;
            }
            if (partner2[partner1.Count - 1] == "לא מאוזן")
            {
                tmp2 = 4;
            }

            if (tmp1 < tmp2)
            {
                values.Add(tmp1);
            }
            else
            {
                values.Add(tmp2);
            }

            #endregion

            resAns = 0;
            for (int i = 0; i < values.Count; i++)
            {
                resAns += values[i];
            }
            resAns = resAns / values.Count;


            //חישוב של שיאים
            List<int> valuesSe = new List<int>();
            int num3 = Convert.ToInt16(partner1[6]);
            int num4 = Convert.ToInt16(partner2[6]);

            valuesSe.Add(GetMarriageValue(num3, num4, LongTerm));

            num3 = Convert.ToInt16(partner1[7]);
            num4 = Convert.ToInt16(partner2[7]);

            valuesSe.Add(GetMarriageValue(num3, num4, LongTerm));

            double resSeAns = 0;
            for (int i = 0; i < valuesSe.Count; i++)
            {
                resSeAns += valuesSe[i];
            }
            resSeAns = resSeAns / valuesSe.Count;

            resAns = (resAns + resSeAns) / 2;

            #region Bonus

            if (isBounus1 == true)
            {
                resAns += 0.5;
            }
            if (isBounus2 == true)
            {
                resAns += 0.5;
            }

            #endregion Bonus

            return resAns;
        }

        // **********

        public double SexualMatchCalc(List<string> partner1, List<string> partner2)
        {
            List<int> values = new List<int>();
            double resAns = -1;

            for (int i = 0; i < (partner1.Count - 2); i++)
            {
                int num1 = Convert.ToInt16(partner1[i]);
                int num2 = Convert.ToInt16(partner2[i]);

                values.Add(GetSexualValue(num1, num2));
            }

            if (Convert.ToInt16(partner1[partner1.Count - 2]) < Convert.ToInt16(partner2[partner1.Count - 2]))
            {
                values.Add(Convert.ToInt16(partner1[partner1.Count - 2]));
            }
            else
            {
                values.Add(Convert.ToInt16(partner2[partner1.Count - 2]));
            }

            resAns = 0;
            for (int i = 0; i < values.Count; i++)
            {
                resAns += values[i];

            }
            resAns = resAns / values.Count;

            return resAns;
        }

        // **********

        private int GetMarriageValue(int num1, int num2, bool isLongTerm)
        {
            bool found = false;
            int resAns = -1;

            string[,] Values;

            if (isLongTerm == true)
            {
                Values = mRelationCompetability_TermLong;
            }
            else
            {
                Values = mRelationCompetability_TermShort;
            }

            for (int i = 0; i < Values.GetLength(0); i++)
            {
                if (((Values[i, 0].Trim() == num1.ToString()) && (Values[i, 1].Trim() == num2.ToString())) || ((Values[i, 0].Trim() == num2.ToString()) && (Values[i, 1].Trim() == num1.ToString())))
                {
                    found = true;
                    resAns = Convert.ToInt16(Values[i, 2].Trim());
                }
            }

            if (found == false)
            {
                if (isLongTerm == true)
                {
                    if (((num1 == 0) && (isMaterNumber(num2))) || ((num2 == 0) && (isMaterNumber(num1))))
                    {
                        resAns = 10;
                    }
                    if (((num1 == 0) && (isCarmaticNumber(num2))) || ((num2 == 0) && (isCarmaticNumber(num1))))
                    {
                        resAns = 6;
                    }

                    if (isNumbersMatch(num1, num2, new int[] { 0 }, new int[] { 1, 33 }) == true)
                    {
                        resAns = 3;
                    }

                    if (isNumbersMatch(num1, num2, new int[] { 0 }, new int[] { 0, 2, 4, 6, 8 }) == true)
                    {
                        resAns = 10;
                    }

                    if (isNumbersMatch(num1, num2, new int[] { 0 }, new int[] { 3, 5, 7 }) == true)
                    {
                        resAns = 4;
                    }

                    if (isNumbersMatch(num1, num2, new int[] { 13, 14, 15, 16, 19 }, new int[] { 14, 15, 19 }) == true)
                    {
                        resAns = 3;
                    }

                    if (((num1 == 13) && (num2 == 13)) || ((num1 == 16) && (num2 == 16)))
                    {
                        resAns = 7;
                    }
                }
                else
                {
                    if ((num1 == 0) || (num2 == 0))
                    {
                        resAns = 4;
                    }

                    if (isNumbersMatch(num1, num2, new int[] { 13, 14, 15, 16, 19 }, new int[] { 13, 15, 16 }) == true)
                    {
                        resAns = 4;
                    }

                    if (((num1 == 14) && (num2 == 14)) || ((num1 == 19) && (num2 == 19)))
                    {
                        resAns = 7;
                    }


                }
            }

            return resAns;
        }

        // **********

        private int GetSexualValue(int num1, int num2)
        {
            bool found = false;
            int resAns = -1;

            Tuple<int, int, int>[] Values;

            Values = mSexualCompatability;

            var res = Values.Where(i => (i.Item1 == num1 && i.Item2 == num2) || (i.Item1 == num2 && i.Item2 == num1)).First();

            if (res != null)
            {
                resAns = res.Item3;
                found = true;

            }

            if (found == false)
            {
                Tuple<int, int, int>[] carmatics = new Tuple<int, int, int>[]
                {
                    Tuple.Create(13, 13, 5), Tuple.Create(13, 14, 6), Tuple.Create(13, 16, 4), Tuple.Create(13, 19, 6),
                    Tuple.Create(14, 13, 6), Tuple.Create(14, 14, 8), Tuple.Create(14, 16, 3), Tuple.Create(14, 19, 6),
                    Tuple.Create(16, 13, 4), Tuple.Create(16, 14, 3), Tuple.Create(16, 16, 5), Tuple.Create(16, 19, 4),
                    Tuple.Create(19, 13, 6), Tuple.Create(19, 14, 6), Tuple.Create(19, 16, 4), Tuple.Create(19, 19, 7)

                };

                var carm = carmatics.Where(i => (i.Item1 == num1 && i.Item2 == num2) || (i.Item1 == num2 && i.Item2 == num1)).First();

                if (carm != null)
                {
                    resAns = carm.Item3;

                }

            }

            return resAns;
        }

        // **********

        private bool isNumbersMatch(int num1, int num2, int[] nums1, int[] nums2)
        {
            bool res = false;
            for (int i = 0; i < nums1.Length; i++)
            {
                for (int j = 0; j < nums2.Length; j++)
                {
                    if (((nums1[i] == num1) && (nums2[j] == num2)) && ((nums1[i] == num2) && (nums2[j] == num1)))
                    {
                        return true;
                    }
                }
            }
            return res;
        }

        #endregion Marriage

        #region LearnSccss_AttPrblm
        public double CalcLearnSusccess(List<string> personalInfo, string BonusTime, bool is14Test)
        {
            int throughtValue = GetCorrectNumberFromSplitedString(personalInfo[1].Split(mDelimiter));

            List<string> LC_Values = personalInfo.GetRange(6, 2);

            List<string> Chkr_Values = personalInfo.GetRange(0, 6);
            //Chkr_Values.Add(personalInfo[personalInfo.Count - 1]);  // leveled status

            List<int> iValues = GetValuesFromTable_Simple(Chkr_Values, mLearnSccssAttPrbl_Chkra);
            iValues.AddRange(GetValuesFromTable_Simple(LC_Values, mLearnSccssAttPrbl_LC));

            double sum1 = 0, sum2 = 0;
            for (int i = 0; i < Chkr_Values.Count; i++)
            {
                sum1 += iValues[i];
            }
            sum1 = sum1 / Chkr_Values.Count;

            for (int i = 0; i < LC_Values.Count; i++)
            {
                sum2 += iValues[i];
            }
            sum2 = sum2 / LC_Values.Count;

            double sum = (sum1 + sum2) / 2;

            //double sum = 0;
            //foreach (int v in iValues)
            //{
            //    sum += (double)v;
            //}
            //sum /= iValues.Count;

            int cChkraMst = 0, cChrkaCrm = 0, cLCMst = 0, cLCCrm = 0;
            bool isLC16Found = false;
            foreach (string str in Chkr_Values)
            {
                int num = GetCorrectNumberFromSplitedString(str.Split(mDelimiter));
                if (isMaterNumber(num) == true)
                {
                    cChkraMst++;
                }
                if (isCarmaticNumber(num) == true)
                {
                    cChrkaCrm++;
                }
            }

            foreach (string str in LC_Values)
            {
                int num = GetCorrectNumberFromSplitedString(str.Split(mDelimiter));
                if (isMaterNumber(num) == true)
                {
                    cLCMst++;
                }
                if (isCarmaticNumber(num) == true)
                {
                    cLCCrm++;
                    if (num == 16)
                    {
                        isLC16Found = true;
                    }
                }
            }

            if (is14Test == true)
            {
                sum -= 1.0;
            }

            if ((cLCCrm + cChrkaCrm) == 2)
            {
                sum -= 1.0;
            }
            if ((cLCCrm + cChrkaCrm) == 3)
            {
                sum -= 1.5;
            }
            if ((cLCCrm + cChrkaCrm) == 4)
            {
                sum -= 2.0;
            }
            if ((cLCCrm + cChrkaCrm) == 5)
            {
                sum -= 2.5;
            }
            if ((cLCCrm + cChrkaCrm) >= 6)
            {
                sum -= 3.0;
            }

            if (isLC16Found == true)
            {
                sum += 1.0;
            }

            if (throughtValue == 8)
            {
                sum += 1.0;
            }

            string[] splt = BonusTime.Split("-".ToCharArray()[0]);
            if (splt.Length == 1)
            {
                if (splt[0].Trim() == "עד 1.5")
                {
                    sum += 0.0;
                }
                if (splt[0].Trim() == "3 ויותר")
                {
                    sum += 2.0;
                }
            }
            if (splt.Length == 2)
            {
                if ((splt[0].Trim() == "1.5") && (isLC16Found == false))
                {
                    sum += 1.0;
                }
                if (splt[0].Trim() == "2")
                {
                    sum += 1.5;
                }
            }


            if ((cLCMst + cChkraMst) > 1)
            {
                sum += 0.5;
            }

            int n1, n2, n3, n4, n5;
            string[] s1, s2, s3, s4, s5;
            bool b1, b2, b3, b4;
            #region Fine Tuning
            s1 = personalInfo[0].Split(mDelimiter);
            n1 = GetCorrectNumberFromSplitedString(s1); // כתר

            s2 = personalInfo[2].Split(mDelimiter);
            n2 = GetCorrectNumberFromSplitedString(s2);//מקלעת שמש

            s3 = personalInfo[4].Split(mDelimiter);
            n3 = GetCorrectNumberFromSplitedString(s3);//שם פרטי

            s4 = personalInfo[5].Split(mDelimiter);
            n4 = GetCorrectNumberFromSplitedString(s4);//מזל אסטרולוגי

            s5 = personalInfo[1].Split(mDelimiter);
            n5 = GetCorrectNumberFromSplitedString(s5); // גרון

            b1 = isCarmaticNumber(GetCorrectNumberFromSplitedString(personalInfo[0].Split(mDelimiter)));//KETER
            b2 = isCarmaticNumber(GetCorrectNumberFromSplitedString(personalInfo[1].Split(mDelimiter)));//גרון
            b3 = isCarmaticNumber(GetCorrectNumberFromSplitedString(personalInfo[2].Split(mDelimiter)));//מקלעת
            b4 = isCarmaticNumber(GetCorrectNumberFromSplitedString(personalInfo[4].Split(mDelimiter)));//שם פרטי

            int Num2Test = 30;
            bool is30Found = false;
            if (isNumInArray(s1, Num2Test) || isNumInArray(s2, Num2Test) || isNumInArray(s3, Num2Test))
            {
                sum += 1.0;
                is30Found = true;
            }

            if ((b1 || b2 || b3 || b4) == true)
            {
                sum -= 1.0;
            }


            if ((n1 == 7) || (n4 == 7) || (n5 == 7) || (n2 == 7) || (n3 == 7))
            {
                sum += 1.0;
            }

            Num2Test = 10;
            bool is10Found = false;
            if (isNumInArray(s1, Num2Test) || isNumInArray(s5, Num2Test) || isNumInArray(s2, Num2Test) || isNumInArray(s3, Num2Test))
            {
                sum += 0.5;
                is10Found = true;
            }

            if (is10Found && is30Found)
            {
                sum += 0.5;
            }


            #endregion
            if (sum > 10.0)
            {
                sum = 10.0;
            }
            return sum;
        }

        public double CalcAttentionProblems(List<string> personalInfo, out int[] MajorCarmaticValues, out int[] MinorCarmaticValues)
        {
            double sum = 0;
            List<int> MjrCrmValues = new List<int>();
            List<int> MnrCrmValues = new List<int>();


            string IZUN = personalInfo[personalInfo.Count - 1];
            //personalInfo.Remove(IZUN);

            for (int i = 0; i < personalInfo.Count - 1; i++)
            {
                string str = personalInfo[i];
                int num = GetCorrectNumberFromSplitedString(str.Split(mDelimiter));

                switch (num)
                {
                    case 1:
                    case 3:
                    case 5:
                        if (isNumInArray(str.Split(mDelimiter), num * 10) == false)
                        {
                            sum += 1.0;
                        }
                        break;
                    case 6:
                        if (isNumInArray(str.Split(mDelimiter), 15) == true)
                        {
                            sum += 0.5;
                        }
                        break;
                    case 13:
                    case 14:
                    case 16:
                    case 19:
                        sum += 1.0;
                        break;
                }

                switch (num)
                {
                    case 1:
                    case 3:
                    case 5:
                        if (isNumInArray(str.Split(mDelimiter), num * 10) == false)
                        {
                            MnrCrmValues.Add(num);
                        }
                        break;
                    case 6:
                        if (isNumInArray(str.Split(mDelimiter), 15) == true)
                        {
                            MnrCrmValues.Add(15);
                        }
                        break;
                    case 13:
                    case 14:
                    case 16:
                    case 19:
                        MjrCrmValues.Add(num);
                        break;
                }
            }

            int num1 = GetCorrectNumberFromSplitedString(personalInfo[3].Split(mDelimiter)); // chakra values
            int num2 = GetCorrectNumberFromSplitedString(personalInfo[4].Split(mDelimiter)); // chakra values
            if ((IZUN == "לא מאוזן") && ((isCarmaticNumber(num1) == false) && (isCarmaticNumber(num2) == false)))
            {
                sum += 1.0;
            }

            if (MjrCrmValues.Count > 0)
            {
                MajorCarmaticValues = new int[MjrCrmValues.Count];
                for (int i = 0; i < MajorCarmaticValues.Length; i++)
                {
                    MajorCarmaticValues[i] = MjrCrmValues[i];
                }
            }
            else
            {
                MajorCarmaticValues = new int[1] { -1 };
            }

            if (MnrCrmValues.Count > 0)
            {
                MinorCarmaticValues = new int[MnrCrmValues.Count];
                for (int i = 0; i < MinorCarmaticValues.Length; i++)
                {
                    MinorCarmaticValues[i] = MnrCrmValues[i];
                }
            }
            else
            {
                MinorCarmaticValues = new int[1] { -1 };
            }


            return sum;
        }
        #endregion

        #region Health
        public double CalcHealth(List<int> personalInfo, int cycle, int crown, int thirdeye)
        {
            double resFinal = 0;
            int tmpnum = 0; int count = 0;
            int CarmaticCounter = 0;

            for (int i = 0; i < personalInfo.Count - 1; i++)
            {
                if (i != 3)
                {
                    for (int j = 0; j < (mHealthTable.GetLength(0)); j++)
                    {
                        if (mHealthTable[j, 0] == personalInfo[i])
                        {
                            tmpnum = mHealthTable[j, 1];
                        }
                    }
                }
                else // מזל אסטרולוגי
                {
                    tmpnum = personalInfo[i];
                }

                if (isCarmaticNumber(personalInfo[i]))
                {
                    CarmaticCounter++;
                }
                resFinal += tmpnum;
                count++;
            }

            if (isCarmaticNumber(crown))
            {
                CarmaticCounter++;
            }
            if (isCarmaticNumber(thirdeye))
            {
                CarmaticCounter++;
            }
            if (isCarmaticNumber(personalInfo[personalInfo.Count - 1]))
            {
                CarmaticCounter++;
            }
            // מין ויצירה
            for (int j = 0; j < (mHealthTable.GetLength(0)); j++)
            {
                if (mHealthTable[j, 0] == personalInfo[(personalInfo.Count - 1)])
                {
                    tmpnum = mHealthTable[j, 3];
                }
            }
            resFinal += tmpnum;
            count++;

            if ((cycle < 4) || isCarmaticNumber(crown))
            {
                for (int j = 0; j < (mHealthTable.GetLength(0)); j++)
                {
                    if (mHealthTable[j, 0] == crown)
                    {
                        tmpnum = mHealthTable[j, 1];
                    }
                }
                resFinal += tmpnum;
                count++;
            }
            if (cycle == 4)
            {
                for (int j = 0; j < (mHealthTable.GetLength(0)); j++)
                {
                    if (mHealthTable[j, 0] == thirdeye)
                    {
                        tmpnum = mHealthTable[j, 2];
                    }
                }
                resFinal += tmpnum;
                count++;
            }

            resFinal = resFinal / count;
            //צאקרה 1, 2, 5
            if ((personalInfo[4] == 13) || (personalInfo[1] == 13) || (personalInfo[0] == 13))
            {
                resFinal = resFinal - 0.4;
            }
            if ((personalInfo[4] == 14) || (personalInfo[1] == 14) || (personalInfo[0] == 14))
            {
                resFinal = resFinal - 0.3;
            }

            if ((personalInfo[4] == 16) && (cycle < 3))
            {
                resFinal = resFinal - 0.3;
            }
            if ((personalInfo[4] == 16) && (cycle > 2))
            {
                resFinal = resFinal - 2;
            }
            if ((personalInfo[1] == 16) && (cycle > 2))
            {
                resFinal = resFinal - 1;
            }
            if (personalInfo[0] == 16)
            {
                resFinal = resFinal - 1;
            }
            if ((personalInfo[6] == 16) && (cycle == 4))
            {
                resFinal = resFinal - 2;
            }

            if ((personalInfo[4] == 16) && (personalInfo[6] == 16))
            {
                resFinal = resFinal - 1;
            }

            if ((personalInfo[4] == 19) && (cycle > 2))
            {
                resFinal = resFinal - 1;
            }

            if (CarmaticCounter == 3)
            {
                resFinal = resFinal - 0.5;
            }
            if (CarmaticCounter > 3)
            {
                resFinal = resFinal - 0.7;
            }


            return resFinal;
        }
        #endregion Health

        #region Combinations

        public string[] GetCombinationResults(Dictionary<string, int> chakras, Dictionary<string, int> lifePeriods, int curLifeCycle)
        {
            StringBuilder sb = new StringBuilder();
            string retVal = string.Empty;
            string[] arrRes =
                {
                "נטייה להסתבך עם רשויות", // 0
                "נטייה להיות עצמאי או מנהל", // 1
                "קושי להגיע לדרגת המאסטר בתעסוקה",// 2
                "נטייה לממש את הפוטנציאל ואת שאיפות החיים",// 3
                "נטייה לממש את הייעוד",// 4
                "קושי לממש את הייעוד",// 5
                "מומלץ לעבור את התיקון האישי באמצעות מסע",// 6
                "קושי להצליח בגדול כעצמאי בודד מומלץ לפעול עם שותפים או להיות קבלן משנה", // 7
                "קושי לממש את החיים וקשה בהרבה תחומים",// 8
                "קושי לממש את יכולות החשיבה של המאסטר", // 9
                "נטייה ללכת נגד הרוח וקושי בתחומים רבים במיוחד במערכות יחסים, זוגיות ועסקים", // 10
                "קושי במערכות יחסים וזוגיות", // 11
                "נטייה לחרדות שיאטו את קצב התקדמות האדם בחייו", // 12
                "נטייה לפחד מאירועים לא מוכרים ולחרדות שיאטו את קצב התקדמות האדם בחייו", // 13
                "נולדת באחד משלושת ימי הלידה הטובים בעולם, בבקשה לנצל מתנה זו בהיבט החיובי, אפשרות להרוויח מליונים. עם זאת יש לשים לב במיוחד בתקופת החיים הרביעית בה תהא ירידה אנרגטית בחיי האדם", // 14
                "מזלך האסטרולוגי מקשה עליך, מומלץ לחזק את המפה באמצעות שם פרטי נוסף או שם משפחה", // 15
                "מזלך האסטרולוגי מקל עליך להצליח בחיים נצל זאת לטובה וחזק מפתך", // 16
                "קושי לממש את שאיפות החיים", // 17
                "שמך הפרטי מקשה עליך לממש את שאיפות חייך", // 18
                "היקום פותח לפניך את כל השפע של העולם בכל התחומים ודוחף אותך להיות עצמאי מצליח", // 19
                "הייקום מקשה עליך, מומלץ לחזק את המפה על מנת לפתוח את השפע", // 20
                "נטייה להפסדים כספיים והסתבכות עם רשויות מומלץ לעסוק בתחומים הקשורים ללימוד/ללמד, לכתוב, ספורט, לעמוד על במה ולהעביר ידע (סטנד אפ ....)", // 21
                "נטייה להפסדים כספיים והסתבכות עם רשויות", // 22
                "הגעת לתקופת חיים עם תהפוכות", // 23
                "המפה מתארת נטייה למספרי השיגעון", // 24
                "במפתך קיימים מספרים מעדנים", // 25
                "יש לך פוטנציאל להרוויח סכומי כסף גדולים מאוד", // 26
                "יש לך פוטנציאל להרוויח סכומי כסף גדולים", // 27
                "קיים קושי להרוויח סכומי כסף גדולים", // 28
                "קיים פוטנציאל בשנים אילו להפסיד סכומי כסף גדולים מאוד ועד כדי פשיטת רגל ואפשרות לגנבת כספים - פעל בזהירות", // 29
                "מפתך מתארת אדם בעל פוטנציאל להפסיד סכומי כסף גדולים, ואפשרות לגנבת כספים - פעל בזהירות", //30
                "קיימת נטייה טבעית להומו/לסבית", // 31
                "נטייה להתנהגות של \"דרמה קווין\"", // 32
                "השם הפרטי אינו מאוזן ומקשה על חיי האדם בכל התחומים", // 33
                "שמך אינו מתאים ליום הלידה וגורם נזק במקום להוות יתרון, על מנת לפתוח את השפע מומלץ להוסיף שם פרטי או משפחה", //34
            };

            if (ChakraContainsValues(chakras, new string[] { "7" }, new int[] { 13, 14, 16, 19 }) &&
                ChakraContainsValues(chakras, new string[] { "5", "1", "2" }, new int[] { 8 }))
                sb.AppendLine(arrRes[0]);

            if (ChakraContainsValues(chakras, new string[] { "6" }, new int[] { 1, 8 }))
                sb.AppendLine(arrRes[1]);

            if (ChakraContainsValues(chakras, new string[] { "6" }, new int[] { 9, 11, 22 }))
                sb.AppendLine(arrRes[2]);

            if (ChakraContainsValues(chakras, new string[] { "6" }, new int[] { 1, 3, 5, 8 }))
                sb.AppendLine(arrRes[3]);

            if (ChakraContainsValues(chakras, new string[] { "5" }, new int[] { 1, 2, 3, 4, 5, 6, 33 }))
                sb.AppendLine(arrRes[4]);

            if (ChakraContainsValues(chakras, new string[] { "5" }, new int[] { 7, 8, 9, 11, 22, 13, 14, 16, 19 }))
                sb.AppendLine(arrRes[5]);

            if (LifePeriodsContainsValues(lifePeriods, new int[] { 13, 14, 16, 19 }, new int[] { 0, 1, 1, 0 }))
                //or all chakras... or - when 7 and 2 in any blue chakra >>>> than...
                sb.AppendLine(arrRes[6]);

            if (ChakraContainsValues(chakras, new string[] { "6" }, new int[] { 2, 4, 6, 7, 9, 11, 22, 13, 14, 16, 19 }))
                sb.AppendLine(arrRes[7]);

            if (ChakraContainsValues(chakras, new string[] { "4" }, new int[] { 0 }))
                sb.AppendLine(arrRes[8]);

            if (ChakraContainsValues(chakras, new string[] { "3" }, new int[] { 11, 22 }))
                sb.AppendLine(arrRes[9]);

            if (ChakraContainsValues(chakras, new string[] { "3" }, new int[] { 2, 7, 11, 13, 14, 16, 19 }))
                sb.AppendLine(arrRes[10]);

            if (ChakraContainsValues(chakras, new string[] { "8" }, new int[] { 7, 13, 14, 16, 19 }))
                // in chakras 5, 3, 1, 9, 8, first climax, current climax
                sb.AppendLine(arrRes[11]);

            if (ChakraContainsValues(chakras, new string[] { "7", "5", "3", "2", "1", "9", "8" }, new int[] { 16 }) ||
                LifePeriodsContainsValues(lifePeriods, new int[] { 16 }, new int[] { 0, 1, 1, 0 }, 1) ||
                LifePeriodsContainsValues(lifePeriods, new int[] { 16 }, new int[] { 0, 1, 1, 0 }, curLifeCycle))
                sb.AppendLine(arrRes[12]);

            if (ChakraContainsValues(chakras, new string[] { "5", "3", "2", "1", "9" }, new int[] { 7 }) ||
                LifePeriodsContainsValues(lifePeriods, new int[] { 7 }, new int[] { 0, 1, 1, 0 }, 1) ||
                LifePeriodsContainsValues(lifePeriods, new int[] { 7 }, new int[] { 0, 1, 1, 0 }, curLifeCycle))
                sb.AppendLine(arrRes[12]);

            if (ChakraContainsValues(chakras, new string[] { "5", "3", "9" }, new int[] { 2, 11 }))
                sb.AppendLine(arrRes[13]);

            if (ChakraContainsValues(chakras, new string[] { "1" }, new int[] { 9, 11, 22 }))
                sb.AppendLine(arrRes[14]);

            if (ChakraContainsValues(chakras, new string[] { "9" }, new int[] { 2, 7, 11 }))
                sb.AppendLine(arrRes[15]);

            if (ChakraContainsValues(chakras, new string[] { "9" }, new int[] { 22 }))
                sb.AppendLine(arrRes[16]);

            if (ChakraContainsValues(chakras, new string[] { "8" }, new int[] { 13, 14, 16, 19 }))
                sb.AppendLine(arrRes[17]);

            if (
                // and if there is any Carmatics in one of chakras or...
                ChakraContainsValues(chakras, new string[] { "3" }, new int[] { 2, 11 }) ||
                ChakraContainsValues(chakras, new string[] { "3", "2" }, new int[] { 7 }))
                sb.AppendLine(arrRes[18]);

            // if current climax contains 1,8,9,11,22
            if (LifePeriodsContainsValues(lifePeriods, new int[] { 1, 8, 9, 11, 22 }, new int[] { 0, 0, 1, 0 }, curLifeCycle))
                sb.AppendLine(arrRes[19]);

            // if current climax or present contains 7,13,14,16,19
            if (LifePeriodsContainsValues(lifePeriods, new int[] { 7, 13, 14, 16, 19 }, new int[] { 0, 0, 1, 0 }, curLifeCycle))
                sb.AppendLine(arrRes[20]);

            // if current climax or present contains 7,16
            if (LifePeriodsContainsValues(lifePeriods, new int[] { 7, 16 }, new int[] { 0, 1, 1, 0 }, curLifeCycle))
                sb.AppendLine(arrRes[21]);

            // if current climax or present contains 8
            if (LifePeriodsContainsValues(lifePeriods, new int[] { 8 }, new int[] { 0, 1, 1, 0 }, curLifeCycle))
                sb.AppendLine(arrRes[22]);

            // if current climax or present contains 9
            if (LifePeriodsContainsValues(lifePeriods, new int[] { 9 }, new int[] { 0, 1, 1, 0 }))
                sb.AppendLine(arrRes[23]);

            // "Madness" Numbers
            if (TendencyToMadnessNumbers(chakras, lifePeriods, curLifeCycle))
                sb.AppendLine(arrRes[24]);

            // "Graceful" Numbers
            if (ChakraContainsValues(chakras, new string[] { "5", "3", "1", "9" }, new int[] { 2, 3, 6, 9, 11, 22 }) &&
                LifePeriodsContainsValues(lifePeriods, new int[] { 2, 3, 6, 9, 11, 22 }, new int[] { 0, 0, 1, 0 }, curLifeCycle))
                sb.AppendLine(arrRes[25]);

            // "Money" Numbers
            if (ChakraContainsValues(chakras, new string[] { "6" }, new int[] { 1, 8 }) ||
                LifePeriodsContainsValues(lifePeriods, new int[] { 1, 8, 9, 11, 22 }, new int[] { 0, 0, 1, 0 }, curLifeCycle))
                sb.AppendLine(arrRes[26]);

            if (ChakraContainsValues(chakras, new string[] { "6" }, new int[] { 3, 5 }) ||
                LifePeriodsContainsValues(lifePeriods, new int[] { 3, 5 }, new int[] { 0, 0, 1, 0 }, curLifeCycle))
                sb.AppendLine(arrRes[27]);

            if (ChakraContainsValues(chakras, new string[] { "6" }, new int[] { 2, 4, 6, 7, 9, 13, 14, 16, 19 }))
                sb.AppendLine(arrRes[28]);

            if (LifePeriodsContainsValues(lifePeriods, new int[] { 7, 16, 19 }, new int[] { 0, 1, 1, 0 }, curLifeCycle))
                sb.AppendLine(arrRes[29]);

            if (ChakraContainsValues(chakras, new string[] { "2", "3", "5" }, new int[] { 7, 16, 19 }) ||
                ChakraContainsValues(chakras, new string[] { "8" }, new int[] { 16 }) ||
                ChakraContainsValues(chakras, new string[] { "9" }, new int[] { 7 }))
                sb.AppendLine(arrRes[30]);

            #region Calc Sexual orientation

            int sexualOrientationMark = 0;

            if (ChakraContainsValues(chakras, new string[] { "1" }, new int[] { 2, 11, 5, 14, 7, 16 })) sexualOrientationMark++;
            if (ChakraContainsValues(chakras, new string[] { "5" }, new int[] { 2, 11, 5, 14, 7, 16 })) sexualOrientationMark++;
            if (LifePeriodsContainsValues(lifePeriods, new int[] { 2, 11, 5, 14, 7, 16 }, new int[] { 0, 1, 0, 0 }, 1)) sexualOrientationMark++;
            if (LifePeriodsContainsValues(lifePeriods, new int[] { 13, 19 }, new int[] { 0, 1, 0, 0 }, 1)) sexualOrientationMark++;
            if (ChakraContainsValues(chakras, new string[] { "9" }, new int[] { 7, 2, 5, 11 })) sexualOrientationMark++;
            if (ChakraContainsValues(chakras, new string[] { "8" }, new int[] { 13, 14, 16, 19 })) sexualOrientationMark++;

            if (sexualOrientationMark >= 3)
            {
                sb.AppendLine(arrRes[31]);

            }

            #endregion Calc Sexual orientation

            if (ChakraContainsValues(chakras, new string[] { "5", "3", "2", "1", "9" }, new int[] { 7, 16 }) ||
                ChakraContainsValues(chakras, new string[] { "7", "8" }, new int[] { 16 }) ||
                ChakraContainsValues(chakras, new string[] { "5", "3", "9" }, new int[] { 2, 11 }) ||
                LifePeriodsContainsValues(lifePeriods, new int[] { 7, 16, 2, 11 }, new int[] { 0, 0, 1, 0 }, 1) ||
                LifePeriodsContainsValues(lifePeriods, new int[] { 7, 16, 2, 11 }, new int[] { 0, 0, 1, 0 }, curLifeCycle))
                sb.AppendLine(arrRes[32]);

            if (ChakraContainsValues(chakras, new string[] { "8" }, new int[] { 13, 14, 16, 19 }))
            {
                sb.AppendLine(arrRes[33]);

            }

            if (ChakraContainsValues(chakras, new string[] { "2" }, new int[] { 7, 13, 14, 16, 19 }))
            {
                sb.AppendLine(arrRes[34]);

            }

            var str = sb.ToString().Split(new char[] { '\n' }, StringSplitOptions.RemoveEmptyEntries).Distinct();

            return str.ToArray();
        }

        // **********

        private bool TendencyToMadnessNumbers(Dictionary<string, int> chakras, Dictionary<string, int> lifePeriods, int curLifeCycle)
        {
            bool isInclined = false;
            List<int> pairList = new List<int>();
            bool hasPair = false;

            #region Chakras

            string[][] chk =
            {
                new string[] { "7", "14", "16" },
                new string[] { "8", "14", "16" },
                new string[] { "5", "5", "7", "14", "16" },
                new string[] { "3", "5", "7", "14", "16" },
                new string[] { "1", "5", "7", "14", "16" },
                new string[] { "2", "7", "14", "16"},
                new string[] { "9", "5", "7" },

            };

            string[][] pairs =
            {
                new string[] { "5", "7" },
                new string[] { "5", "16" },
                new string[] { "14", "7" },
                new string[] { "14", "16" },

            };

            foreach (string[] strArr in chk)
            {
                // Is chakra contains any of madness numbers ? and number is not already inserted to list
                if (strArr.Any(n => n == chakras[strArr[0]].ToString() &&
                                         Array.IndexOf(strArr, n) > 0) &&
                   (pairList.IndexOf(chakras[strArr[0]]) == -1))
                {
                    pairList.Add(chakras[strArr[0]]);

                    if (pairList.Count == 2)
                    {
                        hasPair = true;
                        break;
                    }

                }

            }

            #endregion Chakras

            #region LifeCycles

            if (!hasPair && !isInclined)
            {
                string[][] periods =
                {
                    new string[] { "2" + curLifeCycle, "5", "7", "14", "16" },
                    new string[] { "3" + curLifeCycle, "5", "7", "14", "16" },
                    new string[] { "13", "7", "14", "16" },

                };

                foreach (string[] period in periods)
                {
                    if (period.Any(n => n == lifePeriods.ToString() &&
                                         Array.IndexOf(period, n) > 0) &&
                   (pairList.IndexOf(lifePeriods[period[0]]) == -1))
                    {
                        pairList.Add(lifePeriods[period[0]]);

                        if (pairList.Count == 2)
                        {
                            hasPair = true;
                            break;
                        }

                    }

                }

            }

            #endregion LifeCycles

            if (hasPair)
            {
                foreach (string[] pair in pairs)
                {
                    if (pairList.Sum() == pair.Sum(n => Convert.ToInt32(n)))
                    {
                        isInclined = true;
                        break;

                    }

                }

            }

            return isInclined;

            // אם שני תנאים מתקיימים, יש לבדוק אם המספרים הם לפי הצמדים
            // 5<->7, 5<->16, 14<->7, 14<->16
            //if (ChakraContainsValues(chakras, new string[] { "7", "8" }, new int[] { 14, 16 }) ||
            //    ChakraContainsValues(chakras, new string[] { "5", "3", "1" }, new int[] { 5, 7, 14, 16 }) ||
            //    ChakraContainsValues(chakras, new string[] { "2" }, new int[] { 7, 14, 16 }) ||
            //    ChakraContainsValues(chakras, new string[] { "9" }, new int[] { 5, 7 }) ||
            //    LifePeriodsContainsValues(lifePeriods, new int[] { 5, 7, 14, 16 }, new int[] { 0, 1, 1, 0 }, curLifeCycle) ||
            //    LifePeriodsContainsValues(lifePeriods, new int[] { 7, 14, 16 }, new int[] { 0, 0, 1, 0 }, 1))

        }

        // **********

        public bool ChakraContainsValues(Dictionary<string, int> chakras, string[] arrChakras, int[] values)
        {
            bool retVal = false;

            foreach (string chakraNum in arrChakras)
            {
                retVal |= values.Any(n => n == chakras[chakraNum]);

            }

            return retVal;
        }

        // **********

        public bool LifePeriodsContainsValues(Dictionary<string, int> lifePeriods, int[] values, int[] checkMask, int curLifeCycle = 0)
        {
            bool retVal = false;

            for (int i = 0; i < checkMask.Length; i++)
            {
                if (checkMask[i] == 1)
                {
                    string k = (i + 1).ToString();

                    // Return current period of one of: "age", climax, period, challenge
                    if (curLifeCycle > 0)
                    {
                        retVal |= values.Any(n => n == lifePeriods[curLifeCycle.ToString() + k]);

                    }
                    else
                    {
                        foreach (KeyValuePair<string, int> item in lifePeriods.Where(key => key.Key.StartsWith(k)))
                        {
                            retVal |= values.Any(n => n == item.Value);

                        }

                    }

                }

            }
            return retVal;

        }

        #endregion Combinations

        /// <summary>
        /// test if the map is strong or not
        /// </summary>
        /// <param name="ckra">all chakra values (except Heart, Throught, SexAncCreation)</param>
        /// <param name="throught">Throught chakra value</param>
        /// <param name="sexcreation">SexAncCreation chakra value</param>
        /// <returns>true is map is strong, false otherwise</returns>
        public bool isMapStrong(List<int> ckra, int throught, int sexcreation)
        {
            bool res = true; // assuming strong map...

            if ((throught == 8) || (sexcreation == 8) || (throught == 1) || (sexcreation == 1)) // חריג למפה חזקה
            {
                return true;
            }

            ckra.Add(throught);
            ckra.Add(sexcreation);

            List<int> weaknums = new List<int>() { 2, 3, 6, 7, 33 };

            int count = 0;
            for (int i = 0; i < ckra.Count; i++)
            {
                if ((isCarmaticNumber((ckra[i])) == true) || (isNumInList(ckra[i], weaknums) == true))
                {
                    count++;
                }

            }

            if (count >= 3)
            {
                res = false;
            }

            return res;
        }

    }

    #region Deleted
    ///// <summary> Name Calculation
    ///// 
    ///// </summary>
    ///// <param name="FirstName">String of the First Name</param>
    ///// <param name="LastName">String of the Last Name</param>
    ///// <param name="FirstNameCalc">Out string of the First Name Calculation</param>
    ///// <param name="LastNameCalc">Out string of the Last Name Calculation</param>
    ///// <param name="FullNameCalc">Out string of the Full Name Calculation</param>
    //public void CalculateName(string FirstName, string LastName, out string FirstNameCalc, out string LastNameCalc, out string FullNameCalc)
    //{
    //    FirstNameCalc = CalcName(FirstName);
    //    LastNameCalc = CalcName(LastName);

    //    string[] fName = FirstNameCalc.Split(mDelimiter);
    //    string[] lName = LastNameCalc.Split(mDelimiter);

    //    int ifName = Convert.ToInt32(fName[fName.Length - 1]);
    //    int ilName = Convert.ToInt32(lName[lName.Length - 1]);
    //    FullNameCalc = Nums_2_String(ifName + ilName);
    //}s
    #endregion

    #region Inner Calsses
    internal class AstroLuck
    {
        #region Data Members
        private int[] AstroFromDay = new int[13];
        private int[] AstroToDay = new int[13];
        private string[] AstroLuckName = new string[13];
        private int[] AstroNumber = new int[13];
        #endregion

        #region Constructor
        public AstroLuck()
        {
            AstroLuckName[0] = "אריה";
            AstroFromDay[0] = new DateTime(DateTime.Now.Year, 07, 23).DayOfYear;
            AstroToDay[0] = new DateTime(DateTime.Now.Year, 08, 22).DayOfYear;
            AstroNumber[0] = 1;
            if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
            {
                AstroLuckName[0] = "Lion";
            }

            AstroLuckName[1] = "סרטן";
            AstroFromDay[1] = new DateTime(DateTime.Now.Year, 06, 22).DayOfYear;
            AstroToDay[1] = new DateTime(DateTime.Now.Year, 07, 22).DayOfYear;
            AstroNumber[1] = 2;
            if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
            {
                AstroLuckName[1] = "Cancer";
            }

            AstroLuckName[2] = "קשת";
            AstroFromDay[2] = new DateTime(DateTime.Now.Year, 11, 22).DayOfYear;
            AstroToDay[2] = new DateTime(DateTime.Now.Year, 12, 21).DayOfYear;
            AstroNumber[2] = 3;
            if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
            {
                AstroLuckName[2] = "Sagittarius";
            }

            AstroLuckName[3] = "בתולה";
            AstroFromDay[3] = new DateTime(DateTime.Now.Year, 08, 23).DayOfYear;
            AstroToDay[3] = new DateTime(DateTime.Now.Year, 09, 22).DayOfYear;
            AstroNumber[3] = 4;
            if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
            {
                AstroLuckName[3] = "Virgo";
            }

            AstroLuckName[4] = "תאומים";
            AstroFromDay[4] = new DateTime(DateTime.Now.Year, 05, 21).DayOfYear;
            AstroToDay[4] = new DateTime(DateTime.Now.Year, 06, 21).DayOfYear;
            AstroNumber[4] = 5;
            if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
            {
                AstroLuckName[4] = "Gemini";
            }

            AstroLuckName[5] = "מאזניים";
            AstroFromDay[5] = new DateTime(DateTime.Now.Year, 09, 23).DayOfYear;
            AstroToDay[5] = new DateTime(DateTime.Now.Year, 10, 23).DayOfYear;
            AstroNumber[5] = 6;
            if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
            {
                AstroLuckName[5] = "Libra";
            }

            AstroLuckName[6] = "דגים";
            AstroFromDay[6] = new DateTime(DateTime.Now.Year, 02, 19).DayOfYear;
            AstroToDay[6] = new DateTime(DateTime.Now.Year, 03, 20).DayOfYear;
            AstroNumber[6] = 7;
            if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
            {
                AstroLuckName[6] = "Pisces";
            }

            AstroLuckName[7] = "טלה";
            AstroFromDay[7] = new DateTime(DateTime.Now.Year, 03, 21).DayOfYear;
            AstroToDay[7] = new DateTime(DateTime.Now.Year, 04, 20).DayOfYear;
            AstroNumber[7] = 9;
            if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
            {
                AstroLuckName[7] = "Aries";
            }

            AstroLuckName[8] = "דלי";
            AstroFromDay[8] = new DateTime(DateTime.Now.Year, 01, 21).DayOfYear;
            AstroToDay[8] = new DateTime(DateTime.Now.Year, 02, 18).DayOfYear;
            AstroNumber[8] = 11;
            if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
            {
                AstroLuckName[8] = "Aquarius";
            }

            AstroLuckName[9] = "עקרב";
            AstroFromDay[9] = new DateTime(DateTime.Now.Year, 10, 24).DayOfYear;
            AstroToDay[9] = new DateTime(DateTime.Now.Year, 11, 21).DayOfYear;
            AstroNumber[9] = 22;
            if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
            {
                AstroLuckName[9] = "Scorpio";
            }

            AstroLuckName[10] = "שור";
            AstroFromDay[10] = new DateTime(DateTime.Now.Year, 04, 21).DayOfYear;
            AstroToDay[10] = new DateTime(DateTime.Now.Year, 05, 20).DayOfYear;
            AstroNumber[10] = 33;
            if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
            {
                AstroLuckName[10] = "Taurus";
            }

            AstroLuckName[11] = "גדי";
            AstroFromDay[11] = new DateTime(DateTime.Now.Year, 12, 22).DayOfYear;
            AstroToDay[11] = new DateTime(DateTime.Now.Year, 12, 31).DayOfYear;
            AstroNumber[11] = 8;
            if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
            {
                AstroLuckName[11] = "Capricorn";
            }

            AstroLuckName[12] = "גדי";
            AstroFromDay[12] = new DateTime(DateTime.Now.Year, 01, 01).DayOfYear;
            AstroToDay[12] = new DateTime(DateTime.Now.Year, 01, 20).DayOfYear;
            AstroNumber[12] = 8;
            if (AppSettings.Instance.ProgramLanguage == AppSettings.Language.English)
            {
                AstroLuckName[12] = "Capricorn";
            }
        }
        #endregion

        #region Public Methods
        public int Date2AstroLuck(DateTime date, out string LuckName)
        {
            //bool found = false;
            int LuckNum = -1;
            LuckName = "";

            date = date.AddYears(DateTime.Now.Year - date.Year);

            for (int i = 0; i < AstroFromDay.Length; i++)
            {
                if ((date.DayOfYear >= AstroFromDay[i]) && (date.DayOfYear <= AstroToDay[i]))
                {
                    //found = true;
                    LuckNum = AstroNumber[i];
                    LuckName = AstroLuckName[i];
                }
            }

            return LuckNum;
        }

        public string GetAstroNameByNumber(int num)
        {
            for (int i = 0; i < AstroLuckName.Length; i++)
            {
                if (AstroNumber[i] == num)
                {
                    return AstroLuckName[i];
                }
            }

            return "";
        }

        #endregion
    }

    internal class Proffing
    {
        #region Data Members
        private bool mProofingValue;
        private string mWinSN;

        private string mFirstKey;
        //private string mSecondKey;

        private bool isPatched;
        private string mDeafaltMSsn = @"91369-DMS-8292433-00082";
        //private DateTime mDate = new DateTime(2018, 6, 23);
        #endregion

        #region Constructor
        public Proffing()
        {
            //mWinSN = "5439871";
            mProofingValue = false;

            isPatched = false;
            DirectoryInfo overwiteDir = new DirectoryInfo(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData) + @"\dc");
            if (overwiteDir.Exists == true)
            {
                FileInfo overwriteFile = new FileInfo(Path.Combine(overwiteDir.FullName, @"rlk767rv.xml"));
                if (overwriteFile.Exists == true)
                {
                    isPatched = true;
                }
                else
                {
                    mWinSN = GetWinSN();
                }
            }
            else
            {
                mWinSN = GetWinSN();
            }
        }
        #endregion

        #region Properties
        public bool Proofed
        {
            get
            {
                return mProofingValue;
            }
        }
        #endregion

        #region Public Methods
        public string CreateFirstKey()
        {
            mFirstKey = SetUnique7Digits();
            return mFirstKey;
        }

        public bool IsFirstKeyMatchesPC(string firstKey)
        {
            bool res = true;

            if (isPatched == false)
            {
                string SubSyntheticFirstKey = CreateFirstKey().Substring(2, 5);


                res = (firstKey.Substring(2, 5) == SubSyntheticFirstKey);
            }
            else
            {
                res = true;
            }

            return res;
        }

        public bool TestPassword(string sFirstKey, string SecondKey)
        {

            string synthetic2key = GenerateSecondKEY(sFirstKey);

            bool res1 = SecondKey.Substring(0, 3) == synthetic2key.Substring(0, 3);
            bool res2 = SecondKey.Substring(SecondKey.Length - 3, 3) == synthetic2key.Substring(SecondKey.Length - 3, 3);

            bool res3 = true;// test FirstKey to match PC

            if (isPatched == false)
            {
                if (mDeafaltMSsn == mWinSN)
                {
                    return false;
                }
                else
                {
                    return (res1 & res2 & res3);
                }
            }
            else //(isPatched == true)
            {
                return isPatched;
            }
        }
        #endregion

        #region Private Methods
        /// <summary> Getting the MS Windows Serial Number
        /// 
        /// </summary>
        /// <returns></returns>
        private string GetWinSN()
        {
            string outString = "";

            try
            {
                ManagementObjectSearcher searcher;
                string queryObject = "Win32_OperatingSystem";
                int i = 0;

                searcher = new ManagementObjectSearcher("SELECT * FROM " + queryObject);
                foreach (ManagementObject wmi_HD in searcher.Get())
                {
                    i++;
                    PropertyDataCollection searcherProperties = wmi_HD.Properties;
                    foreach (PropertyData sp in searcherProperties)
                    {
                        if (sp.Name == "SerialNumber")
                        {
                            outString = sp.Value.ToString();
                            break;
                        }
                    }
                }
            }
            catch
            {
                outString = mDeafaltMSsn;
                CreateOverWrite();
                isPatched = true;
            }

            return outString;
        }

        private string GetMacAdress()
        {

            try
            {
                System.Net.NetworkInformation.IPGlobalProperties computerProperties = System.Net.NetworkInformation.IPGlobalProperties.GetIPGlobalProperties();

                System.Net.NetworkInformation.NetworkInterface[] nics = System.Net.NetworkInformation.NetworkInterface.GetAllNetworkInterfaces();

                Console.WriteLine("Interface information for {0}.{1} ",

                computerProperties.HostName, computerProperties.DomainName);

                if (nics == null || nics.Length < 1)
                {
                    return (" No network interfaces found.");
                }

                Console.WriteLine(" Number of interfaces .................... : {0}", nics.Length);

                foreach (System.Net.NetworkInformation.NetworkInterface adapter in nics)
                {
                    System.Net.NetworkInformation.IPInterfaceProperties properties = adapter.GetIPProperties(); // .GetIPInterfaceProperties();

                    System.Net.NetworkInformation.PhysicalAddress address = adapter.GetPhysicalAddress();
                    if (address.ToString() != null)
                    {
                        return (address.ToString());
                    }
                }

                return "err";
            }
            catch
            {
                return "err";
            }
        }

        private string SetUnique7Digits()
        {
            string sFinalNums = "";

            System.Random rndObj = new Random();
            int TwoDigs = rndObj.Next(10, 99);
            long n1, n2;

            try
            {


                mWinSN = GetWinSN();

                string[] snParts = mWinSN.Split("-".ToCharArray());

                bool res1 = long.TryParse(snParts[0], out n1);
                bool res2 = long.TryParse(snParts[2], out n2);

                sFinalNums = (TwoDigs.ToString() + n1.ToString().Substring(0, 3) + n2.ToString().Substring(0, 2));
            }
            catch // no WinSN
            {
                string macAdress = GetMacAdress().ToLower();

                int d1, d2;
                n1 = 0; n2 = 0;
                for (int i = 0; i < macAdress.Length / 2; i++)
                {
                    d1 = Convert.ToInt32(macAdress[2 * i]);
                    d2 = Convert.ToInt32(macAdress[2 * i + 1]);

                    int div;
                    Math.DivRem(i, 2, out div);
                    if (div == 0)
                    {
                        n1 += d1 + d2;
                    }
                    else
                    {
                        n2 += d1 + d2;
                    }

                }


                sFinalNums = TwoDigs.ToString();

                if (n1.ToString().Length >= 3)
                {
                    sFinalNums += n1.ToString().Substring(0, 3);
                }
                else
                {
                    sFinalNums += n1.ToString();
                }


                if (n2.ToString().Length >= 2)
                {
                    sFinalNums += n2.ToString().Substring(0, 2);
                }
                else
                {
                    sFinalNums += n2.ToString();
                }

                while (sFinalNums.Length < 7)
                {
                    sFinalNums += "9";
                }
            }

            return sFinalNums;
        }
        #endregion

        private string GenerateSecondKEY(string FirstKEY)
        {
            System.Random rndObj = new Random();
            int TwoDigs = rndObj.Next(10, 99);

            int n1 = Convert.ToInt16(FirstKEY.Substring(2, 3));
            int n2 = Convert.ToInt16(FirstKEY.Substring(5, 2));

            int sum1 = n1 + n2;

            n1 = Convert.ToInt16(FirstKEY.Substring(3, 3));
            n2 = Convert.ToInt16(FirstKEY.Substring(2, 1) + FirstKEY.Substring(6, 1));
            int sum2 = n1 * n2;

            return sum1.ToString().Substring(0, 3) + TwoDigs.ToString() + sum2.ToString().Substring(0, 3);
        }


        private void CreateOverWrite()
        {
            DirectoryInfo overwiteDir = new DirectoryInfo(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData) + @"\dc");
            if (overwiteDir.Exists == false)
            {
                overwiteDir.Create();
            }

            FileInfo overwriteFile = new FileInfo(Path.Combine(overwiteDir.FullName, @"rlk767rv.xml"));
            if (overwriteFile.Exists == false)
            {
                overwriteFile.Create();
            }
        }
        private bool isOverWrite()
        {
            DirectoryInfo overwiteDir = new DirectoryInfo(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData) + @"\dc");
            FileInfo overwriteFile = new FileInfo(Path.Combine(overwiteDir.FullName, @"rlk767rv.xml"));

            return (overwriteFile.Exists && overwiteDir.Exists);
        }
    }
    #endregion //Inner Calsses
}
