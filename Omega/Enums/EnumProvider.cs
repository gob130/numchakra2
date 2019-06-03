using System;
using System.Collections.Generic;
using System.Text;
using System.Reflection;
using Omega;
using System.Windows.Forms;

namespace Omega.Enums
{
    public class EnumProvider
    {
        public static EnumProvider Instance = new EnumProvider();
        private Form mMainForm;

        public EnumProvider()
        {
            mMainForm = MainForm.ActiveForm; //Reports.ReportDataProvider.Instance.mMainForm;
            //mMainForm = MainForm.ActiveForm;
        }

        public void Init()
        {
        }

        public string GetEnumDescription(Enum en)
        {

            Type type = en.GetType();

            MemberInfo[] memInfo = type.GetMember(en.ToString());

            if (memInfo != null && memInfo.Length > 0)
            {

                object[] attrs = memInfo[0].GetCustomAttributes(typeof(Description),
                false);

                if (attrs != null && attrs.Length > 0)

                    return ((Description)attrs[0]).Text;

            }

            return en.ToString();

        }

        public Enum GetAlignenmentEnumFromDescription(string desc)
        {
            foreach (Alignment item in Enum.GetValues(typeof(Alignment)))
            {
                if (desc.ToLower() == GetEnumDescription(item).ToLower())
                {
                    return item;
                }
            }

            return null;
        }

        public string GetReservedImageEnumFromDescription(string desc)
        {
            foreach (ReservedImages item in Enum.GetValues(typeof(ReservedImages)))
            {
                if (desc.ToLower() == item.ToString().ToLower())
                {
                    return GetEnumDescription(item).ToLower();
                }
            }

            return null;
        }

        public string GetReservedNameEnumFromDescription(string desc)
        {
            foreach (ReservedNames item in Enum.GetValues(typeof(ReservedNames)))
            {
                if (desc.ToLower() == item.ToString().ToLower())
                {
                    return GetEnumDescription(item).ToLower();
                }
            }

            return null;
        }

        public string GetReservedChakraEnumFromDescription(string desc)
        {
            foreach (ReservedChakra item in Enum.GetValues(typeof(ReservedChakra)))
            {
                if (desc.ToLower() == item.ToString().ToLower())
                {
                    return GetEnumDescription(item).ToLower();
                }
            }

            return null;
        }

        public string GetReservedWorkEnumFromDescription(string desc)
        {
            foreach (ReservedWork item in Enum.GetValues(typeof(ReservedWork)))
            {
                if (desc.ToLower() == item.ToString().ToLower())
                {
                    return GetEnumDescription(item).ToLower();
                }
            }

            return null;
        }

        public ReservedWork GetReservedWorkEnumFromString(string desc)
        {
            foreach (ReservedWork item in Enum.GetValues(typeof(ReservedWork)))
            {
                if (desc.ToLower() == item.ToString().ToLower())
                {
                    return item;
                }
            }
            throw new Exception("Internal Enum Error - ReservedWork : " + desc);
            //return ReservedWork._תעסוקה_משני_;
        }

        public string GetReservedLifeCycleEnumFromDescription(string desc)
        {
            foreach (ReservedLifeCycle item in Enum.GetValues(typeof(ReservedLifeCycle)))
            {
                if (desc.ToLower() == item.ToString().ToLower())
                {
                    return GetEnumDescription(item).ToLower();
                }
            }

            return null;
        }

        public string GetReservedOPSEnumFromDescription(string desc)
        {
            foreach (ReservedOPS item in Enum.GetValues(typeof(ReservedOPS)))
            {
                if (desc.ToLower() == item.ToString().ToLower())
                {
                    return GetEnumDescription(item).ToLower();
                }
            }

            return null;
        }

        public string GetReservedCrysisEnumFromDescription(string desc)
        {
            foreach (ReservedCrysis item in Enum.GetValues(typeof(ReservedCrysis)))
            {
                if (desc.ToLower() == item.ToString().ToLower())
                {
                    return GetEnumDescription(item).ToLower();
                }
            }

            return null;
        }

        public ReservedCrysis GetReservedCrysisEnumFromString(string desc)
        {
            foreach (ReservedCrysis item in Enum.GetValues(typeof(ReservedCrysis)))
            {
                if (desc.ToLower() == item.ToString().ToLower())
                {
                    return item;
                }
            }
            throw new Exception("Internal Enum Error - ReservedCrysis : " + desc);
            //return ReservedCrysis._משבר_בסיס_;
        }

        public string GetBalanceEnumFromDescription(string desc)
        {
            foreach (Balance item in Enum.GetValues(typeof(Balance)))
            {
                if (desc.ToLower() == item.ToString().ToLower())
                {
                    return GetEnumDescription(item).ToLower();
                }
            }

            return null;
        }

        public Balance GetReservedBalanceEnumFromString(string desc)
        {
            foreach (Balance item in Enum.GetValues(typeof(Balance)))
            {
                if (desc.ToLower() == item.ToString().ToLower())
                {
                    return item;
                }
            }
            throw new Exception("Internal Enum Error - Balance : " + desc);
            //return Balance._לא_מאוזן_;
        }

        #region Sex
        public enum Sex
        {
            [Description("זכר")]
            Male = 0,
            [Description("נקבה")]
            Female = 1
        }

        public string GetSexEnumFromDescription(string desc)
        {
            foreach (Sex item in Enum.GetValues(typeof(Sex)))
            {
                if (desc.ToLower() == item.ToString().ToLower())
                {
                    return GetEnumDescription(item).ToLower();
                }
            }

            return null;
        }

        public Sex GetSexEnumFromString(string desc)
        {
            foreach (Sex item in Enum.GetValues(typeof(Sex)))
            {
                if ((desc.ToLower() == item.ToString().ToLower()) || (desc.ToLower() == GetSexEnumFromDescription(item.ToString())))
                {
                    return item;
                }
            }
            throw new Exception("Internal Enum Error - Sex : " + desc);
            //return Sex.Male;
        }
        #endregion Sex

        #region Passed Rectification 
        public enum PassedRectification
        {
            [Description("לא")]
            NotPassed = 0,
            [Description("כן")]
            Passed = 1
        }

        public string GetPassedRectificationEnumFromDescription(string desc)
        {
            foreach (PassedRectification item in Enum.GetValues(typeof(PassedRectification)))
            {
                if (desc.ToLower() == item.ToString().ToLower())
                {
                    return GetEnumDescription(item).ToLower();
                }
            }

            return "";
        }

        public PassedRectification GetPassedRectificationEnumFromString(string desc)
        {
            foreach (PassedRectification item in Enum.GetValues(typeof(PassedRectification)))
            {
                if (desc.ToLower() == item.ToString().ToLower())
                {
                    return item;
                }
            }
            throw new Exception("Internal Enum Error - Passed Rectification : " + desc);
            //return PassedRectification.NotPassed;
        }
        #endregion Sex

        #region Reached Master
        public enum ReachedMaster
        {
            [Description("לא")]
            No = 0,
            [Description("כן")]
            Yes = 1
        }

        public string GetIsMasterEnumFromDescription(string desc)
        {
            foreach (ReachedMaster item in Enum.GetValues(typeof(ReachedMaster)))
            {
                if (desc.ToLower() == item.ToString().ToLower())
                {
                    return GetEnumDescription(item).ToLower();
                }
            }

            return "";
        }

        public ReachedMaster GGetIsMasterEnumFromString(string desc)
        {
            foreach (ReachedMaster item in Enum.GetValues(typeof(ReachedMaster)))
            {
                if (desc.ToLower() == item.ToString().ToLower())
                {
                    return item;
                }
            }
            throw new Exception("Internal Enum Error - Reached Master : " + desc);
            //return ReachedMaster.No;
        }
        #endregion Sex

        #region XmlDB Internal Fields
        public enum InnerXmlDBFields
        {
            [Description("מזהה")]
            Id = -1,
            [Description("שם")]
            FirstName = 0,
            [Description("שם משפחה")]
            LastName = 1,
            [Description("שם אב")]
            FatherName = 2,
            [Description("שם אם")]
            MotherName = 3,
            [Description("תאריך לידה")]
            BirthDay = 4,
            [Description("מין")]
            Sex = 5,
            [Description("עיר")]
            City = 6,
            [Description("רחוב")]
            Street = 7,
            [Description("בניין")]
            Building = 8,
            [Description("דירה")]
            Appartment = 9,
            [Description("עבר תיקון")]
            PassedRect = 10,
            [Description("מקצוע")]
            Application = 11,
            [Description("טלפונים")]
            Phones = 12,
            [Description("אי-מייל")]
            EMail = 13,
            [Description("הגיע למאסטר")]
            ReachedMaster = 14

        }

        public string GetInnerXmlDBFieldsEnumFromDescription(string desc)
        {
            foreach (InnerXmlDBFields item in Enum.GetValues(typeof(InnerXmlDBFields)))
            {
                if (desc.ToLower() == item.ToString().ToLower())
                {
                    return GetEnumDescription(item).ToLower();
                }
            }

            return null;
        }

        public InnerXmlDBFields GetInnerXmlDBFieldsEnumFromString(string desc)
        {
            foreach (InnerXmlDBFields item in Enum.GetValues(typeof(InnerXmlDBFields)))
            {
                if (desc.ToLower() == item.ToString().ToLower())
                {
                    return item;
                }
            }
            throw new Exception("Internal Enum Error - InnerXmlDBFields : " + desc);
            //return InnerXmlDBFields.FirstName;
        }
        #endregion XmlDB Internal Fields

        public enum Alignment
        {
            [Description("left")]
            Left = 0,
            [Description("center")]
            Center = 1,
            [Description("right")]
            Right = 2,
        }

        public enum ReservedImages
        {
            [Description("tabPage4")]
            _פיתגורס_ = 0,
            [Description("tabPage1")]
            _ראשי_ = 1,
            [Description("grobxCycles")] //tabPage5
            _מחזורי_חיים_ = 2,
            [Description("tabPage3")]
            _מפה_איטנסיבית_ = 3,
            [Description("tabPage6")]
            _מפה_משולבת_ = 4,
            [Description("ChakraTab")]
            _פתיחת_צאקרות_ = 5,
            [Description("tabPage7")]
            _התאמה_זוגית_ = 6,
            [Description("tabPage2")]
            _הצלחה_בלימודים_קשב_וריכוז_ = 7,
            [Description("Single")]
            _הצלחה_עסקית_אישית_ = 8,
            [Description("Multiple")]
            _הצלחה_עסקית_משותפת_ = 9
        }

        public enum ReservedNames
        {
            [Description("txtPrivateName")]
            _שם_ = 0,
            [Description("txtFamilyName")]
            _שם_משפחה_ = 1,
            [Description("txtFatherName")]
            _שם_אבא_ = 2,
            [Description("txtMotherName")]
            _שם_אמא_ = 3,
            [Description("DateTimePickerFrom")]
            _תאריך_לידה_ = 4,
            [Description("txtCity")]
            _עיר_ = 5,
            [Description("txtStreet")]
            _רחוב_ = 6,
            [Description("txtBiuldingNum")]
            _מספר_בניין_ = 7,
            [Description("txtAppNum")]
            _מספר_דירה_ = 8,
        }

        public enum ReservedChakra
        {
            [Description("txtNum1")] // כתר
            _צאקרה_1_ = 0,
            [Description("txtNum2")] // חכמה
            _צאקרה_2_ = 1,
            [Description("txtNum3")]// בינה
            _צאקרה_3_ = 2,
            [Description("txtNum4")]//חסד
            _צאקרה_4_ = 3,
            [Description("txtNum5")]//גבורה
            _צאקרה_5_ = 4,
            [Description("txtNum6")]//תפארת
            _צאקרה_6_ = 5,
            [Description("txtNum7")]//נצח
            _צאקרה_7_ = 6
        }

        public enum ReservedLifeCycle
        {
            [Description("txt$_4")]
            _אתגר_ = 0,
            [Description("txt$_3")]
            _שיא_ = 1,
            [Description("txt$_2")]
            _מחזור_ = 2
        }

        public enum ReservedWork
        {
            [Description("txtNum6")]
            _תעסוקה_ראשי_ = 0,
            [Description("txtAstroName")]
            _תעסוקה_משני_ = 1,
            [Description("txtPName_Num")]
            _תעסוקה_שם_פרטי_ = 2,
            [Description("")]
            _בדיקה_חריגה_ = 3
        }

        public enum Balance
        {
            [Description("Balanced")]
            _מאוזן_ = 0,
            [Description("Half Balanced")]
            _מאוזן_חלקית_ = 1,
            [Description("Not Balanced")]
            _לא_מאוזן_ = 2
        }

        public enum ReservedCrysis
        {
            [Description("txtNum7")]
            _משבר_בסיס_ = 0,
            [Description("txtNum4")]
            _משבר_לב_ = 1,
            [Description("txtPName_Num")]
            _משבר_שם_פרטי_ = 2
        }

        public enum ReservedOPS
        {
            [Description("Pro True")]
            _הפוך_למקוצר_ = -3,
            [Description("Pro False")]
            _הפוך_לארוך_ = -2,
            [Description("")]
            _הפוך_מומחה_ = -1,
            [Description("txtAstroName")]
            _ערך_מזל_אסטרולוגי_ = 0,
            [Description("txtPName_Num")]
            _ערך_שם_פרטי_ = 1,
            [Description("txtPName_Num")]
            _ערך_שם_פרטי_ייעוד_ = 2,
            [Description("txtPName_Num")]
            _ערך_שם_פרטי_תכונות_אופי_ = 3,
            [Description("txtPowerNum")]
            _ערך_מספר_הכוח_ = 4,
            [Description("txtPYear")]
            _ערך_שנה_אישית_ = 5,
            [Description("txtPMonth")]
            _ערך_חודש_אישי_ = 6,
            [Description("txtPowerNum")]
            _ערך_יום_אישי_ = 7,
            [Description("txtNum5")]
            _ערך_ייעוד_ = 8,
            [Description("")]
            _ערך_תיקון_ראשי_ = 9,
            [Description("")]
            _ערך_חשש_ופחד_ = 10,
            [Description("")]
            _ערך_חרדות_ = 11,
            [Description("")]
            _ערך_מפה_דיכאוני_ = 12,
            [Description("")]
            _ערך_מפה_מאוזן_ = 13,
            [Description("txtPYear")]
            _מספר_שנה_אישית_ = 14,
            [Description("txtPMonth")]
            _מספר_חודש_אישי_ = 15,
            [Description("txtPDay")]
            _מספר_יום_אישי_ = 16,
            [Description("")]
            _מידע_מחזור_ראשון_ = 17,
            [Description("")]
            _מידע_מחזור_שני_ = 18,
            [Description("")]
            _מידע_מחזור_שלישי_ = 19,
            [Description("")]
            _מידע_מחזור_רביעי_ = 20,
            [Description("")]
            _ערכי_פיתגורס_ = 21,
            [Description("txtFinalBusinessMark")]
            _ערך_הצלחה_עסקית_ = 22,
            [Description("txtFinalBusinessMark")]
            _מלל_הצלחה_עסקית_ = 23,
            [Description("")]
            _מלל_המלצות_אישיות_ = 24,
            [Description("txtMrrgMark")]
            _ערך_הצלחה_זוגית_ = 25,
            [Description("txtMrrgMark")]
            _מלל_הצלחה_זוגית_ = 26,
            [Description("")]
            _ערכי_פתיחת_צאקרות_ = 27,
            [Description("")]
            _היום_ = 28,
            [Description("")]
            _אישיות_ = 29,
            [Description("")]
            _מידע_שנים_אישיות_עתיד_ = 30,
            [Description("txtLeanSccss")]
            _ערך_הצלחה_בלימודים_ = 31,
            [Description("txtLearnStory")]
            _מלל_הצלחה_בלימודים_ = 32,
            [Description("txtAttMinor")]
            _ערך_קשב_וריכוז_קלות_ = 33,
            [Description("txtAttMajor")]
            _ערך_קשב_וריכוז_מורכבות_ = 34,
            [Description("")]
            _הפרעות_קשב_וריכוז_ = 35,
            [Description("txtAstroName")]
            _מחלות_ = 36,
            [Description("txtNum6")]
            _יכולת_לשמור_על_הבריאות_ = 37,
            [Description("txtNum6")]
            _יכולת_לטפל_ = 38,
            [Description("txtNum5")]
            _תכונות_בסיסיות_לגוף_ = 39,
            [Description("txtHealthValue")]
            _בריאות_ = 40,
            [Description("")]
            _מידע_מחזור_נוכחי_שיא_בלבד_ = 41,
            [Description("")]
            _מידע_שנים_אישיות_עתיד_2_ = 42,
            [Description("txtPName_Num")]
            _בריאות_שם_פרטי_ = 43,
            [Description("")]
            _בריאות_שיא_ = 44,
            [Description("txtAstroName")]
            _בריאות_מזל_אסטרולוגי_ = 45,
            [Description("txtPYear")]
            _בריאות_שנה_אישית_ = 46,
            [Description("txtHealthStory")]
            _בריאות_סיכום_ = 47,
            [Description("")]
            _פרטי_בן_זוג_ = 48,
            [Description("")]
            _פרטי_שותפים_ = 49,
            [Description("txtMultiPartnerStory")]
            _מלל_שותפות_עסקית_ = 50,
            [Description("txtFinalMultipleBusineesMartk")]
            _ערך_שותפות_עסקית_ = 51,
            [Description("")]
            _ערך_תיקון_מקוצר_ = 52,
            [Description("")]
            _סיכום_דוח_והמלצות_ = 53,
            [Description("")]
            _מידע_מחזור_נוכחי_ = 54,
            [Description("txtMapStrong")]
            _מפה_חזקה_חלשה_ = 55,
            [Description("")]
            _מתנות_יום_הלידה_ = 56,
            [Description("txtNum2")]
            _חרדות_מולד_ = 57,
            [Description("")]
            _אנרגיות_גרון_ומיןויצירה_ = 58,
            [Description("")]
            _בדיקת_צאקרות_קרמאתי_ = 59,
            [Description("")]
            _בדיקת_זוגיות_צאקרות_קארמתי_ = 60,
            [Description("textBox48")]
            _פרטים_אישיים_ = 61,
            [Description("textBox48")]
            _חתימה_ = 62,
            [Description("textBox48")]
            _תיאור_מקצועי_ = 63,
            [Description("txtCombinations")]
            _שילובים_ = 64,
            [Description("txtFinalSalesMark")]
            _ערך_התאמה_למכירות_ = 65,
            [Description("txtSalesStory")]
            _מלל_התאמה_למכירות_ = 66,
            [Description("txtGroupSalesStory")]
            _מלל_התאמה_קבוצתית_למכירות_ = 67,
            [Description("txtSexMatchMark")]
            _ערך_התאמה_מינית_ = 68,
            [Description("txtSexMatchStory")]
            _מלל_התאמה_מינית_ = 69,
            [Description("")]
            _פרטי_בן_זוג_מינית_ = 70
        }

    }
}
