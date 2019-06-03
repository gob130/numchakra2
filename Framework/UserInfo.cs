#region Namespaces

using System;
using System.Collections.Generic;

#endregion Namespaces

// **********

namespace Framework
{
    class UserInfo
    {
        public int ID { get; set; }
        public string FirstName { get; set; }
        public string LastName { get; set; }
        public string MotherName { get; set; }
        public string FatherName { get; set; }
        public DateTime BirthDate { get; set; }
        public string City { get; set; }
        public string Street { get; set; }
        public int HouseNum { get; set; }
        public int ApartmentNum { get; set; }
        public string PhoneNum { get; set; }
        public string EmailAddress { get; set; }
        //public Enums.EnumProvider.Sex MyProperty { get; set; }
        //public Enums.EnumProvider.PassedRectification PassedMainRect { get; set; }
        //public Enums.EnumProvider.ReachedMaster NotReachedToMastery { get; set; }
        public string Profession { get; set; }
        public string HebBirthDate { get; set; }
        public DateTime DateCalcTo { get; set; }
        //public Dictionary<string, string> ChakraValues { get; set; }
        //public Dictionary<int, string> LifeCycles { get; set; }

    }

}

/* Chakras: Enum of chakras and values
 * LifeCycles: Cycle Number and - Enum of cycle type and its value
 */
