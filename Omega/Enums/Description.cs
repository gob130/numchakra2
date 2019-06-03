using System;
using System.Collections.Generic;
using System.Text;

namespace Omega.Enums
{
    class Description : Attribute
    {

        public string Text;

        public Description(string text)
        {

            Text = text;

        }

    }

    
}

