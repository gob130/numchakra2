using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;

namespace Omega.Objects
{
    public class InputProvider
    {
        public static InputProvider Instance = new InputProvider();
        private object mValue;

        private frmGeneralInput frmInput;

        private InputProvider()
        {
            frmInput = new frmGeneralInput("");
        }

        public void GetValueFromUser(string title)
        {
            frmInput = new frmGeneralInput(title);
            frmInput.Show();
        }

        public object Value
        {
            get
            {
                return mValue;
            }
            set
            {
                mValue = (object)value;
            }
        }

    }
}
