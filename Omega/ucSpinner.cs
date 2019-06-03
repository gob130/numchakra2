using System.Windows.Forms;

namespace Omega
{
    public partial class ucSpinner : UserControl
    {
        public ucSpinner(bool initState = false, string spinnerText = "")
        {
            InitializeComponent();
            lblSpinnerText.Text = spinnerText;
            lblSpinnerText.Visible = initState;
        }

    }

}
