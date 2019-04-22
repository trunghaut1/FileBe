using System.Windows.Controls;

namespace NumericUpDownLib
{
    public class NumericUpDown2 : NumericUpDown
    {
        public new TextBox _PART_TextBox
        {
            get
            {
                return base._PART_TextBox;
            }
            set
            {
                base._PART_TextBox = value;
            }
        }
    }
}
