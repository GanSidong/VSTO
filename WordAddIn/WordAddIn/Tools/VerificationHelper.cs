using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace Tools
{
    class VerificationHelper
    {
        static public void textBoxVer(TextBox textBox,Button sureBtn)
        {
            string value = textBox.Text;
            bool isNum = Regex.IsMatch(value, @"^[+-]?\d*[.]?\d*$");

            if (!isNum)
            {
                textBox.BackColor = Macro.verColor;
                sureBtn.Enabled = false;
            }
            else
            {
                textBox.BackColor = Macro.oriColor;
                sureBtn.Enabled = true;
            }
        }
    }
}
