using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using System.Windows.Forms;

namespace WordAddIn
{
    public partial class Ribbon
    {
        private void Ribbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            MessageBox.Show("hello world");
        }

        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            //if (Globals.ThisAddIn._MyCustomTaskPane != null)
            //{
                Globals.ThisAddIn._MyCustomTaskPane.Visible = true;
            //}
        }

        private void button3_Click(object sender, RibbonControlEventArgs e)
        {
            //if (Globals.ThisAddIn._MyCustomTaskPane != null)
            //{
                Globals.ThisAddIn._MyCustomTaskPane.Visible = false;
            //}
        }
    }
}
