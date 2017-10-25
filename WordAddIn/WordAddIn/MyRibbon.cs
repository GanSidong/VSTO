﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using System.Windows.Forms;
using Microsoft.Office.Tools;

namespace WordAddIn
{
    public partial class MyRibbon
    {
        private void Ribbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            CustomTaskPane ctp = WordAddIn.ThisAddIn.MyPaneManager.Instance.GetMyPane();
            ctp.VisibleChanged += MypanelsVisableChange;
            ctp.Visible = this.tgbMy.Checked;
        }

        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            if (Globals.ThisAddIn._MyCustomTaskPane != null)
            {
                Globals.ThisAddIn._MyCustomTaskPane.Visible = true;
            }
        }

        private void button3_Click(object sender, RibbonControlEventArgs e)
        {
            if (Globals.ThisAddIn._MyCustomTaskPane != null)
            {
                Globals.ThisAddIn._MyCustomTaskPane.Visible = false;
            }
        }

        private void MypanelsVisableChange(object sender,EventArgs e)
        {
            CustomTaskPane ctp = sender as CustomTaskPane;
            this.tgbMy.Checked = ctp.Visible;
            if (ctp.Visible)
            {
                tgbMy.Label = "收起";
            }
            else
            {
                tgbMy.Label = "展开";
            }

           
        }

        private void tgbMy_Click(object sender, RibbonControlEventArgs e)
        {
            CustomTaskPane ctp = WordAddIn.ThisAddIn.MyPaneManager.Instance.GetMyPane();
            ctp.VisibleChanged += MypanelsVisableChange;
            ctp.Visible = this.tgbMy.Checked;            
        }
    }
}
