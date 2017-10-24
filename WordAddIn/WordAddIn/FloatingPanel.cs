using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WordAddIn
{
    public partial class FloatingPanel : Form
    {
        private MyBookMark _BookMark = null;
        public FloatingPanel(MyBookMark bookMark)
        {
            InitializeComponent();
            _BookMark = bookMark;
            this.tbTooltip.Text = _BookMark.ToolTip;
        }

        private void lkBtn_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            _BookMark.BookMarkRange.HighlightColorIndex = _BookMark.OrignalColor;
            _BookMark.BookMark.Delete();
            Globals.ThisAddIn._BookMarkList.Remove(_BookMark);
            this.Close();
        }
    }
}
