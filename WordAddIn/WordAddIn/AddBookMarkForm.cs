using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn
{
   
    public partial class AddBookMarkForm : Form
    {
        private Word.Range _Range;
        public AddBookMarkForm(Word.Range range)
        {
            InitializeComponent();
            _Range = range;
        }

        private void sure_Click(object sender, EventArgs e)
        {
            Word.Bookmark mark = _Range.Bookmarks.Add("VSTOBookMark" + DateTime.Now.Ticks.ToString(), _Range);

            MyBookMark bookMark = new MyBookMark()
            {
                ToolTip = tbToolTip.Text.Trim(),
                BookMarkRange = _Range,
                IsHighLighted = chHighlightColor.Checked,
                OrignalColor = _Range.HighlightColorIndex,
                BookMark = mark
            };

            Globals.ThisAddIn._BookMarkList.Add(bookMark);

            if (chHighlightColor.Checked)
            {
                _Range.HighlightColorIndex = Word.WdColorIndex.wdYellow;
            }

            this.Close();
        }

        private void cancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
