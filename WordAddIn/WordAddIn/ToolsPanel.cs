using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn
{
    public partial class ToolsPanel : UserControl
    {
        // 保存修改过的Range和之前的背景色，以便于恢复
        private Word.Range _LastRange = null;
        private Word.WdColorIndex _LastRangeBackColor = default(Word.WdColorIndex);


        public ToolsPanel()
        {
            InitializeComponent();
        }

        private void btnSearch_Click(object sender, EventArgs e)
        {
            // 清除文档中的高亮显示
            ClearMark();

            lvSearchResult.Items.Clear();
            if (string.IsNullOrWhiteSpace(tbSearchText.Text))
            {
                return;
            }

            // 按段落检索
            Word.Document currentDocument = Globals.ThisAddIn.Application.ActiveDocument;
            if (currentDocument.Paragraphs != null &&
                currentDocument.Paragraphs.Count != 0)
            {
                foreach (Word.Paragraph paragraph in currentDocument.Paragraphs)
                {
                    MatchCollection mc = Regex.Matches(paragraph.Range.Text, tbSearchText.Text.Trim(), RegexOptions.IgnoreCase);
                    if (mc.Count > 0)
                    {
                        foreach (Match m in mc)
                        {
                            try
                            {
                                int startIndex = paragraph.Range.Start + m.Index;
                                int endIndex = paragraph.Range.Start + m.Index + m.Length;

                                Word.Range keywordRange = currentDocument.Range(startIndex, endIndex);

                                // 获取上下文信息
                                // 获取前两个单词的位置（如果有）
                                startIndex = GetStartPositionForView(paragraph, m, startIndex);

                                // 获取后两个单词的位置（如果有）
                                endIndex = GetEndPositionForView(paragraph, m, endIndex);

                                // 在ListView中展示检索的关键字以及其上下文
                                Word.Range range = currentDocument.Range(startIndex, endIndex);
                                ListViewItem item = new ListViewItem(range.Text);
                                item.Tag = keywordRange;
                                lvSearchResult.Items.Add(item);
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show(ex.Message);
                            }
                        }
                    }
                }
            }
        }
        private void lvSearchResult_ItemSelectionChanged(object sender, ListViewItemSelectionChangedEventArgs e)
        {
            ClearMark();
            if (lvSearchResult.SelectedItems.Count > 0)
            {
                Word.Range range = lvSearchResult.SelectedItems[0].Tag as Word.Range;

                // 为了可以恢复被修改的Range，我先将该Range和原本的Color放入Class的成员
                _LastRange = range;
                _LastRangeBackColor = range.HighlightColorIndex;
                range.HighlightColorIndex = Word.WdColorIndex.wdYellow;
            }
        }

        private void ClearMark()
        {
            if (_LastRange != null)
            {
                _LastRange.HighlightColorIndex = _LastRangeBackColor;
            }
        }

        #region 其他方法
        private static int GetEndPositionForView(Word.Paragraph paragraph, Match m, int endIndex)
        {
            string suffixPart = paragraph.Range.Text.Substring(m.Index + m.Length);
            MatchCollection suffixMC = Regex.Matches(suffixPart, "\\s");
            if (suffixMC.Count >= 3)
            {
                endIndex = endIndex + suffixMC[2].Index;
            }
            else
            {
                if (suffixMC.Count >= 2)
                {
                    endIndex = endIndex + suffixMC[1].Index;
                }
                else if (suffixMC.Count >= 1)
                {
                    endIndex = endIndex + suffixMC[0].Index;
                }
            }
            return endIndex;
        }

        private static int GetStartPositionForView(Word.Paragraph paragraph, Match m, int startIndex)
        {
            string prefixPart = paragraph.Range.Text.Substring(0, m.Index);
            MatchCollection preficMC = Regex.Matches(prefixPart, "\\s");
            if (preficMC.Count >= 3)
            {
                startIndex = paragraph.Range.Start + preficMC[preficMC.Count - 3].Index;
            }
            else
            {
                if (preficMC.Count >= 2)
                {
                    startIndex = paragraph.Range.Start + preficMC[preficMC.Count - 2].Index;
                }
                else if (preficMC.Count >= 1)
                {
                    startIndex = paragraph.Range.Start + preficMC[preficMC.Count - 1].Index;
                }
            }
            return startIndex;
        }
        #endregion

        private void button2_Click(object sender, EventArgs e)
        {
            Selection selection = Globals.ThisAddIn.Application.Selection; ;
            Word.Range r = selection.Range;
            r.Text = this.textBox1.Text;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            OpenFileDialog fd = new OpenFileDialog();
            fd.Filter = "txt文件(*.txt)|*.txt;|所有文件|*.*";
            fd.ValidateNames = true;
            fd.CheckPathExists = true;
            fd.CheckFileExists = true;
            if (fd.ShowDialog() == DialogResult.OK)
            {
                string strFileName = fd.FileName;
                String line;
                StreamReader sr = new StreamReader(strFileName, Encoding.Default);
                List<string> strList = new List<string>();
                while ((line = sr.ReadLine()) != null)
                {
                    strList.Add(line);
                }

                for (int i = strList.Count - 1; i >= 0; i--)
                {
                    Selection selection = Globals.ThisAddIn.Application.Selection;
                    Word.Range r = selection.Range;
                    r.Text = strList[i] + Environment.NewLine;
                }

            }

        }

        private void button4_Click(object sender, EventArgs e)
        {
            SaveFileDialog sf = new SaveFileDialog();
            sf.Filter = "txt文件(*.txt)|*.txt;|所有文件|*.*";
            sf.RestoreDirectory = true;
            Word.Range range = Globals.ThisAddIn.Application.Selection.Range;
            StreamWriter myStream;
            if (sf.ShowDialog() == DialogResult.OK)
            {
                string str;
                str = sf.FileName;
                myStream = new StreamWriter(sf.FileName);
                myStream.Write(range.Text);
                myStream.Flush();
                myStream.Close();
            }
        }
    }
}
