using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;
using Microsoft.Office.Tools;
using System.Windows.Forms;
using System.Drawing;

namespace WordAddIn
{
    public partial class ThisAddIn
    {
        public CustomTaskPane _MyCustomTaskPane = null;
        public List<MyBookMark> _BookMarkList = new List<MyBookMark>();
        public FloatingPanel _FloatingPanel = null;
        private Office.CommandBarButton addBtn = null;
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            initView();   
        }
        private void initView()
        {
            RemoveRightBtns();
            UserControl1 taskPane = new UserControl1();
            _MyCustomTaskPane = this.CustomTaskPanes.Add(taskPane, "My Task Pane");
            _MyCustomTaskPane.Width = 200;
            _MyCustomTaskPane.Visible = true;
            
            addBtn = (Office.CommandBarButton)Application.CommandBars["Text"].Controls.Add(Office.MsoControlType.msoControlButton, missing, missing, missing, false);

            // 开始一个新Group，即在我们添加的Menu前加一条分割线   
            addBtn.BeginGroup = true;
            // 为按钮设置Tag
            addBtn.Tag = "BookMarkAddin";
            // 添加按钮上的文字
            addBtn.Caption = "Add Bookmark";
            // 将按钮初始设为不激活状态
            addBtn.Enabled = false;
            this.Application.WindowBeforeRightClick += new Word.ApplicationEvents4_WindowBeforeRightClickEventHandler(Application_WindowBeforeRightClick);
            this.Application.WindowSelectionChange += new Word.ApplicationEvents4_WindowSelectionChangeEventHandler(Application_WindowSelectionChange);
        }

        private void Application_WindowSelectionChange(Word.Selection Sel)
        {
            MyBookMark bookmark = _BookMarkList.FirstOrDefault(mark => Sel.Range.Start >= mark.BookMarkRange.Start && Sel.Range.Start <= mark.BookMarkRange.End);
            if (bookmark != null)
            {
                Globals.ThisAddIn._FloatingPanel = new FloatingPanel(bookmark);

                Point currentPos = GetPositionForShowing(Sel);

                Globals.ThisAddIn._FloatingPanel.Location = currentPos;
                Globals.ThisAddIn._FloatingPanel.Show();
            }
        }

        private void Application_WindowBeforeRightClick(Word.Selection Sel, ref bool Cancel)
        {
            // 根据之前添加的Tag来找到我们添加的右键菜单
            // 注意：我这里没有通过全局变量来控制右键菜单，而是通过findcontrol来取得按钮，因为这里的VSTO和COM对象处理有问题，使用全局变量来控制右键按钮不稳定
           
            addBtn.Enabled = false;
            addBtn.Click -= new Office._CommandBarButtonEvents_ClickEventHandler(_RightBtn_Click);

            if (!string.IsNullOrWhiteSpace(Sel.Range.Text) && Sel.Range.Text.Length > 2)
            {
                addBtn.Enabled = true;

                // 这里是另外一个注意点，每次Click事件都需要重新绑定，你需要在之前先取消绑定。
                addBtn.Click += new Office._CommandBarButtonEvents_ClickEventHandler(_RightBtn_Click);
            }
        }

        private void _RightBtn_Click(Office.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            Point currentPos = GetPositionForShowing(this.Application.Selection);
            AddBookMarkForm form = new AddBookMarkForm(this.Application.Selection.Range);
            form.Location = currentPos;

            form.ShowDialog();
        }
        private static Point GetPositionForShowing(Word.Selection Sel)
        {
            int left = 0;
            int top = 0;
            int width = 0;
            int height = 0;
            Globals.ThisAddIn.Application.ActiveDocument.ActiveWindow.GetPoint(out left, out top, out width, out height, Sel.Range);

            Point currentPos = new Point(left, top);
            if (Screen.PrimaryScreen.Bounds.Height - top > 340)
            {
                currentPos.Y += 20;
            }
            else
            {
                currentPos.Y -= 320;
            }
            return currentPos;
        }
        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            try
            {
                RemoveRightBtns();

                this.Application.WindowBeforeRightClick -= new Word.ApplicationEvents4_WindowBeforeRightClickEventHandler(Application_WindowBeforeRightClick);
                this.Application.WindowSelectionChange -= new Word.ApplicationEvents4_WindowSelectionChangeEventHandler(Application_WindowSelectionChange);
            }
            catch { }
        }

        private void RemoveRightBtns()
        {
            Office.CommandBarControls siteBtns = Application.CommandBars.FindControls(Office.MsoControlType.msoControlButton, missing, "BookMarkAddin", false);
            if (siteBtns != null)
            {
                foreach (Office.CommandBarControl btn in siteBtns)
                {
                    btn.Delete(true);
                }
            }
            if (addBtn != null)
            {
                addBtn.Delete(true);
            }

        }

        #region VSTO 生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
