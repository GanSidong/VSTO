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
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Ribbon;


namespace WordAddIn
{
    public partial class ThisAddIn
    {
        public CustomTaskPane _MyCustomTaskPane = null;
        public List<MyBookMark> _BookMarkList = new List<MyBookMark>();
        public FloatingPanel _FloatingPanel = null;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            initView();
            
        }
        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            try
            {
                RemoveRightBtns();
                this.Application.WindowBeforeRightClick -= new Word.ApplicationEvents4_WindowBeforeRightClickEventHandler(Application_WindowBeforeAddBtnClick);
            }
            catch { }
        }
        private void initView()
        {

            initToolPanel();
            initRightBtn();
            
            
        }

        void Application_WindowActivate(Word.Document Doc, Window Wn)
        {
            CustomTaskPane ctp = MyPaneManager.Instance.GetMyPane();
            
            MyRibbon rm = Globals.Ribbons.GetRibbon<MyRibbon>() as MyRibbon;

            ctp.Visible = rm.tgbMy.Checked;


        }

        private void DocumentChange()
        {
            MyPaneManager.Instance.CollectMyPanes();
        }

        /// <summary>初始化工具面板
        /// </summary>
        private void initToolPanel()
        {  
            ToolsPanel taskPane = new ToolsPanel();
            //_MyCustomTaskPane = this.CustomTaskPanes.Add(taskPane, "我的工具");
           
            MyPaneManager mm = MyPaneManager.Instance;

            _MyCustomTaskPane = mm.GetMyPane();


            _MyCustomTaskPane.Width = 235;
            _MyCustomTaskPane.Visible = true;

        }
        /// <summary> 初始化右键菜单
        /// </summary>
        private void initRightBtn()
        {
            RemoveRightBtns();
            Office.CommandBarControls siteBtns = Application.CommandBars.FindControls(Office.MsoControlType.msoControlButton, missing, "BookMarkAddin", false);
            if (siteBtns != null)
            {
                foreach (Microsoft.Office.Core.CommandBarControl temp_contrl in siteBtns)
                {
                    //如果已经存在就删除
                    if (temp_contrl.Tag == "BookMarkAddin" || temp_contrl.Tag == "checkOutline")
                    {
                        temp_contrl.Delete();
                    }
                }
            }

            // 添加右键按钮
            Office.CommandBarButton addBtn = (Office.CommandBarButton)Application.CommandBars["Text"].Controls.Add(Office.MsoControlType.msoControlButton, missing, missing, missing, false);
            addBtn.BeginGroup = true;
            addBtn.Tag = "BookMarkAddin";
            addBtn.Caption = "查看概要";
            addBtn.Enabled = true;
            this.Application.WindowBeforeRightClick += new Word.ApplicationEvents4_WindowBeforeRightClickEventHandler(Application_WindowBeforeAddBtnClick);

            //// 添加右键按钮
            //Office.CommandBarButton checkBtn = (Office.CommandBarButton)Application.CommandBars["Text"].Controls.Add(Office.MsoControlType.msoControlButton, missing, missing, missing, false);
            //checkBtn.BeginGroup = true;
            //checkBtn.Tag = "checkOutline";
            //checkBtn.Caption = "查看概要";
            //checkBtn.Enabled = true;
            //checkBtn.Click += checkBtn_Click;
           
        }
        private void Application_WindowBeforeAddBtnClick(Word.Selection Sel, ref bool Cancel)
        {
            //这里没有通过全局变量来控制右键菜单里面的按钮，而是通过findcontrol来取得按钮
            //因为这里的VSTO和COM对象处理有问题，使用全局变量来控制右键按钮不稳定
            Office.CommandBarButton addBtn = (Office.CommandBarButton)Application.CommandBars.FindControl(Office.MsoControlType.msoControlButton, missing, "BookMarkAddin", false);
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
            //System.Drawing.Point currentPos = GetPositionForShowing(this.Application.Selection);
            //AddBookMarkForm form = new AddBookMarkForm(this.Application.Selection.Range);
            //form.Location = currentPos;
            //form.ShowDialog();
            (_MyCustomTaskPane.Control as ToolsPanel).textBox2.Text = "ddddd";

        }
        private static System.Drawing.Point GetPositionForShowing(Word.Selection Sel)
        {
            int left = 0;
            int top = 0;
            int width = 0;
            int height = 0;
            Globals.ThisAddIn.Application.ActiveDocument.ActiveWindow.GetPoint(out left, out top, out width, out height, Sel.Range);

            System.Drawing.Point currentPos = new System.Drawing.Point(left, top);
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
        private void RemoveRightBtns()
        {
            Office.CommandBarControls siteBtns = Application.CommandBars.FindControls(Office.MsoControlType.msoControlButton, missing, "BookMarkAddin", false);
            if (siteBtns ==null)
            {
                return;
            }
            foreach (Microsoft.Office.Core.CommandBarControl temp_contrl in siteBtns)
            {
                //如果已经存在就删除
                if (temp_contrl.Tag == "BookMarkAddin" ||temp_contrl.Tag == "checkOutline")
                {
                    temp_contrl.Delete();
                }
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
            this.Application.WindowActivate += Application_WindowActivate;
            this.Application.DocumentChange += new Word.ApplicationEvents4_DocumentChangeEventHandler(DocumentChange);
        }

        #endregion

        #region MyPaneManager
        public class MyPaneManager
        {
            private MyPaneManager()
            {
                m_MyPanes = Globals.ThisAddIn.CustomTaskPanes;
            }
            private static MyPaneManager m_Instance = new MyPaneManager();
            public static MyPaneManager Instance
            {
                get
                {
                    return m_Instance;
                }
            }
            
            private CustomTaskPaneCollection m_MyPanes;
            public CustomTaskPane GetMyPane()
            {
                Word.Window window = Globals.ThisAddIn.Application.ActiveWindow;
                foreach (CustomTaskPane ctp in m_MyPanes)
                {
                    if (ctp.Title == "我的工具" && ctp.Window == window)
                    {
                        return ctp;
                    }
                }
                return m_MyPanes.Add(new ToolsPanel(), "我的工具", window);
            }
            public void CollectMyPanes()
            {
                for (int i = m_MyPanes.Count - 1; i >= 0; i--)
                {
                    if (m_MyPanes[i].Window == null)
                    {
                        m_MyPanes.RemoveAt(i);
                    }
                }
            }
        }
        #endregion
    }
}
