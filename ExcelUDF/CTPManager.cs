using System;
using System.Collections.Generic;
using System.Windows.Forms;
using ExcelDna.Integration.CustomUI;

namespace ExcelUDF
{
    //http://www.cnblogs.com/yangecnu/archive/2013/10/18/3375338.html Excel 自定义任务窗格

    //考虑到 Excel2013改成了single document interface (SDI)，因此需要在application事件中处理任务窗格，以保证在当前窗体中能够显示。
    //https://msdn.microsoft.com/en-us/library/office/dn251093(v=office.15).aspx#odc_xl15_ta_ProgrammingtheSDIinExcel2013_TaskPanes

    //http://www.jkp-ads.com/Articles/keepuserformontop02.asp  Keeping Userforms On Top Of SDI Windows In Excel 2013 And Up
    //https://www.add-in-express.com/creating-addins-blog/2013/02/28/excel2013-single-document-interface-task-panes/
    /// <summary>
    /// 任务窗格管理类
    /// </summary>
    internal class CTPManager : IDisposable
    {
        //显式静态构造函数告诉C＃编译器不要将类型标记为beforefieldinit 
        static CTPManager()
        {
            disposed = false;
            Instance = new CTPManager();
        }
        public static CTPManager Instance { get; private set; } = null;

        #region Dispose

        static bool disposed;
        public void Dispose()
        {
            this.Dispose(true);
            GC.SuppressFinalize(this);
        }
        protected virtual void Dispose(bool disposing)
        {
            if (disposed) return;
            if (disposing)
            {
                //TODO:释放那些实现IDisposable接口的托管对象
            }
            //TODO:释放非托管资源，设置对象为null
            disposed = true;
        }
        ~CTPManager()
        {
            this.Dispose(false);
        }
        # endregion Dispose

        //https://jingyan.baidu.com/article/cbcede071f4d9f02f40b4dcf.html
        //https://msdn.microsoft.com/zh-cn/VBA/Office-Shared-VBA/articles/ictpfactory-createctp-method-office
        private Dictionary<string, CustomTaskPane> DicCustomCTP = new Dictionary<string, CustomTaskPane>();
        public void ShowCTP(string hwnd)
        {
            //Office 2013 is SDI(single document interface) 
            if (DicCustomCTP.ContainsKey(hwnd))
            {
                CustomTaskPane ctp = DicCustomCTP[hwnd];
                if (ctp != null) ctp.Visible = true;
            }
            else
            {
                CustomTaskPane ctp = CustomTaskPaneFactory.CreateCustomTaskPane(typeof(CTPControls), "Custom CTP");
                ctp.DockPosition = MsoCTPDockPosition.msoCTPDockPositionRight;
                ctp.DockPositionStateChange += ctp_DockPositionStateChange;
                ctp.VisibleStateChange += ctp_VisibleStateChange;
                ctp.Visible = true;
                DicCustomCTP.Add(hwnd, ctp);
            }
        }
        public void DeleteCTP(string hwnd)
        {
            if (DicCustomCTP.ContainsKey(hwnd))
            {
                CustomTaskPane ctp = DicCustomCTP[hwnd];
                ctp.DockPositionStateChange -= ctp_DockPositionStateChange;
                ctp.VisibleStateChange -= ctp_VisibleStateChange;
                ctp.Delete();
                ctp = null;
                DicCustomCTP.Remove(hwnd);
            }
        }

        void ctp_VisibleStateChange(CustomTaskPane CustomTaskPaneInst)
        {
            //MessageBox.Show("CTP visible: " + CustomTaskPaneInst.Visible);
        }

        void ctp_DockPositionStateChange(CustomTaskPane CustomTaskPaneInst)
        {
            if (CustomTaskPaneInst != null)
            {
                CTPControls control = CustomTaskPaneInst.ContentControl as CTPControls;
                control.TheLabel.Text = "CTP DockPosition: " + CustomTaskPaneInst.DockPosition.ToString();
            }
        }
    }
}
