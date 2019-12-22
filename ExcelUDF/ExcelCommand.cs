
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using ExcelDna.Integration;

namespace ExcelUDF
{
    public class ExcelCommand
    {

        [ExcelCommand(MenuName = "功能示例", MenuText = "显示版本号",
                      ShortCut = "^1",            //https://msdn.microsoft.com/zh-tw/vba/excel-vba/articles/application-onkey-method-excel
                      Name = "ShowVer")]
        public static void ShowMajorVersion()
        {
            //下面是C API接口代码	
            XlCall.Excel(XlCall.xlcAlert, ExcelDna.Integration.ExcelDnaUtil.ExcelVersion.ToString());
        }

        private static string xllfullname = string.Empty;

        [ExcelCommand(MenuName = "管理XLL", MenuText = "load XLL")]
        public static void LoadXLL()
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Multiselect = true;//该值确定是否可以选择多个文件
            dialog.Title = "请选择文件夹";
            dialog.Filter = "所有文件(*.*)|*.*";
            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                xllfullname = dialog.FileName;
            }
            ExcelIntegration.RegisterXLL(xllfullname);
        }
        [ExcelCommand(MenuName = "管理XLL", MenuText = "Unload XLL")]
        public static void UnloadXLL()
        {
            if (string.IsNullOrEmpty(xllfullname)) return;
            ExcelIntegration.UnregisterXLL(xllfullname);
        }
    }
}
