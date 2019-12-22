using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using ExcelDna.Integration.CustomUI;
using System.Windows.Forms;
using ExcelDna.Integration;
using Excel = Microsoft.Office.Interop.Excel;
using ExcelDna.IntelliSense;
using VB = Microsoft.VisualBasic;
// TODO:   按照以下步骤启用功能区(XML)项: 

// 1. 将以下代码块复制到 ThisAddin、ThisWorkbook 或 ThisDocument 类中。

//  protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
//  {
//      return new Ribbon1();
//  }

// 2. 在此类的“功能区回调”区域中创建回调方法，以处理用户
//    操作(如单击某个按钮)。注意: 如果已经从功能区设计器中导出此功能区，
//    则将事件处理程序中的代码移动到回调方法并修改该代码以用于
//    功能区扩展性(RibbonX)编程模型。

// 3. 向功能区 XML 文件中的控制标记分配特性，以标识代码中的相应回调方法。  

// 有关详细信息，请参见 Visual Studio Tools for Office 帮助中的功能区 XML 文档。


namespace ExcelUDF
{
    [ComVisible(true)]
    public class Ribbon1 : ExcelRibbon
    {
        public void AutoOpen()
        {

        }

        public void AutoClose()
        {
        }

        public static void CalRd()
        {
            try
            {
                Excel.Range selection = AddIn.ExcelApp.Selection;
                decimal[] rng = new decimal[2];
                int i = 0;
                foreach (Excel.Range item in selection)
                {
                    rng[i]= (decimal)item.Value;
                    i++;
                }
                AddIn.ExcelApp.StatusBar = rng[0] + "和"+ rng[1] + "的相对偏差为："+Math.Round(Math.Abs((rng[0] - rng[1]) /(rng[0] + rng[1]) *100),2)+"%";  
            }
            catch (Exception)
            {
                AddIn.ExcelApp.StatusBar = "请选择两个数字……";
            }
            
            
        }

        public static void Round5()
        {
            try
            {
                Excel.Range selection = AddIn.ExcelApp.Selection;
                string i = VB.Interaction.InputBox("请输入要保留的小数位数：","保留位数","2");
                foreach (Excel.Range item in selection)
                {
                    if (item.Value == null)
                    {
                        continue;
                    }
                    item.Value = Math.Round((decimal)item.Value, int.Parse(i));
                }
            }
            catch (Exception)
            {
                AddIn.ExcelApp.StatusBar = "发生错误……";
            }


        }
    }
}
