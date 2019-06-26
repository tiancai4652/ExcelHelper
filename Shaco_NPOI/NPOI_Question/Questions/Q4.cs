using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace NPOI_Question
{
    /// <summary>
    /// NPOI:跨workbook复制sheet后，源sheet正好2页，复制后行高变高导致超过2页
    /// 如果是一个跟源模板不相关的模板，Aspose会变矮变宽
    /// 如果是从源模板继承过来的模板（拷贝过来修改的），就会没问题
    /// 如果手动在excel上移动或拷贝工作簿，也会复现上述两个case
    /// 具体原因未继续深入
    /// </summary>
    public class Q4
    {
        public static void Run()
        {
            string sourceFileName = Application.StartupPath + "\\Folder\\Q4\\4.xls";
            string targetFileName1 = Application.StartupPath + "\\Folder\\Q4\\target1_ori.xlsx";
            string targetFileName2 = Application.StartupPath + "\\Folder\\Q4\\target2_ori.xls";
            string targetFileName3 = Application.StartupPath + "\\Folder\\Q4\\4copy_ori.xls";

            Workbook wbSource = new Workbook(sourceFileName);
            Workbook wbTarget1 = new Workbook(targetFileName1);
            Workbook wbTarget2 = new Workbook(targetFileName2);
            Workbook wbTarget3 = new Workbook(targetFileName3);

            Worksheet ws = wbSource.Worksheets["Sheet1"];

            Worksheet ws1 = wbTarget1.Worksheets[0];
            Worksheet ws2 = wbTarget2.Worksheets[0];
            Worksheet ws3 = wbTarget3.Worksheets[0];

            ws1.Name = ws2.Name = ws3.Name="MySheet";

            ws1.Copy(ws);
            ws2.Copy(ws);
            ws3.Copy(ws);

            wbTarget1.Save(targetFileName1.Replace("_ori", ""));
            wbTarget2.Save(targetFileName2.Replace("_ori", ""));
            wbTarget3.Save(targetFileName3.Replace("_ori", ""));
        }
    }
}
