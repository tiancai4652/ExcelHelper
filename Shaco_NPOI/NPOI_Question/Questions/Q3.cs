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
    /// NPOI:跨workbook复制sheet 出现图片丢失的问题:jpg/png(带透明度)/bmp
    /// 源文件为xlsx格式，目标文件为xls，xlsx格式都可以
    /// 源文件为xls格式，目标文件为xls格式可以，xlsx格式打不开，报错
    /// 关于图片，用董修伟的文件会发现有些艺术字和图形颜色，样式没有了，但是如果用我自己做的excel导出没有问题，怀疑跟谁做的源excel有关
    /// 暂未解决
    /// </summary>
    public class Q3
    {
        public static void Run()
        {
            string sourceFileName = Application.StartupPath + "\\Folder\\Q3\\Cover-picture-inside.xls";
            //string sourceFileName = Application.StartupPath + "\\Folder\\Q3\\1.xls";
            //string sourceFileName = Application.StartupPath + "\\Folder\\Q3\\1.xlsx";
            string targetFileName1 = Application.StartupPath + "\\Folder\\Q3\\target1_ori.xlsx";
            string targetFileName2 = Application.StartupPath + "\\Folder\\Q3\\target2_ori.xls";

            Workbook wbSource = new Workbook(sourceFileName);
            Workbook wbTarget1 = new Workbook(targetFileName1);
            Workbook wbTarget2 = new Workbook(targetFileName2);

            Worksheet ws = wbSource.Worksheets["Sheet1"];

            Worksheet ws1 = wbTarget1.Worksheets[0];
            Worksheet ws2 = wbTarget2.Worksheets[0];

            ws1.Name = ws2.Name = "MySheet";

            ws1.Copy(ws);
            ws2.Copy(ws);

            wbTarget1.Save(targetFileName1.Replace("_ori", ""));
            wbTarget2.Save(targetFileName2.Replace("_ori", ""));
        }
    }
}
