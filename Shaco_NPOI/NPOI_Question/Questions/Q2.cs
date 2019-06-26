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
    /// 跨Excel复制sheet表，NPOI只支持源文件格式为XLS的文件，不支持源文件格式为XLSX的文件
    /// 结论，可以复制源文件为xlsx的表，其中，图片可以完整的复制过来，但是复制后如果原单元格有引用，新的引用变为形如：='G:\GitHub2019\ExcelHelper\Shaco_NPOI\NPOI_Question\bin\Debug\Folder\Q2\[source.xlsx]BaseInfo'!A100
    /// √
    /// </summary>
    public class Q2
    {
        public static void Run()
        {
            string sourceFileName = Application.StartupPath + "\\Folder\\Q2\\source.xlsx";
            string targetFileName1 = Application.StartupPath + "\\Folder\\Q2\\target1_ori.xlsx";
            string targetFileName2 = Application.StartupPath + "\\Folder\\Q2\\target2_ori.xls";

            Workbook wbSource = new Workbook(sourceFileName);
            Workbook wbTarget1 = new Workbook(targetFileName1);
            Workbook wbTarget2 = new Workbook(targetFileName2);

            Worksheet ws = wbSource.Worksheets["Certificate"];

            Worksheet ws1 = wbTarget1.Worksheets[0];
            Worksheet ws2 = wbTarget2.Worksheets[0];

            ws1.Name = ws2.Name= "MySheet";

            ws1.Copy(ws);
            ws2.Copy(ws);

            wbTarget1.Save(targetFileName1.Replace("_ori",""));
            wbTarget2.Save(targetFileName2.Replace("_ori", ""));
        }


    }
}
