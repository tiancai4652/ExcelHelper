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
    ///  NPOI:单元格公式引用了外部Excel中某个单元格的内容，公式无法计算
    ///  C5有公式
    ///  结论:Aspose单元格公式引用了外部Excel中某个单元格的内容不报错，可以计算其他公式
    /// </summary>
    public class Q7
    {
        public static void Run()
        {   
            string targetFileName1 = Application.StartupPath + "\\Folder\\Q7\\7_ori.xlsx";
            string targetFileName2 = Application.StartupPath + "\\Folder\\Q7\\7_ori.xls";

            Workbook wbTarget1 = new Workbook(targetFileName1);
            Workbook wbTarget2 = new Workbook(targetFileName2);

            wbTarget1.CalculateFormula();
            wbTarget2.CalculateFormula();

            Cells cells = wbTarget2.Worksheets[0].Cells;
            Cell cell = cells["C5"];
            ///Formula "=SUM(A5,B5)"
            ///Value 2
            var x = cell.Value;

            wbTarget1.Save(targetFileName1.Replace("_ori", ""));
            wbTarget2.Save(targetFileName2.Replace("_ori", ""));
        }
    }
}
