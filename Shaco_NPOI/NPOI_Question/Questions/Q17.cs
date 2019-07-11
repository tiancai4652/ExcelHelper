using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace NPOI_Question.Questions
{
    /// <summary>
    /// NPOI:拷贝WPS做的封面模板以后，列宽信息丢失
    /// 结论:like Q4，不会出现宽度丢失的情况
    /// </summary>
    public static class Q17
    {
        public static void Run()
        {
            string sourceFileName = Application.StartupPath + "\\Folder\\Q17\\封面.xls";
            string targetFileName = Application.StartupPath + "\\Folder\\Q17\\target.xls";
       
            Workbook wbSource = new Workbook(sourceFileName);
            Workbook wbTarget = new Workbook();

            Worksheet ws = wbSource.Worksheets["证书封面"];
            Worksheet ws1 = wbTarget.Worksheets[0];
            ws1.Name = "MySheet";
            ws1.Copy(ws);


            wbTarget.Save(targetFileName);
        }
    }
}
