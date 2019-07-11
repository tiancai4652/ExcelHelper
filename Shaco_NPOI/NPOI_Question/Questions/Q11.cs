using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Xmas11.NPOI.Excel;

namespace NPOI_Question.Questions
{
    /// <summary>
    /// NPOI:删除Info隐藏表时抛空引用异常
    /// 过程:是用替换法做的，决定替代后测试
    /// </summary>
    public class Q11
    {
        public static void Run()
        {
            string sourceFileName = Application.StartupPath + "\\Folder\\Q11\\11.xls";
            string targetFileName1 = Application.StartupPath + "\\Folder\\Q11\\target1_ori.xlsx";
            string targetFileName2 = Application.StartupPath + "\\Folder\\Q11\\target2_ori.xlsx";

            //NPOIDodDelete(sourceFileName,)
        }


        static void NPOIDodDelete(string sourceFileName, string targetFileName, List<string> DeleteList)
        {
            NPOIExcelApplication excel = new NPOIExcelApplication(sourceFileName);
            for (int i = 0; i < DeleteList.Count; i++)
            {
                excel.RemoveSheet(DeleteList[i]);
            }
            excel.SetActiveSheet(excel.Workbook.GetSheetAt(0).SheetName);
            excel.WriteToFile(targetFileName);
        }
    }
}
