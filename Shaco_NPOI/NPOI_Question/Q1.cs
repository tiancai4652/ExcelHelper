using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Xmas11.NPOI.Excel;

namespace NPOI_Question
{

    /// <summary>
    /// 使用NPOI对excel进行一些删除表和隐藏表的操作后，excel的引用会出现错乱或报错
    /// </summary>
   public class Q1
    {

        static void AsposeDoHideAndDelete(string sourceFileName, string targetFileName, List<string> HideList, List<string> DeleteList)
        {

            Aspose.Cells.License li = new Aspose.Cells.License();
            string path = Application.StartupPath + "\\" + @"Aspose.Cells.lic";
            li.SetLicense(path);

            Workbook wb = new Workbook(sourceFileName);


            foreach (var sheet in wb.Worksheets)
            {
                System.Diagnostics.Debug.WriteLine(sheet.Name + ":" + sheet.Index);
            }


            System.Diagnostics.Debug.WriteLine("-------------------------------");


            foreach (var hideItem in HideList)
            {
                wb.Worksheets[hideItem].IsVisible = false;
            }
            foreach (var deleteItem in DeleteList)
            {
                wb.Worksheets.RemoveAt(deleteItem);
            }

            wb.Worksheets.ActiveSheetIndex = 0;
            wb.Save(targetFileName);


            foreach (var sheet in wb.Worksheets)
            {
                System.Diagnostics.Debug.WriteLine(sheet.Name + ":" + sheet.Index);
            }
        }
        static void NPOIDoHideAndDelete(string sourceFileName, string targetFileName, List<string> HideList, List<string> DeleteList)
        {
            NPOIExcelApplication excel = new NPOIExcelApplication(sourceFileName);
            excel.SetSheetsHidden(HideList);
            for (int i = 0; i < DeleteList.Count; i++)
            {
                excel.RemoveSheet(DeleteList[i]);
            }
            excel.SetActiveSheet(excel.Workbook.GetSheetAt(0).SheetName);

            excel.WriteToFile(targetFileName);
        }

        public void Run()
        {
            string sourceFileName = Application.StartupPath + "\\Folder\\Q1\\OrderMistake.xls";
            string sourceFileName2 = Application.StartupPath + "\\Folder\\Q1\\OrderMistake2.xls";

            string sourceFileName3 = Application.StartupPath + "\\Folder\\Q1\\OrderMistake3.xls";


            string targetAspose3 = Application.StartupPath + "\\Folder\\Q1\\asposeTarget3.xls";
            AsposeDoHideAndDelete(sourceFileName3, targetAspose3, new List<string>(), new List<string>() { "Sheet2", "Sheet3" });


            string targetNPOI3 = Application.StartupPath + "\\Folder\\Q1\\npoiTarget3.xls";
            NPOIDoHideAndDelete(sourceFileName3, targetNPOI3, new List<string>(), new List<string>() { "Sheet2", "Sheet3" });


            //NPOI
            string targetNPOI = Application.StartupPath + "\\Folder\\Q1\\npoiTarget.xls";
            string targetNPOI2 = Application.StartupPath + "\\Folder\\Q1\\npoiTarget2.xls";
            ///Aspose
            string targetAspose = Application.StartupPath + "\\Folder\\Q1\\asposeTarget.xls";
            string targetAspose2 = Application.StartupPath + "\\Folder\\Q1\\asposeTarget2.xls";


            List<string> HideList = new List<string>() { "HideSheet1", "HideSheet2", "HideSheet3", "HideSheet4", "HideSheet5" };
            List<string> DeleteList = new List<string>() { "Other1", "Other2", "Other3", "Other4", "Other5" };

            NPOIDoHideAndDelete(sourceFileName, targetNPOI, HideList, DeleteList);
            AsposeDoHideAndDelete(sourceFileName, targetAspose, HideList, DeleteList);



            List<string> deleteList2 = new List<string>() { "Certificate_AFAL", "Certificate(Blank)", "Certificate_AFAL(Blank)" };
            List<string> hideList2 = new List<string>() { "BaseInfo", "CalData1", "CalData2", "Initial", "Certificate_Tapped", "Certificate_AFAL_Tapped", "Certificate_Tapped(Blank)", "Certificate_AFAL_Tapped(Blank)" };
            NPOIDoHideAndDelete(sourceFileName2, targetNPOI2, hideList2, deleteList2);
            AsposeDoHideAndDelete(sourceFileName2, targetAspose2, hideList2, deleteList2);

        }
    }
}
