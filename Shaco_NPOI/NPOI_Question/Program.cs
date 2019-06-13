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
    class Program
    {
        static void Main(string[] args)
        {
            string sourceFileName = Application.StartupPath + "\\Folder\\OrderMistake.xls";
            string sourceFileName2 = Application.StartupPath + "\\Folder\\OrderMistake2.xls";


            //NPOI
            string targetNPOI = Application.StartupPath + "\\Folder\\npoiTarget.xls";
            string targetNPOI2 = Application.StartupPath + "\\Folder\\npoiTarget2.xls";
            ///Aspose
            string targetAspose = Application.StartupPath + "\\Folder\\asposeTarget.xls";
            string targetAspose2 = Application.StartupPath + "\\Folder\\asposeTarget2.xls";

            List<string> HideList = new List<string>() { "HideSheet1", "HideSheet2", "HideSheet3", "HideSheet4", "HideSheet5" };
            List<string> DeleteList = new List<string>() { "Other1", "Other2", "Other3", "Other4", "Other5" };

            NPOIDo(sourceFileName, targetNPOI, HideList, DeleteList);
            AsposeDo(sourceFileName, targetAspose, HideList, DeleteList);



            List<string> deleteList2 = new List<string>() { "Certificate_AFAL", "Certificate(Blank)", "Certificate_AFAL(Blank)" };
            List<string> hideList2 = new List<string>() { "BaseInfo", "CalData1", "CalData2","Initial","Certificate_Tapped","Certificate_AFAL_Tapped", "Certificate_Tapped(Blank)","Certificate_AFAL_Tapped(Blank)" };
            NPOIDo(sourceFileName2, targetNPOI2, hideList2, deleteList2);
            AsposeDo(sourceFileName2, targetAspose2, hideList2, deleteList2);





        }

        static void NPOIDo(string sourceFileName, string targetFileName,List<string> HideList, List<string> DeleteList)
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

        static void AsposeDo(string sourceFileName, string targetFileName, List<string> HideList, List<string> DeleteList)
        {
            Aspose.Cells.License li = new Aspose.Cells.License();
            string path = Application.StartupPath + "\\" + @"Aspose.Cells.lic";
            li.SetLicense(path);

            Workbook wb = new Workbook(sourceFileName);
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
        }

    }
}
