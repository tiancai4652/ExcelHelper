using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace NPOI_Question.Questions
{
    /// <summary>
    /// 验证Aspose各种常见格式的图片（尤其是xlsx版本），且能保留透明度
    /// 1.jpg 
    /// 2.png 带透明
    /// 3.bmp
    /// 结论：支持png，jpg，bmp格式，png保留透明
    /// </summary>
    public class Q16
    {
        static string targetPic1 = Application.StartupPath + "\\Folder\\Q16\\2.jpg";
        static string targetPic2 = Application.StartupPath + "\\Folder\\Q16\\1.png";
        static string targetPic3 = Application.StartupPath + "\\Folder\\Q16\\3.bmp";

        static string targetFileName1 = Application.StartupPath + "\\Folder\\Q16\\target1.xlsx";
        static string targetFileName1x = Application.StartupPath + "\\Folder\\Q16\\target1.xls";

        /// <summary>
        /// 根据代码里的数据生成表
        /// </summary>
        public static void Run2()
        {
            //Instantiating a Workbook object
            Workbook workbook = new Workbook();

            //Adding a new worksheet to the Workbook object
            Worksheet worksheet = workbook.Worksheets[0];

            //Insert a string value to a cell
            worksheet.Cells["C2"].Value = "Image";

            //Set the 4th row height
            worksheet.Cells.SetRowHeight(3, 150);
            //Set the C column width
            worksheet.Cells.SetColumnWidth(2, 50);
            //Set the C column width
            worksheet.Cells.SetColumnWidth(3, 50);
            worksheet.Cells.SetColumnWidth(4, 50);

            //Add a picture to the C4 cell
            int index = worksheet.Pictures.Add(3, 2, 4, 3, targetPic1);
            //Add a picture to the C4 cell
            int index2 = worksheet.Pictures.Add(3, 3, 4, 4, targetPic2);
            int index3 = worksheet.Pictures.Add(3, 4, 4, 5, targetPic3);

            //Get the picture object
            //Picture pic = worksheet.getPictures().get(index);
            Picture pic = worksheet.Pictures[index];
            Picture pic2 = worksheet.Pictures[index2];
            Picture pic3 = worksheet.Pictures[index3];

            //Set the placement type
            pic.Placement = PlacementType.FreeFloating;
            pic2.Placement = PlacementType.FreeFloating;
            pic3.Placement = PlacementType.FreeFloating;

            // Save the workbook
            workbook.Save(targetFileName1, Aspose.Cells.SaveFormat.Xlsx);
            workbook.Save(targetFileName1x);
        }
    }
}
