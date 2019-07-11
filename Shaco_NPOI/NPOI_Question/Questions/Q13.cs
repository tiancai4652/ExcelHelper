using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace NPOI_Question.Questions
{
    /// <summary>
    /// NPOIExcelReport 类的 CreateScatterChart() 方法，绘制散点图（目前仅 *.xlsx 格式的 Excel 文件才支持）
    /// Aspose测试结论:目前只能根据excel数据绘制表，可以绘制后将表保存成图片再重新插入，问题不大.
    /// Aspose两种格式都支持
    /// </summary>
    public class Q13
    {
        static string targetFileName1 = Application.StartupPath + "\\Folder\\Q13\\target1.xlsx";
        static string targetFileName1x = Application.StartupPath + "\\Folder\\Q13\\target1.xls";
        static string targetPic1 = Application.StartupPath + "\\Folder\\Q13\\pic.jpg";
        static string targetFileName2 = Application.StartupPath + "\\Folder\\Q13\\target2.xlsx";
        static string targetFileName2x = Application.StartupPath + "\\Folder\\Q13\\target2.xls";

        /// <summary>
        /// 根据excel数据生成表
        /// </summary>
        public static void Run1()
        {
            // ExStart:1
            // Instantiate a workbook
            Workbook workbook = new Workbook();

            // Access first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Set columns title 
            worksheet.Cells[0, 0].Value = "X";
            worksheet.Cells[0, 1].Value = "Y";

            // Random data shall be used for generating the chart
            Random R = new Random();

            // Create random data and save in the cells
            for (int i = 1; i < 21; i++)
            {
                worksheet.Cells[i, 0].Value = i;
                worksheet.Cells[i, 1].Value = 0.8;
            }

            for (int i = 21; i < 41; i++)
            {
                worksheet.Cells[i, 0].Value = i - 20;
                worksheet.Cells[i, 1].Value = 0.9;
            }
            // Add a chart to the worksheet
            int idx = worksheet.Charts.Add(ChartType.LineWithDataMarkers, 1, 3, 20, 20);

            // Access the newly created chart
            Chart chart = worksheet.Charts[idx];

            // Set chart style
            chart.Style = 3;

            // Set autoscaling value to true
            chart.AutoScaling = true;

            // Set foreground color white
            chart.PlotArea.Area.ForegroundColor = Color.White;

            // Set Properties of chart title
            chart.Title.Text = "Sample Chart";

            // Set chart type
            chart.Type = ChartType.LineWithDataMarkers;

            // Set Properties of categoryaxis title
            chart.CategoryAxis.Title.Text = "Units";

            //Set Properties of nseries
            int s2_idx = chart.NSeries.Add("A2: A2", true);
            int s3_idx = chart.NSeries.Add("A22: A22", true);

            // Set IsColorVaried to true for varied points color
            chart.NSeries.IsColorVaried = true;

            // Set properties of background area and series markers
            chart.NSeries[s2_idx].Area.Formatting = FormattingType.Custom;
            chart.NSeries[s2_idx].Marker.Area.ForegroundColor = Color.Yellow;
            chart.NSeries[s2_idx].Marker.Border.IsVisible = false;

            // Set X and Y values of series chart
            chart.NSeries[s2_idx].XValues = "A2: A21";
            chart.NSeries[s2_idx].Values = "B2: B21";

            // Set properties of background area and series markers
            chart.NSeries[s3_idx].Area.Formatting = FormattingType.Custom;
            chart.NSeries[s3_idx].Marker.Area.ForegroundColor = Color.Green;
            chart.NSeries[s3_idx].Marker.Border.IsVisible = false;

            // Set X and Y values of series chart
            chart.NSeries[s3_idx].XValues = "A22: A41";
            chart.NSeries[s3_idx].Values = "B22: B41";

            var m = chart.ToImage();
            m.Save(targetPic1, ImageFormat.Jpeg);

            // Save the workbook
            workbook.Save(targetFileName1, Aspose.Cells.SaveFormat.Xlsx);
            workbook.Save(targetFileName1x);
        }

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

            //Add a picture to the C4 cell
            int index = worksheet.Pictures.Add(3, 2, 4, 3, targetPic1);

            //Get the picture object
            //Picture pic = worksheet.getPictures().get(index);
            Picture pic = worksheet.Pictures[index];

            //Set the placement type
            //pic.Placement = PlacementType.FreeFloating;

            // Save the workbook
            workbook.Save(targetFileName2, Aspose.Cells.SaveFormat.Xlsx);
            workbook.Save(targetFileName2x);
        }
    }
}
