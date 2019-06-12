using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xunit;

namespace Shaco_Aspose
{
    public class Test
    {
        /// <summary>
        /// 注册
        /// </summary>
        static Test()
        {
            Aspose.Cells.License li = new Aspose.Cells.License();
            string path = @"C:\Users\zr644\Desktop\Test" + "\\" + @"Aspose.Cells.lic";
            li.SetLicense(path);
        }

        public static string fileName = @"C:\Users\zr644\Desktop\Test\1.xls";
        AsposeExcelApplication workbook = new AsposeExcelApplication(fileName);

        [Trait("Workbook", "OpenExcelFile")]
        [Theory]
        [InlineData(@"C:\Users\zr644\Desktop\Test\1.xls")]
        [InlineData(@"C:\Users\zr644\Desktop\Test\1.xlsx")]
        public void TestCreatExcel(string fileName)
        {
            AsposeExcelApplication x = new AsposeExcelApplication(fileName);
            int count = x.Workbook.Worksheets.Count;
            Assert.True(count > 0);
        }

        [Trait("Cell", "SetValue")]
        [Theory]
        [InlineData("Sheet1", 0, 0, "string")]
        [InlineData("Sheet1", 1, 1, 1)]
        [InlineData("Sheet1", 2, 2, 2.2)]
        [InlineData("Sheet1", 3, 3, "datetime")]
        public void TestSetValueInCell(string sheetNmae, int rowIndex, int ColumnIndex, object value)
        {
            AsposeExcelApplication aea = new AsposeExcelApplication(fileName);
            if (value.Equals("datetime"))
            {
                var dateValue = DateTime.Now;
                aea.SetCellValue(sheetNmae, rowIndex, ColumnIndex, dateValue);
                Assert.Equal(dateValue.ToShortDateString(), aea.GetCellValue(sheetNmae, rowIndex, ColumnIndex).ToString());
            }
            else
            {
                aea.SetCellValue(sheetNmae, rowIndex, ColumnIndex, value);
                Assert.Equal(value.ToString(), aea.GetCellValue(sheetNmae, rowIndex, ColumnIndex).ToString());
            }
        }

        [Trait("Workbook", "SetWorkbookDocementProperties")]
        [Fact]
        public void TestSetWorkbookProperties()
        {
            workbook.SetWorkbookDocementProperties("Author", "金庸");
            Assert.Equal("金庸", workbook.GetWorkbookDocementProperties("Author").ToString());
            workbook.Save();
        }


        [Trait("Workbook", "SetWorkbookDocCustomProperties")]
        [Fact]
        public void TestSetWorkbookCustomProperties()
        {
            workbook.SetWorkbookDocCustomProperties("CustomProperty1", "浮萍漂泊本无根");
            Assert.Equal("浮萍漂泊本无根", workbook.GetWorkbookDocCustomProperties("CustomProperty1").ToString());
            workbook.Save();
        }

        [Trait("Worksheet", "CreateSheet")]
        [Fact]
        public void TestCreateSheet()
        {
            string name = Guid.NewGuid().ToString("N").Substring(0,10);
            workbook.CreateSheet(name);
            workbook.Save();
            string path = workbook.Workbook.FileName;
            workbook.Workbook.Dispose();
            AsposeExcelApplication workbooktemp = new AsposeExcelApplication(path);
            bool x = workbooktemp.Workbook.Worksheets.Find(t=>t.Name.Equals(name))!=null;
            Assert.True(x);
        }
    }
}
