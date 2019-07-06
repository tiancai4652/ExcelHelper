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
    /// NPOI:某些模板生成证书后出现样式错乱，整个Sheet都被设置成黑色边框网格
    /// 即:NPOI跨Workbook复制Sheet，NPOI计算公式，NPOI删除无用表，NPOI隐藏表
    /// 测试内容:NPOI最有可能跨workbook复制sheet出问题，其他项目需要等Xmals.Aspose出来才能验证
    /// 测试结果:跨workbook复制封面没有问题
    /// </summary>
    public class Q9
    {
        public static void Run()
        {
            string sourceFileName = Application.StartupPath + "\\Folder\\Q9\\9.xls";
            string targetFileName = Application.StartupPath + "\\Folder\\Q9\\target_ori.xlsx";


            Workbook wbSource = new Workbook(sourceFileName);
            Workbook wbTarget = new Workbook(targetFileName);


            Worksheet ws = wbSource.Worksheets["封面"];

            Worksheet ws1 = wbTarget.Worksheets[0];


            ws1.Name = "MySheet";

            ws1.Copy(ws);

            wbTarget.Save(targetFileName.Replace("_ori", ""));

        }
    }
}
