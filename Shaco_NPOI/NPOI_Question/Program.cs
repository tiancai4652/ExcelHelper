using Aspose.Cells;
using NPOI_Question.Questions;
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

            Aspose.Cells.License li = new Aspose.Cells.License();
            string path = Application.StartupPath + "\\" + @"Aspose.Cells.lic";
            li.SetLicense(path);


            //Q1.Run();
            //Q13.Run1();
            //Q13.Run2();
            Q17.Run();



        }

      

        
    }
}
