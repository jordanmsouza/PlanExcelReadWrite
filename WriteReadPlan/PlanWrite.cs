using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace WriteReadPlan
{
    public class PlanWrite
    {
        public static void writePlan()
        {
            _Application exc = new Microsoft.Office.Interop.Excel.Application();

            Workbook excWb;
            Worksheet excWs;

            // show the FolderBrowserDialog

            exc.Visible = true;
            excWb = exc.Workbooks.Open(@"C:\Users\Public\Downloads\TesteCsharp.xlsx");
            excWs = excWb.Worksheets[1];
            excWs.Activate();

            for (int i = 1; i <=2; i++)
            {
                var coluna = excWs.Columns[i].Value;
                if (coluna[1,1] == "Idbdc")
                {
                    excWs.Cells[2, 1] = 102020;
                }
                else if (coluna[1,1] == "CGI")
                {
                    excWs.Cells[2, 2] = 9800000;
                }
            }

            //excWs.Cells[2, 2] = 102020;
            //excWs.Cells[2, 1] = 9800000;

            Console.WriteLine(excWs.Cells[2, 1].Value);
            Console.WriteLine(excWs.Cells[2, 1].Value);


            //foreach (var item in excWs.Columns.Value)
            //{
            //    if (item == "Idbdc")
            //    {
            //        excWs.Cells[2, 2] = 102020;
            //    }
            //    else if (item == "CGI")
            //    {
            //        excWs.Cells[2, 1] = 9800000;
            //    }

            //}
            excWb.Close(true, "TesteCsharpAtualizado");
            //excWb.Save();

        }
    }
}
