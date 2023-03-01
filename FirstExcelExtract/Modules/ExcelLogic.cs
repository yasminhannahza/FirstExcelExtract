using FirstExcelExtract.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using ExcelHelperExtension;

namespace FirstExcelExtract
{
    partial class frmFirstExcel
    {
        List<ClientInfo> listClientInfo;

        private async Task ExcelLogic()
        {
            await UpdateConsole("Entering Jeng Jeng Routine...");
            Excel.Application xlApp = new Excel.Application();
            //xlApp.Visible = true;
            Excel.Workbooks workbooks = xlApp.Workbooks;
            Excel.Workbook workbook = workbooks.Open("C:\\Users\\yasmi\\Desktop\\Random excel file.xlsx");
            Excel.Worksheet worksheet = workbook.Sheets[2];


            await UpdateConsole("Extracting columns...");

            int columnQty = 3;

            List<object> columnAFlat = new List<object>();
            List<object> columnBFlat = new List<object>();
            List<object> columnCFlat = new List<object>();

            object[,] columnA = worksheet.Columns[1].Value;
            await UpdateProgressBar(1, columnQty, true);
            columnAFlat = columnA.FlattenArray<object>();

            object[,] columnB = worksheet.Columns[2].Value;
            await UpdateProgressBar(2, columnQty, true);
            columnBFlat = columnB.FlattenArray<object>();

            object[,] columnC = worksheet.Columns[3].Value;
            await UpdateProgressBar(3, columnQty, true);
            columnCFlat = columnC.FlattenArray<object>();


            await UpdateProgressBar(0, 1);


            int firstNullLoc = columnAFlat
                .FindIndex(x => x == null);

            listClientInfo = new List<ClientInfo>();

            for (int i = 1; i < firstNullLoc; i++)
            {
                await UpdateProgressBar(i, firstNullLoc - 1, true);

                ClientInfo theClient = new ClientInfo()
                {
                    Name = columnAFlat[i].ToString(),
                    Gender = columnBFlat[i].ToString(),
                    Age = (double)columnCFlat[i]
                };

                listClientInfo.Add(theClient);
            }

            await UpdateProgressBar(0, 1);
            //object[,] columnA = worksheet.Columns[1].Value;
            //await UpdateProgressBar(1, columnQty, true);

            //object[,] columnB = worksheet.Columns[2].Value;
            //await UpdateProgressBar(2, columnQty, true);

            //object[,] columnC = worksheet.Columns[C].Value;
            //await UpdateProgressBar(3, columnQty, true);

            //foreach (object item in columnA)
            //{
            //    columnAFlat.Add(item);
            //}

            //foreach (object item in columnB)
            //{
            //    columnBFlat.Add(item);
            //}

            //foreach (object item in columnC)
            //{
            //    columnCFlat.Add(item);
            //}

            workbook.Close(SaveChanges: false);
            xlApp.Quit();

            Marshal.ReleaseComObject(worksheet);
            Marshal.ReleaseComObject(workbook);
            Marshal.ReleaseComObject(workbooks);
            Marshal.ReleaseComObject(xlApp);

            await UpdateConsole("Jeng Jeng Routine done!!");
        }
    }
}
