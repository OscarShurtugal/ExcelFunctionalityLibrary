using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace ExcelFunctionality
{
    public class ExcelEasyFunctionality
    {
        public string getExcelActiveSheetName(string pathToExcelFile)
        {


            //I create an instance of a Microsoft Excel Application
            Microsoft.Office.Interop.Excel.Application myExcel = new Microsoft.Office.Interop.Excel.Application();
            //Set visible to false to let the app work "behind".. a.k.a not OPEN the excel application
            myExcel.Visible = false;
            //Using the excel application instance, I open the book with the path to my excel file as parameter
            Microsoft.Office.Interop.Excel.Workbook workbook = myExcel.Workbooks.Open(pathToExcelFile);
            //Later I use a Worksheet Instance to get the actual sheet in the Excel File
            Worksheet worksheet = myExcel.ActiveSheet;
            //In this worksheet, in the Name property we can find the name of the Excel active Worksheet
            
            string excelSheet= worksheet.Name;

            //We close our workbook at the end of our process
            workbook.Close();
            
            //As a sideNote, Excel InteropServices don't release the object in memory, there's an instance of Excel still running
            //We use Marchal Final Release to release the object and close the Excel Process
            Marshal.FinalReleaseComObject(worksheet);
            Marshal.FinalReleaseComObject(workbook);
            Marshal.FinalReleaseComObject(myExcel);

            return excelSheet;

        }

    }
}
