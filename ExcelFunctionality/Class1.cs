using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace ExcelFunctionality
{
    /// <summary>
    /// This class holds key Excel Functionality for Excel Integration 
    /// </summary>
    public class ExcelEasyFunctionality
    {
        
        /// <summary>
        /// This method returns the name of the active sheet in the excel Document provided
        /// </summary>
        /// <param name="pathToExcelFile"></param>
        /// <returns>String with the name of the Active Sheet</returns>
        public string GetExcelActiveSheetName(string pathToExcelFile)
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



        /// <summary>
        /// This method clears the Invisible trash in an Excel file provided
        /// By trash we mean: you write in an excel cell, delete the content of the cell... but the excel
        /// holds some kind of trash that can only be erased by DELETING the cell, that's the purpose of this function
        /// </summary>
        /// <param name="pathToExcelFile"></param>
        /// <returns>String with Confirmation Message: "Invisible Trash Removed"</returns>
        public string EraseExcelInvisibleTrash(string pathToExcelFile)
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

            //variable to Hold the last used row
            int lastUsedRow = 0;
            //variable to hold the last used column
            int lastUsedColumn = 0;


            //Find the real last row used in the excel file
            lastUsedRow = worksheet.Cells.Find("*", System.Reflection.Missing.Value,
                               System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                               Microsoft.Office.Interop.Excel.XlSearchOrder.xlByRows, Microsoft.Office.Interop.Excel.XlSearchDirection.xlPrevious,
                               false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Row;

            // Find the last real column
            lastUsedColumn = worksheet.Cells.Find("*", System.Reflection.Missing.Value,
                                           System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                                           Microsoft.Office.Interop.Excel.XlSearchOrder.xlByColumns, Microsoft.Office.Interop.Excel.XlSearchDirection.xlPrevious,
                                           false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Column;

            //This variable holds the last available row/column in an Excel File: XFD 1048576        
            string lastColumnAvailable = "XFD1048576";

            //with this lines, Im creating both ranges, the horizontal and vertical to basically have the "empty ranges of the excel"
            Microsoft.Office.Interop.Excel.Range horizontalEmptyRange = worksheet.Range[ColumnNumberToName(lastUsedColumn + 1) + "1", lastColumnAvailable];
            Microsoft.Office.Interop.Excel.Range verticalEmptyRange = worksheet.Range["A" + lastUsedRow + 1, lastColumnAvailable];

            //I delete both empty ranges to clear the computational trash the excel files holds, so its "clean" and can be exported to a database without data type conversion problems
            horizontalEmptyRange.Delete(XlDeleteShiftDirection.xlShiftUp);
            verticalEmptyRange.Delete(XlDeleteShiftDirection.xlShiftUp);

            //I change the status of saved to true
            workbook.Saved = true;
            //next i save the workbook
            workbook.Save();

            //We close our workbook at the end of our process
            workbook.Close();
            myExcel.Quit();
            //As a sideNote, Excel InteropServices don't release the object in memory, there's an instance of Excel still running
            //We use Marchal Final Release to release the object and close the Excel Process
            Marshal.FinalReleaseComObject(worksheet);
            Marshal.FinalReleaseComObject(workbook);
            Marshal.FinalReleaseComObject(myExcel);
            string resultado = "Invisible Trash Removed";
            return resultado;
        }



        // Return the column name for this column number.
        private string ColumnNumberToName(int col_num)
        {
            // See if it's out of bounds.
            if (col_num < 1) return "A";

            // Calculate the letters.
            string result = "";
            while (col_num > 0)
            {
                // Get the least significant digit.
                col_num -= 1;
                int digit = col_num % 26;

                // Convert the digit into a letter.
                result = (char)((int)'A' + digit) + result;

                col_num = (int)(col_num / 26);
            }

            return result;
        }

    }
}
