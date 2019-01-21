using System;
using System.Collections.Generic;
using System.IO;
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
        /// This method will return all the excel sheet names in a string separated by this character --> ¬
        /// </summary>
        /// <param name="pathToExcelFile"></param>
        /// <returns>A string with the Excel sheet names separated by ¬</returns>
        public static string getExcelSheetNames(string pathToExcelFile)
        {


            //I create an instance of a Microsoft Excel Application
            Microsoft.Office.Interop.Excel.Application myExcel = new Microsoft.Office.Interop.Excel.Application();
            //Set visible to false to let the app work "behind".. a.k.a not OPEN the excel application
            myExcel.Visible = false;
            //Using the excel application instance, I open the book with the path to my excel file as parameter
            Microsoft.Office.Interop.Excel.Workbook workbook = myExcel.Workbooks.Open(pathToExcelFile);
            //Later I use a Worksheet Instance to get the actual sheet in the Excel File
            Worksheet worksheet = myExcel.ActiveSheet;

            //This string will hold all the sheet's names
            string excelSheet = "";

            //In this loop I concatenate the worksheet names into the string
            foreach (Worksheet hoja in myExcel.Worksheets)
            {
                excelSheet += hoja.Name.ToString();
                //This is the separator I chose to delimitate the sheet names
                excelSheet += "¬";
            }

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
        /// This method will return the total amount of rows of a given excel that are not null
        /// </summary>
        /// <param name="pathToExcel"></param>
        /// <returns>Integer with the total amount of rows</returns>
        public int GetTotalAmountOfRows(string pathToExcel)
        {
            //Se crea una instancia de una aplicación de Excel
            Microsoft.Office.Interop.Excel.Application myExcel = new Microsoft.Office.Interop.Excel.Application();
            //False para que no abra la aplicación, sino que lo haga "por atrás"
            myExcel.Visible = false;
            //Aquí usando la instancia de Aplicación de excel, abro el libro mandando como parámetro la ruta a mi archivo
            Microsoft.Office.Interop.Excel.Workbook workbook = myExcel.Workbooks.Open(pathToExcel);
            //Después uso una instancia de Worksheet (clase de Interop) para obtener la Hoja actual del archivo Excel
            Worksheet worksheet = myExcel.ActiveSheet;
            //En ese worksheet, en la propiedad de Name, tenemos el nombre de la hoja actual, que mando en el query 1 como parámetro
            //Console.WriteLine("WorkSheet.Name: " + worksheet.Name);


            bool exceptionDetected = false;
            int initialRow = 1;
            int totalRows = 0;

            while (!exceptionDetected)
            {
                try
                {
                    if (worksheet.Cells[initialRow, 1].Value2 != null)
                    {
                        //initialRow++;
                        totalRows++;
                        initialRow++;
                        //Console.WriteLine(worksheet.Cells[initialRow, 1].Value2 + " " + initialRow);
                    }
                    else
                    {

                        break;
                    }
                }
                catch (Exception e)
                {
                    exceptionDetected = true;
                }

            }



            //Al finalizar tu proceso debes cerrar tu workbook

            workbook.Close();
            myExcel.Quit();

            //Con esto de Marshal se libera de manera completa el objeto desde Interop Services, si no haces esto
            //El objeto sigue en memoria, no lo libera C#
            Marshal.FinalReleaseComObject(worksheet);
            Marshal.FinalReleaseComObject(workbook);
            Marshal.FinalReleaseComObject(myExcel);
            GC.Collect();
            GC.WaitForPendingFinalizers();


            return totalRows;

        }

        /// <summary>
        /// This method will return the total amount of columns of a given excel file that are not null
        /// </summary>
        /// <param name="pathToExcel"></param>
        /// <returns>Integer with the total amount of columns</returns>
        public static int GetTotalAmountOfColumns(string pathToExcel)
        {


            //Se crea una instancia de una aplicación de Excel
            Microsoft.Office.Interop.Excel.Application myExcel = new Microsoft.Office.Interop.Excel.Application();
            //False para que no abra la aplicación, sino que lo haga "por atrás"
            myExcel.Visible = false;
            //Aquí usando la instancia de Aplicación de excel, abro el libro mandando como parámetro la ruta a mi archivo
            Microsoft.Office.Interop.Excel.Workbook workbook = myExcel.Workbooks.Open(pathToExcel);
            //Después uso una instancia de Worksheet (clase de Interop) para obtener la Hoja actual del archivo Excel
            Worksheet worksheet = myExcel.ActiveSheet;
            //En ese worksheet, en la propiedad de Name, tenemos el nombre de la hoja actual, que mando en el query 1 como parámetro
            //Console.WriteLine("WorkSheet.Name: " + worksheet.Name);

            string hojaExcel = worksheet.Name;



            bool errorDetected = false;
            int initialColumn = 1;
            int totalColumns = 0;




            while (!errorDetected)
            {
                try
                {
                    if (worksheet.Cells[1, initialColumn].Value2 != null)
                    {
                        //initialRow++;
                        totalColumns++;
                        initialColumn++;
                        //Console.WriteLine(worksheet.Cells[1, initialColumn].Value2 + " " + initialColumn);
                    }
                    else
                    {

                        break;
                    }
                }
                catch (Exception e)
                {
                    errorDetected = true;
                }

            }



            //Al finalizar tu proceso debes cerrar tu workbook

            workbook.Close();
            myExcel.Quit();

            //Con esto de Marshal se libera de manera completa el objeto desde Interop Services, si no haces esto
            //El objeto sigue en memoria, no lo libera C#
            Marshal.FinalReleaseComObject(worksheet);
            Marshal.FinalReleaseComObject(workbook);
            Marshal.FinalReleaseComObject(myExcel);
            GC.Collect();
            GC.WaitForPendingFinalizers();

            return totalColumns;
        }


        /// <summary>
        /// This method returns the name of the active sheet in the excel Document provided
        /// </summary>
        /// <param name="pathToExcelFile"></param>
        /// <returns>String with the name of the Active Sheet</returns>
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



        /// <summary>
        /// This method deletes the Invisible trash from a given excel File
        /// </summary>
        /// <param name="pathToExcelFile"></param>
        /// <returns>String with confirmation Text</returns>
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
        public string ColumnNumberToName(int col_num)
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
