using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using System.IO;
using System.Data;
using System.Globalization;
using System.Reflection;

namespace PruebasExcelFunctions
{
    class ProyectoPruebaExcel
    {
        static void Main(string[] args)
        {

            string pathToExcel = @"C:\Users\oscarsanchez2\Desktop\Input.xlsx";




            Console.WriteLine("Columnas: " + GetTotalAmountOfColumns(pathToExcel));
            Console.WriteLine("Filas: " + GetTotalAmountOfRows(pathToExcel));

            Console.WriteLine("Columna a letra: " + ColumnNumberToName(GetTotalAmountOfColumns(pathToExcel)));


            Console.ReadLine();

        }


        public static int GetTotalAmountOfRows(string pathToExcel)
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


        public static string EraseExcelCells(string pathToExcelFile)
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

            string excelSheet = worksheet.Name;
            //I get the starting row (that will always be 1
            int startRow = worksheet.UsedRange.Row;
            //The end row (It will allow me to know the total rows the excel got)
            int endRow = startRow + worksheet.UsedRange.Rows.Count - 1;
            //This rowCount says How many rows the excel has (basically the same value as EndRow)
            int rowCount = worksheet.UsedRange.Rows.Count;
            int lastUsedRow = 0;
            int lastUsedColumn = 0;

            lastUsedRow = worksheet.Cells.Find("*", System.Reflection.Missing.Value,
                               System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                               Microsoft.Office.Interop.Excel.XlSearchOrder.xlByRows, Microsoft.Office.Interop.Excel.XlSearchDirection.xlPrevious,
                               false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Row;

            // Find the last real column
            lastUsedColumn = worksheet.Cells.Find("*", System.Reflection.Missing.Value,
                                           System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                                           Microsoft.Office.Interop.Excel.XlSearchOrder.xlByColumns, Microsoft.Office.Interop.Excel.XlSearchDirection.xlPrevious,
                                           false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Column;

            Console.WriteLine("Last used Row: " + lastUsedRow);
            Console.WriteLine("Last used Column: " + lastUsedColumn);

            int columnCount = worksheet.UsedRange.Columns.Count;

            Console.WriteLine("Ultima Columna (letra):" + ColumnNumberToName(columnCount + 1));
            Console.WriteLine("Ultima Columna (letra):" + ColumnNumberToName(lastUsedColumn + 1));

            string lastColumnAvailable = "XFD1048576";

            Microsoft.Office.Interop.Excel.Range horizontalEmptyRange = worksheet.Range[ColumnNumberToName(lastUsedColumn + 1) + "1", lastColumnAvailable];
            Microsoft.Office.Interop.Excel.Range verticalEmptyRange = worksheet.Range["A" + lastUsedRow + 1, lastColumnAvailable];

            horizontalEmptyRange.Delete(XlDeleteShiftDirection.xlShiftUp);
            verticalEmptyRange.Delete(XlDeleteShiftDirection.xlShiftUp);


            Console.WriteLine("Total de filas: " + rowCount);
            Console.WriteLine("Total de Columnas: " + columnCount);

            Console.WriteLine("Fila fin: " + endRow);

            workbook.Saved = true;

            workbook.Save();

            //We close our workbook at the end of our process
            workbook.Close();
            myExcel.Quit();
            //As a sideNote, Excel InteropServices don't release the object in memory, there's an instance of Excel still running
            //We use Marchal Final Release to release the object and close the Excel Process
            Marshal.FinalReleaseComObject(worksheet);
            Marshal.FinalReleaseComObject(workbook);
            Marshal.FinalReleaseComObject(myExcel);
            string resultado = "fin";
            return resultado;
        }



        // Return the column name for this column number.
        public static string ColumnNumberToName(int col_num)
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

        public static string getExcelActiveSheetName(string pathToExcelFile)
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

            string excelSheet = worksheet.Name;

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





