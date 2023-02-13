using System;
using Excel = Microsoft.Office.Interop.Excel;

namespace MatrixMultiplication
{
    class Program
    {
        static void Main(string[] args)
        {
            int[,] matrixA = new int[,] { { 1, 2, 3 }, { 4, 5, 6 }, { 7, 8, 9 } };
            int[,] matrixB = new int[,] { { 10, 11, 12 }, { 13, 14, 15 }, { 16, 17, 18 } };
            int[,] resultMatrix = new int[3, 3];

            for (int i = 0; i < matrixA.GetLength(0); i++)
            {
                for (int j = 0; j < matrixB.GetLength(1); j++)
                {
                    int res = 0;

                    for (int k = 0; k < matrixA.GetLength(1); k++)
                    {
                        res += matrixA[i, k] * matrixB[k, j];
                    }

                    resultMatrix[i, j] = res;
                }
            }

            // Create an Excel application instance
            Excel.Application excelApp = new Excel.Application();

            // Create an Excel workbook instance and open it from the predefined location
            Excel.Workbook excelWorkBook = excelApp.Workbooks.Add();

            // Add a new worksheet to workbook with the Datatable name
            Excel.Worksheet excelWorkSheet = (Excel.Worksheet)excelWorkBook.ActiveSheet;
            excelWorkSheet.Name = "Matrix";

            // Add column names
            excelWorkSheet.Cells[1, 1] = "Matrix A";
            excelWorkSheet.Cells[1, 4] = "Matrix B";
            excelWorkSheet.Cells[1, 7] = "Result Matrix";

            // Add row names
            excelWorkSheet.Cells[2, 1] = "Row 1";
            excelWorkSheet.Cells[3, 1] = "Row 2";
            excelWorkSheet.Cells[4, 1] = "Row 3";

            // Add columns
            excelWorkSheet.Cells[2, 2] = matrixA[0, 0];
            excelWorkSheet.Cells[2, 3] = matrixA[0, 1];
            excelWorkSheet.Cells[2, 4] = matrixA[0, 2];
            excelWorkSheet.Cells[3, 2] = matrixA[1, 0];
            excelWorkSheet.Cells[3, 3] = matrixA[1, 1];
            excelWorkSheet.Cells[3, 4] = matrixA[1, 2];
            excelWorkSheet.Cells[4, 2] = matrixA[2, 0];
            excelWorkSheet.Cells[4, 3] = matrixA[2, 1];
            excelWorkSheet.Cells[4, 4] = matrixA[2, 2];

            // Add columns
            excelWorkSheet.Cells[2, 5] = matrixB[0, 0];
            excelWorkSheet.Cells[2, 6] = matrixB[0, 1];
            excelWorkSheet.Cells[2, 7] = matrixB[0, 2];
            excelWorkSheet.Cells[3, 5] = matrixB[1, 0];
            excelWorkSheet.Cells[3, 6] = matrixB[1, 1];
            excelWorkSheet.Cells[3, 7] = matrixB[1, 2];
            excelWorkSheet.Cells[4, 5] = matrixB[2, 0];
            excelWorkSheet.Cells[4, 6] = matrixB[2, 1];
            excelWorkSheet.Cells[4, 7] = matrixB[2, 2];

            // Add columns
            excelWorkSheet.Cells[2, 8] = resultMatrix[0, 0];
            excelWorkSheet.Cells[2, 9] = resultMatrix[0, 1];
            excelWorkSheet.Cells[2, 10] = resultMatrix[0, 2];
            excelWorkSheet.Cells[3, 8] = resultMatrix[1, 0];
            excelWorkSheet.Cells[3, 9] = resultMatrix[1, 1];
            excelWorkSheet.Cells[3, 10] = resultMatrix[1, 2];
            excelWorkSheet.Cells[4, 8] = resultMatrix[2, 0];
            excelWorkSheet.Cells[4, 9] = resultMatrix[2, 1];
            excelWorkSheet.Cells[4, 10] = resultMatrix[2, 2];

            // Save and close
            excelWorkBook.SaveAs(@"D:\matrix.xlsx");
            excelWorkBook.Close();
            excelApp.Quit();
            Console.WriteLine("Excel file created");
        }
    }
}