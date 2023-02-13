using System;
using System.Linq;
using System.IO;
using System.Data;
using System.Data.OleDb;

namespace MatrixMultiplication
{
    class Program
    {

        static void Main(string[] args)
        {

            // Создание двух матриц
            int[,] matrix1 = new int[3, 3] { { 1, 2, 3 }, { 4, 5, 6 }, { 7, 8, 9 } };
            int[,] matrix2 = new int[3, 3] { { 9, 8, 7 }, { 6, 5, 4 }, { 3, 2, 1 } };

            // Создание матрицы результата
            int[,] resultMatrix = new int[3, 3];

            // Перемножение матриц
            for (int i = 0; i < 3; i++)
            {
                for (int j = 0; j < 3; j++)
                {
                    for (int k = 0; k < 3; k++)
                    {
                        resultMatrix[i, j] += matrix1[i, k] * matrix2[k, j];
                    }
                }
            }

            // Создание и заполнение Excel документа
            DataTable dt = new DataTable("Matrix");
            dt.Columns.Add("Matrix1", typeof(int));
            dt.Columns.Add("Matrix2", typeof(int));
            dt.Columns.Add("Result", typeof(int));

            for (int i = 0; i < 3; i++)
            {
                for (int j = 0; j < 3; j++)
                {
                    dt.Rows.Add(matrix1[i, j], matrix2[i, j], resultMatrix[i, j]);
                }
            }

            string fileName = "MatrixMultiplication.xlsx";
            string filePath = Path.Combine(Environment.CurrentDirectory, fileName);
            string connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=MyExcel.xls;Extended Properties=\"Excel 12.0;HDR=Yes;\"";

            using (OleDbConnection conn = new OleDbConnection(connectionString))
            {
                conn.Open();

                using (OleDbCommand cmd = new OleDbCommand())
                {
                    cmd.Connection = conn;

                    //Создать таблицу в книге
                    cmd.CommandText = "CREATE TABLE [Table] (Column1 INT, Column2 INT)";
                    cmd.ExecuteNonQuery();

                    //Заполнение
                    cmd.CommandText = "INSERT INTO [Table$] (Column1, Column2) VALUES (1, 2)";
                    cmd.ExecuteNonQuery();
                }
                conn.Close();
            } 
        } 
    }
}
