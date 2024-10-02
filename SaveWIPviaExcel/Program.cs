using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace SaveWIPviaExcel
{
    class Program
    {
        static void Main(string[] args)
        {
            Excel.Application excelApp = null;
            Excel.Workbook workbook = null;
            Excel.Worksheet worksheet = null;

            try
            {
                if (args.Length == 0)
                {
                    MessageBox.Show("Please provide the path to the Excel file.");
                    return;
                }

                string filePath = args[0];

                // Start Excel application
                excelApp = new Excel.Application();
                excelApp.Visible = false;

                // Open the workbook
                workbook = excelApp.Workbooks.Open(filePath, ReadOnly: true);
                worksheet = (Excel.Worksheet)workbook.Sheets["Sheet1"];

                // Find the last row with data in column A
                Excel.Range lastCell = worksheet.Cells[worksheet.Rows.Count, "A"].End[Excel.XlDirection.xlUp];
                int lastRow = lastCell.Row;

                // Read data into an array
                Excel.Range dataRange = worksheet.Range["A1:J" + lastRow];
                object[,] data = dataRange.Value2;

                workbook.Close(false);
                Marshal.ReleaseComObject(workbook);
                excelApp.Quit();
                Marshal.ReleaseComObject(excelApp);

                // Create a list to hold the threads
                List<Thread> threads = new List<Thread>();
                int batchSize = 1000;

                for (int i = 2; i <= data.GetLength(0); i += batchSize)
                {
                    int start = i;
                    int end = Math.Min(i + batchSize - 1, data.GetLength(0));
                    Thread thread = new Thread(() => ProcessBatch(data, start, end));
                    threads.Add(thread);
                    Console.WriteLine($"Thread Started");
                    thread.Start();
                }

                // Wait for all threads to complete
                foreach (Thread thread in threads)
                {
                    thread.Join();
                }

                // Show a message box when process is completed
                Console.WriteLine("WIP to DB completed.");
            }
            catch (SqlException ex)
            {
                Console.WriteLine("An SQL error occurred: " + ex.Message);
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }

        private static void ProcessBatch(object[,] data, int startRow, int endRow)
        {
            string connectionString = "Data Source=130.171.191.142;Initial Catalog=PCL;Persist Security Info=True;User ID=passreg;Password=;";
            using (SqlConnection cnn = new SqlConnection(connectionString))
            {
                cnn.Open();

                int checksum = 0;
                for (int i = startRow; i <= endRow; i++)
                {
                    // Validate array values
                    string pn = string.IsNullOrWhiteSpace(data[i, 1]?.ToString()) ? "0" : data[i, 1].ToString();
                    string location = string.IsNullOrWhiteSpace(data[i, 10]?.ToString()) ? "Not Filled" : data[i, 10].ToString();
                    string projeto1 = string.IsNullOrWhiteSpace(data[i, 3]?.ToString()) ? "Not Filled" : data[i, 3].ToString();
                    string projeto2 = string.IsNullOrWhiteSpace(data[i, 4]?.ToString()) ? "Not Filled" : data[i, 4].ToString();
                    string tipo = string.IsNullOrWhiteSpace(data[i, 5]?.ToString()) ? "Not Filled" : data[i, 5].ToString();
                    string processo = string.IsNullOrWhiteSpace(data[i, 6]?.ToString()) ? "Not Filled" : data[i, 6].ToString();
                    string projeto3 = string.IsNullOrWhiteSpace(data[i, 7]?.ToString()) ? "Not Filled" : data[i, 7].ToString();
                    string projeto4 = string.IsNullOrWhiteSpace(data[i, 8]?.ToString()) ? "Not Filled" : data[i, 8].ToString();
                    string obs = string.IsNullOrWhiteSpace(data[i, 9]?.ToString()) ? "Not Filled" : data[i, 9].ToString();
                    string description = string.IsNullOrWhiteSpace(data[i, 2]?.ToString()) ? "Not Filled" : data[i, 2].ToString();
                     

                    // Insert data into the database
                    InsertData(cnn, pn, location, projeto1, projeto2, tipo, processo, projeto3, projeto4, obs, description);
                    checksum++;
                }
            }
        }

        private static void InsertData(SqlConnection cnn, string pn, string location, string projeto1, string projeto2, string tipo, string processo, string projeto3, string projeto4, string obs, string description)
        {
            using (SqlCommand cmd = new SqlCommand("InsertPartnr_location", cnn))
            {
                cmd.CommandType = System.Data.CommandType.StoredProcedure;

                // Add parameters with validated values
                cmd.Parameters.AddWithValue("@pn", pn);
                cmd.Parameters.AddWithValue("@location", location);
                cmd.Parameters.AddWithValue("@projeto1", projeto1);
                cmd.Parameters.AddWithValue("@projeto2", projeto2);
                cmd.Parameters.AddWithValue("@tipo", tipo);
                cmd.Parameters.AddWithValue("@processo", processo);
                cmd.Parameters.AddWithValue("@projeto3", projeto3);
                cmd.Parameters.AddWithValue("@projeto4", projeto4);
                cmd.Parameters.AddWithValue("@obs", obs);
                cmd.Parameters.AddWithValue("@description", description);

                // Execute command
                cmd.ExecuteNonQuery();
            }
        }
    }
}
