using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace ExceltoCSV
{
    class Program
    {
        public static void Main(string[] args)
        {
            var excelFilePath = @"C:/Reports/HR Flexible Schedule Reports/Flex Report.xlsx";
            string output = Path.ChangeExtension(excelFilePath, DateTime.Now.ToString("yyyy-MM-dd_hh-mm_tt") + ".csv");
            bool saveComplete = SaveAsCsv(excelFilePath, output);
            Environment.Exit(1);
        }

        public static bool SaveAsCsv(string excelFilePath, string destinationCsvFilePath)
        {
            using (var stream = new FileStream(excelFilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                IExcelDataReader reader = null;
                Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
                try
                {
                    using (reader = ExcelReaderFactory.CreateReader(stream))
                    {
                        var ds = reader.AsDataSet(new ExcelDataSetConfiguration()
                        {
                            ConfigureDataTable = (tableReader) => new ExcelDataTableConfiguration()
                            {
                                UseHeaderRow = false
                            }
                        });

                        var csvContent = string.Empty;
                        int row_no = 0;
                        while (row_no < ds.Tables[0].Rows.Count)
                        {
                            var arr = new List<string>();
                            for (int i = 0; i < ds.Tables[0].Columns.Count; i++)
                            {
                                arr.Add(ds.Tables[0].Rows[row_no][i].ToString());
                            }
                            row_no++;
                            csvContent += string.Join(",", arr) + "\n";
                        }
                        StreamWriter csv = new StreamWriter(destinationCsvFilePath, false);
                        csv.Write(csvContent);
                        csv.Close();
                        return true;
                    }
                }
                catch (Exception e)
                {
                    Console.WriteLine(e.Message);
                    return false;                                        
                }                            
            }
        }
    }
}
