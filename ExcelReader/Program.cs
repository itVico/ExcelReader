using OfficeOpenXml;
using System;
using System.IO;
using System.Linq;
using System.Text;

namespace ExcelReader
{
    public class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Enter path to excel");
            var filePath = Console.ReadLine();

            FileInfo file = new FileInfo(filePath);

            using (ExcelPackage package = new ExcelPackage(file))
            {
                StringBuilder sb = new StringBuilder();
                ExcelWorksheet worksheet = package.Workbook.Worksheets.First();
                int rowCount = worksheet.Dimension.Rows;
                int ColCount = worksheet.Dimension.Columns;

                var rawText = string.Empty;
                for (int row = 1; row <= rowCount; row++)
                {
                    for (int col = 1; col <= ColCount; col++)
                    {
                        if (worksheet.Cells[row, col].Value != null)
                        {
                            rawText += worksheet.Cells[row, col].Value.ToString() + "\t";
                        }
                        else
                        {
                            rawText += "Null\t";
                        }

                    }
                    rawText += "\r\n";
                }
                Console.WriteLine(rawText);
            }
            Console.ReadKey();

        }
    }
}
