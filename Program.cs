using System;
using System.IO;
using OfficeOpenXml;

class Program
{
    static void Main(string[] args)
    {
        string filePath = @"C:\Users\Ajayprasath\OneDrive\Desktop\Soustr\Sales-Dashboard-practice-file.xlsx";
        
        if (!File.Exists(filePath))
        {
            Console.WriteLine("File not found.");
            return;
        }
        
        using (var package = new ExcelPackage(new FileInfo(filePath)))
        {
            ExcelWorksheet worksheet = package.Workbook.Worksheets[0]; 
            
            int rowCount = worksheet.Dimension.Rows;
            int colCount = worksheet.Dimension.Columns;
            
            for (int row = 1; row <= rowCount; row++)
            {
                for (int col = 1; col <= colCount; col++)
                {
                    object cellValue = worksheet.Cells[row, col].Value;
                    Console.Write(cellValue + "\t");
                }
                Console.WriteLine(); 
            }
        }
    }
}
