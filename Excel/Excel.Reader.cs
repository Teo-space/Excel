using OfficeOpenXml;
using System.Data;

public static partial class Excel
{
    public static class Reader
    {
        static Reader()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }

        public static DataTable ReadDataTableFromExcelPackage(OfficeOpenXml.ExcelPackage excelPackage, string? sheetName)
        {
            if(!excelPackage.Workbook.Worksheets.Any()) 
            {
                throw new Exception("ExcelPackage has not WorkSheets");
            }

            var sheet = string.IsNullOrEmpty(sheetName)
                ? excelPackage.Workbook.Worksheets.FirstOrDefault(s => s.Name == sheetName)
                : excelPackage.Workbook.Worksheets.FirstOrDefault();
            if (sheet == null)
            {
                throw new Exception(string.IsNullOrEmpty(sheetName) ? "Workbook.Worksheets Is Empty" : $"WorkSheet '{sheetName}' not found");
            }

            DataTable dt = new DataTable();
            dt.TableName = sheetName ?? "Table";
            //Заполняем имена и типы столбцов
            int columnNameIndex = 1;
            foreach (var cell in sheet.Cells[1, 1, 1, sheet.Dimension.End.Column])
            {
                //dt.Columns.Add(cell.Text);
                var columnType = sheet.Cells[2, columnNameIndex, sheet.Dimension.End.Row, columnNameIndex]
                    .Where(x => x.Value != null)
                    .Select(x => x.Value.GetType())
                    .GroupBy(x => x)
                    .OrderByDescending(group => group.Count())
                    .Select(x => x.Key)
                    .FirstOrDefault();
                dt.Columns.Add(cell.Text, columnType ?? typeof(string));

                columnNameIndex++;
            }
            //Заполняем Rows
            for (int rowNum = 2; rowNum <= sheet.Dimension.End.Row; rowNum++)
            {
                var wsRow = sheet.Cells[rowNum, 1, rowNum, sheet.Dimension.End.Column];
                DataRow row = dt.Rows.Add();
                foreach (var cell in wsRow)
                {
                    row[cell.Start.Column - 1] = cell.Value;
                    Console.WriteLine($"[{cell.Value}, ({cell.Value.GetType()})]");
                }
            }
            return dt;
        }

        public static DataTable ReadDataTableFromStream(Stream stream, string? sheetName)
        {
            using (var excelPackage = new OfficeOpenXml.ExcelPackage())
            {
                excelPackage.Load(stream);
                return ReadDataTableFromExcelPackage(excelPackage, sheetName);
            }
        }

        public static DataTable ReadDataTableFromFile(string filePath, string? sheetName)
        {
            if (!File.Exists(filePath))
            {
                throw new FileNotFoundException(filePath);
            }
            using (var stream = File.OpenRead(filePath))
            {
                return ReadDataTableFromStream(stream, sheetName);
            }
        }


    }
}