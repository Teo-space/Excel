using OfficeOpenXml;
using System.Data;


public static partial class Excel
{
    public static class Writer
    {
        static Writer()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }

        public static void WriteDataTableToExcelPackage(ExcelPackage excelPackage, string sheetName, DataTable dt)
        {
            if (string.IsNullOrEmpty(sheetName))
            {
                throw new ArgumentNullException(nameof(sheetName));
            }
            ExcelWorksheet excelWorksheet = excelPackage.Workbook.Worksheets.FirstOrDefault(x => x.Name == sheetName)
                ?? excelPackage.Workbook.Worksheets.Add(sheetName);

            excelWorksheet.Cells["A1"].LoadFromDataTable(dt, true);
            excelPackage.Save();
        }

    }
}