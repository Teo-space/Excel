using OfficeOpenXml;
using System.Data;


public static partial class Excel
{
    /// <summary>
    /// DataTable таблице в Excel файл в виде byte[]
    /// </summary>
    public static class BinaryWriter
    {
        static BinaryWriter() => ExcelPackage.LicenseContext = LicenseContext.NonCommercial;


        public static byte[]? GetAsBytes(DataTable dt, string sheetName)
        {
            if (string.IsNullOrEmpty(sheetName))
            {
                throw new ArgumentNullException(nameof(sheetName));
            }
            using (var excelPackage = new ExcelPackage())
            {
                Excel.Writer.WriteDataTableToExcelPackage(excelPackage, sheetName, dt);
                var bytes = excelPackage.GetAsByteArray();
                return bytes;
            }
        }

        public static byte[]? GetAsBytes(DataTable dt, string sheetName, string templateFilePath)
        {
            if (string.IsNullOrEmpty(sheetName))
            {
                throw new ArgumentNullException(nameof(sheetName));
            }
            if (!File.Exists(templateFilePath))
            {
                throw new InvalidOperationException($"Файл шаблона '{templateFilePath}' не существует");
            }

            using (var templateStream = File.OpenRead(templateFilePath))
            using (var excelPackage = new ExcelPackage())
            {
                excelPackage.Load(templateStream);
                Excel.Writer.WriteDataTableToExcelPackage(excelPackage, sheetName, dt);
                var bytes = excelPackage.GetAsByteArray();
                return bytes;
            }
        }

    }
}