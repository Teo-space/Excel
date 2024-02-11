using OfficeOpenXml;
using System.Data;


public static partial class Excel
{
    /// <summary>
    /// Запись DataTable таблицы в Stream
    /// </summary>
    public static class StreamWriter
    {
        static StreamWriter() => ExcelPackage.LicenseContext = LicenseContext.NonCommercial;


        public static void WriteDataTableToStream(Stream writeStream, string sheetName, DataTable dt)
        {
            if (string.IsNullOrEmpty(sheetName))
            {
                throw new ArgumentNullException(nameof(sheetName));
            }
            using (var excelPackage = new ExcelPackage(writeStream))
            {
                Excel.Writer.WriteDataTableToExcelPackage(excelPackage, sheetName, dt);
            }
        }

        public static void WriteDataTableToFile(string filePath, string sheetName, DataTable dt)
        {
            if (File.Exists(filePath))
            {
                throw new InvalidOperationException($"Файл '{filePath}' уже существует");
            }
            using (var writeStream = File.OpenWrite(filePath))
            {
                WriteDataTableToStream(writeStream, sheetName, dt);
            }
        }

        public static void WriteDataTableToStream(Stream writeStream, Stream templateReadStream, string sheetName, DataTable dt)
        {
            using (var excelPackage = new ExcelPackage(writeStream, templateReadStream))
            {
                Excel.Writer.WriteDataTableToExcelPackage(excelPackage, sheetName, dt);
            }
        }

        public static void WriteDataTableToFile(string filePath, string templatePath, string sheetName, DataTable dt)
        {
            if (File.Exists(filePath))
            {
                throw new InvalidOperationException($"Файл назначения '{filePath}' уже существует");
            }
            if (!File.Exists(templatePath))
            {
                throw new InvalidOperationException($"Файл шаблона '{templatePath}' не существует");
            }
            using (var templateStream = File.OpenRead(templatePath))
            {
                using (var writeStream = new FileStream(filePath, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None))
                {
                    WriteDataTableToStream(writeStream, templateStream, sheetName, dt);
                }
            }
        }

    }
}