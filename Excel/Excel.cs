using OfficeOpenXml;

public static partial class Excel
{
    static Excel() => ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
}
