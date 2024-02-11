/// <summary>
/// Проверяем чтение из Excel файла в Dt, запись dt в Excel файл
/// </summary>
internal class Test1
{
    public static void Run()
    {
        Console.WriteLine("Start");

        var path = Path.Combine(Directory.GetCurrentDirectory(), "TestSheets.xlsx");
        Console.WriteLine(path);
        Console.WriteLine(File.Exists(path));


        var dt = Excel.StreamReader.ReadDataTableFromFile(path, "FirstList");
        Console.WriteLine(dt.Head());


        var pathToWrite = Path.Combine(Directory.GetCurrentDirectory(), "TestSheetsWriten.xlsx");
        var pathToTemplate = Path.Combine(Directory.GetCurrentDirectory(), "TestSheetsTemplate.xlsx");
        Excel.StreamWriter.WriteDataTableToFile(pathToWrite, pathToTemplate, "FirstList", dt);


        var dtWriten = Excel.StreamReader.ReadDataTableFromFile(pathToWrite, "FirstList");
        Console.WriteLine(dtWriten.Head());
    }
}