var path = Path.Combine(Directory.GetCurrentDirectory(), "TestSheets.xlsx");
Console.WriteLine(path);
Console.WriteLine(File.Exists(path));


var dt = Excel.Reader.ReadDataTableFromFile(path, "FirstList");
Console.WriteLine(dt.Head());


var pathToWrite = Path.Combine(Directory.GetCurrentDirectory(), "TestSheetsWriten.xlsx");
var pathToTemplate = Path.Combine(Directory.GetCurrentDirectory(), "TestSheetsTemplate.xlsx");
Excel.Writer.WriteDataTableToFile(pathToWrite, pathToTemplate, "FirstList", dt);


var dtWriten = Excel.Reader.ReadDataTableFromFile(pathToWrite, "FirstList");
Console.WriteLine(dtWriten.Head());