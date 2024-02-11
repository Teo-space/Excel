/// <summary>
/// Запись коллекции объектов в Excel файл
/// </summary>
internal class Test2
{
    public static void Run()
    {
        Console.WriteLine("Start");
        if (File.Exists("tests_collection.xlsx"))
        {
            File.Delete("tests_collection.xlsx");
        }

        var tests = new List<Test>();

        tests.Add(new Test()
        {
            TestId = 1,
            TestName = "Test 1",
        });
        tests.Add(new Test()
        {
            TestId = 1,
            TestName = "Test 2",
        });

        var dt = tests.ToDataTable();
        Console.WriteLine(dt.Head());

        Excel.StreamWriter.WriteDataTableToFile("tests_collection.xlsx", "Sheet", dt);
    }
}
public class Test
{
    public int TestId { get; init; }
    public string TestName { get; init; }
    public DateTime CreatedAt { get; init; } = DateTime.Now;
}