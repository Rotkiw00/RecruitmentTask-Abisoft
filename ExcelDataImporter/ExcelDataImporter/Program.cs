namespace ExcelDataImporter;

internal class Program
{
    static void Main(string[] args)
    {
        CostService costService = new CostService();
        costService.ProcessData();

        Console.ReadKey();
    }
}
