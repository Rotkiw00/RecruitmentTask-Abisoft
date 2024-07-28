namespace ExcelDataImporter;
internal class Program
{
    static void Main(string[] args)
    {
        // Wywołanie metody przetwarzające dane kosztów.
        // Proces przetwarzania jest logowany do okna konsoli oraz w różnych przypadkach do pliku.

        CostService costService = new CostService();
        costService.ProcessData();

        Console.ReadKey();
    }
}
