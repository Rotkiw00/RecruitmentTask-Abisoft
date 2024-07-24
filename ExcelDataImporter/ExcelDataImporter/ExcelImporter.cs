using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelDataImporter;

public sealed class ExcelImporter
{
    private static readonly Lazy<ExcelImporter> _instance = new Lazy<ExcelImporter>(() => new ExcelImporter());
    public static ExcelImporter Instance => _instance.Value;
    private ExcelImporter() { }

    public void ImportData()
    {
        Excel.Application xlsxApp = new Excel.Application();
        Excel.Workbook xlsxWorkbook = xlsxApp.Workbooks.Open(@"C:\Users\Wiktor Kalaga\Documents\WORK\Pliki\Rekrutacja\RecruitmentTask-Abisoft\Dane_rekrutacja.xlsx"); // dodać ścieżkę => jakoś ją wstrzyknąć
        Excel.Worksheet xlsxWorksheet = xlsxWorkbook.Sheets[1]; // są dwa arkusze, więc trzeba będzie to sparametryzować
        Excel.Range xlsxRange = xlsxWorksheet.UsedRange;

        var rawXlsxData = new List<RawData>();

        for (int row = 1; row < xlsxRange.Rows.Count; row++)
        {
            for (int col = 1; col < xlsxRange.Columns.Count; col++)
            {

            }
        }
    }
}

public record RawData(string Konto, string Nazwa, string SaldoOkresu);
