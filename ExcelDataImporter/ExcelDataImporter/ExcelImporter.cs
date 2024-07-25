using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using NLog;
using ExcelDataImporter.EntityModels;

namespace ExcelDataImporter;

public sealed class ExcelImporter
{
    private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

    private static readonly Lazy<ExcelImporter> _instance = new Lazy<ExcelImporter>(() => new ExcelImporter());
    public static ExcelImporter Instance => _instance.Value;
    private ExcelImporter() { }
        
    private Excel.Application xlsxApp;
    private Excel.Workbook xlsxWorkbook;
    private Excel.Worksheet xlsxWorksheet;
    private Excel.Range xlsxRange;

    public List<RawCostData> ImportData()
    {
        InitWorkbook();
        /* potrzebne są trzy kolumny
         * Konto, Nazwa, Saldo okresu
         * Można pominąć nagłówki kolumn, więc iterujemy od 2 
         * , bo w Excelu nie ma 0-based indices
        */
        var rawXlsxData = new List<RawCostData>();
        for (int row = 2; row < xlsxRange.Rows.Count; row++)
        {
            if (xlsxRange.Cells[row, 1].Value is null ||
                xlsxRange.Cells[row, 2].Value is null ||
                xlsxRange.Cells[row, 25].Value is null) { continue; }
            // tu trzeba będzie uważać, bo co jeśli nie będzie kolumny pod tym indeksem
            // albo będzie inna
            // mozna by napisac metodę, która będzie zwracała indeks kolumny po jej nazwie
            string konto = xlsxRange.Cells[row, 1].Value.ToString();
            string nazwa = xlsxRange.Cells[row, 2].Value.ToString();
            string saldoOkresu = xlsxRange.Cells[row, 25].Value.ToString();

            rawXlsxData.Add(new RawCostData(konto, nazwa, saldoOkresu));
        }

        CleanupUnmanagedResources();

        return rawXlsxData;
    }

    private void InitWorkbook()
    {
        ImporterConfigStartup importerStartup = new ImporterConfigStartup();

        xlsxApp = new Excel.Application();
        xlsxWorkbook = xlsxApp.Workbooks.Open(importerStartup.ExcelImporterSettings.XlsxFilePath); 
        // dodać ścieżkę => jakoś ją wstrzyknąć
        xlsxWorksheet = xlsxWorkbook.Sheets[1]; // są dwa arkusze, więc trzeba będzie to sparametryzować
        xlsxRange = xlsxWorksheet.UsedRange;
    }

    private void CleanupUnmanagedResources()
    {
        GC.Collect();
        GC.WaitForPendingFinalizers();

        Marshal.ReleaseComObject(xlsxRange);
        Marshal.ReleaseComObject(xlsxWorksheet);
        xlsxWorkbook.Close();
        Marshal.ReleaseComObject(xlsxWorkbook);
        xlsxApp.Quit();
        Marshal.ReleaseComObject(xlsxApp);
    }
}
