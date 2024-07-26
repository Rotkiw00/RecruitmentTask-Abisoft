using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using NLog;
using ExcelDataImporter.EntityModels;

namespace ExcelDataImporter;

/// <summary>
/// Importuje dane z arkusza Excel w formacie .xlsx i zwraca surowe dane kosztów.
/// </summary>
/// <remarks>
/// Klasa zaimplementowana jako Singleton zaciągające dane do <see cref="CostService"/>
/// </remarks>
public sealed class ExcelImporter
{
    private static readonly Logger _logger = LogManager.GetCurrentClassLogger();

    private static readonly Lazy<ExcelImporter> _instance = new Lazy<ExcelImporter>(() => new ExcelImporter());
    public static ExcelImporter Instance => _instance.Value;
    private ExcelImporter() { }

    private readonly ImporterConfigStartup _config = new();

    private Excel.Application xlsxApp;
    private Excel.Workbook xlsxWorkbook;
    private Excel.Worksheet xlsxWorksheet;
    private Excel.Range xlsxRange;

    public List<RawCostData> ImportData()
    {
        string excelFileName = Path.GetFileName(_config.ExcelImporterSettings.XlsxFilePath);
        _logger.Info($"Rozpoczęcie importu danych z pliku: {excelFileName}");

        InitWorkbook();

        var rawXlsxData = new List<RawCostData>();
        try
        {
            for (int row = 2; row < xlsxRange.Rows.Count; row++)
            {
                if (xlsxRange.Cells[row, 1].Value is null ||
                    xlsxRange.Cells[row, 2].Value is null ||
                    xlsxRange.Cells[row, 25].Value is null)
                {
                    continue;
                }

                string konto = xlsxRange.Cells[row, 1].Value.ToString();
                string nazwa = xlsxRange.Cells[row, 2].Value.ToString();
                string saldoOkresu = xlsxRange.Cells[row, 25].Value.ToString();

                rawXlsxData.Add(new RawCostData(konto, nazwa, saldoOkresu));
            }
        }
        catch (Exception ex)
        {
            _logger.Error($"Wystąpił błąd! | {ex}");
        }

        _logger.Info("Import danych z Excela zakończony pomyślnie.");

        CleanupUnmanagedResources();

        return rawXlsxData;
    }

    /// <summary>
    /// Inicjalizacja pól klasy reprezentujących arkusz kalkulacyjny Excel.
    /// </summary>
    /// <remarks>
    /// Pobranie parametrów konfiguracyjnych z <c>appsettings.json</c> do <see cref="ImporterConfigStartup"/>
    /// </remarks>
    private void InitWorkbook()
    {
        try
        {
            xlsxApp = new Excel.Application();
            xlsxWorkbook = xlsxApp.Workbooks.Open(_config.ExcelImporterSettings.XlsxFilePath);
            xlsxWorksheet = xlsxWorkbook.Sheets[_config.ExcelImporterSettings.WorksheetIndex];
            xlsxRange = xlsxWorksheet.UsedRange;
        }
        catch (Exception ex)
        {
            _logger.Error($"Wystąpił błąd! | {ex}");
        }
    }

    private void CleanupUnmanagedResources()
    {
        _logger.Info("Rozpoczęcie zwalaniania zasobów...");

        GC.Collect();
        GC.WaitForPendingFinalizers();

        Marshal.ReleaseComObject(xlsxRange);
        xlsxRange = null;

        Marshal.ReleaseComObject(xlsxWorksheet);
        xlsxWorksheet = null;

        xlsxWorkbook.Close();
        Marshal.ReleaseComObject(xlsxWorkbook);
        xlsxWorkbook = null;

        xlsxApp.Quit();
        Marshal.ReleaseComObject(xlsxApp);
        xlsxApp = null;

        _logger.Info("Zasoby zwolnione.");
    }
}
