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
    // Utworzenie loggera NLog do logowania informacji i błędów
    private static readonly Logger _logger = LogManager.GetCurrentClassLogger();

    // Instancja singletona klasy ExcelImporter.
    private static readonly Lazy<ExcelImporter> _instance = new Lazy<ExcelImporter>(() => new ExcelImporter());
    public static ExcelImporter Instance => _instance.Value;

    // Prywatny konstruktor klasy ExcelImporter
    private ExcelImporter() { }

    // Konfiguracja importera zaczytana z pliku appsettings.json
    private readonly ImporterConfigStartup _config = new();

    // Pola reprezentujące aplikację Excel oraz różne elementy arkusza kalkulacyjnego
    private Excel.Application xlsxApp;
    private Excel.Workbook xlsxWorkbook;
    private Excel.Worksheet xlsxWorksheet;
    private Excel.Range xlsxRange;

    /// <summary>
    /// Zaczytuje arkusz Excel oraz importuje znajdujące się w nim dane.
    /// Mechanizm importu jest skupiony tylko na trzech kolumnach: Konto, Nazwa i Saldo okresu.
    /// </summary>
    /// <returns>
    /// Lista surowych danych kosztów składających się z Konta, Nazwy oraz Salda okresu.
    /// Dane przekazane do <see cref="CostService"/>.
    /// </returns>
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
    /// Inicjalizuje pola klasy reprezentujące arkusz kalkulacyjny Excel.
    /// </summary>
    /// <remarks>
    /// Pobiera parametry konfiguracyjne z <c>appsettings.json</c> do <see cref="ImporterConfigStartup"/>.
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

    /// <summary>
    /// Zwolnienie niezarządzanych domyślnie przez GC zasobów tj. otwartej aplikacji Excel i arkuszy
    /// </summary>
    private void CleanupUnmanagedResources()
    {
        _logger.Info("Rozpoczęcie zwalaniania zasobów...");

        GC.Collect();
        GC.WaitForPendingFinalizers();

        if (xlsxRange is not null)
        {
            Marshal.ReleaseComObject(xlsxRange);
            xlsxRange = null;
        }

        if (xlsxWorksheet is not null)
        {
            Marshal.ReleaseComObject(xlsxWorksheet);
            xlsxWorksheet = null;
        }

        if (xlsxWorkbook is not null)
        {
            xlsxWorkbook.Close();
            Marshal.ReleaseComObject(xlsxWorkbook);
            xlsxWorkbook = null;
        }

        if (xlsxApp is not null)
        {
            xlsxApp.Quit();
            Marshal.ReleaseComObject(xlsxApp);
            xlsxApp = null;
        }

        _logger.Info("Zasoby zwolnione.");
    }
}
