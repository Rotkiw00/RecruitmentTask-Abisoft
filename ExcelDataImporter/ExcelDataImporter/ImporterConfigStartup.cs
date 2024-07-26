using ExcelDataImporter.ConfigModels;
using Microsoft.Extensions.Configuration;

namespace ExcelDataImporter;

/// <summary>
/// Pobranie konfiguracji do odczytania pliku Excel.
/// </summary>
internal class ImporterConfigStartup
{
    public ExcelImporterSettings ExcelImporterSettings { get; private set; }

    public ImporterConfigStartup()
    {
        var builder = new ConfigurationBuilder()
                  .SetBasePath(Directory.GetCurrentDirectory())
                  .AddJsonFile("appsettings.json", optional: false);

        IConfiguration config = builder.Build();

        ExcelImporterSettings = config.GetSection("ExcelImporterSettings").Get<ExcelImporterSettings>();
    }
}

