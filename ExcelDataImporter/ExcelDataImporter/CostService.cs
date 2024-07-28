using ExcelDataImporter.EntityModels;
using NLog;

namespace ExcelDataImporter;

/// <summary>
/// Klasa odpowiedzialna za przetwarzanie danych kosztów zaimportowanych z pliku Excel.
/// </summary>
public class CostService
{
    // Utworzenie loggera NLog do logowania informacji i błędów
    private static readonly Logger _logger = LogManager.GetCurrentClassLogger();

    // Inicjalizacja listy danych kosztów zaimportowanych z pliku Excel
    private readonly List<RawCostData> XlsxData = ExcelImporter.Instance.ImportData();

    // Inicjalizacja listy kategorii kosztów i przypisanie do nich kodów
    private readonly List<(string, int)> CostCategories = GetCostNames();

    /// <summary>
    /// Statyczna metoda prywatna zwracająca listę nazw kategorii kosztów wraz z przypisanymi kodami.
    /// </summary>
    /// <returns>
    /// Lista krotek (<seealso cref="string"/>, <seealso cref="int"/>), 
    /// gdzie pierwszy element to nazwa kosztu, a drugi to odpowiadający mu kod.
    /// </returns>
    private static List<(string, int)> GetCostNames()
    {
        return new List<(string, int)>
        {
            ("40-amortyzacja", 401),

            ("paliwo", 410),
            ("energia", 413),
            ("materiały biurowe", 411),
            ("materiały do rem.", 412),
            ("części samoch.", 0),
            ("zużycie mat. inne", 419),

            ("telekomunikacja", 424),
            ("naprawy samoch.", 422),
            ("remonty budowlane",0),
            ("prowizje bankowe", 425),
            ("usługi obce inne", 429),

            ("od nieruchomości", 432),
            ("za wiecz. użytk. gruntów", 430),
            ("od środków transp.",0),
            ("podatki pozostałe", 433),

            ("osobowe", 441),
            ("bezosobowe", 443),
            ("ZFN",0),
            ("wypłaty jednorazowe",0),
            ("wynagrodzenia inne", 0),

            ("ZUS", 451),
            ("ZFSS",0),
            ("świadczenia inne", 455),

            ("ryczałt samochodowy", 464),
            ("delegacje", 461),
            ("delegacje inne",0),

            ("ubezp. majątku", 481),
            ("reklama kursów",0),
            ("reprezentacja",0),
            ("pozostałe",0)
        };
    }

    /// <summary>
    /// Główny komponent przetwarzający dane kosztów
    /// </summary>
    /// <remarks>
    /// Metoda typu <c>void</c>, ale można by było zwrócić listę obiektów typu <see cref="CostDto"/>
    /// w celu dalszego przetwarzania.
    /// </remarks>
    public void ProcessData()
    {
        var costs = new List<CostDto>();

        // Pętla przetwarzająca każdy rekord danych kosztów pobranych z Excela
        foreach (var rawCostData in XlsxData)
        {
            // Wyciągnięcie kategorii i kodu wyszczególnienia z rekordu danych kosztów
            (int category, int detailCodeToValidate) = ExtractCategoryAndDetailCodeFromRecord(rawCostData);

            // Znalezienie pasujących kategorii kosztów na podstawie kodu wyszczególnienia
            var matchingCostCategories = CostCategories
                .Where(cc => cc.Item2 == detailCodeToValidate)
                .ToList();

            // Jeśli znaleziono pasujące kategorie, utworzenie rekordów kosztów
            if (matchingCostCategories.Count != 0)
            {
                foreach (var (detail, _) in matchingCostCategories)
                {
                    Cost cost = CreateCost(rawCostData, detail, category);
                    _logger.Info($"Utworzono rekord kosztu dla '{detail}' o wartości {rawCostData.SaldoOkresu}.");

                    costs.Add(new CostDto(cost));
                }
            }
            else
            {
                _logger.Warn($"Wyszczególnienie dla kategorii '{detailCodeToValidate}' nie pasuje do żadnej kategorii i zostało zignorowane.");
            }
        }

        _logger.Info("Przetwarzanie danych zakończone pomyślnie.");
    }

    /// <summary>
    /// Wyciąga kategorię i kod szczegółowy z rekordu danych kosztów.
    /// </summary>
    /// <param name="rawCostData">Rekord danych kosztów</param>
    /// <returns>
    /// Krotka zawierająca kod kategorii oraz kod wyszczególnienia.
    /// </returns>
    private static (int, int) ExtractCategoryAndDetailCodeFromRecord(RawCostData rawCostData)
    {
        string[] parts = rawCostData.Konto.Split('-');

        if (int.TryParse(parts.First(), out int category) &&
            int.TryParse(parts.Last(), out int detailCode))
        {
            return (category, detailCode);
        }
        else
        {
            return (0, 0);
        }
    }

    /// <summary>
    /// Tworzy obiekt Cost na podstawie danych kosztów, szczegółu i kategorii.
    /// </summary>
    /// <param name="rawCostData">Rekord danych kosztów.</param>
    /// <param name="detail">Wyszczególnienie.</param>
    /// <param name="category">Kategoria.</param>
    /// <returns>
    /// Obiekt <see cref="Cost"/>
    /// </returns>
    private static Cost CreateCost(RawCostData rawCostData, string detail, int category) => category switch
    {
        520 => new Cost() { Wyszczegolnienie = detail, Wynajem_pomieszczen = rawCostData.SaldoOkresu },
        527 => new Cost() { Wyszczegolnienie = detail, Stacje_diagnostyczne = rawCostData.SaldoOkresu },
        540 => new Cost() { Wyszczegolnienie = detail, Utrzymanie_biura = rawCostData.SaldoOkresu },
        _ => new Cost() { Wyszczegolnienie = detail, Pozostale_uslugi = rawCostData.SaldoOkresu }
    };
}
