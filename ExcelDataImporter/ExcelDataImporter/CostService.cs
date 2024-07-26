using ExcelDataImporter.EntityModels;
using NLog;

namespace ExcelDataImporter;
public class CostService
{
    private static readonly Logger _logger = LogManager.GetCurrentClassLogger();

    private readonly List<RawCostData> XlsxData = ExcelImporter.Instance.ImportData();

    private readonly List<(string, int)> CostCategories = GetCostNames();
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

    public List<CostDto> ProcessData()
    {
        var costs = new List<CostDto>();

        foreach (var rawCostData in XlsxData)
        {
            int categoryToValidate = ExtractCategoryFromRecord(rawCostData);

            var matchingCostCategories = CostCategories
                .Where(cc => cc.Item2 == categoryToValidate)
                .ToList();

            if (matchingCostCategories.Count != 0)
            {
                foreach (var (detail, costCategory) in matchingCostCategories)
                {
                    Cost cost = CreateCost(rawCostData, detail, categoryToValidate);
                    costs.Add(new CostDto(cost));
                }
            }
            else
            {
                _logger.Warn($"Wyszczególnienie dla kategorii '{categoryToValidate}' nie pasuje do żadnej kategorii i zostało zignorowane.");
            }
        }

        _logger.Info("Przetwarzanie danych zakończone pomyślnie.");

        return costs;
    }

    private static int ExtractCategoryFromRecord(RawCostData rawCostData)
    {
        string[] parts = rawCostData.Konto.Split('-');
        return int.Parse(parts.Last());
    }

    private static Cost CreateCost(RawCostData rawCostData, string detail, int category)
    {
        var cost = new Cost
        {
            Wyszczegolnienie = detail
        };

        switch (category)
        {
            case 401:
                cost.Wynajem_pomieszczen = rawCostData.SaldoOkresu;
                break;
                // Add other categories here
                // ...
        }

        _logger.Info($"Utworzono rekord kosztu dla '{detail}' o wartości {rawCostData.SaldoOkresu}.");

        return cost;
    }
}
