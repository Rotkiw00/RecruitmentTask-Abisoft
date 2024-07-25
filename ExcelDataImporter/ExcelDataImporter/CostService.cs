using ExcelDataImporter;
using ExcelDataImporter.EntityModels;
using NLog;

public class CostService
{
    private static readonly Dictionary<string, int> CostCategories = GetCostNames();
    private static readonly ILogger logger = LogManager.GetCurrentClassLogger();
    private static readonly List<RawCostData> data = ExcelImporter.Instance.ImportData();

    private static Dictionary<string, int> GetCostNames()
    {
        return new Dictionary<string, int>
        {
            {"40-amortyzacja", 401},
            {"paliwo", 410},
            {"energia", 413},
            {"materiały biurowe", 411},
            {"materiały do rem.", 412},
            {"części samoch.", 0},
            {"zużycie mat. inne", 419},
            {"telekomunikacja", 424},
            {"naprawy samoch.", 422},
            {"remonty budowlane", 0},
            {"prowizje bankowe", 425},
            {"usługi obce inne", 429},
            {"od nieruchomości", 432},
            {"za wiecz. użytk. gruntów", 430},
            {"od środków transp.", 0},
            {"podatki pozostałe", 433},
            {"osobowe", 441},
            {"bezosobowe", 443},
            {"ZFN", 0},
            {"wypłaty jednorazowe", 0},
            {"wynagrodzenia inne", 0},
            {"ZUS", 451},
            {"ZFSS", 0},
            {"świadczenia inne", 455},
            {"ryczałt samochodowy", 464},
            {"delegacje", 461},
            {"delegacje inne", 0},
            {"ubezp. majątku", 481},
            {"reklama kursów", 0},
            {"reprezentacja", 0},
            {"pozostałe", 0}
        };
    }

    public List<CostDto> ProcessData()
    {
        logger.Info("Rozpoczęcie przetwarzania danych.");
        var costs = new List<CostDto>();

        foreach (var data in data)
        {
            var wyszczegolnienie = ExtractWyszczegolnienie(data.Konto);
            if (CostCategories.TryGetValue(wyszczegolnienie, out int category))
            {
                var cost = CreateCost(data, wyszczegolnienie, category);
                costs.Add(new CostDto(cost));
            }
            else
            {
                logger.Warn($"Wyszczególnienie '{wyszczegolnienie}' nie pasuje do żadnej kategorii i zostało zignorowane.");
            }
        }

        logger.Info("Przetwarzanie danych zakończone pomyślnie.");
        return costs;
    }

    private string ExtractWyszczegolnienie(string konto)
    {
        var parts = konto.Split('-');
        return parts.Length > 2 ? parts[2].Trim() : string.Empty;
    }

    private Cost CreateCost(RawCostData data, string wyszczegolnienie, int category)
    {
        var cost = new Cost
        {
            Wyszczegolnienie = wyszczegolnienie
        };

        switch (category)
        {
            case 401:
                cost.Wynajem_pomieszczen = data.SaldoOkresu;
                break;
                // Add other categories here
                // ...
        }

        logger.Debug($"Utworzono rekord kosztu dla {wyszczegolnienie} o wartości {data.SaldoOkresu}.");

        return cost;
    }
}
