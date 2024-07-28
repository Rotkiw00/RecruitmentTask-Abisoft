namespace ExcelDataImporter.EntityModels;

/// <summary>
/// Reprezentuje szczegółowe dane kosztów.
/// </summary>
/// <remarks>
/// Pola kosztów powinny mieć typ wartościowy, np. <see cref="decimal"/>, lecz wymagałoby to dodatkowej konwersji ze stringa.
/// Na aktualne potrzeby wystarczy typ <see cref="string"/>
/// </remarks>
public class Cost
{
    public string Wyszczegolnienie { get; set; }
    public string Utrzymanie_biura { get; set; }
    public string Szkolenie_masowe { get; set; }
    public string Wynajem_pomieszczen { get; set; }
    public string Handel_hurt { get; set; }
    public string Uslugi_turystyczne { get; set; }
    public string Szkolenie_komercyjne { get; set; }
    public string Dzialalnosc_wydawnicza { get; set; }
    public string Pozostale_uslugi { get; set; }
    public string Stacje_diagnostyczne { get; set; }
}

/// <summary>
/// DTO dla klasy <see cref="Cost"/>. Inicjalizuje nową instancję klasy <see cref="CostDto"/> na podstawie obiektu <see cref="Cost"/>.
/// </summary>
/// <param name="entity">Obiekt <see cref="Cost"/>, z którego dane są kopiowane.</param>
public class CostDto(Cost entity)
{
    public string Wyszczegolnienie { get; set; } = entity.Wyszczegolnienie;
    public string Utrzymanie_biura { get; set; } = entity.Utrzymanie_biura;
    public string Szkolenie_masowe { get; set; } = entity.Szkolenie_masowe;
    public string Wynajem_pomieszczen { get; set; } = entity.Wynajem_pomieszczen;
    public string Handel_hurt { get; set; } = entity.Handel_hurt;
    public string Uslugi_turystyczne { get; set; } = entity.Uslugi_turystyczne;
    public string Szkolenie_komercyjne { get; set; } = entity.Szkolenie_komercyjne;
    public string Dzialalnosc_wydawnicza { get; set; } = entity.Dzialalnosc_wydawnicza;
    public string Pozostale_uslugi { get; set; } = entity.Pozostale_uslugi;
    public string Stacje_diagnostyczne { get; set; } = entity.Stacje_diagnostyczne;
}