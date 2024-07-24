namespace ExcelDataImporter.EntityModels;

public class Cost
{
    public string? Wyszczegolnienie { get; set; }
    public decimal? Utrzymanie_biura { get; set; }
    public decimal? Szkolenie_masowe { get; set; }
    public decimal? Wynajem_pomieszczen { get; set; }
    public decimal? Handel_hurt { get; set; }
    public decimal? Uslugi_turystyczne { get; set; }
    public decimal? Szkolenie_komercyjne { get; set; }
    public decimal? Dzialalnosc_wydawnicza { get; set; }
    public decimal? Pozostale_uslugi { get; set; }
    public decimal? Stacje_diagnostyczne { get; set; }
}


public class CostDto(Cost entity)
{
    public string? Wyszczegolnienie { get; set; } = entity.Wyszczegolnienie;
    public decimal? Utrzymanie_biura { get; set; } = entity.Utrzymanie_biura;
    public decimal? Szkolenie_masowe { get; set; } = entity.Szkolenie_masowe;
    public decimal? Wynajem_pomieszczen { get; set; } = entity.Wynajem_pomieszczen;
    public decimal? Handel_hurt { get; set; } = entity.Handel_hurt;
    public decimal? Uslugi_turystyczne { get; set; } = entity.Uslugi_turystyczne;
    public decimal? Szkolenie_komercyjne { get; set; } = entity.Szkolenie_komercyjne;
    public decimal? Dzialalnosc_wydawnicza { get; set; } = entity.Dzialalnosc_wydawnicza;
    public decimal? Pozostale_uslugi { get; set; } = entity.Pozostale_uslugi;
    public decimal? Stacje_diagnostyczne { get; set; } = entity.Stacje_diagnostyczne;
}

