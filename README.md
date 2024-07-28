## INFORMACJE
Program realizuje zadanie importu danych z akruszy z pliku Excel, tj. `Dane_rekrutacja.xlsx`. Mechanizm importu głównie opiera się na zaciągnięciu danych z trzech kolumn *Konto*, *Nazwa* oraz *Saldo okresu*. Są one zapisywane do tymczasowego obiektu `RawCostData`, który reprezentuje surowe dane - bezpośrednio zaciągnięte z Excela. Na potrzeby zadania zdecydowano się pominąć pozostałe kolumny.<br>**Klasa** `ExcelImporter` jest zaimplementowana przy użyciu wzorca Singleton. <br>**Klasa** `CostService` korzysta z instancji `ExcelImportera`, który pobiera dane z Excela. W tej klasie odbywa się przetwarzanie surowych danych na docelowe obiekty, tj. `CostDto`.<br> 
Aby nie przekazywać bezpośrednio ścieżki do Excela stworzona została klasa *konfiguracyjna* - `ImporterConfigStartup`, która zaczytuje z pliku `appsettings.json` ścieżkę do arkuszy (parametr: `XlsxFilePath`) oraz aktualnie rozpatrywany arkusz (parametr: `WorksheetIndex`), gdzie przekazywana liczba odpowiada indeksowi arkusza, który jest w pliku (zaczynając od 1).<br>
Wykorzytana została biblioteka NLog do logowania informacji, błędów, czy ostrzeżeń. Dzięki gotowej implementacji możliwe jest utworzenie instancji logera w danej klasie przy pomocy ```LogManager.GetCurrentClassLogger();```.<br>
**Kod programu został udokumentowany.**<br>
1. **plik `appsettings.json` zlokalizowany jest w folderze //bin/Debug/net8.0-windows**,
2. **logi zapisywane są również w tej samej lokalizacji co ww. plik konfiguracyjny w folderze *logs***,
3. **użyte biblioteki: Microsoft.Office.Interop.Excel, NLog, Microsoft.Extensions.Configuration**

## OPIS ZADANIA

Zadaniem jest zaimportowanie danych z pliku XLSX, zawierającego informacje o `Saldo okresu`, do odpowiedniej kategorii i wyszczególnienia. Plik XLSX zawiera kolumny, które należy przetworzyć w następujący sposób:

1. **Import danych**: Odczytaj dane z pliku XLSX, zwracając szczególną uwagę na kolumnę `Saldo okresu`.
2. **Kategoryzacja danych**: Przypisz każdą wartość `Saldo okresu` do odpowiedniej kategorii, zgodnie z wcześniej ustalonymi zasadami kategoryzacji.
3. **Wyszczególnienie (GetCostNames)**: Sprawdź, czy każde wyszczególnienie (`GetCostNames`) zawarte w pliku pasuje do jednej z wcześniej zdefiniowanych kategorii. Jeśli nie pasuje, można je zignorować.
4. **Implementacja mechanizmu importu**: Zaprojektuj i zaimplementuj mechanizm importu, który automatycznie odczyta dane z pliku XLSX, dokona kategoryzacji oraz wyszczególnienia. Mechanizm ten powinien być elastyczny, umożliwiający łatwe modyfikacje zasad kategoryzacji i sprawdzania wyszczególnienia.

**Wymagania techniczne:**

- Projekt może być wykonany jako osobny projekt lub jako zestaw klas do implementacji.
- Mechanizm importu powinien być dobrze udokumentowany, z opisem działania i instrukcjami dotyczącymi konfiguracji.
- Kod powinien być czytelny i zgodny z dobrymi praktykami programistycznymi.
- Przewidziana jest możliwość, że niektóre wyszczególnienia nie będą pasować do żadnej kategorii – w takim przypadku należy je zignorować, ale odpowiednio logować te przypadki.

**Dodatkowe informacje:**

- Plik XLSX zostanie dostarczony w załączniku.

**Ocena zadania:**

Podczas oceny zadania brane będą pod uwagę:

- Poprawność zaimportowanych danych.
- Efektywność i optymalizacja procesu importu.
- Jakość i czytelność kodu.
- Dokumentacja i łatwość wprowadzania zmian w mechanizmie importu.

**Termin wykonania zadania:**

Zadanie należy wykonać w ciągu 7 dni od momentu jego otrzymania. Po zakończeniu, gotowy projekt należy przesłać na wskazany adres e-mail lub umieścić w repozytorium GitHub (lub innym systemie kontroli wersji) i udostępnić link.

Powodzenia!


```c#
public CostDto(Cost entity)
{
    Wyszczegolnienie = entity.Wyszczegolnienie;
    Utrzymanie_biura = entity.Utrzymanie_biura;
    Szkolenie_masowe = entity.Szkolenie_masowe;
    Wynajem_pomieszczen = entity.Wynajem_pomieszczen;
    Handel_hurt = entity.Handel_hurt;
    Uslugi_turystyczne = entity.Uslugi_turystyczne;
    Szkolenie_komercyjne = entity.Szkolenie_komercyjne;
    Dzialalnosc_wydawnicza = entity.Dzialalnosc_wydawnicza;
    Pozostale_uslugi = entity.Pozostale_uslugi;
    Stacje_diagnostyczne = entity.Stacje_diagnostyczne;
}
```


```c#
private static List<(string, int)> GetCostNames() { return new List<(string, int)> { ("40-amortyzacja", 401), ("paliwo", 410), ("energia", 413), ("materiały biurowe", 411), ("materiały do rem.", 412), ("części samoch.", 0), ("zużycie mat. inne", 419), ("telekomunikacja", 424), ("naprawy samoch.", 422), ("remonty budowlane",0), ("prowizje bankowe", 425), ("usługi obce inne", 429), ("od nieruchomości", 432), ("za wiecz. użytk. gruntów", 430), ("od środków transp.",0), ("podatki pozostałe", 433), ("osobowe", 441), ("bezosobowe", 443), ("ZFN",0), ("wypłaty jednorazowe",0), ("wynagrodzenia inne", 0), ("ZUS", 451), ("ZFSS",0), ("świadczenia inne", 455), ("ryczałt samochodowy", 464), ("delegacje", 461), ("delegacje inne",0), ("ubezp. majątku", 481), ("reklama kursów",0), ("reprezentacja",0), ("pozostałe",0) }; } ``` Przykład: | Konto | Nazwa | Saldo okresu | | --------------- | ------------------------------------------------------------- | ------------ | | 520 - 215 - 401 | Wynajem pomieszczeń - Biłgoraj - amortyzacja środków trwałych | 772,98 | `new CostDto { Wyszczegolnienie="40-amortyzacja" , Wynajem_pomieszczen=772,98 }`
