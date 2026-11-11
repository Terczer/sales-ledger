# Generator ewidencji sprzedaży

Lokalna aplikacja, która odczytuje transakcje TPAY i eBok PlanetPay, uzupełnia stały
szablon ewidencji i zapisuje gotowe skoroszyty w folderze `exports`.

## Układ plików

- `template\ewidencja sprzedaży bezrachunkowej template.xlsx` — szablon ewidencji.
- `dane\transakcje <data>.xlsx` — skoroszyty z arkuszami `Tpay Transakcje` i
  `Ebok PlanetPay Transakcje`. W `data` podaj datę, na przykład
  `dane\transakcje 10-2026.xlsx`.
- `exports\` — folder na wygenerowane pliki.

Aplikacja wyświetla listę pasujących plików z folderu `dane`; wybrany skoroszyt jest
łączony ze stałym szablonem. Jeśli dane zawierają transakcje z kilku miesięcy, powstaje
osobny plik dla każdego miesiąca. Nazwa eksportu składa się z nazwy szablonu i miesiąca
transakcji.

## Uruchomienie

W terminalu VS Code, w folderze projektu:

```powershell
npm.cmd install
npm.cmd start
```

Otwórz `http://127.0.0.1:3000`, wybierz plik transakcji i kliknij **Utwórz ewidencję**.
Pliki są zapisywane w folderze `exports` i można je pobrać z aplikacji.

## Testy

```powershell
npm.cmd test
```
