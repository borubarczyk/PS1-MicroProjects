# PS1-MicroProjects

Zbior mikroprojektow PowerShell wspomagajacych codzienna administracje Active Directory oraz systemami Windows. W katalogu znajdziesz rozbudowane narzedzia GUI (WinForms/WPF) i skrypty tekstowe pokrywajace scenariusze od tworzenia kont, poprzez audyty NTFS, po diagnostyke stacji roboczych.

## Wymagania ogolne

- Windows PowerShell 5.1 lub PowerShell 7 uruchamiany na Windowsie.
- RSAT ActiveDirectory (modul `ActiveDirectory`) dla wszystkich skryptow zaczynajacych sie od `AD-`.
- Uprawnienia administracyjne odpowiadajace operacji (np. tworzenie OU, modyfikacja NTFS, zapytania do dziennikow RDP).
- Sesja w trybie STA dla narzedzi WinForms/WPF: `powershell.exe -STA -File .\NazwaSkryptu.ps1`.

## Jak uruchomic skrypt

1. Sklonuj repozytorium lub skopiuj katalog z repo:
   ```powershell
   git clone https://github.com/borubarczyk/PS1-MicroProjects.git
   cd PS1-MicroProjects
   ```
2. (Opcjonalnie) zezwol na uruchamianie: `Set-ExecutionPolicy -Scope Process Bypass -Force`.
3. Uruchom interesujacy plik: `powershell.exe -STA -ExecutionPolicy Bypass -File .\AD-BulkUserCreator.ps1`.
4. Jezeli narzedzie korzysta z GUI, uruchamiaj je w sesji desktopowej z odpowiednimi uprawnieniami.

## Pliki pomocnicze

- `.AD-BulkUserCreator.json` &mdash; zapisuje ustawienia zakladek (formaty, domyslne OU/UPN) kreatora kont.
- `AD-BulkUserCreator.log` &mdash; przykladowy log z ostatniego uruchomienia kreatora.
- `LICENSE` &mdash; informacja o licencji zbioru.

## Lista skryptow

| Skrypt | Opis |
| --- | --- |
| `AD-BulkAccountStatusChecker.ps1` | WinForms do wklejania listy loginow, szybkie sprawdzanie czy konta istnieja i sa wlaczone oraz eksport wynikow do CSV. |
| `AD-BulkAddUsersToGroups.ps1` | GUI z siatka dla wielu uzytkownikow i grup AD; rozwiazuje identyfikatory, dodaje czlonkow, loguje operacje i pozwala wklejac dane prosto z Excela. |
| `AD-BulkGroupCreator.ps1` | Tworzenie wielu grup w wybranym OU z poziomu jednej tabeli (Name, opis, scope, kategoria) wraz z normalizacja sAMAccountName. |
| `AD-BulkOUGroupCreator.ps1` | Kreator WinForms do budowy wielu OU oraz zestawu grup w kazdym OU; wspiera tryb WhatIf, transliteracje nazw i dbanie o unikalne sAM. |
| `AD-BulkUserCreator.ps1` | Rozbudowany kreator masowego tworzenia kont (role: uczen, student, pracownik, wykladowca, inne), walidacja danych, generowanie loginow/e-maili, wybor OU i UPN, zapisywanie ustawien w `.AD-BulkUserCreator.json`. |
| `AD-DeleteNewAccoutsGUI.ps1` | Okno z lista kont utworzonych w ostatnich X minutach; pozwala filtrowac, zaznaczac i usuwac wylapane konta, zapisujac log na Pulpicie. |
| `AD-ManagerDiamond.ps1` | Kompleksowe narzedzie Domain Ops do zarzadzania stacjami w AD (zdalne polecenia, BitLocker, dyski, uslugi, konta lokalne, udzialy, instalacja softu, GPUpdate, Windows Update, logi zdarzen, Defender, zmiana nazwy, firewall, LAPS, diagnostyka sieci). |
| `AD-MoveDisabledUsersGUI.ps1` | Wykrywa konta uzytkownikow z `Enabled=$false`, pozwala je przefiltrowac i przeniesc do wybranego OU za pomoca wbudowanego przegladacza drzewa AD. |
| `AD-NTFS-AuditGUI.ps1` | WPF-owy audytor NTFS: rekursywnie czyta ACL, wykrywa bezposrednie wpisy uzytkownikow, prezentuje je w interfejsie z filtrami i eksportem wynikow. |
| `AD-NTFS-BulkGroupPermissions.ps1` | WinForms do hurtowego nadawania uprawnien NTFS grupom (mapy praw, zakres dziedziczenia, opcja protect/clear) z wklejaniem identyfikatorow ze schowka. |
| `AD-PremissionAudit.ps1` | Prosty panel GUI uruchamiajacy zadanie w tle, ktore przeszukuje drzewo katalogow i wypisuje przypadki kiedy pojedynczy uzytkownik ma ACL, zamiast grupy. |
| `AD-RDP-LoginEvents.ps1` | Skrypt CLI analizujacy zdarzenia 4624/4634 (RemoteInteractive) z ostatnich N dni, paruje logowania z wylogowaniami i buduje raport czasu trwania sesji RDP. |
| `AD-RemoveDisabledUsersFromGroups.ps1` | Narzedzie tekstowe (z opcjonalnym Out-GridView) usuwajace wylaczone konta z niekrytycznych grup AD z kopia zapasowa czlonkostw w CSV. |
| `AD-SwapLogin.ps1` | GUI przyspieszajace naprawianie kont z zamienionym imieniem/nazwiskiem; pozwala hurtowo zaktualizowac pola GivenName/Surname/DisplayName i dziala w trybie testowym. |
| `AD-User-Manager.ps1` | Lekki manager uzytkownikow: wyszukuje po SamAccountName, wyswietla najwazniejsze atrybuty, pozwala blokowac/odblokowac konto i resetowac haslo. |
| `Get-PremissionReport.ps1` | Generator raportu ACL (CSV) dla folderow i plikow wraz z wersja "clean" bez wbudowanych kont; wczytuje katalog przez okno wyboru. |
| `PC-CheckHashFoldersIntegrity.ps1` | WinForms do indeksowania katalogow (hash SHA256) oraz porownywania aktualnej struktury z referencja, co ulatwia wykrycie zmian i sabotażu. |
| `PC-Folder-TreeView.ps1` | Wizualizuje drzewo folderu w ASCII, wspiera ukryte pliki oraz eksport do TXT. |

## Kontakt

Masz uwagi lub pomysl na nowy scenariusz? Otworz Issue na GitHub lub skontaktuj sie z autorem.

**Autor:** Boris
