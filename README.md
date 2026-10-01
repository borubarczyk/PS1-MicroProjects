# PS1-MicroProjects

Zbiór mikroprojektów PowerShell wspomagających codzienną administrację Active Directory oraz systemami Windows. W katalogu znajdziesz rozbudowane narzędzia GUI (WinForms/WPF) i skrypty tekstowe: od tworzenia kont, przez audyty NTFS, po zdalną diagnostykę stacji roboczych.

## Wymagania ogólne

- Windows PowerShell 5.1 lub PowerShell 7 uruchamiany na Windowsie.
- RSAT ActiveDirectory (moduł `ActiveDirectory`) dla wszystkich skryptów zaczynających się od `AD-`.
- Uprawnienia administracyjne odpowiadające operacji (np. tworzenie OU, modyfikacja NTFS, odczyt dziennika Security, zdalne polecenia).
- Sesja desktopowa w trybie STA dla narzędzi WinForms/WPF: `powershell.exe -STA -File .\NazwaSkryptu.ps1`.
- Dla zdalnych operacji (`AD-ManagerDiamond.ps1`, `AD-RDP-LoginEvents.ps1`) włączony WinRM/WMI na hostach docelowych.

## Jak uruchomić skrypt

1. Sklonuj repozytorium lub skopiuj katalog z repo:
   ```powershell
   git clone https://github.com/borubarczyk/PS1-MicroProjects.git
   cd PS1-MicroProjects
   ```
2. (Opcjonalnie) zezwól na uruchamianie w bieżącej sesji: `Set-ExecutionPolicy -Scope Process Bypass -Force`.
3. Uruchom interesujący plik, np.: `powershell.exe -STA -ExecutionPolicy Bypass -File .\AD-BulkUserCreator.ps1`.
4. Narzędzia GUI uruchamiaj w sesji desktopowej z odpowiednimi uprawnieniami.

`AD-DomainGroupTree.ps1` definiuje funkcję, którą trzeba najpierw wczytać (dot-source), a potem wywołać:

```powershell
. .\AD-DomainGroupTree.ps1
Get-ADDomainGroupTree -ExportHtmlPath C:\Temp\GrupyAD.html              # tylko relacje grupa-grupa (szybko)
Get-ADDomainGroupTree -ExportHtmlPath C:\Temp\GrupyAD.html -ShowMembers -IncludeDescription
```

## Kodowanie plików

Skrypty zawierają polskie znaki i są zapisane jako **UTF-8 z BOM**. Windows PowerShell 5.1 bez BOM czyta pliki w kodowaniu systemowym (np. cp1250), przez co polskie znaki psują napisy, a część skryptów w ogóle się nie parsuje. Przy edycji zachowuj kodowanie „UTF-8 with BOM” (w VS Code: *Save with Encoding → UTF-8 with BOM*).

## Pliki pomocnicze i dane zapisywane przez skrypty

| Plik / lokalizacja | Skrypt | Zawartość |
| --- | --- | --- |
| `.AD-BulkUserCreator.json` (katalog skryptu) | `AD-BulkUserCreator.ps1` | Ustawienia zakładek: formaty loginu i nazwy wyświetlanej, domyślne OU/UPN, lista miast. |
| `AD-BulkUserCreator.log` (katalog skryptu) | `AD-BulkUserCreator.ps1` | Log operacji kreatora kont. |
| `.move-disabled-users-ad.config.json` (katalog skryptu) | `AD-MoveDisabledUsersGUI.ps1` | Bazowe i docelowe OU, wyjątki użytkowników, grupy chronione. |
| `Dokumenty\AD-BulkAttributeUpdater\Logs`, `...\Rollback` | `AD-UpdateUserTitleDepartment.ps1` | Logi oraz pliki JSON pozwalające cofnąć wprowadzone zmiany. |
| `Pulpit\AD_Delete_Log_*.log` | `AD-DeleteNewAccoutsGUI.ps1` | Log usuniętych/wyłączonych kont. |
| `Pulpit\NTFS_Audit_*.csv/.xlsx` | `AD-NTFS-AuditGUI.ps1` | Eksport wyników audytu NTFS. |
| `%TEMP%\DisabledUsersGroupsBackup.csv` | `AD-RemoveDisabledUsersFromGroups.ps1` | Kopia członkostw w grupach przed usunięciem. |
| `%TEMP%\PC-CheckHashFoldersIntegrity.log` | `PC-CheckHashFoldersIntegrity.ps1` | Log diagnostyczny porównania. |
| `%APPDATA%\CyberGenPro\CyberGenConfig.json`, `%TEMP%\CyberGenHistory.json` | `PasswordGenerator.ps1` | Ustawienia generatora oraz historia haseł z ostatnich 24 h (zapisana jawnym tekstem). |
| `%APPDATA%\AD-ManagerDiamond\settings.json` | `AD-ManagerDiamond.ps1` | Ustawienia: OU i filtr listy komputerów, kontroler domeny, liczba równoległych hostów, limit połączenia, ostatni moduł, rozmiar okna (bez poświadczeń). |
| `%LOCALAPPDATA%\AD-ManagerDiamond\Logs\DomainOps_RRRRMMDD.log` | `AD-ManagerDiamond.ps1` | Dziennik wszystkich operacji (kto, co, na których hostach, wynik). |
| `%SystemRoot%\Temp\DomainOps` na hostach docelowych | `AD-ManagerDiamond.ps1` | Kopiowane instalatory i ich logi, skrypt i log instalacji aktualizacji (`WU.log`); zadanie Harmonogramu `DomainOps-WindowsUpdate`. |
| `LICENSE` | — | Informacja o licencji zbioru. |

## Lista skryptów

| Skrypt | Opis |
| --- | --- |
| `AD-BulkAccountStatusChecker.ps1` | WinForms do wklejania listy loginów: sprawdza, czy konta istnieją i są włączone, filtruje/sortuje wyniki i eksportuje je do CSV. Duplikaty loginów (bez względu na wielkość liter) są pomijane. |
| `AD-BulkAddUsersToGroups.ps1` | GUI z dwiema siatkami (użytkownicy i grupy, wklejanie z Excela). Rozpoznaje identyfikatory (sAM/UPN/DN/SID, nazwa grupy) i dodaje członków w trybie „każdy do każdej grupy” albo „wiersz do wiersza”; ma tryb dry-run i log. |
| `AD-BulkGroupCreator.ps1` | Tworzenie wielu grup w wybranym OU z jednej tabeli (nazwa, sAMAccountName, opis, zakres, kategoria, e-mail) z normalizacją sAMAccountName i wklejaniem z Excela. Puste zakres/kategoria oznaczają Global/Security. |
| `AD-BulkOUGroupCreator.ps1` | Kreator WinForms budujący wiele OU (miasta) i zestaw grup w każdym OU według szablonów nazw/opisów, z grupami zbiorczymi ALL (miasto, rola, globalna), transliteracją i unikalnym sAM. Ma podgląd WhatIf. |
| `AD-BulkUserCreator.ps1` | Rozbudowany kreator masowego tworzenia kont (zakładki: Uczen, Student, Pracownik, Wykladowca, Inne): walidacja i normalizacja danych, generowanie loginów/e-maili/haseł, wybór domeny UPN i OU, wykrywanie kolizji loginu i nazwy (CN), miasto zapisywane w atrybucie City, tryb WhatIf, kopiowanie danych do schowka, eksport CSV. Po utworzeniu kont uruchamia `Start-ADSyncSyncCycle -PolicyType Delta`. |
| `AD-DeleteNewAccoutsGUI.ps1` | Lista kont utworzonych w ostatnich X minutach; pozwala zaznaczyć i usunąć albo (bezpieczniej) wyłączyć konta. Log trafia na Pulpit. |
| `AD-DomainGroupTree.ps1` | Funkcja `Get-ADDomainGroupTree` generująca raport HTML z drzewem zagnieżdżeń grup AD (wyszukiwarka, rozwijanie/zwijanie), opcjonalnie z członkami i opisami. Wykrywa rzeczywiste pętle zagnieżdżeń. |
| `AD-LastActiviti.ps1` | Interfejs WPF z raportem ostatniego logowania użytkowników (filtr nieaktywności 1/3/6/12 miesięcy, wyszukiwanie po nazwie/loginie/mieście, tylko aktywne), wyłączanie zaznaczonego konta i eksport widoku do CSV. |
| `AD-ManagerDiamond.ps1` | Kompleksowe narzędzie „Domain Ops” do zdalnej administracji komputerami domenowymi: 24 moduły w 7 kategoriach, operacje wykonywane w tle i równolegle na wielu hostach, wyniki w tabeli z filtrem i eksportem CSV — szczegóły poniżej. |
| `AD-MoveDisabledUsersGUI.ps1` | Wyszukuje wyłączone konta spoza bazowego OU, pozwala je zaznaczyć i przenieść do wybranego OU (z opcją usunięcia z grup niechronionych). Ma przeglądarkę drzewa OU, tryb WhatIf (domyślnie włączony) i zakładkę ustawień z wyjątkami. |
| `AD-NTFS-AuditGUI.ps1` | Audytor NTFS w WPF: rekursywnie czyta ACL, rozpoznaje typ podmiotu (użytkownik/grupa), wskazuje foldery z uprawnieniami nadanymi bezpośrednio użytkownikom; widok tabeli i drzewa, filtry, eksport CSV/XLSX. |
| `AD-NTFS-BulkGroupPermissions.ps1` | WinForms do hurtowego nadawania uprawnień NTFS wielu grupom (odczyt … pełna kontrola, zakres dziedziczenia, opcje purge/zablokowania dziedziczenia, WhatIf). Uprawnienia nadawane są po SID grupy. |
| `AD-PremissionAudit.ps1` | Prosty panel GUI: po wybraniu folderu uruchamia zadanie w tle, które przeszukuje drzewo katalogów i wypisuje miejsca, gdzie uprawnienia ma pojedynczy użytkownik zamiast grupy. |
| `AD-RDP-LoginEvents.ps1` | GUI do pobierania zdarzeń z dziennika Security (domyślnie 4624/4634) z ostatnich N dni z wybranych komputerów lub wszystkich kontrolerów domeny. Działa w tle, zapisuje raport CSV do wskazanego folderu; konfigurację można zapisać/wczytać z JSON. |
| `AD-RemoveDisabledUsersFromGroups.ps1` | Skrypt tekstowy (z opcjonalnym wyborem w Out-GridView) usuwający wyłączone konta z niekrytycznych grup AD, z kopią członkostw w CSV. **Domyślnie działa tylko w trybie podglądu** — `Remove-ADGroupMember` ma na sztywno `-WhatIf`, który trzeba usunąć, aby faktycznie wykonać zmiany. |
| `AD-SwapLogin.ps1` | GUI do naprawiania kont z zamienionym imieniem i nazwiskiem: sprawdza listę loginów, podpowiada konta do zamiany, hurtowo zamienia GivenName/Surname (opcjonalnie DisplayName); ma tryb testowy i eksport CSV. |
| `AD-UpdateUserTitleDepartment.ps1` | Kreator WinForms do hurtowej aktualizacji atrybutów (title, department i inne) z CSV: podgląd pliku, mapowanie kolumn, weryfikacja kont, podgląd zmian „było/będzie”, tryb WHATIF (domyślnie włączony) i cofanie zmian z pliku rollback. |
| `AD-User-Manager.ps1` | Lekki manager użytkowników: wyszukuje po SamAccountName, wyświetla najważniejsze atrybuty (hasło, logowania, blokady), pozwala wyłączyć/włączyć konto, odblokować po lockoucie i zresetować hasło. |
| `File-HashChecker.ps1` | Aplikacja WinForms licząca hashe MD5/SHA1/SHA256 dla wskazanych plików lub folderów (rekurencyjnie, również przez przeciągnij-i-upuść), z paskiem postępu i zakładką historii. |
| `Get-PremissionReport.ps1` | Generator raportu ACL (CSV, separator `;`) dla wybranego folderu i jego podfolderów, wraz z wersją „clean” bez kont wbudowanych i nierozwiązanych SID. Folder źródłowy i miejsce zapisu wybierane są w oknach dialogowych. |
| `PasswordGenerator.ps1` | Generator haseł (WPF, kryptograficzne losowanie) z konfigurowalnymi regułami, generowaniem paczek do schowka, historią z ostatnich 24 h, motywem jasnym/ciemnym i wysyłką przez `mailto:`. |
| `PC-CheckHashFoldersIntegrity.ps1` | WinForms do porównywania dwóch folderów (np. lokalnego z udziałem sieciowym): indeksuje pliki (także ukryte), porównuje rozmiar/datę i opcjonalnie SHA-256, pokazuje różnice i eksportuje je do CSV. |
| `PC-Folder-TreeView.ps1` | Podgląd drzewa folderu (leniwe wczytywanie, opcjonalnie pliki ukryte) z eksportem do TXT w formie ASCII. Nie wchodzi w junctiony/symlinki. |

## AD-ManagerDiamond — Domain Ops

Uruchomienie: `powershell.exe -STA -ExecutionPolicy Bypass -File .\AD-ManagerDiamond.ps1` (skrypt sam przełączy się w tryb STA, jeśli trzeba; działa też w PowerShell 7 na Windows).

**Układ okna**

- **Lewy panel — komputery docelowe.** Lista z Active Directory (OU wybierane z drzewa, filtr nazwy z `*`, opcja „tylko włączone konta”), dopisywana ręcznie albo wczytywana z pliku TXT/CSV. Szybkie wyszukiwanie, zaznaczanie checkboxami, spacją, kliknięciem nagłówka lub z menu kontekstowego. Operacje dotyczą zaznaczonych komputerów.
- **Drzewo modułów** pogrupowanych w kategorie; ostatnio używany moduł otwiera się przy starcie.
- **Tabela wyników** w każdym module: sortowanie po kolumnie, filtr po wszystkich kolumnach, eksport widocznych wierszy do CSV (UTF-8, separator z ustawień regionalnych — otwiera się poprawnie w Excelu), kopiowanie do schowka, podgląd szczegółów wiersza dwuklikiem. Akcje typu „zatrzymaj usługę”, „usuń regułę” działają na wierszach zaznaczonych w tabeli — każdy na właściwym hoście.
- **Pasek górny:** bieżący użytkownik albo poświadczenia alternatywne (używane także w poleceniach AD), liczba hostów obsługiwanych równolegle, limit czasu połączenia WinRM, opcjonalny kontroler domeny.
- **Dziennik operacji** na dole (kolorowany, równolegle zapisywany do pliku) i **pasek stanu** z postępem oraz przyciskiem anulowania.

Wszystkie operacje wykonywane są w tle (pula wątków PowerShell), równolegle dla wielu hostów — okno nie zawiesza się, a wyniki pojawiają się na bieżąco. Host niedostępny lub zwracający błąd pojawia się w tabeli jako wiersz ze statusem „Błąd” i opisem przyczyny. Operacje zmieniające stan wymagają potwierdzenia z listą hostów/obiektów.

**Moduły**

| Kategoria | Moduł | Możliwości |
| --- | --- | --- |
| Diagnostyka | Łączność | DNS, ping (czas), porty TCP (lista do edycji), test sesji PowerShell Remoting — wykonywane lokalnie. |
| | Inwentaryzacja | Producent, model, numer seryjny, BIOS, system i kompilacja, procesor, RAM, dysk systemowy, zalogowany użytkownik, IP/MAC, data instalacji i ostatniego startu. |
| | Zasilanie i uptime | Czas pracy, oczekujący restart (CBS, Windows Update, operacje na plikach, zmiana nazwy, SCCM); restart/wyłączenie z opóźnieniem i komunikatem, anulowanie zaplanowanego. |
| Zdalne wykonanie | Polecenia | PowerShell lub cmd.exe (polskie znaki w wyniku cmd), szablony typowych poleceń, pełny wynik w podglądzie wiersza. |
| | Instalacja oprogramowania | MSI/MSP/MSU/EXE: kopiowanie (udział ADMIN$ albo WinRM), cicha instalacja, kod wyjścia z opisem (np. 3010 = wymagany restart), log MSI, opcjonalne usunięcie instalatora. |
| | Aktualizacja zasad grupy | `gpupdate` dla komputera i/lub użytkownika, z `/force` lub bez. |
| System | Usługi | Filtr nazwy i stanu (m.in. „automatyczne, ale zatrzymane”), start/stop/restart, zmiana typu uruchamiania. |
| | Procesy | Pamięć, właściciel, wiersz poleceń; kończenie zaznaczonych procesów. |
| | Dyski | Zajętość; czyszczenie Windows\Temp, TEMP profili i Kosza (pliki starsze niż N dni) z raportem odzyskanego miejsca. |
| | Dziennik zdarzeń | Dzienniki System/Application/Security, poziomy, ID zdarzeń, zakres godzin, limit na host. |
| | Harmonogram zadań | Podgląd (opcjonalnie bez zadań \Microsoft\), uruchom/zatrzymaj/włącz/wyłącz/usuń, tworzenie prostych zadań SYSTEM (logowanie, start, codziennie, jednorazowo, na żądanie). |
| | Sterowniki i urządzenia | Sterowniki PnP z filtrem oraz urządzenia zgłaszające problem w Menedżerze urządzeń. |
| Oprogramowanie | Zainstalowane programy | Rejestr 64/32-bit i profili zalogowanych użytkowników, opcjonalnie aktualizacje i składniki systemowe; ciche odinstalowanie (MSI lub `QuietUninstallString`). |
| | Windows Update | Wyszukiwanie dostępnych aktualizacji i historia (API Windows Update), instalacja przez zadanie SYSTEM z opcjonalnym restartem, podgląd postępu z logu — bez modułu PSWindowsUpdate. |
| Bezpieczeństwo | Microsoft Defender | Stan ochrony i sygnatur, wykryte zagrożenia, aktualizacja sygnatur, szybki i pełny skan. |
| | BitLocker | Stan woluminów (także na systemach bez modułu BitLocker), kopia kluczy odzyskiwania do AD, odczyt kluczy zapisanych w AD i kopiowanie hasła odzyskiwania. |
| | Zapora Windows | Stan profili, reguły z portami/programami/adresami, włączanie/wyłączanie/usuwanie zaznaczonych, tworzenie reguł (grupa „Domain Ops”). |
| | Certyfikaty komputera | Magazyny LocalMachine, filtr, wygasające w ciągu N dni, eksport zaznaczonych do `.cer`. |
| | Lokalni administratorzy | Członkowie grupy wyznaczanej po SID (działa w każdym języku systemu, także z osieroconymi SID-ami), dodawanie i usuwanie; wbudowane konto Administrator i Domain Admins są chronione przed usunięciem. |
| | Konta lokalne | Stan, ostatnie logowanie, wiek hasła; włączanie, wyłączanie, ustawianie hasła. |
| Udostępnianie | Udziały sieciowe | Podgląd (opcjonalnie z administracyjnymi), uprawnienia udziałów, tworzenie z uprawnieniami udziału i opcjonalnie NTFS, usuwanie (bez udziałów administracyjnych). |
| Active Directory | Konto komputera | Informacje z AD, test i naprawa kanału zaufania, włączanie/wyłączanie, przenoszenie do OU, reset konta. |
| | LAPS | Windows LAPS (także hasła szyfrowane) i LAPS legacy, wymuszenie zmiany hasła z opcjonalnym przetworzeniem zasad na hoście. |
| | Zmiana nazwy komputerów | Wczytanie zaznaczonych, autonumeracja (prefiks, numer, liczba cyfr, sufiks), walidacja nazw NetBIOS i duplikatów, edycja w tabeli, opcjonalny restart. |

## Uwagi bezpieczeństwa

- Przed pierwszym użyciem na produkcji korzystaj z trybów podglądu (WhatIf / tryb testowy), które mają m.in. `AD-BulkUserCreator`, `AD-BulkOUGroupCreator`, `AD-BulkAddUsersToGroups`, `AD-NTFS-BulkGroupPermissions`, `AD-MoveDisabledUsersGUI`, `AD-SwapLogin` i `AD-UpdateUserTitleDepartment`.
- `AD-DeleteNewAccoutsGUI` usuwa konta nieodwracalnie (poza przywróceniem z Kosza AD) — jeśli nie masz pewności, użyj przycisku wyłączenia konta.
- „Resetuj konto” w module *Konto komputera* `AD-ManagerDiamond` działa jak polecenie z konsoli ADUC: ustawia hasło konta komputera na domyślne (nazwa komputera małymi literami, bez `$`); komputer trzeba potem ponownie dołączyć do domeny lub naprawić kanał zaufania.
- `AD-ManagerDiamond` maskuje hasła LAPS i klucze odzyskiwania BitLocker w tabeli, eksporcie i kopiowaniu (do czasu zaznaczenia „Pokaż wartości poufne”), a hasło skopiowane przyciskiem „Kopiuj hasło…” usuwa ze schowka po 60 s. Zmiana nazwy komputera i naprawa kanału zaufania wymagają poświadczeń domenowych — program poprosi o nie, jeśli nie ustawiono poświadczeń alternatywnych.
- `PasswordGenerator` i `AD-BulkUserCreator` przechowują wygenerowane hasła jawnym tekstem (historia w `%TEMP%`, kolumna w siatce/schowek) — czyść historię i schowek po pracy.

## Kontakt

Masz uwagi lub pomysł na nowy scenariusz? Otwórz Issue na GitHub lub skontaktuj się z autorem.

**Autor:** Boris
