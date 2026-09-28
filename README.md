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

`NET-SharePermissionAudit.ps1` to skrypt konsolowy z parametrami (bez GUI) — wynik to interaktywny raport HTML oraz pliki CSV:

```powershell
.\NET-SharePermissionAudit.ps1                                            # udziały komputera lokalnego
.\NET-SharePermissionAudit.ps1 -ComputerName FS01, FS02 -Depth 4 -ExpandGroups
.\NET-SharePermissionAudit.ps1 -FromAD -SearchBase 'OU=Serwery,DC=firma,DC=local' -ShareOnly
.\NET-SharePermissionAudit.ps1 -Path '\\FS01\Dzialy\Kadry' -Depth -1 -UserAccess jkowalski
Get-Help .\NET-SharePermissionAudit.ps1 -Full                               # opis wszystkich parametrów
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
| `Pulpit\ShareAudit_<data>\` (lub folder z `-OutputPath`) | `NET-SharePermissionAudit.ps1` | `Raport.html`, pliki CSV (ustalenia, udziały, uprawnienia udziałów i NTFS, członkowie grup, dostęp użytkowników, serwery, błędy), opcjonalnie `Raport.json`/`Raport.xlsx` oraz `Audyt.log`. |
| `%TEMP%\DisabledUsersGroupsBackup.csv` | `AD-RemoveDisabledUsersFromGroups.ps1` | Kopia członkostw w grupach przed usunięciem. |
| `%TEMP%\PC-CheckHashFoldersIntegrity.log` | `PC-CheckHashFoldersIntegrity.ps1` | Log diagnostyczny porównania. |
| `%APPDATA%\CyberGenPro\CyberGenConfig.json`, `%TEMP%\CyberGenHistory.json` | `PasswordGenerator.ps1` | Ustawienia generatora oraz historia haseł z ostatnich 24 h (zapisana jawnym tekstem). |
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
| `AD-ManagerDiamond.ps1` | Kompleksowe narzędzie „Domain Ops” do zdalnego zarządzania komputerami domenowymi — lista zakładek poniżej. |
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
| `NET-SharePermissionAudit.ps1` | Audyt uprawnień udziałów sieciowych na wielu serwerach (lista, OU w AD lub konkretne ścieżki): uprawnienia udziału (NetShareEnum — wystarczy port 445, bez WMI/WinRM), właściwości udziału (ABE, szyfrowanie, offline), NTFS do zadanej głębokości (równolegle, opcjonalnie po stronie serwera przez WinRM). Rozwiązuje SID-y w kontekście serwera, wykrywa ryzyka z wagą (m.in. zapis dla Everyone/Domain Users, dostęp anonimowy, pełna kontrola dla zwykłych kont, uprawnienia bezpośrednio dla użytkowników, osierocone SID-y, wyłączone konta, Deny, przerwane dziedziczenie), opcjonalnie rozwija grupy i liczy efektywny dostęp wskazanych użytkowników. Raport HTML z filtrami + CSV/JSON/XLSX. Tylko odczyt. |
| `PasswordGenerator.ps1` | Generator haseł (WPF, kryptograficzne losowanie) z konfigurowalnymi regułami, generowaniem paczek do schowka, historią z ostatnich 24 h, motywem jasnym/ciemnym i wysyłką przez `mailto:`. |
| `PC-CheckHashFoldersIntegrity.ps1` | WinForms do porównywania dwóch folderów (np. lokalnego z udziałem sieciowym): indeksuje pliki (także ukryte), porównuje rozmiar/datę i opcjonalnie SHA-256, pokazuje różnice i eksportuje je do CSV. |
| `PC-Folder-TreeView.ps1` | Podgląd drzewa folderu (leniwe wczytywanie, opcjonalnie pliki ukryte) z eksportem do TXT w formie ASCII. Nie wchodzi w junctiony/symlinki. |

## AD-ManagerDiamond — zakładki

Po lewej stronie ładujesz komputery z AD (opcjonalnie SearchBase i filtr nazwy) i zaznaczasz hosty docelowe; na górze możesz podać inne poświadczenia. Dostępne zakładki:

- **Polecenia** — uruchamianie poleceń PowerShell/cmd na zaznaczonych hostach.
- **BitLocker** — status woluminów i backup kluczy odzyskiwania do AD.
- **Dyski** — zajętość dysków, czyszczenie TEMP i Kosza.
- **Usługi** — lista usług z filtrem, start/stop/restart.
- **Konta lokalne**, **Udziały (shary)**, **Programy (zainstalowane)**, **Sterowniki (PnP)** (z eksportem CSV), **Certyfikaty (LM)** (z eksportem `.cer`) — podgląd.
- **Instalacja softu** — kopiowanie i ciche uruchamianie instalatora MSI/EXE.
- **GPUpdate**, **Windows Update** (PSWindowsUpdate, jeśli dostępny), **Defender** (status, szybki skan).
- **Zdarzenia (24h)** — błędy i ostrzeżenia z ostatniej doby.
- **Zasilanie/Uptime** — czas pracy, restart, wyłączenie.
- **Konto komputera (AD)** — test/naprawa kanału zaufania, reset konta komputera w AD, włączenie/wyłączenie konta.
- **Zmiana nazwy** — wsadowa zmiana nazw z autonumeracją.
- **Udziały (zarządzanie)** — tworzenie i usuwanie udziałów.
- **Lokalni administratorzy** — podgląd, dodawanie i usuwanie członków (grupa wskazywana po SID, działa na systemach w dowolnym języku).
- **Zadania (Harmonogram)** — podgląd, uruchamianie, włączanie/wyłączanie, usuwanie i tworzenie prostych zadań.
- **Zapora Windows** — reguły przychodzące, przełączanie, tworzenie i usuwanie.
- **LAPS (AD)** — odczyt hasła (LAPS legacy i Windows LAPS), wymuszenie rotacji, kopiowanie do schowka.
- **Diagnostyka łączności** — ping, WSMan, porty 5986/445/3389/135.

## Uwagi bezpieczeństwa

- Przed pierwszym użyciem na produkcji korzystaj z trybów podglądu (WhatIf / tryb testowy), które mają m.in. `AD-BulkUserCreator`, `AD-BulkOUGroupCreator`, `AD-BulkAddUsersToGroups`, `AD-NTFS-BulkGroupPermissions`, `AD-MoveDisabledUsersGUI`, `AD-SwapLogin` i `AD-UpdateUserTitleDepartment`.
- `AD-DeleteNewAccoutsGUI` usuwa konta nieodwracalnie (poza przywróceniem z Kosza AD) — jeśli nie masz pewności, użyj przycisku wyłączenia konta.
- „Reset konta w AD” w `AD-ManagerDiamond` ustawia hasło konta komputera na jego `sAMAccountName` małymi literami; komputer trzeba potem ponownie połączyć z domeną lub naprawić kanał zaufania.
- `PasswordGenerator` i `AD-BulkUserCreator` przechowują wygenerowane hasła jawnym tekstem (historia w `%TEMP%`, kolumna w siatce/schowek) — czyść historię i schowek po pracy.
- Raport `NET-SharePermissionAudit` to pełna mapa „kto ma dostęp do czego” (w tym słabo zabezpieczone lokalizacje) — przechowuj go w miejscu z ograniczonym dostępem i usuń, gdy nie jest już potrzebny.

## Kontakt

Masz uwagi lub pomysł na nowy scenariusz? Otwórz Issue na GitHub lub skontaktuj się z autorem.

**Autor:** Boris
