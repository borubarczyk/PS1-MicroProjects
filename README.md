# PS1-MicroProjects

Zbiór mikroprojektów PowerShell wspomagających codzienną administrację Active Directory oraz systemami Windows. W katalogu znajdziesz rozbudowane narzędzia GUI (WinForms/WPF) i skrypty tekstowe: od tworzenia kont, przez audyty NTFS, po zdalną diagnostykę stacji roboczych.

## Wymagania ogólne

- Windows PowerShell 5.1 lub PowerShell 7 uruchamiany na Windowsie.
- RSAT ActiveDirectory (moduł `ActiveDirectory`) dla wszystkich skryptów zaczynających się od `AD-`.
- Uprawnienia administracyjne odpowiadające operacji (np. tworzenie OU, modyfikacja NTFS, odczyt dziennika Security, zdalne polecenia).
- Sesja desktopowa w trybie STA dla narzędzi WinForms/WPF: `powershell.exe -STA -File .\NazwaSkryptu.ps1`.
- Dla zdalnych operacji (`AD-ManagerDiamond.ps1`, `AD-RDP-LoginEvents.ps1`) włączony WinRM/WMI na hostach docelowych. `AD-ManagerDiamond.ps1` działa w Windows PowerShell 5.1 (uruchomiony z PowerShell 7 sam przełączy się na `powershell.exe`).

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
| `%APPDATA%\AD-ManagerDiamond\settings.json` | `AD-ManagerDiamond.ps1` | Ustawienia: OU i filtry list komputerów i użytkowników, kontroler domeny, liczba równoległych operacji, limit połączenia, próg nieaktywności, ostatnia przestrzeń i moduły, rozmiar okna, panel dziennika i szczegółów (bez poświadczeń). |
| `AD-ManagerDiamond.Modules\*.ps1` (katalog skryptu) | `AD-ManagerDiamond.ps1` | Opcjonalne własne moduły wczytywane przy starcie. |
| `%LOCALAPPDATA%\AD-ManagerDiamond\Logs\DomainOps_RRRRMMDD.log` | `AD-ManagerDiamond.ps1` | Dziennik wszystkich operacji (kto, co, na których komputerach/kontach, wynik). |
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
| `AD-ManagerDiamond.ps1` | Centrum administracji domeną „Domain Ops” (WPF, ciemny motyw): trzy przestrzenie robocze — zarządzanie zdalne komputerami, użytkownicy AD i komputery AD — z 41 modułami, m.in. weryfikacją i czyszczeniem profili użytkowników względem AD, audytem bezpieczeństwa, raportami kont i źródłem blokad. Operacje w tle i równolegle, wyniki z filtrem, eksportem CSV/HTML i akcjami pod prawym przyciskiem; możliwość dodawania własnych modułów — szczegóły poniżej. |
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

Uruchomienie: `powershell.exe -ExecutionPolicy Bypass -File .\AD-ManagerDiamond.ps1`. Interfejs jest napisany w WPF dla Windows PowerShell 5.1 — skrypt sam uruchomi się ponownie w trybie STA, a z PowerShell 7 przełączy się na `powershell.exe`.

**Przestrzenie robocze** (przełącznik na górnym pasku, skróty `Ctrl+1…3`):

| Przestrzeń | Lista po lewej | Do czego służy |
| --- | --- | --- |
| **Zarządzanie zdalne** | komputery | Operacje na komputerach przez PowerShell Remoting (WinRM): diagnostyka, profile i sesje użytkowników, grupy i konta lokalne, usługi, procesy, zdarzenia, oprogramowanie, aktualizacje, bezpieczeństwo, udziały. |
| **Użytkownicy AD** | konta użytkowników | Konta w Active Directory: szczegóły, hasła i blokady, stan konta, atrybuty, grupy, raporty i źródło blokad. |
| **Komputery AD** | komputery | Konta komputerów w AD: informacje i kanał zaufania, LAPS, klucze BitLocker, zmiana nazw, grupy, raporty. |

**Układ okna**

- **Lewy panel — obiekty docelowe.** Komputery: z AD (OU z drzewa, filtr nazwy z `*`, tylko włączone), wpisane ręcznie albo z pliku TXT/CSV. Użytkownicy: wyszukiwanie w AD (login, nazwisko, e-mail, gwiazdka), stan konta (aktywne / wyłączone / zablokowane), OU, loginy wpisane ręcznie lub z pliku. Zaznaczanie polami wyboru (także kilku wierszy naraz i spacją), szybkie wyszukiwanie, kolorowa kropka stanu (np. wynik testu łączności, konto zablokowane), menu kontekstowe (pulpit zdalny, konsola zarządzania, usługi, podgląd zdarzeń, `C$`, kopiowanie nazw / adresów e-mail).
- **Nawigacja modułów** pogrupowanych w kategorie — każda przestrzeń pamięta ostatnio otwarty moduł; kropka przy module oznacza trwającą operację.
- **Moduł:** nagłówek z opisem, panel parametrów (zwijany), kafelki z podsumowaniem (np. kandydaci do usunięcia, średni wynik audytu), tabela wyników z sortowaniem, filtrem po wszystkich kolumnach (`Ctrl+F`, słowo z minusem wyklucza wiersze), kolorowymi „pigułkami” stanu, panelem szczegółów wiersza, menu pod prawym przyciskiem myszy (akcje modułu, kopiowanie komórki/wierszy) oraz eksportem do **CSV** lub **raportu HTML**. `F5` uruchamia główną akcję modułu.
- **Górny pasek:** konto używane do operacji (bieżące albo alternatywne — także dla poleceń AD) i ustawienia (kontroler domeny, liczba równoległych operacji, limit połączenia WinRM, domyślny próg nieaktywności, foldery dziennika, ustawień i modułów).
- **Dziennik operacji** (`Ctrl+L`, z licznikiem nowych ostrzeżeń), **pasek stanu** z postępem i przyciskiem „Przerwij” oraz **powiadomienia** w rogu okna.

Wszystkie operacje wykonywane są w tle (pula wątków PowerShell), równolegle dla wielu komputerów/kont — okno nie zawiesza się, a wyniki pojawiają się na bieżąco. Niedostępny komputer lub błąd pojawia się w tabeli jako wiersz „Błąd” z przyczyną. Operacje zmieniające stan wymagają potwierdzenia z listą obiektów (przy niszczących domyślnym przyciskiem jest „Anuluj”).

**Profile użytkowników — co można usunąć.** Moduł *Profile użytkowników* zbiera profile z zaznaczonych komputerów (ostatnie użycie z `ProfileList`, opcjonalnie rozmiar) i sprawdza każde konto w AD przez ADSI (bez RSAT). Kandydaci do usunięcia: konta **usunięte z AD**, **wyłączone**, **wygasłe**, usunięte konta lokalne, profile **tymczasowe/uszkodzone** oraz — opcjonalnie — profile nieużywane dłużej niż N dni. Profile załadowane (użytkownik zalogowany) i systemowe nigdy nie są kandydatami. Przycisk *Zaznacz kandydatów* → *Usuń zaznaczone profile* usuwa profile przez `Win32_UserProfile` (folder i wpis w rejestrze), z potwierdzeniem pokazującym ścieżki, rozmiary i ostrzeżeniem, jeśli zaznaczono profil, który nie jest kandydatem.

**Moduły — Zarządzanie zdalne**

| Kategoria | Moduł | Możliwości |
| --- | --- | --- |
| Diagnostyka | Łączność | DNS, ping, porty TCP, test sesji PowerShell — lokalnie; koloruje kropki na liście, przycisk „zaznacz tylko dostępne”. |
| | Inwentaryzacja | Producent, model, numer seryjny, BIOS, system i kompilacja, procesor, RAM, dysk systemowy, TPM, Secure Boot, zalogowany użytkownik, IP/MAC, daty instalacji i startu. |
| | Wydajność | Obciążenie procesora i pamięci, wolne miejsce, kolejka dysku, procesy zużywające najwięcej CPU i pamięci, ocena. |
| | Zasilanie i uptime | Czas pracy, oczekujący restart (CBS, Windows Update, pliki, zmiana nazwy, SCCM); restart/wyłączenie z opóźnieniem i komunikatem, anulowanie. |
| Użytkownicy i dostęp | Profile użytkowników | Weryfikacja profili względem AD, kandydaci do usunięcia, rozmiar, usuwanie (opis wyżej). |
| | Sesje użytkowników | Zalogowani użytkownicy (konsola/RDP), bezczynność; wiadomość, rozłączenie, wylogowanie, podgląd sesji (shadow). |
| | Grupy lokalne | Administratorzy, Użytkownicy pulpitu zdalnego, zarządzania zdalnego (WinRM), dziennika zdarzeń, kopii zapasowych — po SID (każdy język systemu); dodawanie (także wybór grup z AD) i usuwanie; wbudowany Administrator i Domain Admins chronione. |
| | Konta lokalne | Stan, ostatnie logowanie, wiek hasła; włączanie, wyłączanie, ustawianie hasła. |
| | Pulpit zdalny | Stan RDP, NLA, port, reguły zapory, członkowie grupy RDP; włączanie/wyłączanie RDP razem z zaporą, NLA. |
| Zdalne wykonanie | Polecenia | PowerShell lub cmd.exe (polskie znaki), szablony typowych poleceń (DNS, Kerberos, gpresult, kanał zaufania, czas, DISM/SFC). |
| | Instalacja oprogramowania | MSI/MSP/MSU/EXE: kopiowanie (ADMIN$ albo WinRM), cicha instalacja, kod wyjścia z opisem, log MSI. |
| | Aktualizacja zasad grupy | `gpupdate` dla komputera i/lub użytkownika, opcjonalnie `/force`. |
| System | Usługi | Filtr nazwy i stanu, start/stop/restart, typ uruchamiania (także z menu wiersza). |
| | Procesy | Pamięć, właściciel, sesja, wiersz poleceń; kończenie procesów. |
| | Dyski | Zajętość z oceną; czyszczenie Temp, TEMP profili, Kosza i pobranych aktualizacji z raportem odzyskanego miejsca. |
| | Dziennik zdarzeń | Gotowe zestawy (nieoczekiwane wyłączenia, błędy dysków, awarie aplikacji, nieudane logowania, blokady, GPO, Windows Update) albo własne dzienniki/poziomy/ID. |
| | Harmonogram zadań | Podgląd, uruchom/zatrzymaj/włącz/wyłącz/usuń, tworzenie prostych zadań SYSTEM. |
| | Autostart | Klucze Run/RunOnce (komputer i zalogowani użytkownicy) i foldery Autostart; usuwanie wpisów. |
| | Drukarki | Drukarki, sterowniki, porty, zadania w kolejce; czyszczenie kolejek, usuwanie, restart bufora wydruku. |
| | Sterowniki i urządzenia | Sterowniki PnP z filtrem i urządzenia z problemami. |
| Oprogramowanie | Zainstalowane programy | Rejestr 64/32-bit i profili zalogowanych; ciche odinstalowanie; „pokaż ten program na wszystkich komputerach”. |
| | Windows Update | Dostępne aktualizacje, historia, **sprawdzenie obecności poprawek KB**, instalacja przez zadanie SYSTEM z opcjonalnym restartem, postęp z logu. |
| Bezpieczeństwo | Szybki audyt bezpieczeństwa | Zapora, antywirus i sygnatury, BitLocker, SMBv1, NLA, UAC, LAPS, świeżość aktualizacji, oczekujący restart, Secure Boot, TPM, WDigest, ochrona LSA, LLMNR, konto Gość — wynik procentowy i lista problemów. |
| | Microsoft Defender | Stan ochrony i sygnatur, zagrożenia, aktualizacja sygnatur, szybki i pełny skan. |
| | BitLocker | Stan woluminów, kopia kluczy odzyskiwania do AD. |
| | Zapora Windows | Profile, reguły z portami/programami/adresami, włączanie/wyłączanie/usuwanie, nowe reguły. |
| | Certyfikaty komputera | Magazyny LocalMachine, ważność z oceną, wygasające w ciągu N dni, eksport `.cer`. |
| Udostępnianie | Udziały sieciowe | Podgląd, uprawnienia, tworzenie (z uprawnieniami udziału i opcjonalnie NTFS), usuwanie. |

**Moduły — Użytkownicy AD**

| Kategoria | Moduł | Możliwości |
| --- | --- | --- |
| Konta | Szczegóły konta | Stan, kontakt, przełożony, logowania, hasło, wygaśnięcie, profil, OU, SID. |
| | Hasło i blokada | Stan haseł i blokad; odblokowanie; reset hasła — **losowe, inne dla każdego konta** (widoczne jako wartości poufne, kopiowanie z czyszczeniem schowka) albo wpisane; wymuszenie zmiany przy logowaniu; „hasło nigdy nie wygasa”. |
| | Stan konta | Włączanie i wyłączanie (z dopiskiem daty/autora/powodu w opisie i przeniesieniem do OU), data wygaśnięcia, przenoszenie do OU. |
| | Edycja atrybutów | Odczyt i masowa zmiana atrybutów (stanowisko, dział, firma, biuro, telefony, e-mail, przełożony, extensionAttribute1–15…) z polami `{login}`, `{imie}`, `{nazwisko}`, `{nazwa}`. |
| Grupy | Członkostwo w grupach | Grupy bezpośrednie i zagnieżdżone, dodawanie (wyszukiwarka grup), usuwanie, kopiowanie członkostwa z konta wzorcowego. |
| Raporty | Raporty kont | Zablokowane, wyłączone, nieaktywne, nigdy nie logowane, hasło wygasa / wygasło / nigdy nie wygasa, konta wygasające, nowe, uprzywilejowane (`adminCount`); wyniki można zaznaczyć na liście kont i od razu wykonać na nich operacje. |
| | Źródło blokady konta | Zdarzenia 4740 z emulatora PDC (lub wszystkich DC) — komputer, z którego przyszły błędne hasła; opcjonalnie 4771/4776 z adresem IP/stacją. |

**Moduły — Komputery AD**

| Kategoria | Moduł | Możliwości |
| --- | --- | --- |
| Konta komputerów | Konto komputera | Informacje, test i naprawa kanału zaufania, włączanie/wyłączanie, opis, przenoszenie do OU, reset, usuwanie, tworzenie nowych kont (pre-staging). |
| | Zmiana nazwy komputerów | Autonumeracja, mapowanie z listy (np. z Excela), edycja w tabeli, walidacja NetBIOS i duplikatów, opcjonalny restart. |
| | Członkostwo w grupach | Jak dla użytkowników — dla kont komputerów. |
| Hasła i klucze | LAPS | Windows LAPS (także szyfrowane) i LAPS legacy, kopiowanie hasła (dwuklik), wymuszenie zmiany z przetworzeniem zasad. |
| | Klucze BitLocker (AD) | Klucze zaznaczonych komputerów oraz **wyszukiwanie komputera po identyfikatorze klucza** z ekranu odzyskiwania. |
| Raporty | Raporty komputerów | Nieaktywne, wyłączone, nowe, podsumowanie systemów, nieobsługiwane systemy, bez LAPS, bez klucza BitLocker w AD, serwery; zaznaczanie na liście, wyłączanie, przenoszenie, usuwanie. |

**Rozbudowa — własne moduły.** Pliki `*.ps1` z folderu `AD-ManagerDiamond.Modules` (obok skryptu) są wczytywane przy starcie i mogą dodawać przestrzenie robocze oraz moduły bez zmiany głównego pliku (folder otworzysz z okna Ustawienia). Przykład:

```powershell
# AD-ManagerDiamond.Modules\Czas.ps1
Register-Module -Workspace 'Remote' -Category 'Diagnostyka' -Key 'TimeSync' -Title 'Synchronizacja czasu' -Icon 'E916' `
    -Description 'Źródło czasu i przesunięcie zegara (w32tm).' -Build {
    param($m)
    $row = Add-ToolbarRow -Module $m -Title 'Akcje'
    Add-Button -Parent $row -Text 'Sprawdź' -Icon 'E72C' -Module $m -Primary -OnClick {
        param($m)
        $targets = @(Get-TargetComputers)
        if (-not $targets) { return }
        Start-HostOperation -Module $m -Name 'Czas' -Targets $targets -ScriptBlock {
            param($P)
            [pscustomobject]@{ 'Źródło' = (w32tm /query /source).Trim(); 'Czas' = Get-Date }
        }
    } | Out-Null
}
```

`Register-Workspace -Key -Title -Icon -Target Computer|User|None` dodaje nową przestrzeń; `Start-HostOperation` (zdalnie lub `-Local`) i `Start-AdOperation` (moduł ActiveDirectory z gotową hashtablą `$ad`) uruchamiają operacje w tle.

## Uwagi bezpieczeństwa

- Przed pierwszym użyciem na produkcji korzystaj z trybów podglądu (WhatIf / tryb testowy), które mają m.in. `AD-BulkUserCreator`, `AD-BulkOUGroupCreator`, `AD-BulkAddUsersToGroups`, `AD-NTFS-BulkGroupPermissions`, `AD-MoveDisabledUsersGUI`, `AD-SwapLogin` i `AD-UpdateUserTitleDepartment`.
- `AD-DeleteNewAccoutsGUI` usuwa konta nieodwracalnie (poza przywróceniem z Kosza AD) — jeśli nie masz pewności, użyj przycisku wyłączenia konta.
- „Resetuj konto” w module *Konto komputera* `AD-ManagerDiamond` działa jak polecenie z konsoli ADUC: ustawia hasło konta komputera na domyślne (nazwa komputera małymi literami, bez `$`); komputer trzeba potem ponownie dołączyć do domeny lub naprawić kanał zaufania.
- `AD-ManagerDiamond` maskuje hasła LAPS, klucze odzyskiwania BitLocker i nowo wygenerowane hasła użytkowników w tabeli, panelu szczegółów, eksporcie CSV/HTML i kopiowaniu (do czasu zaznaczenia „Pokaż poufne”); nie są też przeszukiwane filtrem. Hasło skopiowane akcją „Kopiuj hasło…” jest usuwane ze schowka po 60 s. Zmiana nazwy komputera i naprawa kanału zaufania wymagają poświadczeń domenowych — program poprosi o nie, jeśli nie ustawiono poświadczeń alternatywnych.
- Usuwanie profili w `AD-ManagerDiamond` jest nieodwracalne (folder profilu i wpis w rejestrze). Przed usunięciem sprawdź kolumnę „Ocena” — profile oznaczone jako kandydaci to m.in. konta usunięte/wyłączone/wygasłe w AD; nieużywane profile aktywnych kont są kandydatami tylko przy włączonej opcji. Profile zalogowanych użytkowników są zawsze pomijane.
- Moduły z folderu `AD-ManagerDiamond.Modules` są wykonywane z uprawnieniami użytkownika programu — trzymaj tam tylko zaufane pliki.
- `PasswordGenerator` i `AD-BulkUserCreator` przechowują wygenerowane hasła jawnym tekstem (historia w `%TEMP%`, kolumna w siatce/schowek) — czyść historię i schowek po pracy.

## Kontakt

Masz uwagi lub pomysł na nowy scenariusz? Otwórz Issue na GitHub lub skontaktuj się z autorem.

**Autor:** Boris
