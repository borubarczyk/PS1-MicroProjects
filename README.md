# PS1-MicroProjects

Zbiór mikroprojektów PowerShell wspomagających codzienną administrację Active Directory oraz systemami Windows. W katalogu znajdziesz rozbudowane narzędzia GUI (WinForms/WPF) i skrypty tekstowe: od tworzenia kont, przez audyty NTFS, po zdalną diagnostykę stacji roboczych.

## Wymagania ogólne

- Windows PowerShell 5.1 lub PowerShell 7 uruchamiany na Windowsie.
- RSAT ActiveDirectory (moduł `ActiveDirectory`) dla wszystkich skryptów zaczynających się od `AD-`.
- Uprawnienia administracyjne odpowiadające operacji (np. tworzenie OU, modyfikacja NTFS, odczyt dziennika Security, zdalne polecenia).
- Sesja desktopowa w trybie STA dla narzędzi WinForms/WPF: `powershell.exe -STA -File .\NazwaSkryptu.ps1`.
- Dla zdalnych operacji (`AD-ManagerDiamond.ps1`, `AD-RDP-LoginEvents.ps1`) włączony WinRM/WMI na hostach docelowych (w `AD-ManagerDiamond.ps1` także opcjonalnie na serwerze plików dla uprawnień NTFS). `AD-ManagerDiamond.ps1` działa w Windows PowerShell 5.1 (uruchomiony z PowerShell 7 sam przełączy się na `powershell.exe`).

## Jak uruchomić skrypt

1. Sklonuj repozytorium lub skopiuj katalog z repo:
   ```powershell
   git clone https://github.com/borubarczyk/PS1-MicroProjects.git
   cd PS1-MicroProjects
   ```
2. (Opcjonalnie) zezwól na uruchamianie w bieżącej sesji: `Set-ExecutionPolicy -Scope Process Bypass -Force`.
3. Uruchom interesujący plik, np.: `powershell.exe -ExecutionPolicy Bypass -File .\AD-ManagerDiamond.ps1`.
4. Narzędzia GUI uruchamiaj w sesji desktopowej z odpowiednimi uprawnieniami.

## Kodowanie plików

Skrypty zawierają polskie znaki i są zapisane jako **UTF-8 z BOM**. Windows PowerShell 5.1 bez BOM czyta pliki w kodowaniu systemowym (np. cp1250), przez co polskie znaki psują napisy, a część skryptów w ogóle się nie parsuje. Przy edycji zachowuj kodowanie „UTF-8 with BOM” (w VS Code: *Save with Encoding → UTF-8 with BOM*).

## Pliki pomocnicze i dane zapisywane przez skrypty

| Plik / lokalizacja | Skrypt | Zawartość |
| --- | --- | --- |
| `Pulpit\NTFS_Audit_*.csv/.xlsx` | `AD-NTFS-AuditGUI.ps1` | Eksport wyników audytu NTFS. |
| `%TEMP%\PC-CheckHashFoldersIntegrity.log` | `PC-CheckHashFoldersIntegrity.ps1` | Log diagnostyczny porównania. |
| `%APPDATA%\CyberGenPro\CyberGenConfig.json`, `%TEMP%\CyberGenHistory.json` | `PasswordGenerator.ps1` | Ustawienia generatora oraz historia haseł z ostatnich 24 h (zapisana jawnym tekstem). |
| `%APPDATA%\AD-ManagerDiamond\settings.json` | `AD-ManagerDiamond.ps1` | Ustawienia: OU i filtry list komputerów, użytkowników i grup, kontroler domeny, liczba równoległych operacji i zapytań AD, limit połączenia, próg nieaktywności, profile tworzenia kont, OU dla wyłączonych kont, wyjątki i grupy chronione, konta systemowe ukrywane w raporcie NTFS, ostatnio użyte wartości pól modułów, ostatnia przestrzeń i moduły, rozmiar okna (bez poświadczeń i haseł). |
| `%APPDATA%\AD-ManagerDiamond\Backup\Czlonkostwa_wylaczonych_*.csv` | `AD-ManagerDiamond.ps1` | Kopia członkostw w grupach zapisywana przed usunięciem wyłączonych kont z grup (moduł *Wyłączone konta*); z niej działa przywracanie. |
| `%APPDATA%\AD-ManagerDiamond\Rollback\Import_*.json` | `AD-ManagerDiamond.ps1` | Poprzednie wartości atrybutów zmienionych modułem *Import atrybutów z CSV* — do cofnięcia zmian. |
| `%APPDATA%\AD-ManagerDiamond\AclBackup\Uprawnienia_*.json` | `AD-ManagerDiamond.ps1` | Kopie uprawnień NTFS (właściciel, wpisy i dziedziczenie każdego zmienianego folderu w postaci SDDL) zapisywane przed każdą zmianą narzędzi NTFS i na żądanie — z nich działa przywracanie w module *Kopie uprawnień*. |
| `AD-ManagerDiamond.Modules\*.ps1` (katalog skryptu) | `AD-ManagerDiamond.ps1` | Opcjonalne własne moduły wczytywane przy starcie. |
| `%LOCALAPPDATA%\AD-ManagerDiamond\Logs\DomainOps_RRRRMMDD.log` | `AD-ManagerDiamond.ps1` | Dziennik wszystkich operacji (kto, co, na których komputerach/kontach, wynik). |
| `%SystemRoot%\Temp\DomainOps` na hostach docelowych | `AD-ManagerDiamond.ps1` | Kopiowane instalatory i ich logi, skrypt i log instalacji aktualizacji (`WU.log`); zadanie Harmonogramu `DomainOps-WindowsUpdate`. |
| `LICENSE` | — | Informacja o licencji zbioru. |

## Lista skryptów

| Skrypt | Opis |
| --- | --- |
| `AD-BulkAccountStatusChecker.ps1` | WinForms do wklejania listy loginów: sprawdza, czy konta istnieją i są włączone, filtruje/sortuje wyniki i eksportuje je do CSV. Duplikaty loginów (bez względu na wielkość liter) są pomijane. |
| `AD-LastActiviti.ps1` | Interfejs WPF z raportem ostatniego logowania użytkowników (filtr nieaktywności 1/3/6/12 miesięcy, wyszukiwanie po nazwie/loginie/mieście, tylko aktywne), wyłączanie zaznaczonego konta i eksport widoku do CSV. |
| `AD-ManagerDiamond.ps1` | Centrum administracji domeną „Domain Ops” (WPF, ciemny motyw): sześć przestrzeni roboczych — zarządzanie zdalne komputerami, użytkownicy AD, grupy i OU, komputery AD, pliki i uprawnienia oraz domena — z 69 modułami, m.in. weryfikacją i czyszczeniem profili użytkowników względem AD, hurtowym tworzeniem kont i grup, duplikowaniem grup, budowaniem i klonowaniem drzew OU, uprawnieniami NTFS (nadawanie, odbieranie, uprawnienia efektywne z grupami zagnieżdżonymi i udziałem, raport ryzyk, porównanie ze wzorcem, naprawa z kopią uprawnień) oraz audytem bezpieczeństwa komputerów i domeny AD. Zastępuje dawne osobne skrypty (lista poniżej). Operacje w tle i równolegle, wyniki z filtrem, eksportem CSV/HTML i akcjami pod prawym przyciskiem; możliwość dodawania własnych modułów — szczegóły poniżej. |
| `AD-NTFS-AuditGUI.ps1` | Audytor NTFS w WPF: rekursywnie czyta ACL, rozpoznaje typ podmiotu (użytkownik/grupa), wskazuje foldery z uprawnieniami nadanymi bezpośrednio użytkownikom; widok tabeli i drzewa, filtry, eksport CSV/XLSX. |
| `AD-PremissionAudit.ps1` | Prosty panel GUI: po wybraniu folderu uruchamia zadanie w tle, które przeszukuje drzewo katalogów i wypisuje miejsca, gdzie uprawnienia ma pojedynczy użytkownik zamiast grupy. |
| `AD-RDP-LoginEvents.ps1` | GUI do pobierania zdarzeń z dziennika Security (domyślnie 4624/4634) z ostatnich N dni z wybranych komputerów lub wszystkich kontrolerów domeny. Działa w tle, zapisuje raport CSV do wskazanego folderu; konfigurację można zapisać/wczytać z JSON. |
| `AD-SwapLogin.ps1` | GUI do naprawiania kont z zamienionym imieniem i nazwiskiem: sprawdza listę loginów, podpowiada konta do zamiany, hurtowo zamienia GivenName/Surname (opcjonalnie DisplayName); ma tryb testowy i eksport CSV. |
| `AD-User-Manager.ps1` | Lekki manager użytkowników: wyszukuje po SamAccountName, wyświetla najważniejsze atrybuty (hasło, logowania, blokady), pozwala wyłączyć/włączyć konto, odblokować po lockoucie i zresetować hasło. |
| `PasswordGenerator.ps1` | Generator haseł (WPF, kryptograficzne losowanie) z konfigurowalnymi regułami, generowaniem paczek do schowka, historią z ostatnich 24 h, motywem jasnym/ciemnym i wysyłką przez `mailto:`. |
| `PC-CheckHashFoldersIntegrity.ps1` | WinForms do porównywania dwóch folderów (np. lokalnego z udziałem sieciowym): indeksuje pliki (także ukryte), porównuje rozmiar/datę i opcjonalnie SHA-256, pokazuje różnice i eksportuje je do CSV. |
| `PC-Folder-TreeView.ps1` | Podgląd drzewa folderu (leniwe wczytywanie, opcjonalnie pliki ukryte) z eksportem do TXT w formie ASCII. Nie wchodzi w junctiony/symlinki. |

## AD-ManagerDiamond — Domain Ops

Uruchomienie: `powershell.exe -ExecutionPolicy Bypass -File .\AD-ManagerDiamond.ps1`. Interfejs jest napisany w WPF dla Windows PowerShell 5.1 — skrypt sam uruchomi się ponownie w trybie STA, a z PowerShell 7 przełączy się na `powershell.exe`.

**Przestrzenie robocze** (przełącznik na górnym pasku, skróty `Ctrl+1…6`):

| Przestrzeń | Lista po lewej | Do czego służy |
| --- | --- | --- |
| **Zarządzanie zdalne** | komputery | Operacje na komputerach przez PowerShell Remoting (WinRM): diagnostyka (w tym analiza wydajności i logi aplikacji z wielu komputerów), profile i sesje użytkowników, grupy i konta lokalne, usługi, procesy, zdarzenia, oprogramowanie, aktualizacje, bezpieczeństwo, udziały. |
| **Użytkownicy AD** | konta użytkowników | Konta w Active Directory: szczegóły, hasła i blokady, stan konta, atrybuty, grupy, hurtowe tworzenie kont, import atrybutów z CSV, porządki (ostatnio utworzone, wyłączone konta), raporty i źródło blokad. |
| **Grupy i OU** | grupy | Grupy i jednostki organizacyjne: szczegóły i członkowie grup, członkostwo hurtowe, tworzenie i duplikowanie grup, drzewa OU, klonowanie OU, lokalizacje i role, drzewo zagnieżdżeń, raporty grup. |
| **Komputery AD** | komputery | Konta komputerów w AD: informacje i kanał zaufania, LAPS, klucze BitLocker, zmiana nazw, grupy, raporty. |
| **Pliki i uprawnienia** | — | Uprawnienia NTFS: nadawanie wielu grupom naraz, uprawnienia efektywne osoby, „gdzie ma dostęp”, odbieranie dostępu, grupy dostępu do folderów, raport uprawnień; kontrola i naprawa: raport ryzyk, udziały i NTFS, porównanie ze wzorcem, naprawa dziedziczenia i własności, kopie uprawnień z przywracaniem; sumy kontrolne plików. |
| **Domena** | — | Cała domena naraz: audyt bezpieczeństwa Active Directory z oceną punktową. |

**Układ okna**

- **Lewy panel — obiekty docelowe.** Komputery: z AD (OU z drzewa, filtr nazwy z `*`, tylko włączone), wpisane ręcznie albo z pliku TXT/CSV. Użytkownicy: wyszukiwanie w AD (login, nazwisko, e-mail, gwiazdka), stan konta (aktywne / wyłączone / zablokowane), OU, loginy wpisane ręcznie lub z pliku. Grupy: wyszukiwanie w AD (nazwa, sAMAccountName, opis, e-mail), OU, rodzaj (zabezpieczeń, dystrybucyjne, bez członków, uprzywilejowane), nazwy wpisane ręcznie; przy grupie widać zakres i typ. Zaznaczanie polami wyboru (także kilku wierszy naraz i spacją), szybkie wyszukiwanie, kolorowa kropka stanu (np. wynik testu łączności, konto zablokowane), menu kontekstowe (pulpit zdalny, konsola zarządzania, usługi, podgląd zdarzeń, `C$`, kopiowanie nazw / adresów e-mail).
- **Nawigacja modułów** pogrupowanych w kategorie — każda przestrzeń pamięta ostatnio otwarty moduł; kropka przy module oznacza trwającą operację.
- **Moduł:** nagłówek z opisem, panel parametrów (zwijany), kafelki z podsumowaniem (np. kandydaci do usunięcia, średni wynik audytu), tabela wyników z sortowaniem, filtrem po wszystkich kolumnach (`Ctrl+F`, słowo z minusem wyklucza wiersze), **filtrami kolumn**, kolorowymi „pigułkami” stanu, panelem szczegółów wiersza, menu pod prawym przyciskiem myszy (akcje modułu, kopiowanie komórki/wierszy) oraz eksportem do **CSV** lub **raportu HTML**. `F5` uruchamia główną akcję modułu.
- **Filtry kolumn** (jak autofiltr w Excelu) w każdej tabeli wyników, także w oknach z tabelą: lejek w nagłówku kolumny otwiera okienko z listą wartości (z liczbą wierszy, wyszukiwaniem i „Zaznacz wszystkie”) oraz warunkiem: *zawiera*, *nie zawiera*, *równa się*, *różne od*, *zaczyna się od*, *kończy się na*, *większe/mniejsze niż* (liczby i daty), *puste*, *niepuste*. Filtry kilku kolumn i pole „Filtruj wyniki” działają razem; lista wartości pokazuje to, co zostaje po filtrach pozostałych kolumn. Aktywny filtr wyróżnia lejek, a przycisk „Filtry kolumn: N” obok licznika wierszy czyści je jednym kliknięciem. Pod prawym przyciskiem na komórce: *Pokaż tylko «wartość»* / *Ukryj «wartość»*; na nagłówku: sortowanie, filtr, ukrycie kolumny i przywrócenie ukrytych. Eksport, kopiowanie i akcje dotyczą widocznych wierszy (raport HTML opisuje użyte filtry); nowe wyniki modułu zaczynają bez filtrów kolumn.
- **Górny pasek:** konto używane do operacji (bieżące albo alternatywne — także dla poleceń AD) i ustawienia (kontroler domeny, liczba równoległych operacji na komputerach i osobno zapytań do AD, limit połączenia WinRM, domyślny próg nieaktywności, foldery dziennika, ustawień i modułów).
- **Dziennik operacji** (`Ctrl+L`, z licznikiem nowych ostrzeżeń), **pasek stanu** z postępem i przyciskiem „Przerwij” (wszystkie operacje) oraz **powiadomienia** w rogu okna. Trwająca operacja modułu ma też własny przycisk **„Przerwij”** obok wskaźnika w nagłówku modułu. Zadanie, które nie zareaguje na przerwanie (np. zawieszone wywołanie WMI), jest po 10 s porzucane z informacją w tabeli — okno i kolejne operacje nie czekają na nie.
- **Wybór obiektów przy potwierdzeniu.** Okno potwierdzenia operacji na obiektach z listy po lewej (stan konta, edycja atrybutów, hasła, operacje na komputerach, grupach, polecenia zdalne, instalacje, tworzenie kont i grup z podglądu itd.) ma listę z polami wyboru, filtrem i przyciskami *Zaznacz / Odznacz / Odwróć* — można wykonać operację np. dla 4 z 20 zaznaczonych kont bez zmiany zaznaczenia na liście.
- **Wartości wielowierszowe** (wynik polecenia, wpis logu ze stosem wywołań) tabela pokazuje w jednej linii; całość jest w panelu szczegółów wiersza (dwuklik).

Wszystkie operacje wykonywane są w tle (pula wątków PowerShell), równolegle dla wielu komputerów/kont — okno nie zawiesza się, a wyniki pojawiają się na bieżąco. Zapytania do Active Directory mają osobny, mniejszy limit równoległości (domyślnie 4, *Ustawienia → Równoległe zapytania AD*), bo usługa ADWS na kontrolerze domeny odrzuca zbyt wiele żądań naraz; przejściowe błędy ADWS przy odczytach („A connection to the directory on which to process the request was unavailable”, „invalid enumeration context”) są automatycznie ponawiane. Jeśli mimo to się pojawiają, zmniejsz ten limit. Niedostępny komputer lub błąd pojawia się w tabeli jako wiersz „Błąd” z przyczyną. Operacje zmieniające stan wymagają potwierdzenia z listą obiektów (przy niszczących domyślnym przyciskiem jest „Anuluj”).

**Zmiany hurtowe: podgląd → poprawki → wykonanie.** Moduły tworzące lub zmieniające wiele obiektów (konta, grupy, OU, członkostwa, uprawnienia NTFS, import atrybutów) najpierw budują **podgląd** sprawdzony w AD: kolumna „Stan” mówi, co zostanie utworzone, co już istnieje, gdzie jest konflikt nazwy albo błąd danych. Kolumny oznaczone ołówkiem można poprawić w tabeli (dwuklik lub `F2`), także w wielu zaznaczonych wierszach naraz; po zmianie wiersz jest sprawdzany ponownie. Dane można wkleić prosto z Excela (kolumny rozpoznawane po nagłówkach). Wykonanie raportuje wynik dla każdego wiersza — wiersze z błędem można poprawić i powtórzyć.

**Raporty HTML kont, komputerów i grup.** W modułach *Szczegóły konta*, *Konto komputera* i *Szczegóły grup* wiersz „Raport” tworzy z obiektów zaznaczonych na liście po lewej jeden plik HTML — kartę każdego obiektu ze wszystkimi danymi w sekcjach, zamiast szerokiej tabeli:

- **konta:** tożsamość (login, UPN, e-mail i aliasy, numer pracownika, SID), organizacja (stanowisko, dział, firma, telefony, przełożony), stan konta i blokady, hasło (ustawione, wiek, wygaśnięcie, flagi), profil i skrypt logowania, **grupy** (bezpośrednie, podstawowa i — opcjonalnie — zagnieżdżone) oraz **podwładni**; ostrzeżenia, np. „hasło nigdy nie wygasa”, konto uprzywilejowane, brak wymogu hasła;
- **komputery:** system, konto komputera, ostatnie logowanie, LAPS i klucze BitLocker w AD (**tylko informacja o obecności i dacie — bez haseł i kluczy**), delegowanie Kerberos, liczba SPN, grupy;
- **grupy:** właściwości, zarządca, liczniki członków, lista **członków** z typem i stanem kont (opcjonalnie także zagnieżdżonych, z oznaczeniem członkostwa) i grupy nadrzędne.

Raport ma kafelki z podsumowaniem stanów, spis obiektów z odnośnikami, wyszukiwarkę (filtruje karty, także po nazwach grup i członków), zwijane listy i styl do druku (jasny, z rozwiniętymi listami); obiekty, których nie udało się odczytać, są wymienione na końcu.

Wszystkie raporty HTML programu (eksport tabeli, karty obiektów, drzewo zagnieżdżeń grup, drzewo uprawnień NTFS) mają wspólne funkcje: **Rozwiń wszystko / Zwiń wszystko**, **filtr pod nagłówkiem każdej kolumny** tabel (tekst; `!tekst` – nie zawiera, `=tekst` – równe, `>10` / `<10` – liczby), sortowanie po kliknięciu nagłówka, **kopiowanie pojedynczego wiersza** (jako „Kolumna: wartość”) albo karty / gałęzi drzewa, **otwarcie wiersza, karty lub gałęzi w osobnym oknie**, kopiowanie widocznych wierszy do Excela i czyszczenie filtrów. Plik jest samodzielny (bez skryptów i czcionek z sieci) i otwiera się w przeglądarce po zapisaniu. Listy dłuższe niż 2000 pozycji są skracane z informacją — pełną listę daje moduł i eksport CSV.

**Profile użytkowników — co można usunąć.** Moduł *Profile użytkowników* zbiera profile z zaznaczonych komputerów (ostatnie użycie z `ProfileList`, opcjonalnie rozmiar) i sprawdza każde konto w AD przez ADSI (bez RSAT). Kandydaci do usunięcia: konta **usunięte z AD**, **wyłączone**, **wygasłe**, usunięte konta lokalne, profile **tymczasowe/uszkodzone** oraz — opcjonalnie — profile nieużywane dłużej niż N dni. Profile załadowane (użytkownik zalogowany) i systemowe nigdy nie są kandydatami. Przycisk *Zaznacz kandydatów* → *Usuń zaznaczone profile* usuwa profile przez `Win32_UserProfile` (folder i wpis w rejestrze), z potwierdzeniem pokazującym ścieżki, rozmiary i ostrzeżeniem, jeśli zaznaczono profil, który nie jest kandydatem.

**Moduły — Zarządzanie zdalne**

| Kategoria | Moduł | Możliwości |
| --- | --- | --- |
| Diagnostyka | Łączność | DNS, ping, porty TCP, test sesji PowerShell — lokalnie; koloruje kropki na liście, przycisk „zaznacz tylko dostępne”. |
| | Inwentaryzacja | Producent, model, numer seryjny, BIOS, system i kompilacja, procesor, RAM, dysk systemowy, TPM, Secure Boot, zalogowany użytkownik, IP/MAC, daty instalacji i startu. |
| | Wydajność | **Co obciąża komputer lub serwer** — pomiar w oknie czasu (np. 10 s, próbka co 2 s) z surowych liczników wydajności (niezależnych od języka systemu): **procesor** (średnio i szczyt, kolejka na rdzeń, czas jądra, przerwania i DPC, przełączenia kontekstu), **pamięć** (dostępna, zadeklarowana względem limitu, stronicowanie z dysku, pule), **dyski** — każdy wolumin: opóźnienie odczytu i zapisu w ms, kolejka, zajętość, IOPS, MB/s, wolne miejsce, **sieć** — każda karta: wykorzystanie łącza, kolejka wyjściowa, błędy i odrzucone pakiety, retransmisje TCP, połączenia, **procesy** z czołówki CPU, pamięci i I/O (z usługami w `svchost`, użytkownikiem i wierszem polecenia) oraz **zdarzenia** świadczące o problemach (resety i błędy dysków, brak pamięci, BSOD, nieoczekiwane restarty, awarie aplikacji). Ocena wskazuje wąskie gardło („dysk D: opóźnienie 48 ms, kolejka 6 – sqlservr 35 MB/s”). Widoki *Podsumowanie / Procesy / Dyski / Sieć / Zdarzenia* z jednego pomiaru, tryb **Monitoruj** (pomiar za pomiarem), kończenie procesów z menu wiersza. |
| | Logi aplikacji | **Pliki logów z wielu komputerów naraz** (np. `%ProgramData%\Firma\Aplikacja\Logs\*.log`, przeglądanie folderów na komputerze docelowym): ostatnie N wpisów, ostatnie minuty/godziny albo przedział dat; filtr tekstu lub wyrażenia regularnego, wykluczenia, tylko błędy i ostrzeżenia. Rozpoznaje kodowanie (UTF-8, UTF-16, ANSI) i format daty, łączy wielowierszowe wpisy (stos wywołań), czyta duże pliki od końca bez blokowania pliku aplikacji. Widoki: **wspólna oś czasu** wszystkich komputerów, **porównanie komputerów** — te same komunikaty (bez dat, liczb i identyfikatorów) zliczone na każdym komputerze z oznaczeniem „wszystkie / część / tylko PC-03”, czyli czy problem dzieje się wszędzie, czy na jednej maszynie — oraz lista przeczytanych plików. Tryb **Śledź** odświeża co 15 s. |
| | Zasilanie i uptime | Czas pracy, oczekujący restart (CBS, Windows Update, pliki, zmiana nazwy, SCCM); restart/wyłączenie z opóźnieniem i komunikatem, anulowanie. |
| Użytkownicy i dostęp | Profile użytkowników | Weryfikacja profili względem AD, kandydaci do usunięcia, rozmiar, usuwanie (opis wyżej). |
| | Sesje użytkowników | Zalogowani użytkownicy (konsola/RDP), bezczynność; wiadomość, rozłączenie, wylogowanie, podgląd sesji (shadow). |
| | Grupy lokalne | Administratorzy, Użytkownicy pulpitu zdalnego, zarządzania zdalnego (WinRM), dziennika zdarzeń, kopii zapasowych — po SID (każdy język systemu); dodawanie (także wybór grup z AD) i usuwanie; wbudowany Administrator i Domain Admins chronione. |
| | Konta lokalne | Stan, ostatnie logowanie, wiek hasła; włączanie, wyłączanie, ustawianie hasła. |
| | Pulpit zdalny | Stan RDP, NLA, port, reguły zapory, członkowie grupy RDP; włączanie/wyłączanie RDP razem z zaporą, NLA. |
| | Gdzie używane jest konto | Usługi, zadania Harmonogramu, pule aplikacji i katalogi wirtualne IIS oraz automatyczne logowanie działające na wskazanych kontach (albo na wszystkich kontach poza wbudowanymi) na zaznaczonych komputerach; rodzaj konta (wbudowane / lokalne / domenowe, rozpoznawany po SID), czy hasło jest zapisane, ostrzeżenie o haśle automatycznego logowania zapisanym jawnym tekstem. Kolumna **Konto w AD**: konto wyłączone, zablokowane, z wygasłym hasłem albo nieistniejące (usługa nie wystartuje). **Zapisz nowe hasło…** – po zmianie hasła konta usługi w AD zapisuje je we wskazanych miejscach jednego konta (usługi, zadania z zapisanym hasłem, IIS, automatyczne logowanie), opcjonalnie z ponownym uruchomieniem usług i odświeżeniem pul; hasło jest najpierw sprawdzane w AD, aby złe hasło nie zablokowało konta. |
| Zdalne wykonanie | Polecenia | PowerShell lub cmd.exe (polskie znaki), szablony typowych poleceń (DNS, Kerberos, gpresult, kanał zaufania, czas, DISM/SFC). Polecenie działa w osobnym procesie **bez klawiatury** — programy czekające na odpowiedź (`cmd`, `pause`, `choice`, `set /p`, `Read-Host`) kończą się zamiast wisieć; **limit czasu** (1 min – 2 h albo bez limitu) i **„Przerwij”** zamykają proces razem z procesami potomnymi. Kilka linii w trybie cmd.exe działa jak plik `.cmd`; wynik ze stanem, kodem wyjścia, czasem i liczbą błędów PowerShell. |
| | Instalacja oprogramowania | MSI/MSP/MSU/EXE: kopiowanie (ADMIN$ albo WinRM), cicha instalacja, kod wyjścia z opisem, log MSI; limit czasu i przerwanie zamykają instalator czekający na odpowiedź. |
| | Aktualizacja zasad grupy | `gpupdate` dla komputera i/lub użytkownika, opcjonalnie `/force` (z limitem czasu i możliwością przerwania). |
| System | Usługi | Filtr nazwy i stanu, start/stop/restart, typ uruchamiania (także z menu wiersza). |
| | Procesy | Pamięć, właściciel, sesja, wiersz poleceń; kończenie procesów. |
| | Dyski | Zajętość z oceną; czyszczenie Temp, TEMP profili, Kosza i pobranych aktualizacji z raportem odzyskanego miejsca. |
| | Dziennik zdarzeń | Gotowe zestawy (nieoczekiwane wyłączenia, błędy dysków, awarie aplikacji, nieudane logowania, blokady, GPO, Windows Update) albo własne dzienniki/poziomy/ID. |
| | Harmonogram zadań | Podgląd, uruchom/zatrzymaj/włącz/wyłącz/usuń, tworzenie prostych zadań SYSTEM. |
| | Autostart | Klucze Run/RunOnce (komputer i zalogowani użytkownicy) i foldery Autostart; usuwanie wpisów. |
| | Drukarki | Drukarki, sterowniki, porty, zadania w kolejce; czyszczenie kolejek, usuwanie, restart bufora wydruku. |
| | Sterowniki i urządzenia | Sterowniki PnP z filtrem i urządzenia z problemami. |
| Oprogramowanie | Zainstalowane programy | Rejestr 64/32-bit i profili zalogowanych; kolumna **„Odinstalowanie zdalne”** mówi, czy i jak program da się usunąć bez użytkownika: MSI, tryb cichy producenta, Inno Setup (zielone), tylko z przełącznikami albo w sesji użytkownika (żółte), brak deinstalatora / zablokowane przez producenta / profil użytkownika bez trybu cichego (czerwone). Odinstalowanie z limitem czasu, sprawdzeniem, czy program zniknął z rejestru, i wynikiem z przyczyną w kolumnie „Stan” (np. „deinstalator czekał na odpowiedź użytkownika”, „MSI 1618 – trwa inna instalacja”, „nadal zainstalowany – wymaga restartu”); programy z profilu zalogowanego użytkownika usuwane zadaniem w jego sesji; *Odinstaluj z przełącznikami…* dla deinstalatorów bez trybu cichego w rejestrze; okno z podsumowaniem, gdy coś się nie udało; „pokaż ten program na wszystkich komputerach”. |
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
| Konta | Szczegóły konta | Stan, kontakt, przełożony, logowania, hasło, wygaśnięcie, profil, OU, SID. **Raport HTML** zaznaczonych kont (opis niżej). |
| | Hasło i blokada | Stan haseł i blokad; odblokowanie; reset hasła — **losowe, inne dla każdego konta** (widoczne jako wartości poufne, kopiowanie z czyszczeniem schowka) albo wpisane; wymuszenie zmiany przy logowaniu; „hasło nigdy nie wygasa”. |
| | Stan konta | Włączanie i wyłączanie (z dopiskiem daty/autora/powodu w opisie i przeniesieniem do OU), data wygaśnięcia, przenoszenie do OU. |
| | Edycja atrybutów | Odczyt i masowa zmiana atrybutów (stanowisko, dział, firma, biuro, telefony, e-mail, przełożony, extensionAttribute1–15…) z polami `{login}`, `{imie}`, `{nazwisko}`, `{nazwa}`. |
| Grupy | Członkostwo w grupach | Grupy bezpośrednie i zagnieżdżone, dodawanie (wyszukiwarka grup), usuwanie, kopiowanie członkostwa z konta wzorcowego. |
| Tworzenie i import | Tworzenie kont | Wiele kont z danych wklejonych z Excela (imię, nazwisko, numer albumu/pracownika, dział, stanowisko, miasto, opis). **Profile** (np. Pracownik, Student, Wykładowca) — ustawiane w osobnym oknie **konfiguratora profilu** (*Edytuj profil…*, *Nowy…*, *Duplikuj*, z podglądem przykładowego konta), a w module zostaje lista profili z jednowierszowym podsumowaniem i miejsce na dane — z OU, sufiksem UPN, domeną e-mail, formatem loginu (`i.nazwisko`, `inazwisko`, `imie.nazwisko`, `nazwisko.imie`, numer…) i nazwy wyświetlanej, opisem, firmą, działem, grupami i opcjami konta (włączone, zmiana hasła przy logowaniu, wygaśnięcie po N dniach, długość hasła). Polskie znaki są zamieniane w loginach, imiona i nazwiska poprawiane (wielkość liter, nazwiska dwuczłonowe). Podgląd sprawdza zajęte loginy, UPN i nazwy CN (z automatyczną numeracją), hasła losowe dla każdego konta; po utworzeniu kopiowanie danych logowania i synchronizacja z Microsoft Entra ID (`Start-ADSyncSyncCycle -PolicyType Delta`, lokalnie lub na serwerze). |
| | Import atrybutów z CSV | Hurtowa zmiana atrybutów (stanowisko, dział, firma, telefony, przełożony, extensionAttribute…) z pliku CSV: automatyczne rozpoznanie separatora, mapowanie kolumn, podgląd „było → będzie” dla każdego konta, zastosowanie zmian z zapisem pliku cofania i przywracanie poprzednich wartości. |
| Porządki | Ostatnio utworzone | Konta użytkowników, komputery, grupy i OU utworzone w ostatnich minutach, godzinach lub dniach (z twórcą obiektu); wyłączanie lub usuwanie obiektów utworzonych przez pomyłkę. |
| | Wyłączone konta | Wyłączone konta poza OU dla nieaktywnych: przeniesienie do tej OU i usunięcie z grup (bez grup chronionych i kont-wyjątków) z kopią członkostw w CSV i przywracaniem z kopii. |
| Raporty | Raporty kont | Zablokowane, wyłączone, nieaktywne, nigdy nie logowane, hasło wygasa / wygasło / nigdy nie wygasa, konta wygasające, nowe, uprzywilejowane (`adminCount`); wyniki można zaznaczyć na liście kont i od razu wykonać na nich operacje. |
| | Źródło blokady konta | Zdarzenia 4740 z emulatora PDC (lub wszystkich DC) — komputer, z którego przyszły błędne hasła; opcjonalnie 4771/4776 z adresem IP/stacją. |

**Moduły — Grupy i OU**

| Kategoria | Moduł | Możliwości |
| --- | --- | --- |
| Grupy | Szczegóły grup | Zakres, typ, opis, zarządca, e-mail i liczba członków; zmiana opisu, zarządcy, adresu, zakresu i typu, zmiana nazwy, przeniesienie do OU, ochrona przed usunięciem, usuwanie. **Raport HTML** zaznaczonych grup. |
| | Członkowie grup | Członkowie bezpośredni i zagnieżdżeni (bez limitu 5000 obiektów) z typem i stanem kont; dodawanie, kopiowanie członków z innej grupy, usuwanie, przeniesienie kont na listę użytkowników. |
| | Członkostwo hurtowe | Dodawanie lub usuwanie wielu kont/grup/komputerów w wielu grupach: każdy do każdej albo wiersz do wiersza (np. tabela z Excela). Rozpoznaje loginy, UPN, e-mail, DN, SID i nazwy; plan pokazuje, co już jest, a co zostanie zmienione. |
| Tworzenie | Tworzenie grup | Wiele grup z tabeli (nazwa, opis, zakres, typ, e-mail, sAMAccountName, grupy nadrzędne, zarządca, OU) — wklejanie z Excela, normalizacja sAMAccountName, podgląd z edycją i sprawdzeniem w AD. |
| | Duplikowanie grup | Kopie zaznaczonych grup z nową nazwą (zamiana tekstu, przedrostek, przyrostek — także wiele kopii naraz, np. `WAW → KRK; GDA`), opcjonalnie z członkami, przynależnością do grup, zarządcą i ochroną; kopie w tej samej lub innej OU. |
| | Lokalizacje i role | OU dla lokalizacji (np. miast) i w każdej te same grupy ról według szablonów nazwy i opisu (`{PREFIKS}`, `{LOKALIZACJA}`, `{KOD_LOK}`, `{ROLA}`, `{KOD_ROLI}`), z grupami zbiorczymi ALL dla lokalizacji, dla roli i globalną; unikalne sAMAccountName (opcjonalnie najwyżej 20 znaków). |
| Struktura OU | Drzewo OU | Całe drzewo jednostek z tekstu: wcięcia albo ścieżki (`Warszawa/Komputery`), opisy po `|`, ochrona przed usunięciem. Istniejącą strukturę można wczytać jako tekst, zmienić i utworzyć w innym miejscu. |
| | Klonowanie OU | Kopia jednostki w nowe miejsce: podjednostki, grupy (z zamianą tekstu w nazwach, np. `WAW=KRK`), zagnieżdżenia między nimi, opcjonalnie członkowie i grupy nadrzędne, linki GPO, opisy i ochrona. Konta użytkowników i komputerów nie są kopiowane. |
| Raporty | Drzewo zagnieżdżeń | Grupy zawarte w zaznaczonych grupach (w dół) albo grupy nadrzędne (w górę), opcjonalnie z członkami, z wykrywaniem zapętleń; raport HTML z rozwijanymi gałęziami i wyszukiwarką. |
| | Raporty grup | Puste, z jednym członkiem, duże, bez opisu lub zarządcy, z wyłączonymi kontami, uprzywilejowane, dystrybucyjne, nowe i zmienione, zapętlone zagnieżdżenia; zaznaczanie wyników na liście grup. |

**Moduły — Pliki i uprawnienia**

| Kategoria | Moduł | Możliwości |
| --- | --- | --- |
| Uprawnienia NTFS | Nadawanie uprawnień | Uprawnienia do folderu dla wielu grup lub kont naraz, każda z własnym poziomem (odczyt … pełna kontrola) i zakresem (ten folder / podfoldery / pliki). Tożsamości rozpoznawane w AD i nadawane po SID (także `BUILTIN\Users`); opcje: zastąpienie dotychczasowych wpisów tych tożsamości, wyłączenie dziedziczenia z kopią wpisów; podgląd bieżących uprawnień. Działa lokalnie, przez UNC albo na serwerze plików (PowerShell Remoting) — po wpisaniu innego komputera przycisk wyboru folderu **przegląda dyski i udziały na tym komputerze** (drzewo folderów wczytywane na bieżąco). |
| | Raport uprawnień | Wpisy ACL folderów (opcjonalnie plików) do zadanej głębokości: tożsamość, uprawnienia, typ, dziedziczenie, zakres, właściciel, foldery z wyłączonym dziedziczeniem i brakiem dostępu. Filtry: tylko jawne wpisy, bez kont systemowych (lista do edycji), bez nierozwiązanych SID. **Tryb kopii zapasowej** czyta uprawnienia i zawartość także tam, gdzie Administratorzy nie mają dostępu (przywileje kopii zapasowej administratora), bez przejmowania własności. **Pomiń miejsca bez dostępu** — lokalizacje, do których nawet administrator nie ma prawa odczytu (np. foldery systemowe), nie przerywają skanu, a każda trafia do dziennika i do raportu; łącza (junction) nie są przechodzone. **Raport HTML (drzewo)** — drzewo folderów z odznakami (jawne wpisy, wyłączone dziedziczenie, brak dostępu, nierozwiązane SID, właściciel): poddrzewa, w których wszystko tylko dziedziczy, są zwinięte do jednej linii „dziedziczy; zawartość tak samo (N)”, a przełącznik *Pokaż foldery i wpisy dziedziczone* je rozwija. Podfoldery skanowane równolegle; członkowie grupy z wiersza; eksport CSV/HTML. |
| | Uprawnienia efektywne | Jaki dostęp do folderu lub pliku ma konkretna osoba albo grupa — liczony jak w Windows: przez wszystkie jej grupy (także **zagnieżdżone**, z grupą podstawową), Wszyscy / Użytkownicy uwierzytelnieni / `BUILTIN\Users`, z kolejnością wpisów (odmowa przed zezwoleniem), uprawnieniami ogólnymi (GENERIC), wpisami „tylko podfoldery” i prawami właściciela. Dla ścieżki `\\serwer\udział` także **uprawnienia udziału** — wynik to mniejsze z dwóch. Tabela 13 praw szczegółowych z kolumnami *NTFS*, *Udział*, *Efektywnie* i *Skąd* (który wpis i przez którą grupę daje lub odbiera prawo). |
| | Gdzie ma dostęp | Wszystkie miejsca w drzewie folderów, w których osoba (bezpośrednio albo przez grupy, także zagnieżdżone) lub grupa ma wpis: poziom, efektywny dostęp, przez którą grupę, typ, dziedziczenie, zakres. Z wiersza: uprawnienia efektywne albo odebranie dostępu. |
| | Odbieranie uprawnień | Usuwa wpisy wskazanych osób lub grup w całym drzewie: podgląd każdego jawnego wpisu (dziedziczone są pokazane, ale usuwa się je w folderze wyżej), wybór miejsc w potwierdzeniu, **kopia uprawnień przed zmianą**; folder, którego uprawnienia zmieniły się od podglądu, jest pomijany. |
| | Grupy dostępu do folderu | Model „grupa na folder”: dla folderu lub każdego jego podfolderu tworzy grupy (odczyt / modyfikacja / opcjonalnie pełna kontrola) według szablonu nazwy (`DL_{Folder}_{Poziom}`, pola `{Nadrzędny}`, `{Udział}`, `{Serwer}`) i opisu, w wybranej OU i zakresie, dodaje członków (`RO: grupa`) i od razu nadaje grupom uprawnienia NTFS (ten folder, podfoldery i pliki). Podgląd sprawdzony w AD (istniejąca grupa jest używana), nazwy i opisy do poprawienia w tabeli. |
| Kontrola i naprawa | Naprawa uprawnień | W drzewie folderów: **włączenie dziedziczenia** (jawne wpisy zostają), **przywrócenie dziedziczenia** (bez jawnych wpisów), **usunięcie wpisów kont usuniętych z AD** (nierozwiązane SID są dodatkowo sprawdzane w AD — konta z domen zaufanych zostają), **przejęcie własności** miejsc bez dostępu (grupa Administratorzy, opcjonalnie z pełną kontrolą). W **trybie kopii zapasowej** przejęcie odbywa się jednym przejściem drzewa (folder dostaje właściciela, zanim zostanie odczytana jego zawartość), a wpis Administratorów trafia tylko tam, gdzie go brakowało — reszta drzewa go dziedziczy; bez tego trybu używany jest `icacls` po SID, niezależnie od języka systemu. Zawsze podgląd, wybór miejsc i kopia uprawnień przed zmianą. |
| | Kopie uprawnień | Lista kopii zapisanych automatycznie przed zmianami (i ręcznych — przycisk *Zrób kopię uprawnień* dla całego drzewa): podgląd zawartości, **przywracanie** całości albo wybranych folderów (wpisy i dziedziczenie; właściciel — w trybie kopii zapasowej albo gdy system pozwala go ustawić, inaczej ostrzeżenie w wyniku); stan sprzed przywrócenia trafia do nowej kopii. |
| | Raport ryzyk | Ryzykowne uprawnienia z oceną (wysokie / średnie / niskie) i zaleceniem: zapis dla Wszystkich, Użytkowników uwierzytelnionych, Użytkowników domeny; odczyt dla Wszystkich i gości; pełna kontrola lub zmiana uprawnień dla zwykłych kont; zbyt wiele kont z pełną kontrolą (próg do ustawienia); wpisy dla użytkowników zamiast grup; wpisy usuniętych kont; odmowy; użytkownik jako właściciel. Z wiersza: gdzie jeszcze ma dostęp, odebranie uprawnień. |
| | Porównanie ze wzorcem | Porównuje uprawnienia folderów (np. wszystkich folderów działów — *Wstaw podfoldery…*) z folderem wzorcowym: brakujące, nadmiarowe i inne wpisy, dziedziczenie, właściciel. **Zastosuj wzorzec…** kopiuje uprawnienia wzorca na wybrane foldery — dokładna kopia jawnych wpisów z ustawieniem dziedziczenia albo tylko brakujące wpisy — z kopią uprawnień. |
| | Udziały i NTFS | Udziały serwera plików (opcjonalnie administracyjne) z uprawnieniami udziału i NTFS folderu udziału obok siebie: dla każdego konta poziom z udziału, z NTFS i **wynikowy dostęp przez sieć**, z informacją, które uprawnienia ograniczają dostęp, i ostrzeżeniem o zapisie dla szerokich grup. Odczyt na serwerze przez PowerShell Remoting. |
| | Zmiany uprawnień | Co się zmieniło: kopia uprawnień porównana z **obecnym stanem** albo z **inną kopią** — nowe, usunięte i zmienione wpisy, dziedziczenie, właściciel, foldery, których już nie ma, a przy pełnej kopii drzewa także **nowe foldery**; nowy zapis dla szerokich grup (Wszyscy, Użytkownicy domeny…) jest wyróżniony. Z wiersza: przywrócenie wybranych folderów ze stanu z kopii. Regularna kopia ręczna drzewa (moduł *Kopie uprawnień*, akcja *Porównaj z obecnym stanem*) pozwala śledzić, gdzie i komu nadano dostęp. |
| Pliki | Sumy kontrolne | MD5, SHA1, SHA256, SHA384, SHA512 dla plików i folderów (także przeciągniętych z Eksploratora), liczone w tle partiami; porównanie z oczekiwanym skrótem, historia w tabeli, kopiowanie w formacie `skrót  plik`. |

**Moduły — Komputery AD**

| Kategoria | Moduł | Możliwości |
| --- | --- | --- |
| Konta komputerów | Konto komputera | Informacje, test i naprawa kanału zaufania, włączanie/wyłączanie, opis, przenoszenie do OU, reset, usuwanie, tworzenie nowych kont (pre-staging). **Raport HTML** zaznaczonych komputerów. |
| | Zmiana nazwy komputerów | Autonumeracja, mapowanie z listy (np. z Excela), edycja w tabeli, walidacja NetBIOS i duplikatów, opcjonalny restart. |
| | Członkostwo w grupach | Jak dla użytkowników — dla kont komputerów. |
| Hasła i klucze | LAPS | Windows LAPS (także szyfrowane) i LAPS legacy, kopiowanie hasła (dwuklik), wymuszenie zmiany z przetworzeniem zasad. |
| | Klucze BitLocker (AD) | Klucze zaznaczonych komputerów oraz **wyszukiwanie komputera po identyfikatorze klucza** z ekranu odzyskiwania. |
| Raporty | Raporty komputerów | Nieaktywne, wyłączone, nowe, podsumowanie systemów, nieobsługiwane systemy, bez LAPS, bez klucza BitLocker w AD, serwery; zaznaczanie na liście, wyłączanie, przenoszenie, usuwanie. |

**Moduły — Domena**

| Kategoria | Moduł | Możliwości |
| --- | --- | --- |
| Bezpieczeństwo | Audyt bezpieczeństwa AD | 34 kontrole w pięciu grupach (uruchamiane równolegle), każda z oceną *wysokie / średnie / niskie / bez uwag*, listą obiektów i zaleceniem; ocena punktowa domeny (100 minus wagi niespełnionych kontroli). **Konta uprzywilejowane** (członkowie grup chronionych po SID, także zagnieżdżeni): SPN na kontach administratorów, hasła bez wygasania i stare hasła, nieużywane i wyłączone konta w grupach, brak ochrony przed delegowaniem (flaga „konto poufne” / Protected Users), członkowie Schema i Enterprise Admins, używanie wbudowanego konta Administrator, pozostałości `adminCount = 1`. **Kerberos i delegowanie:** konta z SPN (kerberoasting), bez wstępnego uwierzytelnienia (AS-REP roasting), nieograniczone delegowanie poza kontrolerami domeny, delegowanie z dowolnym protokołem, RBCD, wiek hasła `krbtgt`, DES. **Hasła:** brak wymaganego hasła, szyfrowanie odwracalne, hasła bez wygasania, zasady haseł domeny (długość, złożoność, blokada, historia) i zasady szczegółowe. **Konta i komputery:** nieaktywne konta i komputery, konto Gość, nieobsługiwane systemy, komputery bez LAPS (według schematu: Windows LAPS i LAPS legacy), SID History. **Domena:** `ms-DS-MachineAccountQuota`, kosz AD, poziom funkcjonalny, grupa Protected Users, „Pre-Windows 2000 Compatible Access” z dostępem anonimowym, zaufania bez filtrowania SID. Progi do ustawienia (nieaktywność, wiek hasła administratora i `krbtgt`, liczba kont uprzywilejowanych). Z wiersza: lista obiektów, **zaznaczenie kont lub komputerów na listach** do dalszych operacji; **raport HTML** z oceną, zestawieniem i kartą każdej kontroli. Tylko odczyt. |

**Zastąpione skrypty.** Dawne osobne narzędzia zostały przeniesione do `AD-ManagerDiamond.ps1` (z podglądem przed zmianami, pracą w tle i wspólnym dziennikiem) i usunięte z repozytorium — w razie potrzeby są dostępne w historii git.

| Dawny skrypt | Obecnie |
| --- | --- |
| `AD-BulkUserCreator.ps1` | Użytkownicy AD → *Tworzenie kont* (zakładki stały się profilami; plik `.AD-BulkUserCreator.json` leżący obok programu jest importowany automatycznie przy pierwszym uruchomieniu, a inny można wczytać z menu *Więcej → Importuj profile z pliku…*) |
| `AD-UpdateUserTitleDepartment.ps1` | Użytkownicy AD → *Import atrybutów z CSV* |
| `AD-DeleteNewAccoutsGUI.ps1` | Użytkownicy AD → *Ostatnio utworzone* |
| `AD-MoveDisabledUsersGUI.ps1`, `AD-RemoveDisabledUsersFromGroups.ps1` | Użytkownicy AD → *Wyłączone konta* |
| `AD-BulkGroupCreator.ps1` | Grupy i OU → *Tworzenie grup* |
| `AD-BulkAddUsersToGroups.ps1` | Grupy i OU → *Członkostwo hurtowe* |
| `AD-BulkOUGroupCreator.ps1` | Grupy i OU → *Lokalizacje i role* (oraz *Drzewo OU*) |
| `AD-DomainGroupTree.ps1` | Grupy i OU → *Drzewo zagnieżdżeń* |
| `AD-NTFS-BulkGroupPermissions.ps1` | Pliki i uprawnienia → *Nadawanie uprawnień* |
| `Get-PremissionReport.ps1` | Pliki i uprawnienia → *Raport uprawnień* |
| `File-HashChecker.ps1` | Pliki i uprawnienia → *Sumy kontrolne* |

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

`Register-Workspace -Key -Title -Icon -Target Computer|User|Group|None` dodaje nową przestrzeń; `Start-HostOperation` (zdalnie lub `-Local`) i `Start-AdOperation` (moduł ActiveDirectory z gotową hashtablą `$ad`) uruchamiają operacje w tle.

## Uwagi bezpieczeństwa

- Przed zmianami hurtowymi w `AD-ManagerDiamond` zawsze przejrzyj podgląd (kolumny „Stan” i „Uwagi”) — wykonanie dotyczy tylko wierszy gotowych do zmiany. `AD-SwapLogin` ma osobny tryb testowy.
- Usuwanie obiektów w module *Ostatnio utworzone* jest nieodwracalne (poza przywróceniem z Kosza AD, jeśli jest włączony) — jeśli nie masz pewności, użyj wyłączenia konta.
- Moduł *Wyłączone konta* zapisuje kopię członkostw w grupach (`%APPDATA%\AD-ManagerDiamond\Backup`) przed usunięciem kont z grup; z tej kopii można je przywrócić. Grupy chronione i konta-wyjątki są pomijane (listy do edycji w module).
- *Nadawanie uprawnień* z opcją „Zastąp dotychczasowe wpisy” usuwa jawne wpisy ACL wskazanych tożsamości przed dodaniem nowych, a „Wyłącz dziedziczenie” zamienia wpisy dziedziczone na jawne kopie — obie opcje są wyróżnione w potwierdzeniu.
- Narzędzia NTFS zmieniające uprawnienia (*Odbieranie uprawnień*, *Naprawa uprawnień*, *Porównanie ze wzorcem*, *Grupy dostępu do folderu*) zapisują przed zmianą kopię uprawnień każdego zmienianego folderu (`%APPDATA%\AD-ManagerDiamond\AclBackup`) — przywrócisz ją w module *Kopie uprawnień*; jeśli kopii nie da się zapisać, zmiana nie jest wykonywana. Folder, którego uprawnienia zmieniły się od podglądu, jest pomijany. Przejęcie własności obejmuje całe drzewo i wymaga uprawnień administratora; bez trybu kopii zapasowej (`icacls /setowner`) po pierwszym przebiegu mogą zostać głębsze miejsca bez dostępu — wtedy uruchom podgląd i naprawę ponownie. **Tryb kopii zapasowej** (moduły NTFS) włącza przywileje SeBackupPrivilege, SeRestorePrivilege i SeTakeOwnershipPrivilege w procesie programu albo w sesji WinRM na serwerze plików, do ich zakończenia — pozwalają czytać i zmieniać uprawnienia z pominięciem list ACL, więc używaj go świadomie.
- „Resetuj konto” w module *Konto komputera* `AD-ManagerDiamond` działa jak polecenie z konsoli ADUC: ustawia hasło konta komputera na domyślne (nazwa komputera małymi literami, bez `$`); komputer trzeba potem ponownie dołączyć do domeny lub naprawić kanał zaufania.
- `AD-ManagerDiamond` maskuje hasła LAPS, klucze odzyskiwania BitLocker i nowo wygenerowane hasła użytkowników w tabeli, panelu szczegółów, eksporcie CSV/HTML i kopiowaniu (do czasu zaznaczenia „Pokaż poufne”); nie są też przeszukiwane filtrem. Hasło skopiowane akcją „Kopiuj hasło…” jest usuwane ze schowka po 60 s. Zmiana nazwy komputera i naprawa kanału zaufania wymagają poświadczeń domenowych — program poprosi o nie, jeśli nie ustawiono poświadczeń alternatywnych.
- Usuwanie profili w `AD-ManagerDiamond` jest nieodwracalne (folder profilu i wpis w rejestrze). Przed usunięciem sprawdź kolumnę „Ocena” — profile oznaczone jako kandydaci to m.in. konta usunięte/wyłączone/wygasłe w AD; nieużywane profile aktywnych kont są kandydatami tylko przy włączonej opcji. Profile zalogowanych użytkowników są zawsze pomijane.
- Moduły z folderu `AD-ManagerDiamond.Modules` są wykonywane z uprawnieniami użytkownika programu — trzymaj tam tylko zaufane pliki.
- `PasswordGenerator` przechowuje wygenerowane hasła jawnym tekstem (historia w `%TEMP%`) — czyść historię i schowek po pracy. Hasła nowych kont w *Tworzeniu kont* są maskowane w tabeli i eksporcie (do czasu zaznaczenia „Pokaż poufne”) i nie są zapisywane w ustawieniach ani dzienniku.

## Kontakt

Masz uwagi lub pomysł na nowy scenariusz? Otwórz Issue na GitHub lub skontaktuj się z autorem.

**Autor:** Boris
