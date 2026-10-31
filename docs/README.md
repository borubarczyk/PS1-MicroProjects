# Domain Ops (AD-ManagerDiamond) — dokumentacja

Domain Ops to centrum administracji domeną Windows w jednym pliku `AD-ManagerDiamond.ps1`. Program zbiera w jednym oknie operacje, które zwykle robi się w kilku konsolach: zarządzanie komputerami przez PowerShell Remoting, konta i grupy w Active Directory, uprawnienia NTFS, przegląd bezpieczeństwa i stanu domeny oraz raporty wysyłane według harmonogramu.

Ta dokumentacja opisuje każdą sekcję programu: **do czego służy, co można w niej zrobić i jak działa**. Ten sam opis program pokazuje w oknie **Jak to działa** przy każdym module (`F1`) — jest osadzany w skrypcie z tych plików przez `dev/Update-EmbeddedHelp.ps1` (uruchom go po każdej zmianie dokumentacji). Krótszy przegląd wszystkich funkcji jest w [README repozytorium](../README.md).

## Spis treści

| Dokument | Zawartość |
|---|---|
| [Podstawy pracy z programem](#podstawy-pracy-z-programem) (ten plik) | uruchomienie, układ okna, praca z modułem, ustawienia, tryb diagnostyczny, zabezpieczenia programu |
| [1. Zarządzanie zdalne](01-zarzadzanie-zdalne.md) | operacje na komputerach: diagnostyka, użytkownicy i dostęp, zdalne wykonanie, system, oprogramowanie, bezpieczeństwo, udostępnianie |
| [2. Użytkownicy AD](02-uzytkownicy-ad.md) | konta użytkowników: szczegóły, hasła, stan, atrybuty, offboarding, tworzenie i import, porządki, raporty |
| [3. Grupy i OU](03-grupy-i-ou.md) | grupy i jednostki organizacyjne: członkowie, tworzenie, duplikowanie, drzewa i klonowanie OU, raporty |
| [4. Komputery AD](04-komputery-ad.md) | konta komputerów: kanał zaufania, LAPS, klucze BitLocker, zmiana nazw, raporty |
| [5. Pliki i uprawnienia](05-pliki-i-uprawnienia.md) | uprawnienia NTFS: nadawanie, raporty, uprawnienia efektywne, odbieranie, naprawa, kopie i porównania; sumy kontrolne |
| [6. Domena](06-domena.md) | audyt bezpieczeństwa AD, delegacje, stan domeny, zasady grupy, serwery, raporty cykliczne |

---

## Podstawy pracy z programem

### Uruchomienie i wymagania

- **Uruchomienie:** `powershell.exe -ExecutionPolicy Bypass -File .\AD-ManagerDiamond.ps1`. Program działa w Windows PowerShell 5.1 (WPF). Sam uruchomi się ponownie w trybie STA, a uruchomiony z PowerShell 7 przełączy się na `powershell.exe`.
- **Active Directory:** moduł `ActiveDirectory` (RSAT) — potrzebny w przestrzeniach AD i do wczytywania list z AD. Bez niego działa zarządzanie zdalne z listą komputerów wpisaną ręcznie.
- **Zasady grupy:** kopie zapasowe GPO, raporty ustawień i metoda Invoke-GPUpdate wymagają modułu `GroupPolicy` (RSAT: Group Policy Management Tools).
- **Komputery docelowe:** większość operacji zdalnych idzie przez PowerShell Remoting (WinRM, TCP 5985/5986) i wymaga uprawnień administratora lokalnego. Kilka modułów umie działać bez WinRM — patrz [Gdy WinRM jest niedostępny](#gdy-winrm-jest-niedostępny).
- **Uprawnienia:** program robi to, na co pozwala konto, na którym działa (albo konto alternatywne ustawione na górnym pasku). Do zmian w AD potrzebne są odpowiednie delegacje, do NTFS — prawa do folderów, do dziennika Security — prawo jego odczytu.
- **Kodowanie pliku:** skrypt musi pozostać zapisany jako UTF-8 z BOM, inaczej Windows PowerShell 5.1 popsuje polskie znaki.

### Układ okna

- **Górny pasek** — przełącznik przestrzeni roboczych (`Ctrl+1…6`), konto używane do operacji (bieżące albo alternatywne, także dla poleceń AD), nazwa operatora i kłódka blokady (`Ctrl+Shift+L`), ustawienia.
- **Lewy panel — obiekty docelowe.** Zależnie od przestrzeni: komputery, użytkownicy albo grupy.
  - *Komputery*: z AD (OU z drzewa, filtr nazwy z `*`, tylko włączone), wpisane ręcznie albo z pliku TXT/CSV.
  - *Użytkownicy*: wyszukiwanie w AD (login, nazwisko, e-mail, gwiazdka), stan konta (aktywne, wyłączone, zablokowane), OU, loginy wpisane ręcznie albo z pliku.
  - *Grupy*: wyszukiwanie (nazwa, sAMAccountName, opis, e-mail), OU, rodzaj (zabezpieczeń, dystrybucyjne, bez członków, uprzywilejowane), nazwy wpisane ręcznie.
  - Zaznaczanie polami wyboru (także kilku wierszy naraz i spacją), szybkie wyszukiwanie, kolorowa kropka stanu (np. wynik testu łączności, konto zablokowane).
  - Menu pod prawym przyciskiem na komputerach: pulpit zdalny, pomoc zdalna (gdy włączona w Ustawieniach), **wiadomość do zalogowanych** na wszystkich zaznaczonych komputerach, Zarządzanie komputerem, Usługi, Podgląd zdarzeń, udział `C$`, kopiowanie nazw.
- **Nawigacja modułów** — moduły pogrupowane w kategorie; kategorie można zwijać kliknięciem nagłówka (prawy przycisk: zwiń albo rozwiń wszystkie). Kropka przy module oznacza trwającą operację. Każda przestrzeń pamięta ostatnio otwarty moduł.
- **Moduł** — nagłówek z opisem i przyciskiem **Jak to działa** (albo `F1`), panel parametrów (zwijany), kafelki z podsumowaniem, tabela wyników i panel szczegółów wiersza.
- **Dziennik operacji** (`Ctrl+L`) — wszystkie operacje z wynikiem; licznik nowych ostrzeżeń na przycisku. Po starcie programu dziennik jest zwinięty.
- **Pasek stanu** — postęp i przycisk *Przerwij* dla wszystkich operacji.

### Praca z modułem

1. Zaznacz obiekty na liście po lewej (komputery, konta, grupy) — chyba że moduł działa na całej domenie albo na ścieżce folderu.
2. Ustaw parametry w panelu modułu.
3. Kliknij główny przycisk (albo `F5`). Operacja działa w tle — możesz pracować dalej, także w innych modułach.
4. Wyniki pojawiają się w tabeli na bieżąco. Niedostępny komputer albo błąd to wiersz „Błąd” z przyczyną.

**Jak to działa** — przycisk obok nazwy modułu (albo klawisz `F1`) otwiera opis modułu z tej dokumentacji: do czego służy, co można w nim zrobić, jak działa i na co uważać. Zakładki w tym oknie pokazują też wstęp przestrzeni roboczej i podstawy pracy z programem.

**Tabela wyników**

- Sortowanie po kliknięciu nagłówka, przeciąganie kolumn, filtr po wszystkich kolumnach (`Ctrl+F`; słowo z minusem wyklucza wiersze).
- **Filtry kolumn** jak autofiltr w Excelu — lejek w nagłówku: lista wartości z liczbą wierszy i warunek (*zawiera*, *nie zawiera*, *równa się*, *różne od*, *zaczyna się od*, *kończy się na*, *większe/mniejsze niż*, *puste*, *niepuste*). Przycisk „Filtry kolumn: N” czyści wszystkie.
- **Kolumny** — wybór widocznych kolumn; kopiowanie i eksport biorą tylko widoczne, w kolejności z tabeli. *Kopiuj wartości kolumny* kopiuje wartości jednej kolumny (także bez powtórzeń i pustych).
- **Panel szczegółów** — wszystkie pola zaznaczonego wiersza; dwuklik otwiera szczegóły. Wartości wielowierszowe (wynik polecenia, stos wywołań) tabela pokazuje w jednej linii, całość jest w panelu.
- **Eksport** do CSV albo raportu HTML (samodzielny plik z filtrami pod nagłówkami, sortowaniem, kopiowaniem wierszy, rozwijaniem i zwijaniem).

**Menu pod prawym przyciskiem na wierszach**

- Działa na klikniętym wierszu albo na kilku zaznaczonych (`Ctrl`, `Shift`) — bez zmiany zaznaczenia po lewej.
- W tabelach z kolumną *Komputer*: *Odśwież tylko te komputery* (ponawia ostatnie zapytanie tylko dla nich), *Zaznacz na liście komputerów*, pulpit zdalny, konsole, `C$`.
- W modułach kont: zaznaczenie kont na liście użytkowników.
- Akcje z paska narzędzi modułu są też w menu wiersza i wykonują się dla obiektów z zaznaczonych wierszy. Po takiej akcji odświeżane są tylko ich wiersze.

**Potwierdzenia i zmiany hurtowe**

- Operacje zmieniające stan pytają o potwierdzenie z **listą obiektów z polami wyboru** — można wykonać operację np. dla 4 z 20 zaznaczonych kont. Przy operacjach niszczących domyślnym przyciskiem jest „Anuluj”.
- Moduły tworzące lub zmieniające wiele obiektów (konta, grupy, OU, członkostwa, uprawnienia NTFS, import atrybutów) najpierw budują **podgląd** sprawdzony w AD: kolumna „Stan” mówi, co zostanie utworzone, co już istnieje, gdzie jest konflikt albo błąd danych. Kolumny z ołówkiem można poprawić w tabeli (dwuklik albo `F2`, także w wielu wierszach naraz). Dane można wkleić prosto z Excela.

### Operacje w tle, przerywanie i powiadomienia

- Wszystkie operacje działają w puli wątków, równolegle dla wielu komputerów i kont. Limity równoległości są w Ustawieniach (osobno dla komputerów i dla zapytań do AD).
- Polecenia modułu ActiveDirectory idą pojedynczo przez jedną blokadę na cały program — tak usługa ADWS na kontrolerze nie gubi połączeń. Przejściowe błędy odczytu z AD są ponawiane, zmiany nie.
- **Przerwij** — przycisk przy wskaźniku w nagłówku modułu (ta operacja) albo na pasku stanu (wszystkie). Zadanie, które nie zareaguje (np. zawieszone wywołanie WMI), jest po 10 s porzucane z informacją w tabeli.
- **Powiadomienie** po każdej operacji w rogu okna: nazwa, liczba udanych i błędnych obiektów, czas i pierwszy błąd. Kliknięcie przenosi do modułu, a przy błędach otwiera dziennik.

### Ustawienia

Okno ustawień (ikona na górnym pasku):

- kontroler domeny (puste — wybór automatyczny), liczba równoległych operacji na komputerach i zapytań do AD, limit czasu połączenia WinRM, domyślny próg nieaktywności kont i komputerów;
- **blokada po bezczynności** w minutach (0 wyłącza);
- **nadawca wiadomości** do użytkowników (tytuł okna *Wiadomość od: …*, domyślnie *Administrator*);
- **pomoc zdalna** w menu (domyślnie wyłączona) — z opisem wymagań;
- **tryb diagnostyczny** — opis niżej;
- **MFA – konta i klucze** — opis niżej;
- **lokalizacje plików programu** (ustawienia, szablony poleceń, dziennik, moduły własne) z przyciskiem *Pokaż* w Eksploratorze.

Ustawienia są zapisywane w `%APPDATA%\AD-ManagerDiamond\settings.json` — bez haseł i poświadczeń.

### Tryb diagnostyczny

Włączony w Ustawieniach dodaje do dziennika (i pliku dziennika) wpisy **DEBUG**:

- przy starcie każdej operacji — tryb, pula wątków, wiersz bloku w pliku programu i **parametry**; hasła, klucze, PIN-y i poświadczenia są ukrywane (`***`), długie teksty i listy skracane;
- dla każdego obiektu — **czas wykonania**, liczba wyników i błędów, informacja o przerwanym bloku;
- dla każdego błędu — **typ wyjątku, wiersz bloku z treścią i stos wywołań**; także błędy obsługi wyników, akcji okna i nieobsłużone błędy interfejsu;
- kolejne kroki wybranych operacji (np. podłączanie i odłączanie dysków sieciowych, przegląd stanu domeny).

Tryb przydaje się, gdy coś „niby działa, ale nic się nie dzieje” — wpisy pokazują, co program wysłał i co wróciło. Plik dziennika: `%LOCALAPPDATA%\AD-ManagerDiamond\Logs\DomainOps_RRRRMMDD.log`.

### Zabezpieczenia programu

**PIN przy starcie.** Zanim pokaże się okno, program prosi o PIN (klawiatura na ekranie albo cyfry z klawiatury). Po 5 błędnych próbach program się zamyka. W skrypcie zapisany jest tylko skrót SHA-256 PIN-u. To ochrona przed przypadkowym uruchomieniem przez osobę postronną, nie przed kimś, kto może edytować skrypt. Raporty cykliczne (tryb bez okna) nie pytają o PIN.

**Blokada w trakcie pracy.** Kłódka albo `Ctrl+Shift+L` ukrywa zawartość okna; odblokowanie tym samym PIN-em. Operacje w tle trwają dalej, a okna, które pojawiłyby się w tym czasie, pokazują się po odblokowaniu. Program blokuje się sam po 10 minutach bez aktywności w jego oknach (do zmiany w Ustawieniach). Przycisk *Zakończ awaryjnie* zamyka program bez PIN-u (drugie kliknięcie w ciągu 10 s).

**MFA — kod z aplikacji w telefonie.** Do PIN-u można dodać kod TOTP (Microsoft Authenticator, Google Authenticator i zgodne): *Ustawienia → MFA – konta i klucze… → Włącz MFA dla konta podstawowego*. Klucz pokazywany jest raz, jako kod QR i tekst; zapisuje się dopiero po poprawnym kodzie z aplikacji.

- **Konta dodatkowe** — własna nazwa, własny PIN i własny klucz w aplikacji. Na ekranie blokady pojawia się lista kont.
- **Uprawnienia kont dodatkowych** — drzewo przestrzeni i modułów; odznaczone są dla konta zablokowane i znikają z nawigacji.
- **Reset PIN-u konta dodatkowego** — tylko z konta podstawowego, zatwierdzony jego kodem.
- Dodawanie, usuwanie, nowe klucze i uprawnienia zatwierdza się kodem konta podstawowego.
- Dane MFA są w `%APPDATA%\AD-ManagerDiamond\Mfa.json`, zaszyfrowane DPAPI — działają tylko na tym koncie Windows i tym komputerze. Utrata telefonu z kontem podstawowym: usuń ten plik (program wróci do samego PIN-u).

**Wartości poufne.** Hasła LAPS, klucze odzyskiwania BitLocker i nowe hasła są maskowane w tabeli, panelu, eksporcie i kopiowaniu (do zaznaczenia „Pokaż poufne”). Skopiowane hasło znika ze schowka po 60 s.

### Gdy WinRM jest niedostępny

Większość modułów zarządzania zdalnego działa przez WinRM. Bez niego zostają:

- **DCOM / WMI** (TCP 135 i porty dynamiczne) — *Licencja i wersja Windows* i *Diagnostyka WinRM* przechodzą na DCOM same; *Diagnostyka WinRM* potrafi przez DCOM włączyć WinRM;
- **RPC** — *Aktualizacja zasad grupy* metodą Invoke-GPUpdate;
- **konsole MMC i udziały** z menu (Zarządzanie komputerem, Usługi, Podgląd zdarzeń, `C$`);
- **SMB** (kopiowanie instalatorów przez `ADMIN$`), **Pulpit zdalny** (TCP 3389).

Najczęstszą przyczynę „usługa WinRM działa, a z innej podsieci nie” — reguła zapory ograniczona do podsieci lokalnej i karta w profilu Publicznym — wskazuje moduł [Diagnostyka WinRM](01-zarzadzanie-zdalne.md#diagnostyka-winrm), razem z zaleceniem dla GPO.

### Pliki programu i moduły własne

- Lista plików, które program zapisuje (ustawienia, kopie, dzienniki, raporty), jest w [README — Pliki pomocnicze](../README.md#pliki-pomocnicze-i-dane-zapisywane-przez-skrypty).
- Własne moduły: pliki `*.ps1` w folderze `AD-ManagerDiamond.Modules` obok skryptu są wczytywane przy starcie i mogą dodawać przestrzenie i moduły. Przykład jest w README w sekcji *Rozbudowa — własne moduły*. Uruchamiają się z uprawnieniami użytkownika programu — trzymaj tam tylko zaufane pliki.
