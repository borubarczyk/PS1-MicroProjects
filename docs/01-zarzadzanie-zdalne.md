# 1. Zarządzanie zdalne

[← Spis treści](README.md)

**Do czego jest ta przestrzeń.** Operacje na komputerach i serwerach domeny bez logowania się na nie: diagnostyka, użytkownicy i dostęp, uruchamianie poleceń i instalacji, usługi, procesy, dyski, zdarzenia, oprogramowanie, aktualizacje, bezpieczeństwo i udostępnianie.

**Lista po lewej:** komputery — z AD (drzewo OU, filtr nazwy, tylko włączone), wpisane ręcznie albo wczytane z pliku TXT/CSV. Operacje działają na komputerach zaznaczonych na liście, chyba że akcję wywołasz z menu wiersza w tabeli wyników.

**Jak to działa ogólnie.** Program łączy się z komputerami przez PowerShell Remoting (WinRM), równolegle, i uruchamia na nich krótkie skrypty. Wyniki wracają do tabeli na bieżąco. Wymagane są uprawnienia administratora lokalnego na komputerze docelowym. Wyjątki (DCOM, RPC, odczyt lokalny) są opisane przy modułach.

**Spis modułów**

- [Diagnostyka](#diagnostyka): [Łączność](#łączność) · [Inwentaryzacja](#inwentaryzacja) · [Wydajność](#wydajność) · [Monitor obciążenia](#monitor-obciążenia) · [Logi aplikacji](#logi-aplikacji) · [Zasilanie i uptime](#zasilanie-i-uptime) · [Licencja i wersja Windows](#licencja-i-wersja-windows) · [Diagnostyka WinRM](#diagnostyka-winrm)
- [Użytkownicy i dostęp](#użytkownicy-i-dostęp): [Profile użytkowników](#profile-użytkowników) · [Sesje użytkowników](#sesje-użytkowników) · [Grupy lokalne](#grupy-lokalne) · [Zgodność grup lokalnych](#zgodność-grup-lokalnych) · [Konta lokalne](#konta-lokalne) · [Pulpit zdalny](#pulpit-zdalny) · [Gdzie używane jest konto](#gdzie-używane-jest-konto)
- [Zdalne wykonanie](#zdalne-wykonanie): [Polecenia](#polecenia) · [Instalacja oprogramowania](#instalacja-oprogramowania) · [Aktualizacja zasad grupy](#aktualizacja-zasad-grupy)
- [System](#system): [Usługi](#usługi) · [Procesy](#procesy) · [Dyski](#dyski) · [Dziennik zdarzeń](#dziennik-zdarzeń) · [Harmonogram zadań](#harmonogram-zadań) · [Rejestr – odczyt i eksport](#rejestr--odczyt-i-eksport) · [Autostart](#autostart) · [Drukarki](#drukarki) · [Sterowniki i urządzenia](#sterowniki-i-urządzenia)
- [Oprogramowanie](#oprogramowanie): [Zainstalowane programy](#zainstalowane-programy) · [Windows Update](#windows-update) · [Outlook i Microsoft 365](#outlook-i-microsoft-365)
- [Bezpieczeństwo](#bezpieczeństwo): [Szybki audyt bezpieczeństwa](#szybki-audyt-bezpieczeństwa) · [Microsoft Defender](#microsoft-defender) · [BitLocker](#bitlocker) · [Zapora Windows](#zapora-windows) · [Certyfikaty komputera](#certyfikaty-komputera)
- [Udostępnianie](#udostępnianie): [Udziały sieciowe](#udziały-sieciowe) · [Zasoby sieciowe użytkowników](#zasoby-sieciowe-użytkowników)

---

## Diagnostyka

### Łączność

**Do czego.** Szybko sprawdzić, które komputery są osiągalne i dlaczego któryś nie odpowiada.

**Co możesz zrobić**
- sprawdzić DNS, ping, wybrane porty TCP i sesję PowerShell Remoting;
- pokolorować kropki na liście komputerów według wyniku i jednym przyciskiem zostawić zaznaczone tylko dostępne;
- z menu wiersza uruchomić *Diagnostykę WinRM* dla komputerów, które nie odpowiadają.

**Jak działa.** Test wykonuje Twój komputer, lokalnie — działa więc także dla komputerów bez WinRM.

### Inwentaryzacja

**Do czego.** Podstawowe dane sprzętu i systemu z wielu komputerów w jednej tabeli.

**Co możesz zrobić**
- odczytać producenta, model, numer seryjny, BIOS, system, wydanie i kompilację, architekturę, procesor i rdzenie, RAM, rozmiar i wolne miejsce dysku systemowego, TPM, Secure Boot, zalogowanego użytkownika, domenę, adresy IP i MAC, datę instalacji systemu i ostatniego startu;
- wyeksportować wynik do CSV albo HTML (np. spis sprzętu).

**Jak działa.** Przez WinRM odczytuje klasy WMI/CIM komputera (system, BIOS, procesor, dyski, karty sieciowe), rejestr wersji Windows oraz stan TPM i Secure Boot.

### Wydajność

**Do czego.** Odpowiedzieć na pytanie „co obciąża ten komputer albo serwer” na podstawie pomiaru w krótkim oknie czasu.

**Co możesz zrobić**
- zmierzyć obciążenie w oknie czasu (np. 10 s, próbka co 2 s) albo uruchomić tryb *Monitoruj* (pomiar za pomiarem);
- zobaczyć **procesor** (średnio i szczyt, kolejka na rdzeń, czas jądra, przerwania i DPC, przełączenia kontekstu), **pamięć** (dostępna, zadeklarowana względem limitu, stronicowanie, pule), **dyski** (opóźnienia odczytu i zapisu, kolejka, zajętość, IOPS, MB/s, wolne miejsce), **sieć** (wykorzystanie łącza, kolejka, błędy, retransmisje TCP, połączenia);
- zobaczyć **procesy** z czołówki CPU, pamięci i dysku (z usługami w `svchost`, użytkownikiem i wierszem polecenia) i **zdarzenia** świadczące o problemach (błędy dysków, brak pamięci, BSOD, nieoczekiwane restarty, awarie aplikacji);
- przełączać widoki *Podsumowanie / Procesy / Dyski / Sieć / Zdarzenia* z jednego pomiaru;
- zakończyć proces z menu wiersza.

**Jak działa.** Odczytuje surowe liczniki wydajności (niezależne od języka systemu) i wylicza wartości z różnic między próbkami. Ocena wskazuje wąskie gardło, np. „dysk D: opóźnienie 48 ms, kolejka 6 – sqlservr 35 MB/s”.

### Monitor obciążenia

**Do czego.** Obserwować jeden serwer na żywo, jak w Monitorze zasobów, i dowiedzieć się, kto zużył najwięcej od początku pomiaru.

**Co możesz zrobić**
- wybrać komputer w polu **Komputer** — lista komputerów zaznaczonych po lewej (domyślnie pierwszy); zmiana w trakcie pomiaru **przełącza pomiar** na wybrany komputer;
- uruchomić pomiar z próbką co kilka sekund (domyślnie 5 s) na zadany czas (najdłużej 15 min);
- widzieć **kafelki z wykresami** ostatnich próbek: procesor, pamięć, najwolniejszy dysk, sieć;
- czytać **„Co się dzieje teraz”** — wąskie gardło i jego sprawca, np. „Dysk D: 60 ms na operację – najwięcej I/O: backup 25 MB/s”;
- czytać **„Od startu pomiaru”** — kto zużył najwięcej czasu procesora, odczytu i zapisu, kiedy było najmniej wolnej pamięci;
- przełączać widoki *Procesy* (teraz, średnio, szczyt, łącznie), *Dyski*, *Sieć*, *Przebieg* (próbka po próbce) i *Podsumowanie*;
- skopiować podsumowanie do schowka i zakończyć proces z menu wiersza bez przerywania pomiaru.

**Jak działa.** Jedno połączenie WinRM na cały pomiar: na serwerze działa pętla, która odczytuje liczniki i wysyła tylko różnice od poprzedniego odczytu, a program je sumuje. Monitor obserwuje jeden komputer naraz — jego nazwę widać w polu *Komputer*, w nagłówku „Co się dzieje teraz – …” i na kafelku czasu pomiaru. Przy przełączeniu bieżący pomiar jest kończony, a nowy rusza od zera na wybranym komputerze. Tabela odświeża się w miejscu (sortowanie i zaznaczenie zostają). Koszt po stronie serwera widać w tabeli jako proces monitora — zwykle ułamek procenta CPU.

**Uwagi.** Ruch sieciowy procesów to tylko liczba połączeń TCP — bajty na proces wymagałyby śledzenia jądra (ETW).

### Logi aplikacji

**Do czego.** Przeglądać pliki logów aplikacji (np. `%ProgramData%\Firma\Aplikacja\Logs\*.log`) z wielu komputerów naraz i sprawdzić, czy problem dzieje się wszędzie, czy na jednej maszynie.

**Co możesz zrobić**
- wskazać folder i maskę plików (z przeglądaniem folderów na komputerze docelowym);
- wybrać zakres: ostatnie N wpisów, ostatnie minuty lub godziny albo przedział dat;
- filtrować tekstem lub wyrażeniem regularnym, wykluczać, pokazywać tylko błędy i ostrzeżenia;
- przeglądać **wspólną oś czasu** wszystkich komputerów;
- **porównać komputery** — te same komunikaty (bez dat, liczb i identyfikatorów) zliczone na każdym komputerze z oznaczeniem „wszystkie / część / tylko PC-03”;
- włączyć tryb **Śledź** (odświeżanie co 15 s).

**Jak działa.** Skrypt na komputerze rozpoznaje kodowanie (UTF-8, UTF-16, ANSI) i format daty, łączy wpisy wielowierszowe (stos wywołań) i czyta duże pliki od końca, bez blokowania pliku aplikacji.

### Zasilanie i uptime

**Do czego.** Sprawdzić, od kiedy komputery działają i czy czekają na restart; zaplanować restart lub wyłączenie.

**Co możesz zrobić**
- odczytać czas pracy i oczekujący restart (obsługa składników, Windows Update, operacje na plikach, zmiana nazwy, SCCM);
- zaplanować restart lub wyłączenie z opóźnieniem i komunikatem dla użytkownika;
- anulować zaplanowany restart (także z menu wiersza).

**Jak działa.** Przez WinRM odczytuje czas startu systemu i znaczniki oczekującego restartu w rejestrze; restart i wyłączenie uruchamia systemowym `shutdown` z opóźnieniem.

### Licencja i wersja Windows

**Do czego.** Sprawdzić wersję i stan aktywacji Windows na serwerach i stacjach — także tam, gdzie WinRM nie działa.

**Co możesz zrobić**
- odczytać system i edycję, wydanie, kompilację z numerem poprawki, typ komputera (stacja, serwer, kontroler domeny, Server Core);
- sprawdzić **stan licencji** (aktywowany, prolongata, tryb powiadomień, bez licencji), **kanał** (Retail, OEM, MAK, KMS, AVMA, ewaluacyjna), końcówkę klucza, serwer KMS, datę ważności i dni do końca, klucz OEM w UEFI i przyczynę problemu po polsku;
- opcjonalnie odczytać licencje Office 2013 i nowszych;
- z menu wiersza **aktywować** system (`slmgr /ato`) z odświeżeniem stanu.

**Jak działa.** Odczytuje usługę licencjonowania systemu przez WinRM, a gdy jest niedostępny — przez DCOM (WMI). Kolumna *Połączenie* mówi, którą drogą się udało.

### Diagnostyka WinRM

**Do czego.** Ustalić, dlaczego WinRM na komputerze jest niedostępny, i doraźnie go włączyć.

**Co możesz zrobić**
- z Twojego komputera: sprawdzić DNS, ping, porty 5985/5986/135/445 (z rozróżnieniem *otwarty / odrzucony / filtrowany*), odpowiedź WSMan i sesję PowerShell (brak uprawnień, Kerberos);
- przez DCOM z komputera docelowego: odczytać usługę WinRM, listener (także z GPO), **profil sieci** karty, przez którą się łączysz, **reguły zapory dla TCP 5985** (lokalne i z GPO, z zakresem adresów porównanym z Twoim adresem) i **filtr IPv4 zasady WinRM**;
- przeczytać **wniosek** (np. „reguła dopuszcza tylko podsieć lokalną, karta w profilu Publicznym”) i **zalecenie dla GPO**;
- **włączyć WinRM przez DCOM** (`Enable-PSRemoting`, na PowerShell 2.0 `winrm quickconfig`), opcjonalnie z regułą zapory dopuszczającą TCP 5985 tylko z Twojego adresu, i ponowić diagnozę.

**Jak działa.** Testy sieciowe wykonuje Twój komputer; konfigurację komputera docelowego czyta przez WMI (DCOM), więc działa bez WinRM. Włączenie uruchamia polecenie na komputerze przez WMI; jego dziennik zostaje w `%SystemRoot%\Temp\DomainOps\EnableWinRM.log`.

---

## Użytkownicy i dostęp

### Profile użytkowników

**Do czego.** Znaleźć profile, które można bezpiecznie usunąć, i usunąć je.

**Co możesz zrobić**
- zebrać profile z zaznaczonych komputerów (ostatnie użycie, opcjonalnie rozmiar);
- zobaczyć kandydatów do usunięcia: konta **usunięte z AD**, **wyłączone**, **wygasłe**, usunięte konta lokalne, profile **tymczasowe/uszkodzone** i opcjonalnie nieużywane dłużej niż N dni;
- *Zaznacz kandydatów* → *Usuń zaznaczone profile* — z potwierdzeniem pokazującym ścieżki i rozmiary.

**Jak działa.** Profile i ostatnie użycie czyta z rejestru `ProfileList` przez WinRM, a każde konto sprawdza w AD przez ADSI (bez RSAT). Usuwa przez `Win32_UserProfile` (folder i wpis w rejestrze).

**Uwagi.** Profile załadowane (użytkownik zalogowany) i systemowe nigdy nie są kandydatami. Potwierdzenie ostrzega, gdy zaznaczysz profil, który nie jest kandydatem.

### Sesje użytkowników

**Do czego.** Zobaczyć, kto jest zalogowany na komputerach i serwerach, i obsłużyć te sesje.

**Co możesz zrobić**
- odczytać sesje konsoli i pulpitu zdalnego: stan, **nazwę i adres IP komputera klienta**, bezczynność, czas logowania, chwilę rozłączenia;
- dla jednej albo kilku sesji: wysłać wiadomość, rozłączyć, wylogować, **zresetować zawieszoną sesję** (`rwinsta`), podejrzeć sesję (shadow), zaoferować pomoc zdalną;
- zobaczyć **procesy w sesji** (od największego zużycia pamięci) i zakończyć proces.

**Jak działa.** Czyta sesje przez WTS API (niezależnie od języka systemu). Gdy na serwerze nie da się uruchomić pomocnika WTS (np. tryb ograniczonego języka), lista pochodzi z `quser`, bez danych klienta. Po zmianie odświeżany jest tylko serwer tej sesji.

### Grupy lokalne

**Do czego.** Sprawdzić i zmienić członków wbudowanych grup lokalnych na wielu komputerach.

**Co możesz zrobić**
- odczytać członków grup Administratorzy, Użytkownicy pulpitu zdalnego, Użytkownicy zarządzania zdalnego (WinRM), Czytelnicy dziennika zdarzeń, Operatorzy kopii zapasowych;
- dodać członków (także wybierając grupy z AD) i usunąć zaznaczonych.

**Jak działa.** Grupy są wyznaczane po SID, więc działa w każdym języku systemu. Wbudowany Administrator i Domain Admins są chronione przed usunięciem.

### Zgodność grup lokalnych

**Do czego.** Pilnować, żeby w grupach lokalnych (np. Administratorzy) byli tylko ci, którzy powinni — na wszystkich komputerach według jednego szablonu.

**Co możesz zrobić**
- przygotować **szablon**: dla grup wbudowanych (Administratorzy, Pulpit zdalny, WinRM, Operatorzy kopii, DCOM, Czytelnicy dziennika i inne) i grup własnych — członkowie **dozwoleni** i **wymagani**; wzorce `DOMENA\nazwa`, sama nazwa, `.\nazwa`, SID albo wzorzec SID (`*-500` — wbudowany Administrator, `*-512` — Administratorzy domeny);
- porównać komputery z szablonem: **niedozwoleni**, **brakujący**, dodatkowi, brak grupy, **osierocone SID-y**;
- **usunąć niedozwolonych** (z kopią członkostwa przed zmianą), **dodać brakujących**, **przywrócić z kopii**;
- z wiersza dopisać członka do dozwolonych w szablonie.

**Jak działa.** Odczyt i zmiany przez WinRM, grupy wbudowane po SID. Kopia przed usunięciem trafia do `%APPDATA%\AD-ManagerDiamond\Backup\GrupyLokalne_*.json`.

**Uwagi.** Wbudowany Administrator i Administratorzy domeny w grupie Administratorzy nigdy nie są usuwani automatycznie. Ten sam przegląd jest dostępny jako [raport cykliczny](06-domena.md#raporty-cykliczne).

### Konta lokalne

**Do czego.** Zarządzać lokalnymi kontami użytkowników na komputerach.

**Co możesz zrobić:** odczytać stan, ostatnie logowanie i wiek hasła; włączyć, wyłączyć i ustawić hasło zaznaczonym kontom.

**Jak działa.** Przez WinRM, poleceniami kont lokalnych systemu.

### Pulpit zdalny

**Do czego.** Sprawdzić i zmienić dostęp przez Pulpit zdalny (RDP).

**Co możesz zrobić**
- odczytać stan RDP, uwierzytelnianie na poziomie sieci (NLA), port, reguły zapory i członków grupy Użytkownicy pulpitu zdalnego;
- włączyć lub wyłączyć RDP razem z regułami zapory i wymagać NLA (także z menu wiersza).

**Jak działa.** Przez WinRM zmienia ustawienia Terminal Services w rejestrze i reguły zapory.

### Gdzie używane jest konto

**Do czego.** Przed zmianą hasła konta usługowego dowiedzieć się, gdzie to konto działa — i potem zapisać nowe hasło we wszystkich tych miejscach.

**Co możesz zrobić**
- znaleźć usługi, zadania Harmonogramu, pule aplikacji i katalogi wirtualne IIS oraz automatyczne logowanie działające na wskazanych kontach (albo na wszystkich poza wbudowanymi);
- zobaczyć rodzaj konta (wbudowane, lokalne, domenowe), czy hasło jest zapisane, ostrzeżenie o haśle automatycznego logowania zapisanym jawnym tekstem;
- w kolumnie **Konto w AD** zobaczyć konto wyłączone, zablokowane, z wygasłym hasłem albo nieistniejące (usługa nie wystartuje);
- **Zapisz nowe hasło…** — zapisać nowe hasło konta we wskazanych miejscach, opcjonalnie z restartem usług i odświeżeniem pul.

**Jak działa.** Przeszukuje konfigurację komputerów przez WinRM. Nowe hasło jest najpierw sprawdzane w AD, żeby złe hasło nie zablokowało konta.

---

## Zdalne wykonanie

### Polecenia

**Do czego.** Uruchomić polecenie PowerShell albo cmd.exe na wielu komputerach naraz i zebrać wynik.

**Co możesz zrobić**
- wpisać polecenie albo wybrać **szablon** (DNS, Kerberos, gpresult, kanał zaufania, czas, DISM/SFC, ostatnie poprawki);
- zapisać bieżące polecenie jako **własny szablon** (plik `CommandTemplates.json` — można go przekazać innym administratorom);
- ustawić **limit czasu** (1 min – 2 h albo bez limitu);
- przeczytać wynik: stan, kod wyjścia, czas, liczbę błędów; pełny tekst w panelu szczegółów;
- uruchomić ponownie na wybranych komputerach z menu wiersza.

**Jak działa.** Polecenie działa w osobnym procesie na komputerze, **bez klawiatury** — programy czekające na odpowiedź (`pause`, `choice`, `set /p`, `Read-Host`) kończą się zamiast wisieć. Limit czasu i *Przerwij* zamykają proces razem z procesami potomnymi. Kilka linii w trybie cmd.exe działa jak plik `.cmd`. Polskie znaki są poprawne także w wyniku programów konsolowych.

**Uwagi.** Polecenie działa w sesji WinRM, nie na pulpicie użytkownika — nie pokaże okna zalogowanej osobie.

### Instalacja oprogramowania

**Do czego.** Zainstalować program albo poprawkę na wielu komputerach w trybie cichym.

**Co możesz zrobić**
- wskazać plik MSI, MSP, MSU albo EXE i dodatkowe argumenty;
- ustawić limit czasu i usuwanie instalatora po instalacji;
- przeczytać wynik: kod wyjścia z opisem, czas, sposób kopiowania, ścieżka logu instalatora;
- zainstalować ponownie na wybranych komputerach z menu wiersza.

**Jak działa.** Kopiuje instalator do `%SystemRoot%\Temp\DomainOps` przez udział `ADMIN$` (gdy działasz na bieżącym koncie), a jeśli się nie da — przez sesję WinRM. MSI/MSP instaluje z `/qn /norestart` i logiem, MSU przez DISM (`wusa.exe` nie działa w sesji WinRM), EXE z podanymi przełącznikami. Limit czasu i *Przerwij* zamykają instalator, który czeka na odpowiedź.

### Aktualizacja zasad grupy

**Do czego.** Wymusić odświeżenie zasad grupy na wielu komputerach i zobaczyć, czy się udało, a gdy nie — która zasada zawiodła.

**Co możesz zrobić**
- wybrać zakres: komputer, komputer i użytkownik albo użytkownik, opcjonalnie `/force`;
- wybrać **metodę**:
  - **WinRM** — `gpupdate` na komputerze z pełnym wynikiem (domyślna);
  - **Invoke-GPUpdate** — bez WinRM;
  - **WinRM, a gdy niedostępny — Invoke-GPUpdate**;
- dla Invoke-GPUpdate ustawić, jak długo czekać na wynik (nie czekaj, 1, 2 albo 5 min);
- przy błędzie zobaczyć kolumny **Problem z zasadą** (GUID zasady) i **Nazwa zasady**;
- uruchomić ponownie na wybranych komputerach z menu wiersza.

**Jak działa**
- *WinRM:* `gpupdate` działa na komputerze z odpowiedzią „N” na pytania o wylogowanie i restart, z limitem czasu. GUID-y zasad z komunikatów (niezależnie od języka) program tłumaczy na nazwy — z zapisanego stanu zasad grupy na komputerze, a brakujące z AD (nieznany GUID — „brak w AD”).
- *Invoke-GPUpdate:* moduł GroupPolicy na Twoim komputerze tworzy na komputerze docelowym, przez RPC, zadanie Harmonogramu, które od razu uruchamia `gpupdate`. Wynik program czyta ze zdarzeń *Microsoft-Windows-GroupPolicy* w dzienniku System komputera: zasady komputera i użytkownika zastosowane albo błąd przetwarzania z numerem zdarzenia.
- Kolumna **Metoda** mówi, którą drogą poszła aktualizacja.

**Wymagania dla Invoke-GPUpdate**
- moduł GroupPolicy (RSAT) na Twoim komputerze;
- na komputerach reguły zapory *Zdalne zarządzanie zaplanowanymi zadaniami* (RPC i RPC-EPMAP) i *Instrumentacja zarządzania Windows (WMI)* — gotowe w początkowym GPO *Group Policy Remote Update Firewall Ports*;
- do odczytu wyniku: reguła *Zdalne zarządzanie dziennikiem zdarzeń*;
- Invoke-GPUpdate działa na koncie programu — nie przyjmuje poświadczeń alternatywnych.

---

## System

### Usługi

**Do czego.** Sprawdzać i obsługiwać usługi Windows na wielu komputerach, łącznie z zawieszonymi.

**Co możesz zrobić**
- filtrować usługi po nazwie i stanie;
- uruchomić, zatrzymać, zrestartować i zmienić typ uruchamiania (z paska albo menu wiersza);
- **Zakończ proces i uruchom** — dla usługi zawieszonej (np. „Zatrzymywanie”, brak reakcji na restart);
- **Pokaż proces usługi** — przejście do modułu *Procesy* z procesami zaznaczonych usług.

**Jak działa.** Akcje dotyczą usług zaznaczonych w tabeli, na właściwych komputerach. *Zakończ proces i uruchom* kończy proces usługi, czeka, aż menedżer usług uzna ją za zatrzymaną, uruchamia ją i ponownie te usługi zależne, które działały. Po operacji lista odświeża się na komputerach, których dotyczyła.

**Uwagi.** Pomijane (z powodem w dzienniku): usługa we wspólnym procesie z innymi (svchost — program wypisuje, które jeszcze tam działają), usługi i procesy krytyczne dla systemu oraz WinRM.

### Procesy

**Do czego.** Zobaczyć, co działa na komputerach, i zakończyć wybrane procesy.

**Co możesz zrobić:** odczytać pamięć, właściciela, sesję, wiersz poleceń i **usługi działające w procesie**; zakończyć zaznaczone procesy.

**Jak działa.** Przez WinRM. Potwierdzenie ostrzega, gdy proces obsługuje usługi.

### Dyski

**Do czego.** Sprawdzić zajętość dysków i zwolnić miejsce, wiedząc najpierw, ile się zwolni.

**Co możesz zrobić**
- odczytać zajętość dysków lokalnych z oceną;
- wybrać kategorie czyszczenia: `Windows\Temp`, TEMP profili użytkowników, Kosz, raporty błędów Windows, pamięć optymalizacji dostarczania, zrzuty pamięci (domyślnie wyłączone — bywają potrzebne do diagnozy), pobrane aktualizacje; tylko pliki starsze niż N dni;
- **Podgląd** — dla każdej kategorii rozmiar i ile miejsca się zwolni, bez usuwania;
- **Wyczyść** — wiersz na kategorię (usunięte, w użyciu, pominięte) i podsumowanie dysku systemowego z odzyskanym miejscem.

**Jak działa.** Przez WinRM. Rozmiary są w czytelnych jednostkach (KB, MB, GB, TB), a sortowanie idzie według liczby bajtów. Pobrane aktualizacje są czyszczone z chwilowym zatrzymaniem usługi Windows Update. Folder programu `Temp\DomainOps` jest pomijany, puste podfoldery usuwane.

### Dziennik zdarzeń

**Do czego.** Przeszukiwać zdarzenia z wielu komputerów naraz, w jednej tabeli, i odpowiadać na pytania typu „kto zrestartował serwer”, „skąd blokada konta”, „czy to dzieje się wszędzie”.

**Co możesz zrobić**
- wybrać **dzienniki**: System, Application, Security, Setup i dowolne inne (*Wybierz…* pokazuje dzienniki z zapisanymi zdarzeniami);
- wybrać **poziomy**, w tym *inspekcja – sukces* i *inspekcja – niepowodzenie*;
- **filtrować**: ostatnie godziny, ID z zakresami i wykluczeniami (`4624, 4720-4767, -4634`), źródło, konto, fragmenty treści (`;` — wystarczy jeden, `!` — wyklucza);
- użyć **gotowych zestawów**: nieoczekiwane restarty, błędy dysków, awarie aplikacji, logowania, nieudane uwierzytelnienia, blokady kont, zmiany kont i grup, sesje RDP, nowe usługi i zadania, wyczyszczenie dziennika, zasady grupy, Windows Update, podejrzane skrypty PowerShell i inne;
- czytać dane zdarzenia w kolumnach *Konto*, *Skąd*, *Logowanie*, *Szczegóły*;
- zobaczyć **Podsumowanie** — zdarzenia pogrupowane według dziennika, ID i źródła: ile, na których komputerach, jakie konta;
- **Śledzić nowe zdarzenia** (co 30 s, z tymi samymi filtrami);
- z menu wiersza: te zdarzenia na zaznaczonych komputerach, zdarzenia tego konta, pomijanie identyfikatora.

**Jak działa.** Przez WinRM, z filtrem po stronie komputera. Brak dziennika, wyłączony dziennik albo brak dostępu to wiersz „Uwaga”, a nie błąd całego komputera; uwaga mówi też, gdy limit na komputer uciął starsze zdarzenia.

### Harmonogram zadań

**Do czego.** Zarządzać zadaniami Harmonogramu na komputerach.

**Co możesz zrobić:** przeglądać zadania; uruchamiać, zatrzymywać, włączać, wyłączać i usuwać; tworzyć proste zadania na koncie SYSTEM.

**Jak działa.** Przez WinRM, poleceniami Harmonogramu zadań.

### Rejestr – odczyt i eksport

**Do czego.** Odczytać wybrane klucze rejestru z wielu komputerów albo je wyeksportować — np. klucze TeamViewera przed reinstalacją.

**Co możesz zrobić**
- wpisać klucze, jeden w wierszu: `HKLM\…`, `HKU\…`, `HKCR\…`, także skopiowane z paska adresu regedit albo w zapisie PowerShell;
- użyć **szablonu TeamViewer** (klucz 64- i 32-bitowy z ClientID);
- **Pokaż wartości** — klucz, nazwa, typ i dane jak w regedit, opcjonalnie z podkluczami, z limitem liczby wartości;
- **Eksportuj do .reg** — jeden plik `<komputer>.reg` na komputer, w nowym folderze z datą; *Otwórz folder*.

**Jak działa.** Na każdym komputerze program uruchamia `reg export`, a treść wraca do programu i trafia do pliku w formacie regedit (UTF-16). Przywrócenie: uruchomienie pliku albo `reg import`.

**Uwagi.** HKCU jest odrzucany z podpowiedzią — przez WinRM to gałąź konta połączenia, a nie użytkownika; dla użytkownika użyj `HKU\<SID>`.

### Autostart

**Do czego.** Sprawdzić, co uruchamia się automatycznie, i usunąć zbędne wpisy.

**Co możesz zrobić:** odczytać klucze Run/RunOnce (komputer i załadowane profile użytkowników) i foldery Autostart; usunąć wybrane wpisy.

**Jak działa.** Przez WinRM, z rejestru i folderów Autostart.

### Drukarki

**Do czego.** Obsługiwać drukarki zainstalowane na komputerach.

**Co możesz zrobić:** odczytać drukarki, sterowniki, porty, udostępnianie i zadania w kolejce; wyczyścić kolejki, usunąć drukarki, zrestartować bufor wydruku.

**Jak działa.** Przez WinRM, poleceniami modułu drukowania.

### Sterowniki i urządzenia

**Do czego.** Sprawdzić sterowniki i znaleźć urządzenia z problemami.

**Co możesz zrobić:** odczytać zainstalowane sterowniki z filtrem; pokazać urządzenia zgłaszające problem w Menedżerze urządzeń (także z menu wiersza).

**Jak działa.** Przez WinRM, z klas WMI sterowników PnP i urządzeń.

---

## Oprogramowanie

### Zainstalowane programy

**Do czego.** Zobaczyć, co jest zainstalowane, i odinstalować programy zdalnie — wiedząc wcześniej, czy to się da zrobić bez użytkownika.

**Co możesz zrobić**
- odczytać programy z rejestru (64- i 32-bitowe oraz z profili zalogowanych użytkowników);
- sprawdzić w kolumnie **„Odinstalowanie zdalne”**, czy i jak program da się usunąć: MSI, tryb cichy producenta, Inno Setup (zielone); tylko z przełącznikami albo w sesji użytkownika (żółte); brak deinstalatora, blokada producenta albo profil bez trybu cichego (czerwone);
- odinstalować zaznaczone programy (z limitem czasu) albo *Odinstaluj z przełącznikami…*;
- *Pokaż ten program na wszystkich komputerach*.

**Jak działa.** Deinstalacja działa przez WinRM. Po niej program sprawdza, czy aplikacja zniknęła z rejestru, i pisze wynik z przyczyną w kolumnie „Stan” (np. „deinstalator czekał na odpowiedź użytkownika”, „MSI 1618 – trwa inna instalacja”, „nadal zainstalowany – wymaga restartu”). Programy z profilu zalogowanego użytkownika są usuwane jednorazowym zadaniem Harmonogramu w jego sesji.

### Windows Update

**Do czego.** Sprawdzić i zainstalować aktualizacje na wielu komputerach oraz zrozumieć, dlaczego coś się nie instaluje.

**Co możesz zrobić**
- wyszukać dostępne aktualizacje (ważność, rozmiar, KB, czy wymagają restartu);
- przejrzeć historię z opisem kodów błędów;
- **sprawdzić obecność konkretnych poprawek KB**;
- zainstalować **wszystkie albo zaznaczone** aktualizacje, z opcjonalnym restartem;
- sprawdzić **stan instalacji**: etap, wynik każdej aktualizacji, czy wymagany jest restart i dlaczego.

**Jak działa.** API Windows Update odmawia dostępu w sesji WinRM, więc wyszukiwanie, historia i instalacja działają jako zadanie Harmonogramu na koncie SYSTEM (zasady wykonywania skryptów z GPO go nie blokują). Wyszukiwanie czeka na wynik do 30 min. Pliki zadania i dzienniki są w `%SystemRoot%\Temp\DomainOps`. Typowe błędy (WSUS niedostępny, proxy, usługa wyłączona) są opisane po polsku.

### Outlook i Microsoft 365

**Do czego.** Sprawdzić Office, Outlooka, OneDrive i Teams na komputerach użytkowników oraz rozwiązać typowe problemy bez łączenia z chmurą.

**Co możesz zrobić** — pięć widoków:
- **Office i Microsoft 365:** produkty Click-to-Run, wersja, kanał aktualizacji, architektura, licencja współdzielona; Office z MSI; dołączenie do **Entra ID** i zapis w **Intune**;
- **Konta i skrzynki:** konta zalogowane w Office, **OneDrive** (konto, folder, foldery znane w OneDrive), profile Outlooka (konta, pliki OST i PST) oraz **skrzynki współdzielone, delegowane, archiwa i skrzynki grup** z automapowania; opcjonalnie także niezalogowani użytkownicy;
- **Pliki PST** (opcjonalnie OST): w profilach albo na wszystkich dyskach lokalnych; rozmiar, dni bez użycia, właściciel, czy plik jest otwarty w Outlooku; ostrzeżenia o dużych i osieroconych plikach; z menu wiersza **kopiowanie plików PST na udział** (np. przed migracją);
- **Dodatki Outlooka:** dodatki COM i VSTO ze stanem (w tym **wyłączone przez Outlooka**); z menu wiersza **włącz na stałe** albo **wyłącz u użytkownika**;
- **Teams, OneDrive, Outlook:** stan synchronizacji OneDrive, zasady grupy OneDrive, czy aplikacje działają i ile zajmuje ich pamięć podręczna; z menu wiersza **wyczyść pamięć Teams**, **wyczyść pamięć Outlooka**, **zresetuj OneDrive**.

**Jak działa.** Wyłącznie dane z komputera (rejestr, pliki, `dsregcmd`), bez łączenia z Exchange Online ani Graph. Zalogowanych użytkowników czyta z ich gałęzi rejestru, niezalogowanych — przez chwilowe załadowanie `NTUSER.DAT`. Skrzynki z automapowania pochodzą z kopii odpowiedzi Autodiscover zapisanej w profilu (skrzynki otwierane ręcznie nie są tam zapisywane). Kopiowanie PST wykonuje Twój komputer przez udział administracyjny; plik otwarty w Outlooku jest pomijany, oryginały zostają. Czyszczenie pamięci najpierw zamyka aplikację tego użytkownika. Reset OneDrive działa w sesji zalogowanego użytkownika przez jednorazowe zadanie Harmonogramu (pliki zostają, pełna synchronizacja może potrwać).

---

## Bezpieczeństwo

### Szybki audyt bezpieczeństwa

**Do czego.** W jednym przebiegu sprawdzić podstawowe zabezpieczenia wielu komputerów.

**Co możesz zrobić:** sprawdzić zaporę, antywirusa i sygnatury, BitLockera, SMBv1, NLA, UAC, LAPS, świeżość aktualizacji, oczekujący restart, Secure Boot, TPM, WDigest, ochronę LSA, LLMNR i konto Gość; zobaczyć wynik procentowy i listę problemów.

**Jak działa.** Przez WinRM, z ustawień systemu i rejestru.

### Microsoft Defender

**Do czego.** Sprawdzić ochronę antywirusową i uruchomić skan.

**Co możesz zrobić:** odczytać stan ochrony i sygnatur, wykryte zagrożenia; zaktualizować sygnatury; uruchomić szybki albo pełny skan (także z menu wiersza).

**Jak działa.** Przez WinRM, poleceniami modułu Defender. Skan trwa w tle — możesz pracować dalej.

### BitLocker

**Do czego.** Sprawdzić szyfrowanie dysków i zabezpieczyć klucze odzyskiwania.

**Co możesz zrobić:** odczytać stan szyfrowania woluminów; zrobić kopię kluczy odzyskiwania do AD (także z menu wiersza).

**Jak działa.** Przez WinRM. Klucze zapisane w AD przeglądasz w przestrzeni [Komputery AD → Klucze BitLocker](04-komputery-ad.md#klucze-bitlocker-ad).

### Zapora Windows

**Do czego.** Sprawdzić i zmienić reguły zapory na wielu komputerach.

**Co możesz zrobić:** odczytać stan profili i reguły (z portami, programami i adresami); włączyć, wyłączyć i usunąć zaznaczone reguły; utworzyć nowe.

**Jak działa.** Przez WinRM, poleceniami modułu NetSecurity.

### Certyfikaty komputera

**Do czego.** Znaleźć wygasające certyfikaty i je wyeksportować.

**Co możesz zrobić:** przejrzeć magazyny LocalMachine z oceną ważności; znaleźć certyfikaty wygasające w ciągu N dni; wyeksportować zaznaczone do plików `.cer`.

**Jak działa.** Przez WinRM, z magazynów certyfikatów komputera. Wygasające certyfikaty serwerów są też dostępne jako [raport cykliczny](06-domena.md#raporty-cykliczne).

---

## Udostępnianie

### Udziały sieciowe

**Do czego.** Sprawdzać, tworzyć i usuwać udziały SMB.

**Co możesz zrobić**
- odczytać udziały i ich uprawnienia;
- utworzyć udział z uprawnieniami: konto lub grupa (wpisane albo wybrane z AD) i **poziom z listy** — *Odczyt*, *Zmiana*, *Pełna kontrola*; opcjonalnie te same uprawnienia w NTFS na folderze;
- usunąć udział.

**Jak działa.** Przez WinRM. Bez żadnych wpisów Windows nadaje *Wszyscy: Odczyt*. Potwierdzenie pokazuje pełną listę uprawnień.

### Zasoby sieciowe użytkowników

**Do czego.** Zobaczyć, co mają podpięte użytkownicy (dyski, drukarki), skąd to pochodzi, oraz podłączać i odłączać dyski sieciowe w imieniu użytkownika.

**Co możesz zrobić** — widoki:
- **Dyski zmapowane** — litera, ścieżka, serwer, konto połączenia; zalogowani i opcjonalnie niezalogowani użytkownicy oraz historia połączeń Eksploratora. Kolumna **Pochodzenie**: *GPO: nazwa*, *skrypt logowania: plik* albo *ręcznie*;
- **Dyski z GPO** — Preferencje zasad grupy zastosowane na komputerze (litera, akcja, etykieta, targetowanie, nazwa GPO);
- **Drukarki sieciowe** — z rejestru i z GPO, z drukarką domyślną;
- **Podłączeni do komputera** — sesje SMB (klient, konto, otwarte pliki, bezczynność) i opcjonalnie otwarte pliki.

Akcje:
- **Podłącz dysk…** — ścieżka `\\serwer\udział`, litera (wybrana albo pierwsza wolna od Z:) i opcjonalna etykieta; dla użytkowników zaznaczonych w tabeli, wszystkich zalogowanych albo jednego po loginie;
- **Odłącz dysk sieciowy…** — tylko dyski zmapowane ręcznie (dyski z GPO i skryptów logowania i tak wróciłyby przy logowaniu);
- zamknięcie sesji SMB lub pliku, otwarcie ścieżki w Eksploratorze.

**Jak działa**
- Dyski zalogowanych program czyta z ich gałęzi rejestru (polecenie w sesji WinRM nie widzi cudzych dysków), niezalogowanych — przez chwilowe załadowanie `NTUSER.DAT`.
- Pochodzenie *skrypt logowania* program ustala, czytając skrypty u siebie, w tle (skrypty z AD i GPO). Ścieżek składanych ze zmiennych nie rozpozna.
- Podłączanie i odłączanie u zalogowanego działa **w jego sesji, bez okna konsoli** — przez jednorazowe zadanie Harmonogramu i krótki skrypt `wscript.exe`. Gdy Windows Script Host jest wyłączony, program używa `net use` (okno konsoli mignie).
- U niezalogowanego program zapisuje albo usuwa wpis trwały w `NTUSER.DAT` — dysk pojawi się albo zniknie przy logowaniu.
- W trybie diagnostycznym każdy krok podłączania i odłączania trafia do dziennika.

**Uwagi**
- Połączenie zawsze działa na koncie użytkownika, bez hasła.
- Programy uruchomione jako administrator nie widzą takiego dysku — to zachowanie UAC.
- Nie używaj opcji „Także niezalogowani”, gdy użytkownicy właśnie się logują — załadowany plik profilu na chwilę blokuje logowanie.
- Zamknięcie sesji SMB lub pliku powoduje utratę niezapisanych zmian użytkownika.
