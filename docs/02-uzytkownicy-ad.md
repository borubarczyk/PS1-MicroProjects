# 2. Użytkownicy AD

[← Spis treści](README.md)

**Do czego jest ta przestrzeń.** Codzienna obsługa kont użytkowników w Active Directory: szczegóły, hasła i blokady, stan konta, atrybuty, grupy, odejście pracownika, hurtowe tworzenie kont i import danych, porządki oraz raporty.

**Lista po lewej:** konta użytkowników — wyszukiwanie w AD (login, nazwisko, e-mail, gwiazdka), filtr stanu (aktywne, wyłączone, zablokowane), OU, loginy wpisane ręcznie albo z pliku. Kropka przy koncie pokazuje stan (np. zablokowane).

**Jak to działa ogólnie.** Moduły korzystają z modułu `ActiveDirectory` (RSAT) na Twoim komputerze i działają z uprawnieniami Twojego konta (albo konta alternatywnego z górnego paska). Kontroler domeny wybierasz w Ustawieniach (puste — automatycznie). Zmiany w AD wymagają potwierdzenia z listą kont; zmiany hurtowe mają podgląd.

**Spis modułów**

- [Konta](#konta): [Szczegóły konta](#szczegóły-konta) · [Hasło i blokada](#hasło-i-blokada) · [Stan konta](#stan-konta) · [Edycja atrybutów](#edycja-atrybutów) · [Offboarding](#offboarding)
- [Grupy](#grupy): [Członkostwo w grupach](#członkostwo-w-grupach)
- [Tworzenie i import](#tworzenie-i-import): [Tworzenie kont](#tworzenie-kont) · [Konta MAB (802.1X)](#konta-mab-8021x) · [Import atrybutów z CSV](#import-atrybutów-z-csv)
- [Porządki](#porządki): [Ostatnio utworzone](#ostatnio-utworzone) · [Wyłączone konta](#wyłączone-konta)
- [Raporty](#raporty): [Raporty kont](#raporty-kont) · [Źródło blokady konta](#źródło-blokady-konta) · [Gdzie zalogowany](#gdzie-zalogowany)

---

## Konta

### Szczegóły konta

**Do czego.** Zobaczyć wszystko o koncie w jednym miejscu albo przygotować raport o kilku kontach.

**Co możesz zrobić**
- odczytać stan, kontakt, przełożonego, logowania, hasło, wygaśnięcie, profil, OU i SID zaznaczonych kont;
- utworzyć **raport HTML** — kartę każdego konta: tożsamość (login, UPN, e-mail i aliasy, SID), organizacja, stan i blokady, hasło (wiek, wygaśnięcie, flagi), profil i skrypt logowania, **grupy** (bezpośrednie, podstawowa i opcjonalnie zagnieżdżone), **podwładni** oraz ostrzeżenia (np. „hasło nigdy nie wygasa”, konto uprzywilejowane).

**Jak działa.** Odczyt z AD. Raport HTML to samodzielny plik z kafelkami, spisem kont, wyszukiwarką i stylem do druku.

### Hasło i blokada

**Do czego.** Obsłużyć zgłoszenia „nie mogę się zalogować”: odblokować konto, zresetować hasło.

**Co możesz zrobić**
- odczytać stan haseł i blokad;
- odblokować konta;
- zresetować hasło — **losowe, inne dla każdego konta**, albo wpisane;
- wymusić zmianę hasła przy logowaniu i ustawić albo zdjąć „hasło nigdy nie wygasa” (także z menu wiersza).

**Jak działa.** Zmiany przez moduł ActiveDirectory. Wygenerowane hasła są wartościami poufnymi: maskowane w tabeli i eksporcie, kopiowane z automatycznym czyszczeniem schowka po 60 s. Program ich nigdzie nie zapisuje.

### Stan konta

**Do czego.** Włączać i wyłączać konta, ustawiać datę wygaśnięcia, przenosić konta między OU.

**Co możesz zrobić**
- wyłączyć konto z dopiskiem w opisie (data, autor, powód) i przeniesieniem do wybranej OU;
- włączyć konto;
- ustawić albo usunąć datę wygaśnięcia;
- przenieść konto do innej OU.

**Jak działa.** Zmiany przez moduł ActiveDirectory, z potwierdzeniem i listą kont. Wszystkie akcje są też w menu wiersza.

### Edycja atrybutów

**Do czego.** Odczytać albo zmienić hurtowo te same atrybuty wielu kont.

**Co możesz zrobić**
- odczytać atrybuty zaznaczonych kont (stanowisko, dział, firma, biuro, telefony, e-mail, przełożony, extensionAttribute1–15 i inne);
- ustawić wartość wielu kontom naraz, z polami `{login}`, `{imie}`, `{nazwisko}`, `{nazwa}` — np. `{login}@firma.pl`.

**Jak działa.** Zmiana przez moduł ActiveDirectory, po potwierdzeniu z listą kont. Pola są podstawiane osobno dla każdego konta.

### Offboarding

**Do czego.** Obsłużyć odejście pracownika w jednym kroku, z kopią stanu konta i możliwością cofnięcia.

**Co możesz zrobić** — dla kont zaznaczonych po lewej, z wyborem kroków (wybór jest zapamiętywany):
- **wyłączyć konto**;
- ustawić **losowe hasło** (24 znaki, nigdzie nie pokazywane ani zapisywane);
- **usunąć z grup** (bez grupy podstawowej i grup chronionych);
- **dopisać do opisu** datę, autora i powód (np. numer zgłoszenia);
- ustawić wygaśnięcie konta z końcem dnia, wyczyścić przełożonego, ukryć w książce adresowej (gdy schemat ma atrybuty Exchange);
- **przenieść do OU**;
- **wylogować sesje** na serwerach farmy RDS (lista z modułu *Sesje RDS*);
- **przywrócić konto z kopii** z wyborem elementów (grupy, OU, opis, przełożony, wygaśnięcie, książka adresowa, włączenie).

**Jak działa**
- **Podgląd** pokazuje liczbę grup do usunięcia, przełożonego i ostrzega o podwładnych i kontach uprzywilejowanych.
- Przed zmianami program zapisuje **stan każdego konta** w pliku JSON (z protokołem wyników kroków) i kopię członkostw w CSV. Bez udanego zapisu kopii nic nie jest zmieniane.
- Każdy krok ma własny wynik — błąd jednego nie przerywa pozostałych.

**Uwagi.** Konta wbudowane (Administrator, krbtgt), konta-wyjątki i konto, na którym działa program, są pomijane. Kopie: `%APPDATA%\AD-ManagerDiamond\Offboarding` i `…\Backup\Czlonkostwa_offboarding_*.csv`.

---

## Grupy

### Członkostwo w grupach

**Do czego.** Sprawdzić i zmienić, do jakich grup należą konta.

**Co możesz zrobić**
- odczytać grupy zaznaczonych kont — bezpośrednie i zagnieżdżone;
- dodać konta do grup (wyszukiwarka grup) i usunąć z zaznaczonych grup;
- **skopiować członkostwo z konta wzorcowego** (np. nowy pracownik „jak Kowalski”).

**Jak działa.** Odczyt i zmiany przez moduł ActiveDirectory. Ten sam moduł jest w przestrzeniach Komputery AD (dla kont komputerów) i Grupy i OU (*Przynależność do grup* — zagnieżdżanie grup).

---

## Tworzenie i import

### Tworzenie kont

**Do czego.** Założyć wiele kont naraz z danych wklejonych z Excela, według ustalonych profili (np. Pracownik, Student, Wykładowca).

**Co możesz zrobić**
- wkleić dane: imię, nazwisko, numer albumu lub pracownika, dział, stanowisko, miasto, opis;
- wybrać **profil** i edytować go w **konfiguratorze profilu** (z podglądem przykładowego konta): OU, sufiks UPN, domena e-mail, format loginu (`i.nazwisko`, `inazwisko`, `imie.nazwisko`, `nazwisko.imie`, numer…) i nazwy wyświetlanej, opis, firma, dział, grupy, opcje konta (włączone, zmiana hasła przy logowaniu, wygaśnięcie po N dniach, długość hasła);
- przejrzeć **podgląd** i poprawić dane w tabeli;
- utworzyć konta, skopiować dane logowania i uruchomić synchronizację z Microsoft Entra ID.

**Jak działa**
- Polskie znaki są zamieniane w loginach, imiona i nazwiska poprawiane (wielkość liter, nazwiska dwuczłonowe).
- Podgląd sprawdza w AD zajęte loginy, UPN i nazwy CN i numeruje je automatycznie; hasło jest losowe dla każdego konta.
- Synchronizacja to `Start-ADSyncSyncCycle -PolicyType Delta` — lokalnie albo na serwerze Entra Connect.

**Uwagi.** Hasła nowych kont są maskowane w tabeli i eksporcie i nie są zapisywane w ustawieniach ani dzienniku.

### Konta MAB (802.1X)

**Do czego.** Zakładać konta do uwierzytelniania urządzeń adresem MAC (MAC Authentication Bypass, np. przez NPS) — z hasłem, które przejdzie przez zasady haseł.

**Co możesz zrobić**
- wkleić adresy MAC w dowolnym zapisie (`00:11:22:AA:BB:CC`, `00-11-…`, `0011.22aa.bbcc`) z opcjonalną nazwą urządzenia i opisem;
- wybrać format loginu i hasła taki, jaki wysyła przełącznik (bez separatorów, z myślnikami, format Cisco z kropkami; małe albo wielkie litery);
- wybrać jednostkę z listy OU (stale widocznej po prawej) i **PSO** (szczegółowe zasady haseł) z listy odczytanej z AD;
- dodać konta do grup, opcjonalnie ustawić pierwszą grupę jako **grupę podstawową** zamiast *Domain Users*;
- zarządzać istniejącymi kontami w OU: ponowne ustawienie hasła i PSO, włączanie, wyłączanie i usuwanie z menu wiersza.

**Jak działa**
- Podgląd wykrywa błędne, zerowe i powtórzone adresy, adresy grupowe, istniejące konta i zajęte nazwy CN.
- Program **sprawdza zgodność hasła ze skutecznymi zasadami konta** tak, jak liczy je AD: PSO przypisane do konta, w przeciwnym razie PSO grup konta, w przeciwnym razie zasady domeny.
- Konto jest tworzone wyłączone, potem dostaje grupy i PSO, flagi (hasło nie wygasa, użytkownik nie może zmienić hasła, opcjonalnie szyfrowanie odwracalne dla CHAP/EAP-MD5), hasło — i dopiero wtedy jest włączane.

**Uwagi.** Złożoność hasła w AD zawsze odrzuci hasło równe nazwie konta — dlatego zwykle potrzebne jest PSO bez wymogu złożoności dla grupy kont MAB.

### Import atrybutów z CSV

**Do czego.** Zaktualizować dane wielu kont (stanowisko, dział, telefon, przełożony…) z pliku CSV, np. z kadr.

**Co możesz zrobić**
- wczytać plik CSV (separator rozpoznawany automatycznie) i przypisać kolumny do atrybutów;
- zobaczyć podgląd **„było → będzie”** dla każdego konta;
- zastosować zmiany;
- **cofnąć** zmiany z zapisanego pliku.

**Jak działa.** Przed zmianą program zapisuje poprzednie wartości w `%APPDATA%\AD-ManagerDiamond\Rollback\Import_*.json` — z niego działa przywracanie.

---

## Porządki

### Ostatnio utworzone

**Do czego.** Sprawdzić, co niedawno powstało w AD i kto to utworzył — i szybko posprzątać pomyłki.

**Co możesz zrobić:** znaleźć konta użytkowników, komputery, grupy i jednostki utworzone w ostatnich minutach, godzinach lub dniach, z twórcą (właścicielem obiektu); wyłączyć albo usunąć zaznaczone obiekty.

**Jak działa.** Zapytanie do AD po dacie utworzenia obiektu; usuwanie i wyłączanie z potwierdzeniem.

### Wyłączone konta

**Do czego.** Uporządkować wyłączone konta: przenieść je do jednej OU i usunąć z grup, zachowując możliwość cofnięcia.

**Co możesz zrobić**
- znaleźć wyłączone konta poza jednostką dla nieaktywnych;
- przenieść je do tej OU i usunąć z grup;
- przywrócić członkostwa z kopii.

**Jak działa.** Przed usunięciem z grup program zapisuje kopię członkostw w CSV (`%APPDATA%\AD-ManagerDiamond\Backup\Czlonkostwa_wylaczonych_*.csv`). Grupy chronione i konta-wyjątki (do ustawienia) są pomijane.

---

## Raporty

### Raporty kont

**Do czego.** Zestawienia kont z całej domeny lub OU — i od razu operacje na wynikach.

**Co możesz zrobić**
- wybrać raport: zablokowane, wyłączone, nieaktywne, nigdy nie logowane, hasło wygasa / wygasło / nigdy nie wygasa, konta wygasające, nowe, uprzywilejowane;
- przenieść wyniki na listę kont po lewej i wykonać na nich operacje w innych modułach.

**Jak działa.** Zapytania do AD (konta uprzywilejowane według `adminCount`). Te same raporty są dostępne jako [raporty cykliczne](06-domena.md#raporty-cykliczne).

### Źródło blokady konta

**Do czego.** Znaleźć komputer, z którego przychodzą błędne hasła blokujące konto (np. stary telefon, zapisane poświadczenia, usługa).

**Co możesz zrobić**
- odczytać zdarzenia blokady 4740 z emulatora PDC albo ze wszystkich kontrolerów — z nazwą komputera źródłowego;
- opcjonalnie dołączyć nieudane uwierzytelnienia Kerberos 4771 i NTLM 4776 z adresem IP lub stacją;
- bez zaznaczonych kont — zobaczyć wszystkie blokady.

**Jak działa.** Odczyt dziennika Security kontrolerów domeny. Wymaga prawa odczytu tego dziennika.

### Gdzie zalogowany

**Do czego.** Ustalić, na których komputerach są zalogowani wskazani użytkownicy — i obsłużyć te sesje.

**Co możesz zrobić**
- podać loginy (albo zaznaczyć konta po lewej);
- wybrać komputery z AD: stacje robocze, serwery albo wszystkie, włączone, aktywne w ostatnich N dniach, opcjonalnie z OU, najwyżej M najświeżej logowanych;
- zobaczyć sesje konsoli i pulpitu zdalnego szukanych kont (stan, klient, bezczynność, czas logowania);
- z menu wiersza: wiadomość, podgląd, pomoc zdalna, procesy, rozłączenie, wylogowanie, reset.

**Jak działa.** Sprawdza sesje na komputerach przez WinRM (jak moduł *Sesje użytkowników*). Kafelki pokazują sprawdzone komputery, komputery bez połączenia (lista w dzienniku) i znalezione sesje.
