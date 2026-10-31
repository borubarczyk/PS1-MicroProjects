# 6. Domena

[← Spis treści](README.md)

**Do czego jest ta przestrzeń.** Spojrzenie na całą domenę naraz: audyt bezpieczeństwa Active Directory, delegacje uprawnień, stan kontrolerów i replikacji, zasady grupy, serwery (farma RDS, weryfikacja według ról) oraz raporty cykliczne działające bez otwierania programu.

**Lista po lewej:** brak — moduły działają na całej domenie albo na serwerach wpisanych w panelu parametrów.

**Jak to działa ogólnie.** Odczyty z AD przez moduł `ActiveDirectory` (zasady grupy także z folderu SYSVOL), a z kontrolerów i serwerów — przez WinRM. Moduły bezpieczeństwa i zasad grupy są **tylko do odczytu** (poza kopią GPO).

**Spis modułów**

- [Bezpieczeństwo](#bezpieczeństwo): [Audyt bezpieczeństwa AD](#audyt-bezpieczeństwa-ad) · [Delegacje uprawnień w AD](#delegacje-uprawnień-w-ad)
- [Stan domeny](#stan-domeny): [Stan domeny](#stan-domeny-1)
- [Zasady grupy](#zasady-grupy): [Obiekty zasad grupy](#obiekty-zasad-grupy) · [Wynikowe zasady (RSoP)](#wynikowe-zasady-rsop)
- [Serwery](#serwery): [Sesje RDS](#sesje-rds) · [Weryfikacja serwerów](#weryfikacja-serwerów)
- [Automatyzacja](#automatyzacja): [Raporty cykliczne](#raporty-cykliczne)

---

## Bezpieczeństwo

### Audyt bezpieczeństwa AD

**Do czego.** Ocenić bezpieczeństwo domeny jedną liczbą i listą konkretnych problemów z zaleceniami.

**Co możesz zrobić**
- uruchomić 34 kontrole w pięciu grupach, każdą z oceną *wysokie / średnie / niskie / bez uwag*, listą obiektów i zaleceniem:
  - **konta uprzywilejowane** — SPN na kontach administratorów, hasła bez wygasania i stare hasła, nieużywane i wyłączone konta w grupach, brak ochrony przed delegowaniem, członkowie Schema i Enterprise Admins, używanie wbudowanego Administratora, pozostałości `adminCount = 1`;
  - **Kerberos i delegowanie** — konta z SPN (kerberoasting), bez wstępnego uwierzytelnienia (AS-REP roasting), nieograniczone delegowanie poza kontrolerami, delegowanie z dowolnym protokołem, RBCD, wiek hasła `krbtgt`, DES;
  - **hasła** — brak wymaganego hasła, szyfrowanie odwracalne, hasła bez wygasania, zasady haseł domeny i zasady szczegółowe;
  - **konta i komputery** — nieaktywne, konto Gość, nieobsługiwane systemy, komputery bez LAPS, SID History;
  - **domena** — limit dołączania komputerów, kosz AD, poziom funkcjonalny, grupa Protected Users, dostęp anonimowy przez „Pre-Windows 2000 Compatible Access”, zaufania bez filtrowania SID;
- ustawić progi (nieaktywność, wiek hasła administratora i `krbtgt`, liczba kont uprzywilejowanych);
- z wiersza: zobaczyć listę obiektów i **zaznaczyć konta lub komputery na listach** do dalszych operacji;
- utworzyć **raport HTML** z oceną, zestawieniem i kartą każdej kontroli.

**Jak działa.** Kontrole działają równolegle, tylko odczyt z AD. Ocena to 100 minus wagi niespełnionych kontroli; gdy któraś kontrola się nie wykonała, ocena jest oznaczona jako niepełna.

### Delegacje uprawnień w AD

**Do czego.** Sprawdzić, kto ma jakie uprawnienia do jednostek organizacyjnych i katalogu głównego domeny — i znaleźć niebezpieczne delegacje (np. DCSync).

**Co możesz zrobić**
- odczytać uprawnienia do **OU**, **katalogu głównego domeny** i **AdminSDHolder** (opcjonalnie kontenerów Users i Computers);
- czytać wpisy opisane po polsku: reset hasła, tworzenie i usuwanie obiektów danej klasy, zapis atrybutów (np. `member`), pełna kontrola, zmiana uprawnień i właściciela — z zakresem;
- zobaczyć ocenę ryzyka:
  - **wysokie** — zapis dla szerokich grup, prawa replikacji poza kontrolerami (**DCSync**), kontrola nad katalogiem głównym, dodatkowe wpisy AdminSDHolder;
  - **średnie** — pełna kontrola, wszystkie prawa rozszerzone albo zapis wszystkich atrybutów w jednostce;
  - **niskie** — delegacja dla konta zamiast grupy, nierozwiązany SID, nietypowy właściciel;
- z wiersza: zobaczyć wszystkie delegacje danej tożsamości.

**Jak działa.** Odczyt list ACL z AD; nazwy klas, atrybutów i praw rozszerzonych program czyta ze schematu. Wpisy domyślne (SYSTEM, administratorzy, kontrolery, SELF, odczyt dla wszystkich) są ukryte. Tylko odczyt.

---

## Stan domeny

### Stan domeny

**Do czego.** Sprawdzić zdrowie domeny w jednym przeglądzie — kontrolery, replikację, role, DNS, kopie zapasowe — z oceną i wskazówką przy każdym punkcie.

**Co możesz zrobić**
- z Twojego komputera sprawdzić:
  - **kontrolery** (system, lokacja, wykaz globalny, RODC; ostrzeżenie przy jednym kontrolerze albo starym systemie);
  - **metadane kontrolerów** — pozostałości po usuniętych kontrolerach;
  - opcjonalny **test portów** (DNS, Kerberos, LDAP, SMB, ADWS, GC);
  - **replikację** każdego kontrolera od każdego partnera i dla każdej partycji;
  - **role FSMO**, **rekordy SRV** w DNS, **ostatnią kopię zapasową AD** każdej partycji, poziomy funkcjonalne;
- na kontrolerach (WinRM) sprawdzić: **usługi** (AD DS, Netlogon, KDC, czas, ADWS, DNS, DFSR), **SYSVOL i NETLOGON** z replikacją DFSR, **odchylenie zegara od emulatora PDC** i źródło czasu, wolne miejsce, **błędy w dziennikach**, czas działania;
- ustawić progi; z wiersza kontrolera uruchomić **dcdiag /q**.

**Jak działa**
- Każdy obszar przeglądu jest sprawdzany osobno — błąd jednego nie przerywa reszty, a komunikat podaje polecenie, które zawiodło.
- Gdy polecenia AD zgłaszają błąd „Specified argument was out of the range of valid values. Parameter name: index” (typowy przy pozostałościach po usuniętym kontrolerze), dane domeny, lasu i role FSMO program czyta wprost z katalogu, z ostrzeżeniem w wynikach.
- Gdy mimo to przegląd się przerwie, wyniki zebrane do tej chwili zostają, a wiersz *przegląd przerwany* pokazuje polecenie, które zawiodło (w trybie diagnostycznym także stos wywołań).

Ten sam przegląd jest dostępny jako [raport cykliczny](#raporty-cykliczne).

---

## Zasady grupy

### Obiekty zasad grupy

**Do czego.** Znaleźć problemy w zasadach grupy: niepodlinkowane, puste, niespójne, źle filtrowane, z niebezpiecznymi uprawnieniami do edycji.

**Co możesz zrobić**
- przeglądać trzy widoki:
  - **obiekty GPO** — stan, linki, ustawienia komputera i użytkownika, wersje, zgodność z SYSVOL, filtr WMI, „stosowany do”, edycja poza administratorami, data zmiany, ocena;
  - **problemy** z zaleceniami;
  - **linki** — kolejność, wymuszenie, link wyłączony, blokada dziedziczenia;
- zobaczyć **dziedziczenie dla jednostki** — kolejność GPO jak w karcie GPMC;
- zrobić **kopię zapasową** zaznaczonych albo wszystkich GPO (format GPMC, z raportem ustawień HTML);
- wygenerować **raport ustawień** wybranego GPO.

**Co wykrywa:** GPO niepodlinkowane, z samymi wyłączonymi linkami, puste, wyłączone a podlinkowane, z wyłączoną częścią zawierającą ustawienia, niezgodną wersję AD i SYSVOL, brak folderu albo `GPT.INI` w SYSVOL, brak prawa „Zastosuj”, filtrowanie bez prawa odczytu dla komputerów (**MS16-072** — ustawienia użytkownika nie zadziałają), edycję dla szerokich grup, nieistniejący filtr WMI, linki do nieistniejących GPO i osierocone foldery w SYSVOL.

**Jak działa.** Odczyt z samego modułu ActiveDirectory (obiekty GPO, linki domeny, jednostek i lokacji, filtry WMI) i folderu Policies w SYSVOL. Kopia i raport ustawień wymagają modułu `GroupPolicy`. Tylko odczyt poza kopią. Problemy z zasadami grupy są też dostępne jako [raport cykliczny](#raporty-cykliczne).

### Wynikowe zasady (RSoP)

**Do czego.** Sprawdzić, które GPO zadziałały na komputerze (i dla użytkownika), a które zostały odrzucone i dlaczego.

**Co możesz zrobić**
- wpisać komputery i opcjonalnie użytkownika;
- zobaczyć informacje (konto, jednostka, lokacja, wolne łącze), GPO **zastosowane** w kolejności z lokalizacją linku i rozszerzeniami, GPO **odrzucone z przyczyną** (filtrowanie zabezpieczeń, filtr WMI, wyłączone, puste) i **błędy rozszerzeń** z kodem HRESULT;
- z wiersza: otworzyć raport HTML z gpresult.

**Jak działa.** `gpresult /x` (opcjonalnie `/h`) uruchamiany na komputerach przez WinRM.

---

## Serwery

### Sesje RDS

**Do czego.** Obsługa sesji na serwerach farmy pulpitu zdalnego: znaleźć użytkownika na całej farmie, rozłączyć, wylogować, zresetować zawieszoną sesję.

**Co możesz zrobić**
- wybrać serwery: wpisane ręcznie (zapamiętywane), **z brokera połączeń RD** (kolekcje i serwery sesji) albo z AD według wzorca nazwy;
- zobaczyć sesje jak w module *Sesje użytkowników* plus kolumnę *Kolekcja*;
- **szukać użytkownika na całej farmie** i filtrować *tylko rozłączone*;
- dla jednej albo kilku sesji: wylogowanie, rozłączenie, **reset** (`rwinsta`), wiadomość, procesy w sesji, podgląd;
- **Wyloguj rozłączone…** — sesje rozłączone dłużej niż podana liczba minut, z listą do potwierdzenia.

**Jak działa.** Sesje przez WTS API na serwerach (WinRM); lista z brokera — modułem RemoteDesktop na brokerze. Dziennik wskazuje serwery, które nie przyjmują nowych połączeń. Moduł *Offboarding* korzysta z tej listy serwerów przy wylogowaniu odchodzącego pracownika.

### Weryfikacja serwerów

**Do czego.** Przegląd serwerów dopasowany do tego, co na nich działa — z listą błędów i ostrzeżeń na górze.

**Co możesz zrobić**
- sprawdzić kontrole **ogólne**: system i koniec wsparcia, czas od restartu, oczekujący restart, ostatnią poprawkę, wolne miejsce, zatrzymane usługi automatyczne, źródło czasu, błędy w dzienniku System z 24 godzin, zaporę;
- sprawdzić kontrole **ról**: kontroler domeny, DNS, DHCP, IIS, serwer plików, serwer wydruku, Hyper-V, host sesji RD, licencjonowanie RD, WSUS, urząd certyfikacji, klaster, kopia zapasowa Windows Server;
- ustawić progi dysków, aktualizacji i kopii; *Ukryj wyniki OK*;
- z wiersza: *Uruchom zatrzymane usługi…* (potem ponowna kontrola tego serwera), *Sesje użytkowników na tych serwerach*, *odśwież tylko te komputery*.

**Jak działa.** Role są rozpoznawane przez `Get-WindowsFeature`, a bez modułu ServerManager — po usługach. Każda kontrola ma ocenę (*błąd / ostrzeżenie / info / OK*), wynik, szczegóły i zalecenie. Brak modułu roli albo uprawnień daje wiersz „Info”, a nie błąd całego serwera. Ten sam przegląd jest dostępny jako [raport cykliczny](#raporty-cykliczne).

---

## Automatyzacja

### Raporty cykliczne

**Do czego.** Uruchamiać wybrane przeglądy według harmonogramu — bez otwierania programu — i dostawać raport pocztą, najlepiej tylko wtedy, gdy coś się zmieniło.

**Rodzaje raportów**
- audyt bezpieczeństwa AD, stan domeny, delegacje uprawnień w AD;
- **członkowie grup uprzywilejowanych** (także przez grupy zagnieżdżone) — z wysyłką „gdy są zmiany” działa jak alarm o dodaniu kogoś do Domain Admins;
- raporty kont (np. hasło wygasa w ciągu N dni, nieaktywne, zablokowane) i komputerów (nieaktywne, nieobsługiwane systemy, bez LAPS…);
- problemy z zasadami grupy;
- zgodność grup lokalnych z szablonem;
- ryzyka uprawnień NTFS;
- wolne miejsce na dyskach, wygasające certyfikaty serwerów, weryfikacja serwerów.

**Co możesz zrobić**
- utworzyć definicję raportu w edytorze: opcje rodzaju, wynik, poczta, harmonogram, konto zadania, kontroler domeny;
- uruchomić raport od razu (*Uruchom teraz*) albo zadanie w Harmonogramie na jego koncie;
- z wiersza: edytować, otworzyć ostatni raport i folder raportów, przejrzeć **historię uruchomień** i **dziennik przebiegów**, wysłać **wiadomość testową**, duplikować, utworzyć zadanie ponownie, włączyć lub wyłączyć zadanie, usunąć.

**Wynik**
- raport HTML (ten sam wygląd i filtry co inne raporty programu) w folderze lokalnym albo na udziale, stała kopia `…_najnowszy.html`, opcjonalnie CSV, retencja N ostatnich raportów;
- każdy przebieg porównuje wyniki z poprzednim: sekcja **„Zmiany od poprzedniego raportu”** (co nowego i co ustąpiło); pierwszy przebieg zapisuje stan odniesienia;
- dla raportów z oceną ryzyka wybierasz, od jakiego poziomu wyniki się liczą.

**Poczta**
- tryb: zawsze, gdy są wyniki albo gdy są zmiany; błąd raportu zawsze wysyła powiadomienie;
- treść z kafelkami i listą nowych i ustąpionych pozycji (czytelna w Outlooku), raport w załączniku, temat z podsumowaniem (np. *ocena 72/100, nowe: 3*);
- **Serwer SMTP** — przekaźnik bez hasła, konto zadania (Windows) albo login i hasło (zaszyfrowane DPAPI);
- **Microsoft Graph** (Microsoft 365, bez SMTP):
  - aplikacja w Entra ID z uprawnieniem aplikacji *Mail.Send* i zgodą administratora; nadawca (*Od*) to skrzynka w Exchange Online;
  - logowanie aplikacji **certyfikatem** (zalecane — działa też dla zadań na SYSTEM i gMSA, gdy konto zadania może odczytać klucz prywatny) albo **kluczem tajnym** (zaszyfrowanym DPAPI — odczyta go tylko konto, które go wpisało);
  - opcjonalny serwer proxy;
  - raport do 3 MB idzie w treści wiadomości; większy — przez sesję przekazywania pliku, co wymaga też *Mail.ReadWrite* (bez tego wiadomość idzie bez załącznika, z informacją, gdzie jest raport);
  - **ogranicz aplikację do skrzynki nadawcy** (zasady dostępu aplikacji w Exchange Online) — bez tego może wysyłać jako dowolna skrzynka w dzierżawie.

**Harmonogram i konto zadania**
- codziennie, w wybrane dni tygodnia albo w dzień miesiąca (także ostatni), o podanej godzinie;
- konto zadania: wskazane konto z hasłem zapisanym w Harmonogramie, bieżący użytkownik (tylko gdy zalogowany), SYSTEM (w domenie działa jako konto komputera) albo gMSA; opcja najwyższych uprawnień.

**Jak działa**
- Definicje, stan i raporty są w folderze zadań (domyślnie `%APPDATA%\AD-ManagerDiamond\Reports`, do zmiany w module). Zadania trafiają do Harmonogramu, folder `\Domain Ops\`.
- Zadanie uruchamia **kopię programu** z folderu zadań w trybie bez okna: `powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass -File "<folder zadań>\Program\AD-ManagerDiamond.ps1" -RunReport "<folder zadań>\<id>.json"`. Kopia jest odświeżana przy każdym zapisie definicji, a folder `Program` dostaje uprawnienia tylko dla administratorów, SYSTEM, Ciebie i (odczyt) konta zadania.
- Kod wyjścia: **0** — raport gotowy, **1** — raport z problemami częściowymi (np. niedostępny serwer) albo niewysłana poczta, **2** — raport nie powstał. Widać go w kolumnie *Wynik ostatniego uruchomienia* Harmonogramu.
- W tym trybie program nie pyta o PIN.

**Uwagi**
- Konto zadania musi mieć dostęp do folderu zadań i folderu raportów. Dla gMSA i innych kont wybierz folder zadań poza swoim profilem (np. `C:\ProgramData\DomainOps\Raporty`).
- Raporty i wiadomości zawierają nazwy kont, grup i ścieżek — wysyłaj je tylko do zaufanych odbiorców.
- Folder zadań powinien być zapisywalny tylko dla administratorów: kto może zmieniać definicje, może zmienić adresatów raportów.
