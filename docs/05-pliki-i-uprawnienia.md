# 5. Pliki i uprawnienia

[← Spis treści](README.md)

**Do czego jest ta przestrzeń.** Uprawnienia NTFS na folderach i udziałach: nadawanie, raportowanie, sprawdzanie, kto ma jaki dostęp, odbieranie, naprawa, porównania i kopie z przywracaniem. Do tego sumy kontrolne plików.

**Lista po lewej:** brak — moduły działają na ścieżkach folderów wpisanych w panelu parametrów.

**Jak to działa ogólnie**
- **Gdzie działa:** ścieżka lokalna, ścieżka sieciowa (`\\serwer\udział\…`) albo — po wpisaniu nazwy serwera plików — na samym serwerze przez PowerShell Remoting. Na serwerze skan idzie po ścieżkach lokalnych, bez ruchu sieciowego dla każdego folderu; przycisk wyboru folderu przegląda wtedy dyski i udziały tego serwera.
- **Tożsamości** są rozpoznawane w AD i zapisywane po SID (także `BUILTIN\Users`), więc działa w każdym języku systemu. Grupy są rozwijane z AD, także zagnieżdżone.
- **Kopia przed zmianą:** każda zmiana uprawnień najpierw zapisuje kopię uprawnień zmienianych folderów (`%APPDATA%\AD-ManagerDiamond\AclBackup`). Jeśli kopii nie da się zapisać, zmiana nie jest wykonywana. Folder, którego uprawnienia zmieniły się od podglądu, jest pomijany.
- **Tryb kopii zapasowej** (w modułach NTFS) włącza przywileje kopii zapasowej, przywracania i przejmowania własności — program czyta i zmienia uprawnienia także tam, gdzie Administratorzy nie mają dostępu, bez przejmowania własności. Używaj go świadomie.

**Spis modułów**

- [Uprawnienia NTFS](#uprawnienia-ntfs): [Nadawanie uprawnień](#nadawanie-uprawnień) · [Raport uprawnień](#raport-uprawnień) · [Uprawnienia efektywne](#uprawnienia-efektywne) · [Gdzie ma dostęp](#gdzie-ma-dostęp) · [Odbieranie uprawnień](#odbieranie-uprawnień) · [Grupy dostępu do folderu](#grupy-dostępu-do-folderu)
- [Kontrola i naprawa](#kontrola-i-naprawa): [Naprawa uprawnień](#naprawa-uprawnień) · [Kopie uprawnień](#kopie-uprawnień) · [Raport ryzyk](#raport-ryzyk) · [Porównanie ze wzorcem](#porównanie-ze-wzorcem) · [Udziały i NTFS](#udziały-i-ntfs) · [Zmiany uprawnień](#zmiany-uprawnień)
- [Pliki](#pliki): [Sumy kontrolne](#sumy-kontrolne)

---

## Uprawnienia NTFS

### Nadawanie uprawnień

**Do czego.** Nadać uprawnienia do folderu wielu grupom albo kontom naraz, każdej z innym poziomem.

**Co możesz zrobić**
- dodać tożsamości (wpisane albo wybrane z AD), każdą z własnym **poziomem** (od odczytu do pełnej kontroli) i **zakresem** (ten folder, podfoldery, pliki);
- zastąpić dotychczasowe wpisy tych tożsamości;
- wyłączyć dziedziczenie z kopią odziedziczonych wpisów;
- podejrzeć bieżące uprawnienia folderu.

**Jak działa.** Lokalnie, przez UNC albo na serwerze plików (WinRM). Wpisy są nadawane po SID, z kopią uprawnień przed zmianą.

### Raport uprawnień

**Do czego.** Zobaczyć, kto ma jakie uprawnienia w drzewie folderów, gdzie przerwano dziedziczenie i gdzie nie ma dostępu.

**Co możesz zrobić**
- przeskanować foldery (opcjonalnie pliki) do zadanej głębokości;
- zobaczyć tożsamość, uprawnienia, typ, dziedziczenie, zakres, właściciela; foldery z wyłączonym dziedziczeniem i bez dostępu;
- filtrować: tylko jawne wpisy, bez kont systemowych (lista do edycji), bez nierozwiązanych SID;
- włączyć **tryb kopii zapasowej** i **pomijanie miejsc bez dostępu** (każde trafia do dziennika i raportu);
- utworzyć **raport HTML w postaci drzewa** — z odznakami (jawne wpisy, wyłączone dziedziczenie, brak dostępu, nierozwiązane SID, właściciel); poddrzewa, w których wszystko dziedziczy, są zwinięte do jednej linii;
- z wiersza: pokazać członków grupy.

**Jak działa.** Podfoldery są skanowane równolegle. Łącza (junction) nie są przechodzone.

### Uprawnienia efektywne

**Do czego.** Odpowiedzieć na pytanie „jaki dostęp do tego folderu ma naprawdę ta osoba” — tak, jak policzy go Windows.

**Co możesz zrobić**
- podać osobę albo grupę i ścieżkę folderu lub pliku;
- zobaczyć tabelę 13 praw szczegółowych z kolumnami *NTFS*, *Udział*, *Efektywnie* i *Skąd* — który wpis i przez którą grupę daje albo odbiera prawo.

**Jak działa.** Program liczy dostęp przez wszystkie grupy osoby (także zagnieżdżone, z grupą podstawową), konta ogólne (Wszyscy, Użytkownicy uwierzytelnieni, `BUILTIN\Users`), kolejność wpisów (odmowa przed zezwoleniem), uprawnienia ogólne (GENERIC), wpisy „tylko podfoldery” i prawa właściciela. Dla ścieżki `\\serwer\udział` liczy też **uprawnienia udziału** — wynik to mniejsze z dwóch.

### Gdzie ma dostęp

**Do czego.** Znaleźć wszystkie miejsca w drzewie folderów, do których dana osoba albo grupa ma dostęp.

**Co możesz zrobić**
- podać osobę albo grupę i folder początkowy;
- zobaczyć każde miejsce z wpisem: poziom, efektywny dostęp, przez którą grupę, typ, dziedziczenie, zakres;
- z wiersza: przejść do uprawnień efektywnych albo odebrać dostęp.

**Jak działa.** Skan drzewa z rozwinięciem grup osoby z AD (także zagnieżdżonych) i kont ogólnych.

### Odbieranie uprawnień

**Do czego.** Usunąć dostęp osoby albo grupy w całym drzewie folderów — np. po odejściu pracownika albo zmianie działu.

**Co możesz zrobić**
- podać tożsamości i folder;
- przejrzeć podgląd każdego jawnego wpisu (wpisy dziedziczone są pokazane, ale usuwa się je w folderze wyżej);
- wybrać miejsca w potwierdzeniu i usunąć wpisy.

**Jak działa.** Kopia uprawnień przed zmianą; folder zmieniony od podglądu jest pomijany.

### Grupy dostępu do folderu

**Do czego.** Wdrożyć model „grupa na folder”: dla folderu (albo każdego podfolderu) grupy odczytu i modyfikacji z nadanymi od razu uprawnieniami.

**Co możesz zrobić**
- wybrać folder albo jego podfoldery;
- ustalić szablon nazwy (np. `DL_{Folder}_{Poziom}`, pola `{Nadrzędny}`, `{Udział}`, `{Serwer}`) i opisu, OU i zakres grup;
- wybrać poziomy: odczyt, modyfikacja, opcjonalnie pełna kontrola;
- dodać członków (np. `RO: grupa`);
- poprawić nazwy i opisy w podglądzie.

**Jak działa.** Podgląd sprawdzony w AD (istniejąca grupa jest używana zamiast tworzenia nowej). Po utworzeniu grup program nadaje im uprawnienia NTFS (ten folder, podfoldery i pliki), z kopią uprawnień.

---

## Kontrola i naprawa

### Naprawa uprawnień

**Do czego.** Uporządkować drzewo folderów: przywrócić dziedziczenie, usunąć wpisy po usuniętych kontach, odzyskać dostęp do miejsc bez dostępu.

**Co możesz zrobić**
- **włączyć dziedziczenie** w podfolderach (jawne wpisy zostają) albo **przywrócić dziedziczenie** (bez jawnych wpisów);
- **usunąć wpisy kont usuniętych z AD** (nierozwiązane SID);
- **przejąć własność** miejsc bez dostępu (grupa Administratorzy, opcjonalnie z pełną kontrolą);
- wybrać miejsca z podglądu.

**Jak działa**
- Nierozwiązane SID-y są dodatkowo sprawdzane w AD — konta z domen zaufanych zostają.
- W trybie kopii zapasowej przejęcie własności idzie jednym przejściem drzewa, a wpis Administratorów trafia tylko tam, gdzie go brakowało.
- Bez tego trybu program używa `icacls` po SID (niezależnie od języka systemu); po pierwszym przebiegu mogą zostać głębsze miejsca bez dostępu — wtedy powtórz podgląd i naprawę.
- Zawsze podgląd, wybór miejsc i kopia uprawnień przed zmianą.

### Kopie uprawnień

**Do czego.** Przeglądać kopie uprawnień i przywracać stan sprzed zmiany.

**Co możesz zrobić**
- zobaczyć kopie zapisane automatycznie przed zmianami i ręczne;
- **zrobić kopię uprawnień** całego drzewa (przycisk *Zrób kopię uprawnień*);
- podejrzeć zawartość kopii;
- **przywrócić** całość albo wybrane foldery (wpisy i dziedziczenie, właściciel);
- porównać kopię z obecnym stanem (moduł *Zmiany uprawnień*).

**Jak działa.** Kopia zapisuje właściciela, wpisy i dziedziczenie każdego folderu (SDDL). Stan sprzed przywrócenia trafia do nowej kopii. Właściciel jest przywracany w trybie kopii zapasowej albo gdy system pozwala go ustawić — inaczej wynik ma ostrzeżenie.

### Raport ryzyk

**Do czego.** Znaleźć ryzykowne uprawnienia w drzewie folderów.

**Co możesz zrobić**
- przeskanować drzewo i zobaczyć ryzyka z oceną (*wysokie / średnie / niskie*) i zaleceniem: zapis dla Wszystkich, Użytkowników uwierzytelnionych i domeny; odczyt dla Wszystkich i gości; pełna kontrola lub zmiana uprawnień dla zwykłych kont; zbyt wiele kont z pełną kontrolą (próg do ustawienia); wpisy dla użytkowników zamiast grup; wpisy usuniętych kont; odmowy; użytkownik jako właściciel;
- z wiersza: sprawdzić, gdzie jeszcze dana tożsamość ma dostęp, albo odebrać uprawnienia.

**Jak działa.** Skan jak w raporcie uprawnień, z regułami oceny. Ten sam raport jest dostępny jako [raport cykliczny](06-domena.md#raporty-cykliczne).

### Porównanie ze wzorcem

**Do czego.** Ujednolicić uprawnienia podobnych folderów (np. folderów działów) według jednego wzorca.

**Co możesz zrobić**
- wskazać folder wzorcowy i foldery do porównania (*Wstaw podfoldery…*);
- zobaczyć brakujące, nadmiarowe i inne wpisy, różnice w dziedziczeniu i właścicielu;
- **Zastosuj wzorzec…** — dokładna kopia jawnych wpisów z ustawieniem dziedziczenia albo tylko brakujące wpisy.

**Jak działa.** Porównanie wpisów po SID; zmiana z kopią uprawnień przed zmianą.

### Udziały i NTFS

**Do czego.** Zobaczyć, jaki dostęp przez sieć wynika z połączenia uprawnień udziału i NTFS.

**Co możesz zrobić**
- odczytać udziały serwera plików (opcjonalnie administracyjne) z uprawnieniami udziału i NTFS folderu udziału obok siebie;
- dla każdego konta zobaczyć poziom z udziału, z NTFS i **wynikowy dostęp przez sieć**, z informacją, które uprawnienia go ograniczają;
- zobaczyć ostrzeżenia o zapisie dla szerokich grup.

**Jak działa.** Odczyt na serwerze przez PowerShell Remoting (uprawnienia udziałów także przez DCOM).

### Zmiany uprawnień

**Do czego.** Sprawdzić, co się zmieniło w uprawnieniach od czasu kopii — gdzie i komu nadano dostęp.

**Co możesz zrobić**
- porównać kopię uprawnień z **obecnym stanem** albo z **inną kopią**;
- zobaczyć nowe, usunięte i zmienione wpisy, zmiany dziedziczenia i właściciela, foldery, których już nie ma, a przy pełnej kopii drzewa także **nowe foldery**;
- z wiersza: przywrócić wybrane foldery ze stanu z kopii.

**Jak działa.** Porównanie po SID. Nowy zapis dla szerokich grup (Wszyscy, Użytkownicy domeny…) jest wyróżniony. Regularna ręczna kopia drzewa (moduł *Kopie uprawnień*) pozwala śledzić zmiany w czasie.

---

## Pliki

### Sumy kontrolne

**Do czego.** Policzyć skróty plików i sprawdzić, czy plik jest taki, jak powinien (np. pobrany obraz ISO, kopia na serwerze).

**Co możesz zrobić**
- dodać pliki i całe foldery (także przeciągnięte z Eksploratora);
- wybrać algorytm: MD5, SHA1, SHA256, SHA384, SHA512;
- porównać z oczekiwanym skrótem;
- skopiować wyniki w formacie `skrót  plik`.

**Jak działa.** Liczenie w tle, partiami — okno nie zamarza przy dużych plikach. Historia zostaje w tabeli.
