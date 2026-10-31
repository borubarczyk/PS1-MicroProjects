# 3. Grupy i OU

[← Spis treści](README.md)

**Do czego jest ta przestrzeń.** Grupy i jednostki organizacyjne w Active Directory: szczegóły i członkowie grup, zmiany hurtowe, tworzenie i duplikowanie grup, budowanie i klonowanie struktur OU, raporty.

**Lista po lewej:** grupy — wyszukiwanie w AD (nazwa, sAMAccountName, opis, e-mail), OU, rodzaj (zabezpieczeń, dystrybucyjne, bez członków, uprzywilejowane), nazwy wpisane ręcznie. Przy grupie widać zakres i typ.

**Jak to działa ogólnie.** Moduł `ActiveDirectory` na Twoim komputerze, uprawnienia Twojego konta (albo alternatywnego). Moduły tworzące wiele obiektów najpierw budują **podgląd** sprawdzony w AD (kolumna „Stan”), który można poprawić w tabeli; dopiero potem wykonują zmiany i raportują wynik dla każdego wiersza.

**Spis modułów**

- [Grupy](#grupy): [Szczegóły grup](#szczegóły-grup) · [Członkowie grup](#członkowie-grup) · [Przynależność do grup](#przynależność-do-grup) · [Członkostwo hurtowe](#członkostwo-hurtowe)
- [Tworzenie](#tworzenie): [Tworzenie grup](#tworzenie-grup) · [Duplikowanie grup](#duplikowanie-grup) · [Lokalizacje i role](#lokalizacje-i-role)
- [Struktura OU](#struktura-ou): [Drzewo OU](#drzewo-ou) · [Klonowanie OU](#klonowanie-ou)
- [Raporty](#raporty): [Drzewo zagnieżdżeń](#drzewo-zagnieżdżeń) · [Raporty grup](#raporty-grup)

---

## Grupy

### Szczegóły grup

**Do czego.** Zobaczyć i zmienić właściwości grup.

**Co możesz zrobić**
- odczytać zakres, typ, opis, zarządcę, e-mail i liczbę członków zaznaczonych grup;
- zmienić opis, zarządcę, adres e-mail, zakres i typ, nazwę;
- przenieść grupę do innej OU, włączyć ochronę przed usunięciem, usunąć grupę;
- utworzyć **raport HTML** — kartę każdej grupy: właściwości, zarządca, liczniki, lista **członków** z typem i stanem kont (opcjonalnie także zagnieżdżonych) i grupy nadrzędne.

**Jak działa.** Odczyt i zmiany przez moduł ActiveDirectory, z potwierdzeniem.

### Członkowie grup

**Do czego.** Sprawdzić i zmienić, kto jest w grupie.

**Co możesz zrobić**
- odczytać członków bezpośrednich i zagnieżdżonych, z typem i stanem kont;
- dodać członków, skopiować członków z innej grupy, usunąć zaznaczonych;
- przenieść konta z wyników na listę użytkowników (do dalszych operacji).

**Jak działa.** Odczyt przez moduł ActiveDirectory **bez limitu 5000 obiektów** (którego `Get-ADGroupMember` nie przeskoczy w dużych grupach).

### Przynależność do grup

**Do czego.** Sprawdzić zagnieżdżanie: do jakich grup należą zaznaczone grupy.

**Co możesz zrobić:** odczytać grupy nadrzędne — bezpośrednio i przez inne grupy; dodać zaznaczone grupy do innych grup i usunąć je z nich.

**Jak działa.** Ten sam mechanizm co *Członkostwo w grupach* w przestrzeni Użytkownicy AD, dla obiektów grup.

### Członkostwo hurtowe

**Do czego.** Dodać albo usunąć wiele obiektów w wielu grupach naraz — np. według tabeli z Excela.

**Co możesz zrobić**
- podać obiekty (konta, grupy, komputery) i grupy;
- wybrać tryb: **każdy do każdej** (A×B) albo **wiersz do wiersza**;
- dodać albo usunąć członkostwa.

**Jak działa.** Obiekty są rozpoznawane po loginie, UPN, e-mailu, DN, SID i nazwie. Najpierw powstaje **plan** — co już jest, co zostanie zmienione, czego nie znaleziono — potem wykonanie z wynikiem dla każdego wiersza.

---

## Tworzenie

### Tworzenie grup

**Do czego.** Utworzyć wiele grup naraz z tabeli.

**Co możesz zrobić**
- wkleić tabelę z Excela: nazwa, opis, zakres, typ, e-mail, sAMAccountName, grupy nadrzędne, zarządca, OU;
- przejrzeć podgląd z edycją komórek;
- utworzyć grupy.

**Jak działa.** sAMAccountName jest normalizowany, podgląd sprawdza w AD konflikty nazw, istnienie OU, grup nadrzędnych i zarządców.

### Duplikowanie grup

**Do czego.** Zrobić kopie istniejących grup z nową nazwą — np. zestaw grup oddziału WAW dla nowego oddziału KRK.

**Co możesz zrobić**
- ustalić nową nazwę: zamiana tekstu, przedrostek, przyrostek — także wiele kopii naraz (np. `WAW → KRK; GDA`);
- opcjonalnie skopiować członków, przynależność do grup, zarządcę i ochronę;
- umieścić kopie w tej samej albo innej OU.

**Jak działa.** Podgląd nazw i konfliktów w AD, potem tworzenie z wynikiem dla każdej kopii.

### Lokalizacje i role

**Do czego.** Zbudować powtarzalną strukturę: jednostki dla lokalizacji (np. miast) i w każdej te same grupy ról.

**Co możesz zrobić**
- podać lokalizacje i role z kodami;
- ustalić szablony nazwy i opisu grup z polami `{PREFIKS}`, `{LOKALIZACJA}`, `{KOD_LOK}`, `{ROLA}`, `{KOD_ROLI}`;
- dodać grupy zbiorcze **ALL**: dla lokalizacji, dla roli i globalną;
- wymusić unikalne sAMAccountName (opcjonalnie najwyżej 20 znaków).

**Jak działa.** Podgląd całej struktury sprawdzony w AD, potem utworzenie jednostek, grup i zagnieżdżeń.

---

## Struktura OU

### Drzewo OU

**Do czego.** Utworzyć całe drzewo jednostek organizacyjnych z tekstu — albo przenieść istniejącą strukturę w inne miejsce.

**Co możesz zrobić**
- opisać drzewo wcięciami albo ścieżkami (np. `Firma/Warszawa/Komputery`), z opisami po znaku `|`;
- włączyć ochronę przed usunięciem;
- wczytać istniejącą strukturę jako tekst, zmienić ją i utworzyć w innym miejscu.

**Jak działa.** Podgląd pokazuje, które jednostki już są, a które powstaną; tworzenie idzie od góry drzewa.

### Klonowanie OU

**Do czego.** Skopiować całą jednostkę w nowe miejsce — np. struktura istniejącego oddziału dla nowego.

**Co możesz zrobić**
- skopiować podjednostki i grupy (z zamianą tekstu w nazwach, np. `WAW=KRK`) oraz zagnieżdżenia między nimi;
- opcjonalnie skopiować członków i grupy nadrzędne, linki GPO, opisy i ochronę.

**Jak działa.** Podgląd sprawdzony w AD, potem tworzenie. **Konta użytkowników i komputerów nie są kopiowane.**

---

## Raporty

### Drzewo zagnieżdżeń

**Do czego.** Zrozumieć, jak grupy są w siebie zagnieżdżone, i znaleźć zapętlenia.

**Co możesz zrobić**
- pokazać drzewo grup zawartych w zaznaczonych grupach (w dół) albo grup nadrzędnych (w górę), opcjonalnie z członkami;
- bez zaznaczenia — wszystkie grupy;
- utworzyć raport HTML z rozwijanymi gałęziami i wyszukiwarką.

**Jak działa.** Odczyt z AD z wykrywaniem zapętleń (grupa pośrednio zawiera samą siebie).

### Raporty grup

**Do czego.** Znaleźć grupy do sprzątania albo do przeglądu.

**Co możesz zrobić**
- wybrać raport: puste, z jednym członkiem, duże, bez opisu lub zarządcy, z wyłączonymi kontami, uprzywilejowane, dystrybucyjne, nowe i zmienione, zapętlone zagnieżdżenia;
- zaznaczyć wyniki na liście grup i wykonać na nich operacje w innych modułach.

**Jak działa.** Zapytania do AD na całej domenie albo w wybranej OU.
