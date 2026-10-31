# 4. Komputery AD

[← Spis treści](README.md)

**Do czego jest ta przestrzeń.** Konta komputerów w Active Directory: informacje i kanał zaufania, hasła LAPS, klucze odzyskiwania BitLocker, zmiana nazw, członkostwo w grupach i raporty.

**Lista po lewej:** komputery z AD (drzewo OU, filtr nazwy, tylko włączone) albo wpisane ręcznie.

**Jak to działa ogólnie.** Moduł `ActiveDirectory` na Twoim komputerze i uprawnienia Twojego konta (albo alternatywnego). Część operacji (kanał zaufania, zmiana nazwy) łączy się też z samym komputerem przez WinRM. Hasła LAPS i klucze BitLocker są **wartościami poufnymi** — maskowane w tabeli, eksporcie i kopiowaniu, dopóki nie zaznaczysz „Pokaż poufne”.

**Spis modułów**

- [Konta komputerów](#konta-komputerów): [Konto komputera](#konto-komputera) · [Zmiana nazwy komputerów](#zmiana-nazwy-komputerów) · [Członkostwo w grupach](#członkostwo-w-grupach)
- [Hasła i klucze](#hasła-i-klucze): [LAPS](#laps) · [Klucze BitLocker (AD)](#klucze-bitlocker-ad)
- [Raporty](#raporty): [Raporty komputerów](#raporty-komputerów)

---

## Konta komputerów

### Konto komputera

**Do czego.** Obsługa obiektu komputera w AD, w tym naprawa „relacji zaufania” bez wyjmowania komputera z domeny.

**Co możesz zrobić**
- odczytać informacje o koncie komputera;
- **przetestować i naprawić kanał zaufania** z domeną;
- włączyć, wyłączyć, zmienić opis, przenieść do OU, zresetować i usunąć konto (także z menu wiersza);
- utworzyć nowe konta komputerów przed dołączeniem do domeny (pre-staging);
- utworzyć **raport HTML** — kartę każdego komputera: system, konto, ostatnie logowanie, LAPS i klucze BitLocker w AD (**tylko informacja o obecności i dacie, bez haseł i kluczy**), delegowanie Kerberos, liczba SPN, grupy.

**Jak działa.** Zmiany konta przez moduł ActiveDirectory. Test i naprawa kanału zaufania (`Test-ComputerSecureChannel`) działają na samym komputerze przez WinRM; naprawa wymaga poświadczeń domenowych — program poprosi o nie, jeśli nie ustawiono konta alternatywnego.

### Zmiana nazwy komputerów

**Do czego.** Zmienić nazwy wielu komputerów naraz — np. według nowego standardu nazewnictwa.

**Co możesz zrobić**
- ustalić nowe nazwy autonumeracją albo mapowaniem z listy (np. z Excela);
- poprawić nazwy w tabeli (kolumna *Nowa nazwa*);
- zmienić nazwy z opcjonalnym restartem.

**Jak działa.** Walidacja przed zmianą: zasady NetBIOS (najwyżej 15 znaków) i duplikaty. Zmiana nazwy działa na komputerze i wymaga poświadczeń domenowych.

### Członkostwo w grupach

**Do czego.** Sprawdzić i zmienić grupy kont komputerów (np. grupy, którym przypisano GPO albo dostęp do sieci).

**Co możesz zrobić:** odczytać grupy bezpośrednie i zagnieżdżone; dodać komputery do grup i usunąć z zaznaczonych grup.

**Jak działa.** Ten sam mechanizm co [Członkostwo w grupach](02-uzytkownicy-ad.md#członkostwo-w-grupach) dla kont użytkowników.

---

## Hasła i klucze

### LAPS

**Do czego.** Odczytać hasło lokalnego administratora komputera zarządzane przez LAPS i wymusić jego zmianę.

**Co możesz zrobić**
- odczytać hasła Windows LAPS (także szyfrowane) i LAPS legacy;
- bezpiecznie skopiować hasło (dwuklik) — znika ze schowka po 60 s;
- wymusić zmianę hasła z przetworzeniem zasad na komputerze (także z menu wiersza).

**Jak działa.** Odczyt atrybutów LAPS z AD (wymaga uprawnień do odczytu hasła; dla szyfrowanych haseł Windows LAPS — prawa odszyfrowania). Hasło jest maskowane w tabeli i eksporcie.

### Klucze BitLocker (AD)

**Do czego.** Znaleźć klucz odzyskiwania BitLocker — także wtedy, gdy znasz tylko identyfikator klucza z ekranu odzyskiwania.

**Co możesz zrobić**
- odczytać klucze odzyskiwania zapisane w AD dla zaznaczonych komputerów;
- **wyszukać komputer po identyfikatorze klucza** (fragment z niebieskiego ekranu odzyskiwania) w całej domenie.

**Jak działa.** Odczyt obiektów kluczy odzyskiwania z AD (wymaga odpowiednich uprawnień). Klucze są wartościami poufnymi. Kopię kluczy do AD robi się w przestrzeni Zarządzanie zdalne → [BitLocker](01-zarzadzanie-zdalne.md#bitlocker).

---

## Raporty

### Raporty komputerów

**Do czego.** Zestawienia komputerów do sprzątania i przeglądu bezpieczeństwa.

**Co możesz zrobić**
- wybrać raport: nieaktywne, wyłączone, nowe, podsumowanie systemów, nieobsługiwane systemy, bez LAPS, bez klucza BitLocker w AD, serwery;
- zaznaczyć wyniki na liście komputerów, wyłączyć, przenieść albo usunąć konta.

**Jak działa.** Zapytania do AD na całej domenie albo w OU. Te same raporty są dostępne jako [raporty cykliczne](06-domena.md#raporty-cykliczne).
