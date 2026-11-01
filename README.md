# PS1-MicroProjects

Zbiór mikroprojektów PowerShell wspomagających codzienną administrację Active Directory oraz systemami Windows. W katalogu znajdziesz niewielkie narzędzia GUI (WinForms/WPF): od sprawdzania i obsługi kont, przez audyty NTFS, po zdarzenia logowania RDP i porównywanie folderów.

## Wymagania ogólne

- Windows PowerShell 5.1 lub PowerShell 7 uruchamiany na Windowsie.
- RSAT ActiveDirectory (moduł `ActiveDirectory`) dla wszystkich skryptów zaczynających się od `AD-`.
- Uprawnienia administracyjne odpowiadające operacji (np. tworzenie OU, modyfikacja NTFS, odczyt dziennika Security, zdalne polecenia).
- Sesja desktopowa w trybie STA dla narzędzi WinForms/WPF: `powershell.exe -STA -File .\NazwaSkryptu.ps1`.
- Dla zdalnych operacji (`AD-RDP-LoginEvents.ps1`) włączony WinRM/WMI na hostach docelowych.

## Jak uruchomić skrypt

1. Sklonuj repozytorium lub skopiuj katalog z repo:
   ```powershell
   git clone https://github.com/borubarczyk/PS1-MicroProjects.git
   cd PS1-MicroProjects
   ```
2. (Opcjonalnie) zezwól na uruchamianie w bieżącej sesji: `Set-ExecutionPolicy -Scope Process Bypass -Force`.
3. Uruchom interesujący plik, np.: `powershell.exe -STA -ExecutionPolicy Bypass -File .\AD-User-Manager.ps1`.
4. Narzędzia GUI uruchamiaj w sesji desktopowej z odpowiednimi uprawnieniami.

## Kodowanie plików

Skrypty zawierają polskie znaki i są zapisane jako **UTF-8 z BOM**. Windows PowerShell 5.1 bez BOM czyta pliki w kodowaniu systemowym (np. cp1250), przez co polskie znaki psują napisy, a część skryptów w ogóle się nie parsuje. Przy edycji zachowuj kodowanie „UTF-8 with BOM” (w VS Code: *Save with Encoding → UTF-8 with BOM*).

## Pliki pomocnicze i dane zapisywane przez skrypty

| Plik / lokalizacja | Skrypt | Zawartość |
| --- | --- | --- |
| `Pulpit\NTFS_Audit_*.csv/.xlsx` | `AD-NTFS-AuditGUI.ps1` | Eksport wyników audytu NTFS. |
| `%TEMP%\PC-CheckHashFoldersIntegrity.log` | `PC-CheckHashFoldersIntegrity.ps1` | Log diagnostyczny porównania. |
| `%APPDATA%\CyberGenPro\CyberGenConfig.json`, `%TEMP%\CyberGenHistory.json` | `PasswordGenerator.ps1` | Ustawienia generatora oraz historia haseł z ostatnich 24 h (zapisana jawnym tekstem). |
| `LICENSE` | — | Warunki licencji: wszelkie prawa zastrzeżone. |

## Lista skryptów

| Skrypt | Opis |
| --- | --- |
| `AD-BulkAccountStatusChecker.ps1` | WinForms do wklejania listy loginów: sprawdza, czy konta istnieją i są włączone, filtruje/sortuje wyniki i eksportuje je do CSV. Duplikaty loginów (bez względu na wielkość liter) są pomijane. |
| `AD-LastActiviti.ps1` | Interfejs WPF z raportem ostatniego logowania użytkowników (filtr nieaktywności 1/3/6/12 miesięcy, wyszukiwanie po nazwie/loginie/mieście, tylko aktywne), wyłączanie zaznaczonego konta i eksport widoku do CSV. |
| `AD-NTFS-AuditGUI.ps1` | Audytor NTFS w WPF: rekursywnie czyta ACL, rozpoznaje typ podmiotu (użytkownik/grupa), wskazuje foldery z uprawnieniami nadanymi bezpośrednio użytkownikom; widok tabeli i drzewa, filtry, eksport CSV/XLSX. |
| `AD-PremissionAudit.ps1` | Prosty panel GUI: po wybraniu folderu uruchamia zadanie w tle, które przeszukuje drzewo katalogów i wypisuje miejsca, gdzie uprawnienia ma pojedynczy użytkownik zamiast grupy. |
| `AD-RDP-LoginEvents.ps1` | GUI do pobierania zdarzeń z dziennika Security (domyślnie 4624/4634) z ostatnich N dni z wybranych komputerów lub wszystkich kontrolerów domeny. Działa w tle, zapisuje raport CSV do wskazanego folderu; konfigurację można zapisać/wczytać z JSON. |
| `AD-SwapLogin.ps1` | GUI do naprawiania kont z zamienionym imieniem i nazwiskiem: sprawdza listę loginów, podpowiada konta do zamiany, hurtowo zamienia GivenName/Surname (opcjonalnie DisplayName); ma tryb testowy i eksport CSV. |
| `AD-User-Manager.ps1` | Lekki manager użytkowników: wyszukuje po SamAccountName, wyświetla najważniejsze atrybuty (hasło, logowania, blokady), pozwala wyłączyć/włączyć konto, odblokować po lockoucie i zresetować hasło. |
| `PasswordGenerator.ps1` | Generator haseł (WPF, kryptograficzne losowanie) z konfigurowalnymi regułami, generowaniem paczek do schowka, historią z ostatnich 24 h, motywem jasnym/ciemnym i wysyłką przez `mailto:`. |
| `PC-CheckHashFoldersIntegrity.ps1` | WinForms do porównywania dwóch folderów (np. lokalnego z udziałem sieciowym): indeksuje pliki (także ukryte), porównuje rozmiar/datę i opcjonalnie SHA-256, pokazuje różnice i eksportuje je do CSV. |
| `PC-Folder-TreeView.ps1` | Podgląd drzewa folderu (leniwe wczytywanie, opcjonalnie pliki ukryte) z eksportem do TXT w formie ASCII. Nie wchodzi w junctiony/symlinki. |

## Uwagi bezpieczeństwa

- `AD-SwapLogin` ma tryb testowy — użyj go przed hurtową zamianą imion i nazwisk.
- `PasswordGenerator` przechowuje wygenerowane hasła jawnym tekstem (historia w `%TEMP%`) — czyść historię i schowek po pracy.

## Licencja

Copyright © 2025–2026 Borys Kaleta. Wszelkie prawa zastrzeżone. Skrypty z tego repozytorium wolno używać, kopiować, modyfikować i rozpowszechniać wyłącznie na podstawie pisemnej zgody albo umowy licencyjnej z autorem i w zakresie w niej określonym — szczegóły w pliku [`LICENSE`](LICENSE).

## Kontakt

Uwagi, pomysły na nowe scenariusze i pytania o licencję kieruj do autora.

**Autor:** Boris
