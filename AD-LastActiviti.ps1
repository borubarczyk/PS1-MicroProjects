# Sprawdzenie i import modułu Active Directory
try {
    Import-Module ActiveDirectory -ErrorAction Stop
} catch {
    Write-Warning "Nie znaleziono modułu ActiveDirectory. Zainstaluj narzędzia RSAT."
    Exit
}

# Ładowanie klas WPF
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase

# Definicja interfejsu w XAML
[xml]$XAML = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Zarządzanie kontami AD - Raport Logowania" Height="700" Width="1150" 
        WindowStartupLocation="CenterScreen" Background="#F4F4F9" FontFamily="Segoe UI">
    <Grid Margin="15">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <!-- Panel górny: Filtry i odświeżanie -->
        <Border Grid.Row="0" Background="White" Padding="10" Margin="0,0,0,15" CornerRadius="5" BorderBrush="#DDDDDD" BorderThickness="1">
            <WrapPanel VerticalAlignment="Center">
                
                <!-- Wyszukiwarka tekstowa -->
                <TextBlock Text="Szukaj:" VerticalAlignment="Center" FontWeight="SemiBold" Margin="0,0,5,0"/>
                <TextBox Name="TxtSearch" Width="150" Padding="5" Margin="0,0,20,0" VerticalContentAlignment="Center" ToolTip="Wpisz imię, nazwisko, login lub miasto"/>

                <!-- Filtr daty -->
                <TextBlock Text="Nieużywane od:" VerticalAlignment="Center" FontWeight="SemiBold" Margin="0,0,5,0"/>
                <ComboBox Name="ComboFilter" Width="140" Padding="5" Margin="0,0,20,0">
                    <ComboBoxItem Content="Wszyscy (Brak)" IsSelected="True"/>
                    <ComboBoxItem Content="Powyżej 1 miesiąca"/>
                    <ComboBoxItem Content="Powyżej 3 miesięcy"/>
                    <ComboBoxItem Content="Powyżej 6 miesięcy"/>
                    <ComboBoxItem Content="Powyżej 12 miesięcy"/>
                </ComboBox>
                
                <!-- Filtr aktywności -->
                <CheckBox Name="ChkOnlyActive" Content="Tylko aktywne" IsChecked="True" VerticalAlignment="Center" Margin="0,0,20,0" FontWeight="SemiBold" Foreground="#333333"/>
                
                <!-- Przyciski i status -->
                <Button Name="BtnRefresh" Content="🔄 Pobierz z AD" Padding="15,5" Background="#0078D7" Foreground="White" FontWeight="Bold" BorderThickness="0" Cursor="Hand"/>
                <TextBlock Name="TxtStatus" Text="Czekam na pobranie danych..." VerticalAlignment="Center" Margin="20,0,0,0" Foreground="#666666" FontStyle="Italic"/>
            
            </WrapPanel>
        </Border>

        <!-- Tabela z danymi -->
        <DataGrid Name="GridUsers" Grid.Row="1" AutoGenerateColumns="False" IsReadOnly="True" 
                  SelectionMode="Single" AlternatingRowBackground="#EBF1FA" Background="White" 
                  BorderBrush="#DDDDDD" BorderThickness="1" GridLinesVisibility="Horizontal">
            <DataGrid.Columns>
                <DataGridTextColumn Header="Imię i Nazwisko" Binding="{Binding Name}" Width="220"/>
                <DataGridTextColumn Header="Login (sAMAccountName)" Binding="{Binding SamAccountName}" Width="180"/>
                <DataGridTextColumn Header="Miasto" Binding="{Binding City}" Width="140"/>
                <DataGridTextColumn Header="Ostatnie Logowanie" Binding="{Binding LastLogon}" Width="140"/>
                <DataGridTextColumn Header="Aktywne" Binding="{Binding Enabled}" Width="70"/>
                <DataGridTextColumn Header="Dni Nieaktywności" Binding="{Binding InactiveDays}" Width="120"/>
            </DataGrid.Columns>
        </DataGrid>

        <!-- Panel dolny: Akcje -->
        <Border Grid.Row="2" Background="White" Padding="10" Margin="0,15,0,0" CornerRadius="5" BorderBrush="#DDDDDD" BorderThickness="1">
            <Grid>
                <StackPanel Orientation="Horizontal" HorizontalAlignment="Left">
                    <Button Name="BtnExport" Content="💾 Eksportuj widok do CSV" Padding="15,8" Background="#107C41" Foreground="White" FontWeight="Bold" BorderThickness="0" Cursor="Hand"/>
                </StackPanel>
                <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
                    <Button Name="BtnDisable" Content="🔒 Wyłącz (Zablokuj) zaznaczone konto" Padding="15,8" Background="#D13438" Foreground="White" FontWeight="Bold" BorderThickness="0" Cursor="Hand"/>
                </StackPanel>
            </Grid>
        </Border>
    </Grid>
</Window>
"@

# Odczyt XAML
$Reader = New-Object System.Xml.XmlNodeReader $XAML
$Window = [Windows.Markup.XamlReader]::Load($Reader)

# Mapowanie kontrolek do zmiennych
$TxtSearch     = $Window.FindName("TxtSearch")
$ComboFilter   = $Window.FindName("ComboFilter")
$ChkOnlyActive = $Window.FindName("ChkOnlyActive")
$BtnRefresh    = $Window.FindName("BtnRefresh")
$BtnExport     = $Window.FindName("BtnExport")
$BtnDisable    = $Window.FindName("BtnDisable")
$GridUsers     = $Window.FindName("GridUsers")
$TxtStatus     = $Window.FindName("TxtStatus")

# Zmienna globalna przechowująca pobranych użytkowników z AD
$Script:AllUsers = @()

# Główna funkcja aplikująca filtry (odpala się przy każdej zmianie na UI)
function Apply-Filters {
    if ($Script:AllUsers.Count -eq 0) { return }

    # Formatowanie podstawowe (dodano Miasto)
    $DisplayList = foreach ($User in $Script:AllUsers) {
        $Days = if ($User.LastLogonDate) { ((Get-Date) - $User.LastLogonDate).Days } else { "Nigdy" }
        
        [PSCustomObject]@{
            Name           = $User.Name
            SamAccountName = $User.SamAccountName
            City           = if ($User.City) { $User.City } else { "-" }
            LastLogon      = if ($User.LastLogonDate) { $User.LastLogonDate.ToString("yyyy-MM-dd HH:mm") } else { "Brak danych" }
            Enabled        = if ($User.Enabled) { "Tak" } else { "Nie" }
            InactiveDays   = $Days
            LastLogonDate  = $User.LastLogonDate
        }
    }

    # Filtr 1: Aktywne/Nieaktywne
    if ($ChkOnlyActive.IsChecked) {
        $DisplayList = $DisplayList | Where-Object { $_.Enabled -eq "Tak" }
    }

    # Filtr 2: Czas logowania
    $ThresholdDate = $null
    switch ($ComboFilter.SelectedItem.Content) {
        "Powyżej 1 miesiąca"  { $ThresholdDate = (Get-Date).AddMonths(-1) }
        "Powyżej 3 miesięcy"  { $ThresholdDate = (Get-Date).AddMonths(-3) }
        "Powyżej 6 miesięcy"  { $ThresholdDate = (Get-Date).AddMonths(-6) }
        "Powyżej 12 miesięcy" { $ThresholdDate = (Get-Date).AddMonths(-12) }
    }

    if ($ThresholdDate) {
        $DisplayList = $DisplayList | Where-Object { 
            ($_.LastLogonDate -lt $ThresholdDate) -or ($_.LastLogonDate -eq $null) 
        }
    }

    # Filtr 3: Wyszukiwarka tekstowa (Imię, Nazwisko, Login LUB MIASTO)
    $SearchPhrase = $TxtSearch.Text.Trim()
    if (![string]::IsNullOrEmpty($SearchPhrase)) {
        $EscapedPhrase = [regex]::Escape($SearchPhrase)
        $DisplayList = $DisplayList | Where-Object {
            ($_.Name -match $EscapedPhrase) -or 
            ($_.SamAccountName -match $EscapedPhrase) -or 
            ($_.City -match $EscapedPhrase)
        }
    }

    # Zabezpieczenie przed błędem Overload - rzutowanie na tablicę [object[]]
    $SafeArray = @($DisplayList | Sort-Object InactiveDays -Descending)
    $GridUsers.ItemsSource = [System.Collections.ObjectModel.ObservableCollection[System.Object]]::new([object[]]$SafeArray)
    
    $Count = $SafeArray.Count
    $TxtStatus.Text = "Wyświetlono: $Count użytkowników."
}

# Pobieranie danych z AD (Zaktualizowano o właściwość City)
$BtnRefresh.Add_Click({
    $TxtStatus.Text = "Pobieranie danych z AD... Proszę czekać."
    
    # Odpytanie AD
    $Script:AllUsers = Get-ADUser -Filter * -Properties LastLogonDate, Enabled, City | 
                       Select-Object Name, SamAccountName, Enabled, LastLogonDate, City
    
    # Zaaplikowanie filtrów po pobraniu
    Apply-Filters
})

# Zmiany filtrów w UI - natychmiastowe odświeżenie bez ponownego pytania AD
$ComboFilter.Add_SelectionChanged({ Apply-Filters })
$ChkOnlyActive.Add_Checked({ Apply-Filters })
$ChkOnlyActive.Add_Unchecked({ Apply-Filters })
$TxtSearch.Add_TextChanged({ Apply-Filters })

# Akcja: Zablokuj
$BtnDisable.Add_Click({
    $SelectedItem = $GridUsers.SelectedItem
    if ($null -eq $SelectedItem) {
        [System.Windows.MessageBox]::Show("Wybierz użytkownika z listy, którego chcesz zablokować.", "Informacja", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
        return
    }

    if ($SelectedItem.Enabled -eq "Nie") {
        [System.Windows.MessageBox]::Show("To konto jest już wyłączone.", "Informacja", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
        return
    }

    $Result = [System.Windows.MessageBox]::Show("Czy na pewno chcesz WYŁĄCZYĆ konto: $($SelectedItem.Name) ($($SelectedItem.SamAccountName))?", "Potwierdzenie blokady", [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Warning)
    
    if ($Result -eq [System.Windows.MessageBoxResult]::Yes) {
        try {
            Disable-ADAccount -Identity $SelectedItem.SamAccountName -ErrorAction Stop
            [System.Windows.MessageBox]::Show("Konto zostało pomyślnie wyłączone.", "Sukces", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
            
            # Aktualizacja lokalnej zmiennej, żeby nie pytać AD na nowo
            $UserToUpdate = $Script:AllUsers | Where-Object { $_.SamAccountName -eq $SelectedItem.SamAccountName }
            if ($UserToUpdate) { $UserToUpdate.Enabled = $false }
            
            Apply-Filters # Odśwież tabelę
        } catch {
            [System.Windows.MessageBox]::Show("Wystąpił błąd podczas blokowania konta: $($_.Exception.Message)", "Błąd", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
        }
    }
})

# Akcja: Eksport do CSV (Uwzględniono kolumnę City)
$BtnExport.Add_Click({
    if ($null -eq $GridUsers.ItemsSource -or $GridUsers.ItemsSource.Count -eq 0) {
        [System.Windows.MessageBox]::Show("Brak danych do wyeksportowania. Najpierw pobierz dane.", "Informacja", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
        return
    }

    # Wywołanie systemowego okna zapisu
    $SaveDialog = New-Object Microsoft.Win32.SaveFileDialog
    $SaveDialog.Filter = "Plik CSV (*.csv)|*.csv"
    $SaveDialog.Title = "Zapisz raport jako..."
    $SaveDialog.FileName = "RaportNieuzywanychKont_$(Get-Date -Format 'yyyyMMdd_HHmm').csv"
    
    if ($SaveDialog.ShowDialog() -eq $true) {
        try {
            # Bierzemy tylko potrzebne kolumny z widocznej tabeli (w tym City)
            $GridUsers.ItemsSource | 
                Select-Object Name, SamAccountName, City, LastLogon, Enabled, InactiveDays | 
                Export-Csv -Path $SaveDialog.FileName -NoTypeInformation -Encoding UTF8 -Delimiter ";"
            
            [System.Windows.MessageBox]::Show("Eksport zakończony pomyślnie!", "Sukces", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
        } catch {
            [System.Windows.MessageBox]::Show("Błąd zapisu do pliku: $($_.Exception.Message)", "Błąd", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
        }
    }
})

# Wyświetlenie okna
$Window.ShowDialog() | Out-Null