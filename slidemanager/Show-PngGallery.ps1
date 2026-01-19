
<# 
.SYNOPSIS
Zeigt alle PNG-Dateien eines Verzeichnisses in einem WPF-Fenster an (scrollbar, 16:9, Checkbox je Bild).
Beim Klick auf "Export" werden die Dateinamen der ausgewählten PNGs in der Konsole ausgegeben.

.PARAMETER Path
Basisverzeichnis, in dem .png gesucht wird (nur oberster Ordner).

.EXAMPLE
.\Show-PngGallery.ps1 -Path "C:\Bilder"
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, Position = 0)]
    [ValidateScript({ Test-Path $_ -PathType 'Container' })]
    [string]$ThumbsPath,
	[string]$PPTXPath
)

# ------------------------------------------------------------------------------------
# STA-Sicherstellung (WPF braucht STA)
# ------------------------------------------------------------------------------------
$needSta = $false
try {
    $apartment = [System.Threading.Thread]::CurrentThread.GetApartmentState()
    if ($apartment -ne [System.Threading.ApartmentState]::STA) { $needSta = $true }
} catch { $needSta = $true }

if ($needSta) {
    Write-Host "[INFO] Kein STA; Neustart im STA-Modus..." -ForegroundColor Yellow
    $psExeCandidates = @(
        "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe",
        "$env:SystemRoot\SysWOW64\WindowsPowerShell\v1.0\powershell.exe",
        "powershell.exe"
    ) | Where-Object { Test-Path $_ }

    $psExe = $psExeCandidates | Select-Object -First 1
    if (-not $psExe) { Write-Error "Keine 'powershell.exe' gefunden."; return }

    Start-Process -FilePath $psExe -ArgumentList @('-STA','-NoProfile','-ExecutionPolicy','Bypass','-File',$PSCommandPath,'-ThumbsPath',$ThumbsPath,' -PPTXPath',$PPTXPath)
    return
}

# ------------------------------------------------------------------------------------
# WPF laden
# ------------------------------------------------------------------------------------
Add-Type -AssemblyName PresentationCore, PresentationFramework, WindowsBase

# ------------------------------------------------------------------------------------
# PNG-Dateien sammeln
# ------------------------------------------------------------------------------------
$pngFiles = Get-ChildItem -Path $ThumbsPath -Filter *.png -File | Sort-Object Name
if (-not $pngFiles) {
    [System.Windows.MessageBox]::Show("Keine PNG-Dateien gefunden in:`n$THumbsPath","Keine Dateien", 'OK', 'Information') | Out-Null
    return
}

# Observable Collection für Binding
$observableCollectionType = 'System.Collections.ObjectModel.ObservableCollection[object]'
$Items = New-Object $observableCollectionType

# ViewModels vorbereiten (simple Datenobjekte)
foreach ($f in $pngFiles) {
    $Items.Add([PSCustomObject]@{
        FileName   = $f.Name
        FullPath   = $f.FullName   # Direkt an Image.Source gebunden
        IsChecked  = $false
    }) | Out-Null
}

# ------------------------------------------------------------------------------------
# XAML-UI (Breite via ElementName-Binding auf den Slider)
# ------------------------------------------------------------------------------------
# Single-quoted Here-String (keine Variablen im XAML expandieren)
$Xaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="PNG-Galerie"
        Width="1200" Height="800" WindowStartupLocation="CenterScreen">
    <DockPanel LastChildFill="True">
        <!-- Toolbar -->
        <StackPanel DockPanel.Dock="Top" Orientation="Horizontal" Margin="8">
            <TextBlock Text="Breite:" VerticalAlignment="Center" Margin="0,0,6,0" />
            <Slider x:Name="TileWidthSlider"
                    Minimum="240" Maximum="1024" Value="360" Width="260"
                    TickFrequency="40" IsSnapToTickEnabled="True"
                    VerticalAlignment="Center" />
            <TextBlock x:Name="TileWidthLabel"
                       Text="360 px" VerticalAlignment="Center" Margin="6,0,12,0" />
            <Button x:Name="ExportBtn" Content="Export" Padding="12,6" />
        </StackPanel>

        <!-- Inhalt -->
        <ScrollViewer VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Disabled">
            <ItemsControl x:Name="ItemsList">
                <ItemsControl.ItemsPanel>
                    <ItemsPanelTemplate>
                        <WrapPanel Orientation="Horizontal" />
                    </ItemsPanelTemplate>
                </ItemsControl.ItemsPanel>
                <ItemsControl.ItemTemplate>
                    <DataTemplate>
                        <!-- Kachel -->
                        <Border BorderBrush="#DDD" BorderThickness="1" CornerRadius="4" Margin="6" Padding="6" Background="#FAFAFA">
                            <!-- WICHTIG: Breite direkt an den Slider binden -->
                            <StackPanel Width="{Binding ElementName=TileWidthSlider, Path=Value}">
                                <CheckBox Content="{Binding FileName}" IsChecked="{Binding IsChecked, Mode=TwoWay}" FontWeight="SemiBold" Margin="0,0,0,6"/>

                                <!-- Bildcontainer: Stretch=Uniform bewahrt 16:9 -->
                                <Border Background="Black" BorderBrush="#EEE" BorderThickness="1">
                                    <Image Source="{Binding FullPath}"
                                           Stretch="Uniform"
                                           SnapsToDevicePixels="True"
                                           RenderOptions.BitmapScalingMode="HighQuality" />
                                </Border>

                                <TextBlock Text="{Binding FullPath}" FontSize="10" Opacity="0.6" TextTrimming="CharacterEllipsis"/>
                            </StackPanel>
                        </Border>
                    </DataTemplate>
                </ItemsControl.ItemTemplate>
            </ItemsControl>
        </ScrollViewer>
    </DockPanel>
</Window>
'@

# ------------------------------------------------------------------------------------
# XAML parsen & Controls referenzieren
# ------------------------------------------------------------------------------------
try {
    $xmlDoc = New-Object System.Xml.XmlDocument
    $xmlDoc.LoadXml($Xaml)
    $reader = New-Object System.Xml.XmlNodeReader $xmlDoc
    $Window = [Windows.Markup.XamlReader]::Load($reader)
} catch {
    Write-Error "XAML konnte nicht geladen werden: $($_.Exception.Message)"
    [System.Windows.MessageBox]::Show("Die UI konnte nicht geladen werden:`n$($_.Exception.Message)","XAML-Fehler",'OK','Error') | Out-Null
    return
}

# Titel mit Pfad (kein String-Interpolation-Problem)
$Window.Title = "PNG-Galerie " + $Path

# Controls
$ExportBtn       = $Window.FindName("ExportBtn")
$ItemsList       = $Window.FindName("ItemsList")
$TileWidthSlider = $Window.FindName("TileWidthSlider")
$TileWidthLabel  = $Window.FindName("TileWidthLabel")

if (-not $ItemsList -or -not $TileWidthSlider) {
    [System.Windows.MessageBox]::Show("UI-Elemente nicht gefunden. Bitte Skript erneut ausführen.","UI-Fehler",'OK','Error') | Out-Null
    return
}

# Items binden
$ItemsList.ItemsSource = $Items

# ------------------------------------------------------------------------------------
# Breite-Label aktualisieren
# ------------------------------------------------------------------------------------
$TileWidthLabel.Text = "$([int]$TileWidthSlider.Value) px"
$TileWidthSlider.Add_ValueChanged({
    $TileWidthLabel.Text = "$([int][Math]::Round($TileWidthSlider.Value)) px"
})

# ------------------------------------------------------------------------------------
# Export-Button: Ausgewählte Dateinamen in Konsole ausgeben
# ------------------------------------------------------------------------------------
$ExportBtn.Add_Click({
    $selected = $Items | Where-Object { $_.IsChecked -eq $true } | Select-Object -ExpandProperty FullPath
    if ($selected) {
        Write-Host "Ausgewählte Dateien:" -ForegroundColor Cyan
        $selected | ForEach-Object { Write-Host $_ }
    } else {
        Write-Host "Keine Dateien ausgewählt." -ForegroundColor Yellow
    }
})

# ------------------------------------------------------------------------------------
# Fenster anzeigen
# ------------------------------------------------------------------------------------
$null = $Window.ShowDialog()
