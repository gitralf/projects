
<# 
.SYNOPSIS
Zeigt alle PNG-Dateien eines Verzeichnisses in einem WPF-Fenster an (scrollbar, 16:9, Checkbox je Bild).
Beim Klick auf "Export" werden die Dateinamen der ausgewählten PNGs in der Konsole ausgegeben
UND eine neue PowerPoint aus den Original-Folien erstellt.

.PARAMETER Path
Ordner mit PNG-Dateien (z. B. "...\Thumbnails").

.PARAMETER PptxRoot
Root-Pfad, in dem die originalen .pptx liegen (Standard: 
- Wenn 'Path' auf '...\Thumbnails' endet → dessen Elternordner
- Sonst: derselbe Ordner wie 'Path').

.PARAMETER RecursePptx
Wenn gesetzt, werden .pptx rekursiv unterhalb von PptxRoot gesucht.

.EXAMPLE
.\Show-PngGallery.ps1 -Path "C:\PPTs\Thumbnails"
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

    $args = @(
        '-STA','-NoProfile','-ExecutionPolicy','Bypass',
        '-File', $PSCommandPath,
        '-ThumbsPath', $ThumbsPath,
		'-PPTXPath', $PPTXPath
    )

    Start-Process -FilePath $psExe -ArgumentList $args
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
    [System.Windows.MessageBox]::Show("Keine PNG-Dateien gefunden in:`n$ThumbsPath","Keine Dateien", 'OK', 'Information') | Out-Null
    return
}

# # PptxRoot ableiten, falls nicht gesetzt:
# if (-not $PptxRoot) {
#     if ((Split-Path -Leaf $Path) -ieq 'Thumbnails') {
#         $PptxRoot = Split-Path -Parent $Path
#     } else {
#         $PptxRoot = $Path
#     }
# }
# Write-Host "[INFO] PptxRoot: $PptxRoot" -ForegroundColor DarkCyan

# ------------------------------------------------------------------------------------
# Observable Collection für Binding
# ------------------------------------------------------------------------------------
$observableCollectionType = 'System.Collections.ObjectModel.ObservableCollection[object]'
$Items = New-Object $observableCollectionType

foreach ($f in $pngFiles) {
    $Items.Add([PSCustomObject]@{
        FileName  = $f.Name
        FullPath  = $f.FullName   # Source fürs Image
        IsChecked = $false
    }) | Out-Null
}

$totalCount = $Items.Count

# ------------------------------------------------------------------------------------
# XAML-UI (Breite via ElementName-Binding auf den Slider)
# ------------------------------------------------------------------------------------
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
            <Button x:Name="ExportBtn" Content="Export (Liste + PPTX bauen)" Padding="12,6" />
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
                            <StackPanel Width="{Binding ElementName=TileWidthSlider, Path=Value}">
                                <CheckBox Content="{Binding FileName}" IsChecked="{Binding IsChecked, Mode=TwoWay}" FontWeight="SemiBold" Margin="0,0,0,6"/>
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

$Window.Title = "PNG-Galerie " + $Path

# Controls
$ExportBtn       = $Window.FindName("ExportBtn")
$ItemsList       = $Window.FindName("ItemsList")
$TileWidthSlider = $Window.FindName("TileWidthSlider")
$TileWidthLabel  = $Window.FindName("TileWidthLabel")

if (-not $ItemsList -or -not $TileWidthSlider -or -not $ExportBtn) {
    [System.Windows.MessageBox]::Show("UI-Elemente nicht gefunden. Bitte Skript erneut ausfuehren.","UI-Fehler",'OK','Error') | Out-Null
    return
}

# Bind Items


$ItemsList.ItemsSource = $Items

# Slider Label
$TileWidthLabel.Text = "$([int]$TileWidthSlider.Value) px"
$TileWidthSlider.Add_ValueChanged({
    $TileWidthLabel.Text = "$([int][Math]::Round($TileWidthSlider.Value)) px"
})

# ------------------------------------------------------------------------------------
# Hilfsfunktionen (PPTX Bauen)
# ------------------------------------------------------------------------------------

function Resolve-PptxPath {
    param(
        [Parameter(Mandatory=$true)][string]$BaseName,        # ohne .pptx
        [Parameter(Mandatory=$true)][string]$Root
    )
    $filter = "$BaseName.pptx"
    $match = Get-ChildItem -Path $Root -Filter $filter -File -ErrorAction SilentlyContinue | Select-Object -First 1
    return $match.FullName
}

function Build-PptxFromSelection {

	# Aufruf         Build-PptxFromSelection -SelectedPngs $selected -PptxRoot $PptxRoot -Recurse:$RecursePptx
    param(
        [Parameter(Mandatory=$true)][string[]]$SelectedPngs,
        [Parameter(Mandatory=$true)][string]$PPTXPath
    )

    if (-not $SelectedPngs -or $SelectedPngs.Count -eq 0) {
        Write-Host "[INFO] Keine Auswahl fuer PPTX-Bau." -ForegroundColor Yellow
        return
	}

    # 1) Zielpfad erfragen (SaveFileDialog)
    $sfd = New-Object Microsoft.Win32.SaveFileDialog
    $sfd.Filter = "PowerPoint-Präsentation (*.pptx)|*.pptx"
    $sfd.OverwritePrompt = $true
    $sfd.FileName = "Auswahl_" + (Get-Date).ToString("yyyyMMdd_HHmmss") + ".pptx"
    $ok = $sfd.ShowDialog()
    if (-not $ok) {
        Write-Host "[INFO] Abbruch: Kein Ziel ausgewaehlt." -ForegroundColor Yellow
        return
    }
    $targetPath = $sfd.FileName

    # 2) Ausgewählte PNGs parsen -> (PPTX-Basename, SlideNum)
    $rx = [regex]::new('^(?<base>.+?)_Slide(?<num>\d+)\.png$', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)

    $map = @()  # Liste von @{'pptxName'='XY'; 'slide'=13; 'png'='...'; 'pptxPath'='...'}
    foreach ($png in $SelectedPngs) {
        $fn = [IO.Path]::GetFileName($png)
        $m = $rx.Match($fn)
        if (-not $m.Success) {
            Write-Host "[WARN] PNG passt nicht zum erwarteten Muster: $fn" -ForegroundColor DarkYellow
            continue
        }
        $base = $m.Groups['base'].Value
        $num  = [int]$m.Groups['num'].Value

        Write-Host "[INFO] Aufruf fuer basename $base und root $PPTXPath" -ForegroundColor DarkYellow

        $pptx = Resolve-PptxPath -BaseName $base -Root $PPTXPath 
        if (-not $pptx) {
            Write-Host "[ERROR] Zu '$fn' wurde keine passende PPTX gefunden unter '$PPTXPath'." -ForegroundColor Red
            continue
        }

        $map += [PSCustomObject]@{
            pptxName = $base
            slide    = $num
            png      = $png
            pptx	 = $pptx
        }
    }

    if (-not $map -or $map.Count -eq 0) {
        Write-Host "[INFO] Keine validen Zuordnungen gefunden: Abbruch." -ForegroundColor Yellow
        return
    }

    # 3) Neue PPTX erstellen & Slides einfügen
    $ppApp = $null
    try {
        $ppApp = New-Object -ComObject PowerPoint.Application
        # Kein Visible setzen
        $newPres = $ppApp.Presentations.Add()

        # Quelle PPTX öffnen (Cache je Datei)
        $openCache = @{} # key: pptxPath, value: pres
        try {
            foreach ($entry in $map) {
                $srcPath = $entry.pptx
                $slideNo = $entry.slide

                if (-not $openCache.ContainsKey($srcPath)) {
                    $openCache[$srcPath] = $ppApp.Presentations.Open($srcPath, $true, $false, $false) # ReadOnly, no window
                }
                $srcPres = $openCache[$srcPath]

                # InsertFromFile: Insert nach Index (0 = am Anfang), enthält SlideStart..SlideEnd
                $insertAfterIndex = $newPres.Slides.Count  # hinten anhängen
                try {
                    # Insert einzelne Folie (SlideStart = SlideEnd = slideNo)
                    [void]$newPres.Slides.InsertFromFile($srcPath, $insertAfterIndex, $slideNo, $slideNo)
                    Write-Host "[INFO] Eingefuegt: $($entry.pptxName) Folie $slideNo" -ForegroundColor Cyan
                }
                catch {
                    Write-Host "[ERROR] InsertFromFile fehlgeschlagen fuer $srcPath Folie $slideNo : $($_.Exception.Message)" -ForegroundColor Red
                }
            }

            # Speichern
            $newPres.SaveAs($targetPath)

            Write-Host "[DONE] Neue Praesentation gespeichert $targetPath" -ForegroundColor Green

        }
        finally {
            # Quellen schließen
            foreach ($kv in $openCache.GetEnumerator()) {
                try { $kv.Value.Close() } catch {}
            }
            # Zielpräs schließen
            try { $newPres.Close() } catch {}
        }
    }
    catch {
        Write-Host "[ERROR] PowerPoint-Automation fehlgeschlagen: $($_.Exception.Message)" -ForegroundColor Red
    }
    finally {
        if ($ppApp -ne $null) {
            try { $ppApp.Quit() } catch {}
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppApp) | Out-Null
            [gc]::Collect()
            [gc]::WaitForPendingFinalizers()
        }
    }
}

# ------------------------------------------------------------------------------------
# Export-Button: Auswahl → Ausgabe + PPTX bauen
# ------------------------------------------------------------------------------------

$ExportBtn.Add_Click({
    $selected = $Items | Where-Object { $_.IsChecked -eq $true } | Select-Object -ExpandProperty FullPath
    if ($selected) {
        Write-Host "Ausgewählte Dateien:" -ForegroundColor Cyan
        $selected | ForEach-Object { Write-Host $_ }

        # PPTX bauen
        Build-PptxFromSelection -SelectedPngs $selected -PPTXPath $PPTXPath
    } else {
        Write-Host "Keine Dateien ausgewählt." -ForegroundColor Yellow
    }
})

write-host "[INFO] loading $totalcount images, please wait..." -ForegroundColor Green  



# ------------------------------------------------------------------------------------
# Fenster anzeigen
# ------------------------------------------------------------------------------------
$null = $Window.ShowDialog()
