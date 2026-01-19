<# 
.SYNOPSIS
Exportiert aus allen .pptx-Dateien in einem Verzeichnis alle Folien als PNG in "<Path>\Thumbnails".
- Flacher Zielordner (keine Unterordner)
- Dateinamen mit dreistelliger Foliennummer (z. B. _Slide001.png)
- Inkrementeller Export: nur wenn PPTX neuer als Ziel-PNG
- Robuster Export mit Fallback (Slide.Export -> Presentation.Export)
- Temporärordner im lokalen %TEMP% (kurz & unkritisch)
- Logfile in Thumbnails (Dateiname = Zeitstempel bis Sekunde)

.PARAMETER Path
Basisverzeichnis, in dem nach .pptx gesucht wird.

.PARAMETER Recurse
Wenn gesetzt, werden Unterordner rekursiv durchsucht.

.PARAMETER Size
Breite, Höhe in Pixeln (Exportgröße). Standard: 1920x1080 (16:9).

.EXAMPLE
.\Export-PptxSlides.ps1 -Path "C:\PPTs" -Recurse -Size 2560,1440
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, Position = 0)]
    [ValidateScript({ Test-Path $_ -PathType 'Container' })]
    [string]$Path,
	[string]$OutPath,
    [switch]$Recurse,

    # Breite, Höhe in Pixeln
    [int[]]$Size = @(1920, 1080)
)

# --- Hilfsfunktion: Logging ---
function Write-Log {
    param(
        [string]$Message,
        [ValidateSet("INFO","WARN","ERROR","DONE","SKIP","DEBUG")]
        [string]$Level = "INFO"
    )
    $timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    $line = "[{0}] [{1}] {2}" -f $timestamp, $Level, $Message
    Write-Host $line
    Add-Content -Path $Global:LogFile -Value $line
}

# --- Setup: Zielordner + Log ---
$thumbroot = $OutPath
# $thumbRoot = Join-Path -Path $Path -ChildPath 'Thumbnails'



if (-not (Test-Path $thumbRoot)) {
    New-Item -Path $thumbRoot -ItemType Directory | Out-Null
}

$logName = (Get-Date).ToString("yyyyMMdd_HHmmss") + ".log"
$Global:LogFile = Join-Path -Path $thumbRoot -ChildPath $logName
New-Item -Path $Global:LogFile -ItemType File -Force | Out-Null
Write-Log -Message "Start Export. Path='$Path', Recurse='$Recurse', Size='$($Size -join "x")'" -Level INFO

# --- PowerPoint COM starten (ohne Visible setzen) ---
$ppApp = $null
try {
    $ppApp = New-Object -ComObject PowerPoint.Application
    Write-Log -Message "PowerPoint COM gestartet. Version: $($ppApp.Version)" -Level INFO
}
catch {
    Write-Log -Message "Konnte PowerPoint nicht starten. Ist es installiert? Fehler: $($_.Exception.Message)" -Level ERROR
    throw
}

# --- Dateien finden ---
$searchParams = @{
    Path        = $Path
    Filter      = '*.pptx'
    File        = $true
    ErrorAction = 'Stop'
}
if ($Recurse) { $searchParams['Recurse'] = $true }

try {
    $pptxFiles = Get-ChildItem @searchParams
}
catch {
    Write-Log -Message "Fehler bei Get-ChildItem: $($_.Exception.Message)" -Level ERROR
    throw
}

if (-not $pptxFiles) {
    Write-Log -Message "Keine .pptx-Dateien gefunden unter: $Path" -Level WARN
    if ($ppApp -ne $null) {
        try { $ppApp.Quit() } catch { }
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppApp) | Out-Null
        [gc]::Collect()
        [gc]::WaitForPendingFinalizers()
    }
    Write-Host "Logfile: $Global:LogFile" -ForegroundColor Yellow
    return
}

# --- Export-Größe bestimmen ---
$width  = if ($Size.Count -ge 1 -and $Size[0] -gt 0) { [int]$Size[0] } else { 1920 }
$height = if ($Size.Count -ge 2 -and $Size[1] -gt 0) { [int]$Size[1] } else { [int][math]::Round($width * 9 / 16) }
Write-Log -Message "Export-Größe: ${width}x${height}" -Level DEBUG

# --- Hauptschleife ---
foreach ($ppt in $pptxFiles) {
    $cleanName     = [IO.Path]::GetFileNameWithoutExtension($ppt.Name)
    $pptLastWrite  = (Get-Item $ppt.FullName).LastWriteTimeUtc

    Write-Log -Message "Verarbeite: $($ppt.FullName) (Quelle geändert am: $($pptLastWrite.ToLocalTime()))" -Level INFO

    $pres = $null
    try {
        # Open(ReadOnly=true, Untitled=false, WithWindow=false) → fensterlos
        $pres = $ppApp.Presentations.Open($ppt.FullName, $true, $false, $false)
    }
    catch {
        Write-Log -Message "Konnte öffnen: $($ppt.FullName). Fehler: $($_.Exception.Message)" -Level ERROR
        continue
    }

    try {
        $slideCount = $pres.Slides.Count
        if ($slideCount -le 0) {
            Write-Log -Message "Keine Folien: $($ppt.FullName)" -Level WARN
            $pres.Close()
            continue
        }

        # Temp-Ordner JE DATEI im lokalen TEMP (kurzer, sicherer Pfad)
        $tmpBase = Join-Path ([IO.Path]::GetTempPath()) ("pptx_export_" + [guid]::NewGuid().ToString("N"))
        [void][System.IO.Directory]::CreateDirectory($tmpBase)
        Write-Log -Message "Temp-Ordner: $tmpBase" -Level DEBUG

        $exported = 0
        $skipped  = 0

        for ($i = 1; $i -le $slideCount; $i++) {
            $slide = $pres.Slides.Item($i)

            $exportName      = '{0}_Slide{1:D3}.png' -f $cleanName, $i
            $exportPathFinal = Join-Path $thumbRoot $exportName

            # --- Inkrementell: nur exportieren, wenn PPTX neuer als existierendes PNG ---
            $needExport = $true
            if (Test-Path $exportPathFinal) {
                try {
                    $pngLastWrite = (Get-Item $exportPathFinal).LastWriteTimeUtc
                    if ($pngLastWrite -ge $pptLastWrite) {
                        $needExport = $false
                    }
                } catch {
                    $needExport = $true
                }
            }

            if (-not $needExport) {
                Write-Log -Message "Übersprungen: $exportName (Thumbnail neuer/gleich alt wie Quelle)" -Level SKIP
                $skipped++
                continue
            }

            # --- Export ---
            $tmpOut    = Join-Path $tmpBase ('{0}_S{1:D3}.png' -f $cleanName, $i)
            $fallback1 = Join-Path $tmpBase ('Slide{0}.PNG' -f $i)
            $fallback2 = Join-Path $tmpBase ('Folie{0}.PNG' -f $i)  # lokalisierte Variante, falls Office das nutzt

            try {
                # 1) Direkter Folien-Export
                Write-Log -Message "Versuche Slide.Export → $tmpOut" -Level DEBUG
                $slide.Export($tmpOut, 'PNG', $width, $height)
                Start-Sleep -Milliseconds 200  # PP braucht manchmal eine kleine Pause

                if (-not (Test-Path $tmpOut)) {
                    Write-Log -Message "Slide.Export hat keine Datei erzeugt, starte Fallback per Presentation.Export" -Level WARN

                    # 2) Fallback: Gesamte Präsentation exportieren
                    Get-ChildItem -Path $tmpBase -Filter '*.png' -File -ErrorAction SilentlyContinue | Remove-Item -Force -ErrorAction SilentlyContinue
                    $pres.Export($tmpBase, 'PNG', $width, $height)
                    Start-Sleep -Milliseconds 300

                    # 2a) Primär -> Slide{i}.PNG
                    if (Test-Path $fallback1) {
                        Move-Item -Path $fallback1 -Destination $tmpOut -Force
                    }
                    # 2b) Sekundär -> Folie{i}.PNG (deutsche Lokalisierung)
                    elseif (Test-Path $fallback2) {
                        Move-Item -Path $fallback2 -Destination $tmpOut -Force
                    }
                    else {
                        # 2c) Letzter Versuch: unter allen PNGs den treffen, der _dieser_ Folie entsprechen sollte:
                        $globCandidates = Get-ChildItem -Path $tmpBase -Filter '*.png' -File -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending
                        if ($globCandidates -and $globCandidates.Count -ge $i) {
                            # naive Zuordnung (nicht perfekt, aber besser als Abbruch):
                            $guess = $globCandidates | Select-Object -First 1
                            Write-Log -Message "Weder 'Slide$i.PNG' noch 'Folie$i.PNG' gefunden; nutze Ersatz: $($guess.FullName)" -Level WARN
                            Move-Item -Path $guess.FullName -Destination $tmpOut -Force
                        } else {
                            throw "Export fehlgeschlagen: Weder '$tmpOut' noch '$fallback1'/'$fallback2' existiert."
                        }
                    }
                }

                # Final verschieben/umbenennen (flaches Zielverz.)
                if (Test-Path $exportPathFinal) { Remove-Item $exportPathFinal -Force }
                Move-Item -Path $tmpOut -Destination $exportPathFinal -Force

                Write-Log -Message "Exportiert: $exportPathFinal" -Level INFO
                $exported++
            }
            catch {
                Write-Log -Message "Fehler beim Export Slide $i aus '$($ppt.Name)': $($_.Exception.Message)" -Level ERROR
            }
        }

        # Aufräumen tmp
        if (Test-Path $tmpBase) {
            try { Remove-Item $tmpBase -Recurse -Force } catch { }
        }

        Write-Log -Message "Fertig: $cleanName | Folien: $slideCount | Exportiert: $exported | Übersprungen: $skipped → $thumbRoot" -Level DONE
    }
    catch {
        Write-Log -Message "Genereller Fehler bei '$($ppt.Name)': $($_.Exception.Message)" -Level ERROR
    }
    finally {
        if ($pres -ne $null) {
            try { $pres.Close() } catch { }
        }
    }
}

# --- Aufräumen PowerPoint ---
if ($ppApp -ne $null) {
    try { $ppApp.Quit() } catch { }
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppApp) | Out-Null
    [gc]::Collect()
    [gc]::WaitForPendingFinalizers()
}

Write-Log -Message "Alle Exporte abgeschlossen." -Level DONE
Write-Host "Logfile: $Global:LogFile" -ForegroundColor Yellow
