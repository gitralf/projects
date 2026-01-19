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

# --- Helpfunction: Logging ---
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

# --- Setup: Target folder + Log ---
$thumbroot = $OutPath

if (-not (Test-Path $thumbRoot)) {
    New-Item -Path $thumbRoot -ItemType Directory | Out-Null
}

$logName = (Get-Date).ToString("yyyyMMdd_HHmmss") + ".log"
$Global:LogFile = Join-Path -Path $thumbRoot -ChildPath $logName
New-Item -Path $Global:LogFile -ItemType File -Force | Out-Null
Write-Log -Message "Start Export. Path='$Path', Recurse='$Recurse', Size='$($Size -join "x")'" -Level INFO

# --- start PowerPoint COM (no Visible) ---
$ppApp = $null
try {
    $ppApp = New-Object -ComObject PowerPoint.Application
    Write-Log -Message "PowerPoint COM started. Version: $($ppApp.Version)" -Level INFO
}
catch {
    Write-Log -Message "Could not start PowerPoint. Is it installed? Error: $($_.Exception.Message)" -Level ERROR
    throw
}

# --- find files ---
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
    Write-Log -Message "Error during Get-ChildItem: $($_.Exception.Message)" -Level ERROR
    throw
}

if (-not $pptxFiles) {
    Write-Log -Message "No .pptx files found under: $Path" -Level WARN
    if ($ppApp -ne $null) {
        try { $ppApp.Quit() } catch { }
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppApp) | Out-Null
        [gc]::Collect()
        [gc]::WaitForPendingFinalizers()
    }
    Write-Host "Logfile: $Global:LogFile" -ForegroundColor Yellow
    return
}

# --- Export size determination ---
$width  = if ($Size.Count -ge 1 -and $Size[0] -gt 0) { [int]$Size[0] } else { 1920 }
$height = if ($Size.Count -ge 2 -and $Size[1] -gt 0) { [int]$Size[1] } else { [int][math]::Round($width * 9 / 16) }
Write-Log -Message "Export size: ${width}x${height}" -Level DEBUG

# --- Main loop ---
foreach ($ppt in $pptxFiles) {
    $cleanName     = [IO.Path]::GetFileNameWithoutExtension($ppt.Name)
    $pptLastWrite  = (Get-Item $ppt.FullName).LastWriteTimeUtc

    Write-Log -Message "Processing: $($ppt.FullName) (Source last modified: $($pptLastWrite.ToLocalTime()))" -Level INFO

    $pres = $null
    try {
        # Open(ReadOnly=true, Untitled=false, WithWindow=false) → fensterlos
        $pres = $ppApp.Presentations.Open($ppt.FullName, $true, $false, $false)
    }
    catch {
        Write-Log -Message "Could not open: $($ppt.FullName). Error: $($_.Exception.Message)" -Level ERROR
        continue
    }

    try {
        $slideCount = $pres.Slides.Count
        if ($slideCount -le 0) {
            Write-Log -Message "No slides: $($ppt.FullName)" -Level WARN
            $pres.Close()
            continue
        }

        # Temp folder PER FILE in local TEMP (short, safe path)
        $tmpBase = Join-Path ([IO.Path]::GetTempPath()) ("pptx_export_" + [guid]::NewGuid().ToString("N"))
        [void][System.IO.Directory]::CreateDirectory($tmpBase)
        Write-Log -Message "Temp folder: $tmpBase" -Level DEBUG

        $exported = 0
        $skipped  = 0

        for ($i = 1; $i -le $slideCount; $i++) {
            $slide = $pres.Slides.Item($i)

            $exportName      = '{0}_Slide{1:D3}.png' -f $cleanName, $i
            $exportPathFinal = Join-Path $thumbRoot $exportName

            # --- Incremental: only export if PPTX is newer than existing PNG ---
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
                Write-Log -Message "Skipped: $exportName (Thumbnail newer/equal age than source)" -Level SKIP
                $skipped++
                continue
            }

            # --- Export ---
            $tmpOut    = Join-Path $tmpBase ('{0}_S{1:D3}.png' -f $cleanName, $i)
            $fallback1 = Join-Path $tmpBase ('Slide{0}.PNG' -f $i)
            $fallback2 = Join-Path $tmpBase ('Folie{0}.PNG' -f $i)  # localized variant, in case Office uses it

            try {
                # 1) Direct slide export
                Write-Log -Message "Trying Slide.Export → $tmpOut" -Level DEBUG
                $slide.Export($tmpOut, 'PNG', $width, $height)
                Start-Sleep -Milliseconds 200  

                if (-not (Test-Path $tmpOut)) {
                    Write-Log -Message "Slide.Export did not produce a file, starting fallback via Presentation.Export" -Level WARN

                    # 2) Fallback: export entire presentation
                    Get-ChildItem -Path $tmpBase -Filter '*.png' -File -ErrorAction SilentlyContinue | Remove-Item -Force -ErrorAction SilentlyContinue
                    $pres.Export($tmpBase, 'PNG', $width, $height)
                    Start-Sleep -Milliseconds 300

                    # 2a) Primary -> Slide{i}.PNG
                    if (Test-Path $fallback1) {
                        Move-Item -Path $fallback1 -Destination $tmpOut -Force
                    }
                    # 2b) Secondary -> Folie{i}.PNG (German localization)
                    elseif (Test-Path $fallback2) {
                        Move-Item -Path $fallback2 -Destination $tmpOut -Force
                    }
                    else {
                        # 2c) Last resort: among all PNGs find the one that should correspond to _this_ slide:
                        $globCandidates = Get-ChildItem -Path $tmpBase -Filter '*.png' -File -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending
                        if ($globCandidates -and $globCandidates.Count -ge $i) {
                            # naive assignment (not perfect, but better than aborting):
                            $guess = $globCandidates | Select-Object -First 1
                            Write-Log -Message "Neither 'Slide$i.PNG' nor 'Folie$i.PNG' found; using fallback: $($guess.FullName)" -Level WARN
                            Move-Item -Path $guess.FullName -Destination $tmpOut -Force
                        } else {
                            throw "Export failed: Neither '$tmpOut' nor '$fallback1'/'$fallback2' exist."
                        }
                    }
                }

                # Final move/rename (flat target directory)
                if (Test-Path $exportPathFinal) { Remove-Item $exportPathFinal -Force }
                Move-Item -Path $tmpOut -Destination $exportPathFinal -Force

                Write-Log -Message "Exported: $exportPathFinal" -Level INFO
                $exported++
            }
            catch {
                Write-Log -Message "Error exporting slide $i from '$($ppt.Name)': $($_.Exception.Message)" -Level ERROR
            }
        }

        # Clean up tmp
        if (Test-Path $tmpBase) {
            try { Remove-Item $tmpBase -Recurse -Force } catch { }
        }

        Write-Log -Message "Done: $cleanName | Slides: $slideCount | Exported: $exported | Skipped: $skipped → $thumbRoot" -Level DONE
    }
    catch {
        Write-Log -Message "General error with '$($ppt.Name)': $($_.Exception.Message)" -Level ERROR
    }
    finally {
        if ($pres -ne $null) {
            try { $pres.Close() } catch { }
        }
    }
}

# --- Clean up PowerPoint ---
if ($ppApp -ne $null) {
    try { $ppApp.Quit() } catch { }
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppApp) | Out-Null
    [gc]::Collect()
    [gc]::WaitForPendingFinalizers()
}

Write-Log -Message "All exports completed." -Level DONE
Write-Host "Logfile: $Global:LogFile" -ForegroundColor Yellow
