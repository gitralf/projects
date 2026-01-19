
<#
.SYNOPSIS
Find and remove duplicate PNG files (content-based) in a flat directory.

.PARAMETER Path
Folder path that contains PNG files (no subfolders).

.PARAMETER Delete
If set, duplicates will be deleted. If not set, this is a dry-run (no deletions).

.PARAMETER Recycle
If set together with -Delete, duplicates will be moved to Recycle Bin (requires WinForms/VisualBasic).

.PARAMETER ReportCsv
Optional: Path to a CSV report listing duplicate groups.

.EXAMPLE
.\Remove-DuplicatePng.ps1 -Path "C:\Images"            # Dry-run (show only)
.\Remove-DuplicatePng.ps1 -Path "C:\Images" -Delete     # Delete duplicates (permanently)
.\Remove-DuplicatePng.ps1 -Path "C:\Images" -Delete -Recycle -ReportCsv "C:\Images\dupes.csv"
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory = $true)]
    [ValidateScript({ Test-Path $_ -PathType 'Container' })]
    [string]$Path,
	[string]$MovetoPath,
    [string]$ReportCsv
)

Write-Host "[INFO] Starting duplicate search in: $Path" -ForegroundColor Cyan

# Collect PNG files (flat, no subfolders)
$pngFiles = Get-ChildItem -Path $Path -Filter *.png -File | Sort-Object Length, Name
if (-not $pngFiles) {
    Write-Warning "No PNG files found in: $Path"
    return
}

# Load assemblies for Recycle Bin if requested
if ($Recycle) {
    try {
        Add-Type -AssemblyName Microsoft.VisualBasic -ErrorAction Stop | Out-Null
        Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop | Out-Null
    } catch {
        Write-Warning "Recycle option is not available (required assemblies could not be loaded). Falling back to permanent delete."
        $Recycle = $false
    }
}

# 1) Group by size (quick pre-filter)
$groupsBySize = $pngFiles | Group-Object Length | Where-Object { $_.Count -gt 1 }

# 2) Hash files within each size group (SHA-256)
$dupeGroups = @()
$hashAlgo = 'SHA256'
$progressIndex = 0
$totalToHash = ($groupsBySize | ForEach-Object { $_.Count }) | Measure-Object -Sum | Select-Object -ExpandProperty Sum
if (-not $totalToHash) { $totalToHash = 0 }

foreach ($sizeGroup in $groupsBySize) {
    $hashMap = @{}  # hash -> List[FileInfo]
    foreach ($file in $sizeGroup.Group) {
        $progressIndex++
        if ($totalToHash -gt 0) {
            $pct = [Math]::Floor(($progressIndex / $totalToHash) * 100)
            Write-Progress -Activity "Hashing PNGs" -Status "$progressIndex / $totalToHash" -PercentComplete $pct
        }

        try {
            $hash = (Get-FileHash -Path $file.FullName -Algorithm $hashAlgo).Hash
            if (-not $hashMap.ContainsKey($hash)) {
                $hashMap[$hash] = New-Object System.Collections.Generic.List[System.IO.FileInfo]
            }
            $null = $hashMap[$hash].Add($file)
        } catch {
            Write-Warning "Hash failed: $($file.FullName) -> $($_.Exception.Message)"
        }
    }

    # Only hashes with multiple files are duplicate groups
    foreach ($kv in $hashMap.GetEnumerator()) {
        if ($kv.Value.Count -gt 1) {
            # Keep one (oldest by LastWriteTime, then by Name), mark others as duplicates
            $sorted = $kv.Value | Sort-Object LastWriteTime, Name
            $keep = $sorted | Select-Object -First 1
            $dupes = $sorted | Select-Object -Skip 1

            $dupeGroups += [PSCustomObject]@{
                Hash       = $kv.Key
                Size       = $keep.Length
                Keep       = $keep.FullName
                Duplicates = ($dupes | Select-Object -ExpandProperty FullName) -join '|'
                Count      = $kv.Value.Count
            }
        }
    }
}
Write-Progress -Activity "Hashing PNGs" -Completed

if (-not $dupeGroups) {
    Write-Host "[INFO] No duplicates found." -ForegroundColor Green
    return
}

# Report to console
Write-Host "[INFO] Duplicate groups found: $($dupeGroups.Count)" -ForegroundColor Yellow
$dupeGroups | ForEach-Object {
    Write-Host "Hash: $($_.Hash.Substring(0,10))...  Size: $($_.Size) bytes  Count: $($_.Count)" -ForegroundColor DarkCyan
    Write-Host "  Keep: $($_.Keep)"
    ($_.Duplicates -split '\|') | ForEach-Object { Write-Host "  Dupe: $_" -ForegroundColor DarkYellow }
}

# Optional CSV report
if ($ReportCsv) {
    try {
        $dupeGroups | Export-Csv -Path $ReportCsv -NoTypeInformation -Encoding UTF8
        Write-Host "[INFO] Report written: $ReportCsv" -ForegroundColor Cyan
    } catch {
        Write-Warning "Failed to write report: $($_.Exception.Message)"
    }
}

if ($MovetoPath){
    Write-Host "[INFO] Moving duplicates ..." -ForegroundColor Red
    $movedCount = 0
    foreach ($grp in $dupeGroups) {
        foreach ($dupe in $grp.Duplicates -split '\|') {
            try {
				$dest=$MovetoPath + "\" + $dupe.name
				Move-Item -Path $dupe -Destination $dest -Force

                Write-Host "Moved: $dupe" -ForegroundColor DarkRed
                $movedCount++
            } catch {
                Write-Warning "Could not delete: $dupe -> $($_.Exception.Message)"
            }
        }
    }
    Write-Host "[DONE] Duplicates moved: $movedCount" -ForegroundColor Green
} else {
    Write-Host "[WHATIF] No files moved. Use -MovetoPath to remove duplicates." -ForegroundColor Yellow
}
