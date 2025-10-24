<#
    hevc.ps1 - Batch and single-file HEVC (H.265) video converter for Windows PowerShell

    Features:
      - Processes a single file or all .mp4/.mkv/.mov files in a folder (optionally recursively)
      - Detects video encoding and duration
      - Converts non-HEVC files to HEVC using ffmpeg
      - Supports CRF (default, software x265) or CQ (NVENC, hardware) for quality control
      - Automatically uses NVENC hardware acceleration if CQ is specified, otherwise uses software x265
      - Progress bar (Write-Progress) for each file
      - Optionally backs up or removes original files after conversion
      - Prints per-file and summary statistics
      - Usage message and error reporting

    Usage:
      hevc.ps1 [-File <file>] [-Folder <folder>] [-Encode] [-Remove] [-Recurse] [-Force] [-BackupFolder <folder>] [-CRF <15-30>] [-CQ <15-30>] [-Compare]

    Parameters:
      -File          Path to a single video file to process
      -Folder        Folder to scan for video files (default: current directory)
      -Encode        Perform HEVC conversion (otherwise just scan/report)
      -Remove        Remove original file after successful conversion (otherwise backup)
      -Recurse       Scan subfolders recursively (for folder mode)
      -Force         Force re-encoding even if file is already HEVC
      -BackupFolder  Where to store backups (default: same as original)
      -CRF           Constant Rate Factor for ffmpeg (range: 15-30, software x265). If set without value, uses 21. If set, overrides CQ.
      -CQ            Constant Quality for NVENC (hardware, range: 15-30). If set without value, uses 23. Used if -CRF is not set.
      -Compare       Compare quality between original and encoded file using SSIM/PSNR (prints results before overwrite)

    Notes:
      - If -CRF is set, software x265 encoding is used (default value 21 if no value supplied).
      - If -CQ is set (and -CRF is not), NVENC hardware encoding is used (default value 23 if no value supplied).
      - If neither -CRF nor -CQ is set, NVENC with CQ=23 is used by default.
      - Progress bar is shown for each file using Write-Progress.
      - Backups are saved in the same folder as the original (or -BackupFolder if specified) with .bak after the extension.
      - If -Remove is used, the original file is deleted after successful conversion.
      - Prints summary statistics at the end.
      - If -Compare is used, prints SSIM/PSNR quality metrics before overwriting the original file.

    Examples:
      hevc.ps1 -File video.mp4 -Encode -CQ 21 -Remove
      hevc.ps1 -Folder C:\Videos -Encode -Recurse -CRF 20
      hevc.ps1 -Folder . -Encode -BackupFolder C:\Backups
      hevc.ps1 -Folder . -Encode -Force
      hevc.ps1 -File video.mp4 -Encode -Compare
#>

param(
    [Parameter(Position=0, Mandatory=$false)]
    [string]$InputPath = $null,
    [string]$File = $null,
    [string]$Folder = $null,
    [switch]$Encode = $false,
    [switch]$Remove = $false,
    [switch]$Recurse = $false,
    [string]$BackupFolder = $null,
    [ValidateRange(15, 30)]
    [int]$CRF = 21,  # Default CRF: 21 (good balance for x265)
    [ValidateRange(15, 30)]
    [int]$CQ = 23,   # Default CQ: 23 (good balance for NVENC)
    [switch]$Force = $false,  # New: Force re-encode even if already HEVC
    [Switch]$Compare = $false  # New: Compare quality between original and encoded file
)

# If InputPath is provided, determine if it's a file or folder and set $File or $Folder accordingly
if ($InputPath) {
    try {
        $resolvedPath = $null
        # Try to resolve the path, but fallback to the raw input if needed
        try {
            $resolved = Resolve-Path -LiteralPath $InputPath -ErrorAction Stop
            $resolvedPath = ($resolved | Select-Object -First 1).Path
        } catch {
            if (Test-Path -LiteralPath $InputPath) {
                $resolvedPath = (Get-Item -LiteralPath $InputPath).FullName
            }
        }
        if (-not $resolvedPath -or -not (Test-Path -LiteralPath $resolvedPath)) {
            Write-Host ("Could not resolve input path: {0}" -f $InputPath)
            Show-Usage
            exit 1
        }
        $item = $null
        try {
            $item = Get-Item -LiteralPath $resolvedPath -ErrorAction Stop
        } catch {
            Write-Host ("Failed to get item for path: {0}" -f $resolvedPath)
            Write-Host ("Exception: {0}" -f $_.Exception.Message)
            Show-Usage
            exit 1
        }
        if ($null -eq $item) {
            Write-Host ("Item is null for path: {0}" -f $resolvedPath)
            Show-Usage
            exit 1
        }
        if ($item.PSIsContainer) {
            $Folder = $item.FullName
            $File = $null
        } else {
            $File = $item.FullName
            $Folder = $null
        }
    } catch {
        Write-Host ("Input path not found or invalid: {0}" -f $InputPath)
        Show-Usage
        exit 1
    }
}
if (-not $File -and -not $Folder) {
    $Folder = (Get-Item .).FullName
}

$SupportedExtensions = @('.mp4', '.mkv', '.mov')

# Show usage message for the script and its parameters
function Show-Usage {
    Write-Host @"
Usage: hevc.ps1 [-File <file>] [-Folder <folder>] [-Encode] [-Remove] [-Recurse] [-BackupFolder <folder>] [-CRF <15-30>] [-CQ <15-30>] [-Force] [-Compare]

  -File          Path to a single video file to process
  -Folder        Folder to scan for video files (default: current directory)
  -Encode        Perform HEVC conversion (otherwise just scan/report)
  -Remove        Remove original file after successful conversion (otherwise backup)
  -Recurse       Scan subfolders recursively (for folder mode)
  -BackupFolder  Where to store backups (default: same as original)
  -CRF           Constant Rate Factor for ffmpeg (range: 15-30, software x265). If set without value, uses 21. If set, overrides CQ.
  -CQ            Constant Quality for NVENC (hardware, range: 15-30). If set without value, uses 23. Used if -CRF is not set.
  -Force         Force re-encoding even if file is already HEVC
  -Compare       Compare quality between original and encoded file using SSIM/PSNR (prints results before overwrite)

Examples:
    hevc.ps1 -File video.mp4 -Encode -CQ 21 -Remove
    hevc.ps1 -Folder C:\Videos -Encode -Recurse -CRF 20
    hevc.ps1 -Folder . -Encode -BackupFolder C:\Backups
    hevc.ps1 -Folder . -Encode -Force
    hevc.ps1 -File video.mp4 -Encode -Compare
"@
}

# Get all video files in the folder (optionally recursively)
function Get-VideoFiles {
    param($VideoFolder, $Recurse)
    $opts = @{'Path' = $VideoFolder; 'File' = $true}
    if ($Recurse) { $opts['Recurse'] = $true }
    Get-ChildItem @opts | Where-Object { $SupportedExtensions -contains $_.Extension.ToLower() }
}

# Detect the video encoding using ffprobe
function Get-Encoding {
    param($VideoFile)
    $cmd = "ffprobe -v error -select_streams v:0 -show_entries stream=codec_name -of default=nokey=1:noprint_wrappers=1 `"$($VideoFile.FullName)`""
    $encoding = Invoke-Expression $cmd
    return $encoding.Trim()
}

# Get the duration of a video file (for reporting and progress bar)
function Get-Duration {
    param($VideoFile)
    try {
        $shell = New-Object -ComObject Shell.Application
        $folder = $shell.Namespace($VideoFile.DirectoryName)
        $duration = $folder.GetDetailsOf($folder.ParseName($VideoFile.Name), 27)
        return $duration
    } catch { return '' }
}

# Convert a video file to HEVC using ffmpeg
# Uses NVENC (hardware) if CQ is specified, otherwise software x265 with CRF
# Shows a progress bar using Write-Progress
function Convert-ToHEVC {
    param($VideoFile, $FinalCRF, $FinalCQ, $UseCRF, $UseCQ)
    $tempPath = [System.IO.Path]::Combine($VideoFile.DirectoryName, ($VideoFile.BaseName + ".tmp" + $VideoFile.Extension))
    $ffmpegArgs = @('-y', '-i', $VideoFile.FullName, '-c:a', 'copy', '-pix_fmt', 'yuv420p', '-vsync', 'cfr')
    if ($UseCQ) {
        $ffmpegArgs += '-c:v', 'hevc_nvenc', '-preset', 'slow', '-cq', $FinalCQ
    } elseif ($UseCRF) {
        $ffmpegArgs += '-vcodec', 'libx265', '-crf', $FinalCRF, '-x265-params', 'log-level=quiet'
    }
    $ffmpegArgs += $tempPath
    $psi = New-Object System.Diagnostics.ProcessStartInfo 'ffmpeg'
    foreach ($arg in $ffmpegArgs) { $psi.ArgumentList.Add($arg) }
    $psi.RedirectStandardError = $true
    $psi.RedirectStandardOutput = $true
    $psi.UseShellExecute = $false
    $psi.CreateNoWindow = $true
    Write-Host "Running ffmpeg command: ffmpeg $($ffmpegArgs -join ' ')"
    $proc = New-Object System.Diagnostics.Process
    $proc.StartInfo = $psi
    $proc.Start() | Out-Null
    $duration = 0
    try { $duration = [double]((Get-Duration $VideoFile) -replace '[^0-9\.]','') } catch {}
    $lastPercent = -1
    $stderr = ""
    while (-not $proc.HasExited) {
        while ($null -ne ($line = $proc.StandardError.ReadLine())) {
            $stderr += $line + "`n"
            if ($line -match "time=([0-9:.]+)") {
                $timeStr = $matches[1]
                $parts = $timeStr -split ':'
                if ($parts.Length -eq 3) {
                    $seconds = [double]$parts[0]*3600 + [double]$parts[1]*60 + [double]$parts[2]
                } elseif ($parts.Length -eq 2) {
                    $seconds = [double]$parts[0]*60 + [double]$parts[1]
                } else {
                    $seconds = [double]$parts[0]
                }
                if ($duration -gt 0) {
                    $percent = [math]::Round(($seconds/$duration)*100)
                    if ($percent -ne $lastPercent) {
                        Write-Progress -Activity "Encoding $($VideoFile.Name)" -Status "$percent% complete" -PercentComplete $percent
                        $lastPercent = $percent
                    }
                }
            }
        }
        Start-Sleep -Milliseconds 200
    }
    Write-Progress -Activity "Encoding $($VideoFile.Name)" -Completed
    $proc.WaitForExit()
    if ($proc.ExitCode -eq 0) {
        Write-Host "Conversion complete: $tempPath"
        return $tempPath
    } else {
        Write-Host "ffmpeg failed for $($VideoFile.FullName)"
        Write-Host "ffmpeg stderr output:\n$stderr"
        return $null
    }
}

# Backup the original video file to <filename>.<ext>.bak
function Backup-OldVideo {
    param($VideoFile)
    $backupPath = $VideoFile.FullName + ".bak"
    Move-Item -Path $VideoFile.FullName -Destination $backupPath -Force
    Write-Host "Backed up $($VideoFile.FullName) to $backupPath"
}

# Compare quality between original and encoded video using ffmpeg
function Compare-Quality {
    param($OriginalPath, $EncodedPath)
    # Use ffmpeg to compute SSIM and PSNR
    Write-Host "Comparing quality (SSIM/PSNR) between original and encoded..."
    $output = & ffmpeg -i $OriginalPath -i $EncodedPath -lavfi "[0:v][1:v]ssim;[0:v][1:v]psnr" -f null - 2>&1
    if ($LASTEXITCODE -eq 0) {
        $ssimLine = $output | Select-String -Pattern 'All:([0-9.]+)' | Select-Object -First 1
        $psnrLine = $output | Select-String -Pattern 'average:([0-9.]+)' | Select-Object -First 1
        if ($ssimLine) {
            $ssim = $ssimLine.Matches[0].Groups[1].Value
            Write-Host "SSIM: $ssim"
        }
        if ($psnrLine) {
            $psnr = $psnrLine.Matches[0].Groups[1].Value
            Write-Host "PSNR: $psnr"
        }
    } else {
        Write-Host "Quality comparison failed. Output:"
        Write-Host $output
    }
}

# Main script logic: handles single file or batch mode, conversion, backup/remove, and reporting
try {
    $Files = @()
    if ($File) {
        if (-not (Test-Path $File -PathType Leaf)) {
            Write-Host "File not found: $File"
            Show-Usage
            exit 1
        }
        $Files = ,(Get-Item $File)
    } else {
        $AbsFolder = (Get-Item $Folder | Resolve-Path).ProviderPath
        $Files = Get-VideoFiles -VideoFolder $AbsFolder -Recurse:$Recurse
    }
    if ($Files.Count -eq 0) {
        Write-Host "No video files found."
        exit 0
    }

    # Parse CRF and CQ logic
    $UseCRF = $false
    $UseCQ = $false
    $FinalCRF = $null
    $FinalCQ = $null

    if ($PSBoundParameters.ContainsKey('CRF')) {
        if ($null -eq $CRF -or $CRF -eq "") {
            $FinalCRF = 21
        } elseif ($CRF -ge 15 -and $CRF -le 30) {
            $FinalCRF = $CRF
        } else {
            Write-Host "CRF value $CRF is out of range (15-30)."
            exit 1
        }
        $UseCRF = $true
    }
    if ($PSBoundParameters.ContainsKey('CQ')) {
        if ($null -eq $CQ -or $CQ -eq "") {
            $FinalCQ = 23
        } elseif ($CQ -ge 15 -and $CQ -le 30) {
            $FinalCQ = $CQ
        } else {
            Write-Host "CQ value $CQ is out of range (15-30)."
            exit 1
        }
        $UseCQ = $true
    }
    if (-not $UseCRF -and -not $UseCQ) {
        $FinalCQ = 23
        $UseCQ = $true
    }

    $Total = 0; $TotalHEVC = 0; $TotalNonHEVC = 0; $TotalProcessed = 0; $TotalSaved = 0
    foreach ($VideoFile in $Files) {
        if ($VideoFile.BaseName -match 'bak') { continue }
        $encoding = Get-Encoding $VideoFile
        if ($encoding -eq 'hevc' -and -not $Force) {
            $TotalHEVC++
            continue
        } elseif ($encoding -eq 'hevc' -and $Force) {
            Write-Host "Forcing re-encode of already HEVC file: $($VideoFile.FullName)"
            $TotalHEVC++
        } else {
            $TotalNonHEVC++
        }
        $duration = Get-Duration $VideoFile
        $sizeMB = [math]::Round($VideoFile.Length/1MB, 2)
        Write-Host "[$encoding] $($VideoFile.FullName) | $sizeMB MB | $duration"
        $Total++
        if ($Encode) {
            $tempPath = Convert-ToHEVC $VideoFile $FinalCRF $FinalCQ $UseCRF $UseCQ
            if ($tempPath -and (Test-Path $tempPath)) {
                if ($Compare) {
                    Compare-Quality $VideoFile.FullName $tempPath
                }
                if (-not $Remove) {
                    Backup-OldVideo $VideoFile
                }
                Move-Item -Path $tempPath -Destination $VideoFile.FullName -Force
                $newSize = (Get-Item $VideoFile.FullName).Length
                $saved = $VideoFile.Length - $newSize
                $TotalProcessed++
                $TotalSaved += $saved
                Write-Host "File size reduced from $sizeMB MB to $([math]::Round($newSize/1MB,2)) MB"
                if ($Remove) {
                    Remove-Item $VideoFile.FullName -Force
                    Write-Host "Deleted original: $($VideoFile.FullName)"
                }
            }
        }
    }
    Write-Host "`nScan complete. $Total files checked. $TotalHEVC already HEVC. $TotalNonHEVC non-HEVC. $TotalProcessed processed."
    if ($TotalProcessed -gt 0) {
        Write-Host "Total space saved: $([math]::Round($TotalSaved/1MB,2)) MB"
    }
}
catch {
    Write-Host "Error: $_"
    Show-Usage
    exit 1
}
