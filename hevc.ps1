<#
    hevc-combined.ps1 - Batch and single-file HEVC (H.265) video converter for Windows PowerShell

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
      pwsh hevc-combined.ps1 [-File <file>] [-Folder <folder>] [-Encode] [-Remove] [-Recurse] [-BackupFolder <folder>] [-CRF <15-30>] [-CQ <15-30>]

    Parameters:
      -File          Path to a single video file to process
      -Folder        Folder to scan for video files (default: current directory)
      -Encode        Perform HEVC conversion (otherwise just scan/report)
      -Remove        Remove original file after successful conversion (otherwise backup)
      -Recurse       Scan subfolders recursively (for folder mode)
      -BackupFolder  Where to store backups (default: same as original)
      -CRF           Constant Rate Factor for ffmpeg (default: 18, range: 15-30, software x265)
      -CQ            Constant Quality for NVENC (hardware, overrides CRF if set)

    Notes:
      - If -CQ is specified, NVENC hardware encoding is used (faster, requires NVIDIA GPU).
      - If -CRF is specified (default), software x265 encoding is used.
      - Progress bar is shown for each file using Write-Progress.
      - Backups are saved in the same folder as the original (or -BackupFolder if specified) with .bak before the extension.
      - If -Remove is used, the original file is deleted after successful conversion.
      - Prints summary statistics at the end.

    Examples:
      pwsh hevc-combined.ps1 -File video.mp4 -Encode -CQ 21 -Remove
      pwsh hevc-combined.ps1 -Folder C:\Videos -Encode -Recurse -CRF 20
      pwsh hevc-combined.ps1 -Folder . -Encode -BackupFolder C:\Backups
#>

param(
    [string]$File = $null,
    [string]$Folder = (Get-Item .),
    [switch]$Encode = $false,
    [switch]$Remove = $false,
    [switch]$Recurse = $false,
    [string]$BackupFolder = $null,
    [ValidateRange(15, 30)]
    [int]$CRF = 21,  # Default CRF: 21 (good balance for x265)
    [ValidateRange(15, 30)]
    [int]$CQ = 23    # Default CQ: 23 (good balance for NVENC)
)

$SupportedExtensions = @('.mp4', '.mkv', '.mov')

# Show usage message for the script and its parameters
function Show-Usage {
    Write-Host @"
Usage: pwsh hevc-combined.ps1 [-File <file>] [-Folder <folder>] [-Encode] [-Remove] [-Recurse] [-BackupFolder <folder>] [-CRF <15-30>] [-CQ <15-30>] [-NVENC]

  -File          Path to a single video file to process
  -Folder        Folder to scan for video files (default: current directory)
  -Encode        Perform HEVC conversion (otherwise just scan/report)
  -Remove        Remove original file after successful conversion (otherwise backup)
  -Recurse       Scan subfolders recursively (for folder mode)
  -BackupFolder  Where to store backups (default: same as original)
  -CRF           Constant Rate Factor for ffmpeg (default: 18, range: 15-30, software x265)
  -CQ            Constant Quality for NVENC (hardware, overrides CRF if -Nvenc is used)
  -NVENC         Use NVIDIA NVENC hardware acceleration (if available)

Examples:
    hevc-combined -File video.mp4 -Encode -Nvenc -CQ 21 -Remove
    hevc-combined -Folder C:\Videos -Encode -Recurse -CRF 20
    hevc-combined -Folder . -Encode -BackupFolder C:\Backups
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
    param($VideoFile, $CRF, $CQ)
    $base = $VideoFile.BaseName -replace '_hevc(\d+)?$',''
    $newName = "${base}_hevc$($CQ -gt 0 ? $CQ : $CRF)$($VideoFile.Extension)"
    $newPath = Join-Path $VideoFile.DirectoryName $newName
    if (Test-Path $newPath) {
        Write-Host "Target file already exists: $newPath"
        return $null
    }
    $ffmpegArgs = @('-y', '-i', $VideoFile.FullName, '-c:a', 'copy', '-pix_fmt', 'yuv420p')
    if ($CQ -gt 0) {
        # Use NVENC with CQ if CQ is specified
        $ffmpegArgs += '-c:v', 'hevc_nvenc', '-preset', 'slow', '-cq', $CQ
    } else {
        # Use software x265 with CRF
        $ffmpegArgs += '-vcodec', 'libx265', '-crf', $CRF, '-x265-params', 'log-level=quiet'
    }
    $ffmpegArgs += $newPath
    $psi = New-Object System.Diagnostics.ProcessStartInfo 'ffmpeg'
    foreach ($arg in $ffmpegArgs) { $psi.ArgumentList.Add($arg) }
    $psi.RedirectStandardError = $true
    $psi.RedirectStandardOutput = $true
    $psi.UseShellExecute = $false
    $psi.CreateNoWindow = $true
    $proc = New-Object System.Diagnostics.Process
    $proc.StartInfo = $psi
    $proc.Start() | Out-Null
    $duration = 0
    try { $duration = [double]((Get-Duration $VideoFile) -replace '[^0-9\.]','') } catch {}
    $lastPercent = -1
    while (-not $proc.HasExited) {
        while ($null -ne ($line = $proc.StandardError.ReadLine())) {
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
        Write-Host "Conversion complete: $newPath"
        return $newPath
    } else {
        Write-Host "ffmpeg failed for $($VideoFile.FullName)"
        return $null
    }
}

# Backup the original video file to the backup folder (or same folder if not specified)
# Backup is named as <basename>.bak<extension>
function Backup-OldVideo {
    param($VideoFile, $BackupFolder)
    $backupDir = $BackupFolder
    if (-not $backupDir) { $backupDir = $VideoFile.DirectoryName }
    if (-not (Test-Path $backupDir)) { New-Item -Path $backupDir -ItemType Directory -Force | Out-Null }
    $backupName = "${($VideoFile.BaseName)}.bak${($VideoFile.Extension)}"
    $backupPath = Join-Path $backupDir $backupName
    Move-Item -Path $VideoFile.FullName -Destination $backupPath -Force
    Write-Host "Backed up $($VideoFile.FullName) to $backupPath"
}

# Main script logic: handles single file or batch mode, conversion, backup/remove, and reporting
try {
    $Files = @()
    if ($File) {
        if (-not (Test-Path $File -PathType Leaf)) { Write-Host "File not found: $File"; Show-Usage; exit 1 }
        $Files = ,(Get-Item $File)
    } else {
        $AbsFolder = (Get-Item $Folder | Resolve-Path).ProviderPath
        $Files = Get-VideoFiles -VideoFolder $AbsFolder -Recurse:$Recurse
    }
    if ($Files.Count -eq 0) { Write-Host "No video files found."; exit 0 }

    $Total = 0; $TotalHEVC = 0; $TotalNonHEVC = 0; $TotalProcessed = 0; $TotalSaved = 0
    foreach ($File in $Files) {
        if ($File.BaseName -match 'bak') { continue }
        $encoding = Get-Encoding $File
        $duration = Get-Duration $File
        $sizeMB = [math]::Round($File.Length/1MB, 2)
        Write-Host "[$encoding] $($File.FullName) | $sizeMB MB | $duration"
        $Total++
        if ($encoding -eq 'hevc') {
            $TotalHEVC++
            continue
        } else {
            $TotalNonHEVC++
        }
        if ($Encode) {
            $newPath = Convert-ToHEVC $File $CRF $CQ
            if ($newPath -and (Test-Path $newPath)) {
                $newSize = (Get-Item $newPath).Length
                $saved = $File.Length - $newSize
                $TotalProcessed++
                $TotalSaved += $saved
                Write-Host "File size reduced from $sizeMB MB to $([math]::Round($newSize/1MB,2)) MB"
                if ($Remove) {
                    Remove-Item $File.FullName -Force
                    Write-Host "Deleted original: $($File.FullName)"
                } else {
                    Backup-OldVideo $File $BackupFolder
                }
            }
        }
    }
    Write-Host "`nScan complete. $Total files checked. $TotalHEVC already HEVC. $TotalNonHEVC non-HEVC. $TotalProcessed processed."
    if ($TotalProcessed -gt 0) {
        Write-Host "Total space saved: $([math]::Round($TotalSaved/1MB,2)) MB"
    }
} catch {
    # Print error and usage on failure
    Write-Host "Error: $_"
    Show-Usage
    exit 1
}
