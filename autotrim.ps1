<#
    rec.ps1 - Automatically trim a video file at the point where it becomes mostly white or mostly black.

    - Scans a video file for frames that are mostly white or black (using ffmpeg scene analysis).
    - Trims the video from the start until the detected point and saves it as a new file.
    - The new filename is appended with the trim end time as -HH-MM-SS.
    - Processes a single video file or all supported files in a folder (non-recursive).
    - Requires ffmpeg to be installed and in PATH.
    - Supports .mkv, .mp4, and .mov files.

    Usage:
        rec -Path <video | folder>
        rec <video | folder>
#>

param(
    [Parameter(Position=0, Mandatory=$false)]
    [Alias('File')]
    [string]$Path
)

function Show-Usage {
    Write-Host "Usage: autotrim [-Path] <video file | folder>"
}

# Helper: Convert seconds to HH-MM-SS (approved verb: Convert)
function ConvertTo-HHMMSS {
    param($seconds)
    $ts = [TimeSpan]::FromSeconds([int]$seconds)
    return ('{0:D2}-{1:D2}-{2:D2}' -f $ts.Hours, $ts.Minutes, $ts.Seconds)
}

$allowedExt = @('.mkv', '.mp4', '.mov')

function Refine-SegmentStart {
    param(
        [string]$File,
        [double]$start,
        [double]$end,
        [double]$minLength,
        [string]$mode  # 'bw' or 'static'
    )
    $low = $start
    $high = $end
    while (($high - $low) -gt 1) {
        $mid = [math]::Floor(($low + $high) / 2)
        $segmentFound = $false
        if ($mode -eq 'bw') {
            $segmentFound = $true
            for ($i = 0; $i -lt $minLength; $i++) {
                $frameTime = $mid + $i
                $result = & ffmpeg -hide_banner -ss $frameTime -i "$File" -vframes 1 -vf "blackdetect=d=0.01:pic_th=0.98,signalstats" -an -f null - 2>&1
                $isBlack = $result | Select-String -Pattern "black_start"
                $yavg = ($result | Select-String -Pattern "lavfi.signalstats.YAVG=([0-9.]+)" | ForEach-Object { $_.Matches[0].Groups[1].Value })
                $isWhite = $false
                if ($yavg) {
                    foreach ($y in $yavg) {
                        if ([double]$y -gt 245) { $isWhite = $true; break }
                    }
                }
                if (-not ($isBlack -or $isWhite)) {
                    $segmentFound = $false
                    break
                }
            }
        } elseif ($mode -eq 'static') {
            $segmentFound = $true
            $prevHash = $null
            for ($i = 0; $i -lt $minLength; $i++) {
                $frameTime = $mid + $i
                $result = & ffmpeg -hide_banner -ss $frameTime -i "$File" -vframes 1 -vf "crop=iw:ih,scale=160:90,format=gray,hash=md5" -f null - 2>&1
                $hash = ($result | Select-String -Pattern "MD5=([0-9a-f]+)" | ForEach-Object { $_.Matches[0].Groups[1].Value })
                if ($i -eq 0) { $prevHash = $hash }
                if ($hash -ne $prevHash) {
                    $segmentFound = $false
                    break
                }
            }
        }
        if ($segmentFound) {
            $high = $mid
        } else {
            $low = $mid + 1
        }
    }
    return $high
}

function Find-SegmentAdaptive {
    param(
        [string]$File,
        [double]$duration,
        [double]$minLength,
        [string]$mode  # 'bw' or 'static'
    )
    $maxStep = 10
    $minStep = 1
    $scanStart = [math]::Floor($duration - $minLength)
    $scanEnd = [math]::Ceiling($duration * 0.2)
    $step = $maxStep
    $bestStart = $null
    while ($step -ge $minStep) {
        $found = $false
        $t = $scanStart
        while ($t -ge $scanEnd) {
            $segment = $true
            if ($mode -eq 'bw') {
                for ($i = 0; $i -lt $minLength; $i++) {
                    $frameTime = $t + $i
                    if ($frameTime -ge $duration) { $segment = $false; break }
                    $result = & ffmpeg -hide_banner -ss $frameTime -i "$File" -vframes 1 -vf "blackdetect=d=0.01:pic_th=0.98,signalstats" -an -f null - 2>&1
                    $isBlack = $result | Select-String -Pattern "black_start"
                    $yavg = ($result | Select-String -Pattern "lavfi.signalstats.YAVG=([0-9.]+)" | ForEach-Object { $_.Matches[0].Groups[1].Value })
                    $isWhite = $false
                    if ($yavg) {
                        foreach ($y in $yavg) {
                            if ([double]$y -gt 245) { $isWhite = $true; break }
                        }
                    }
                    if (-not ($isBlack -or $isWhite)) {
                        $segment = $false
                        break
                    }
                }
            } elseif ($mode -eq 'static') {
                $prevHash = $null
                for ($i = 0; $i -lt $minLength; $i++) {
                    $frameTime = $t + $i
                    if ($frameTime -ge $duration) { $segment = $false; break }
                    $result = & ffmpeg -hide_banner -ss $frameTime -i "$File" -vframes 1 -vf "crop=iw:ih,scale=160:90,format=gray,hash=md5" -f null - 2>&1
                    $hash = ($result | Select-String -Pattern "MD5=([0-9a-f]+)" | ForEach-Object { $_.Matches[0].Groups[1].Value })
                    if ($i -eq 0) { $prevHash = $hash }
                    if ($hash -ne $prevHash) {
                        $segment = $false
                        break
                    }
                }
            }
            if ($segment) {
                $found = $true
                $bestStart = $t
                break
            }
            $t -= $step
        }
        if ($found) {
            $scanStart = $bestStart + [math]::Floor($step/2)
            $scanEnd = [math]::Max($bestStart - $step, $scanEnd)
            $step = [math]::Floor($step / 2)
        } else {
            $step = [math]::Floor($step / 2)
        }
    }
    return $bestStart
}

function Test-BlackWhiteSegment {
    param([string]$File, [double]$duration, [double]$minLength)
    return Find-SegmentAdaptive $File $duration $minLength 'bw'
}

function Test-BlurOrStatic {
    param([string]$File, [double]$duration, [double]$minLength)
    return Find-SegmentAdaptive $File $duration $minLength 'static'
}

function Invoke-RecVideoFile {
    param([string]$File)
    if (-not (Test-Path $File -PathType Leaf)) {
        Write-Host "File not found: $File"
        return
    }
    $ext = [System.IO.Path]::GetExtension($File).ToLower()
    if ($allowedExt -notcontains $ext) {
        Write-Host "Unsupported file type: $ext. Supported file types are: $($allowedExt -join ', ')."
        return
    }
    # Get video duration
    $durationLine = & ffprobe -v error -show_entries format=duration -of default=noprint_wrappers=1:nokey=1 "$File"
    $duration = [double]::Parse($durationLine.Trim())
    $bwStart = Test-BlackWhiteSegment $File $duration 5
    $staticStart = Test-BlurOrStatic $File $duration 10
    $trimTime = $null
    if (($null -ne $bwStart) -and (($null -eq $staticStart) -or ($bwStart -lt $staticStart))) {
        $trimTime = $bwStart
        Write-Host "Detected mostly black/white segment at $trimTime seconds (>=5s)."
    } elseif ($null -ne $staticStart) {
        $trimTime = $staticStart
        Write-Host "Detected static/blurred segment at $trimTime seconds (>=10s)."
    }
    if ($null -eq $trimTime -or $trimTime -lt 1) {
        Write-Host "No valid trim point detected. No trim performed for $File."
        return
    }
    $newName = "{0}-{1}{2}" -f ([System.IO.Path]::GetFileNameWithoutExtension($File)), (ConvertTo-HHMMSS $trimTime), ([System.IO.Path]::GetExtension($File))
    Write-Host "Trimming $File to $trimTime seconds ($(ConvertTo-HHMMSS $trimTime))..."
    $ffmpegArgs = "-i `"$File`" -ss 0 -to $trimTime -c:v copy -c:a copy `"$newName`""
    Write-Host "ffmpeg $ffmpegArgs"
    & ffmpeg -y -hide_banner -loglevel error -i "$File" -ss 0 -to $trimTime -c:v copy -c:a copy "$newName"
    Write-Host "Trimmed file saved as $newName"
}


if (-not $Path) {
    Show-Usage
    exit 1
}

if (Test-Path $Path -PathType Leaf) {
    Invoke-RecVideoFile $Path
} elseif (Test-Path $Path -PathType Container) {
    $files = Get-ChildItem -Path $Path -File | Where-Object { $allowedExt -contains $_.Extension.ToLower() }
    if ($files.Count -eq 0) {
        Write-Host "No supported video files found in folder: $Path"
        exit 1
    }
    foreach ($f in $files) {
        Invoke-RecVideoFile $f.FullName
    }
} else {
    Write-Host "File or folder not found: $Path"
    exit 1
}
