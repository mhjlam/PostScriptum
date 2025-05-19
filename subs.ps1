param(
    [Parameter(Position=0, Mandatory=$false)]
    [string]$Path,
    [switch]$Recurse
)

# Expand cmd-style environment variables in $Path (e.g., %USERPROFILE%)
if ($Path -and $Path -match '%') {
    $Path = [System.Environment]::ExpandEnvironmentVariables($Path)
}

$videoExtensions = @("*.mp4", "*.mkv", "*.avi", "*.mov")

# Improved internet connection check: try ICMP ping, then DNS resolution
function Test-InternetConnection {
    # Try ICMP ping to a public DNS server
    if (Test-Connection -ComputerName 1.1.1.1 -Count 1 -Quiet -ErrorAction SilentlyContinue) {
        return $true
    }
    # If ICMP fails, try DNS resolution of a root DNS server
    try {
        [System.Net.Dns]::GetHostAddresses('a.root-servers.net') | Out-Null
        return $true
    } catch {
        return $false
    }
}

if (-not (Test-InternetConnection)) {
    Write-Host "No internet connection detected. Exiting script."
    exit 1
}

function Install-Pip {
    if (-not (Get-Command pip -ErrorAction SilentlyContinue)) {
        Write-Host "pip not found. Attempting to install pip..."
        $pipInstaller = Join-Path $env:TEMP "get-pip.py"
        try {
            Invoke-WebRequest -Uri "https://bootstrap.pypa.io/get-pip.py" -OutFile $pipInstaller
            python $pipInstaller
            Remove-Item $pipInstaller
            Write-Host "pip installed successfully."
        } catch {
            Write-Host "Failed to download or run get-pip.py. Please install pip manually."
            exit 1
        }
    }
}

function Install-Subliminal {
    if (-not (Get-Command subliminal -ErrorAction SilentlyContinue)) {
        Write-Host "Subliminal not found. Attempting to install via pip..."
        try {
            pip install subliminal
            Write-Host "Subliminal installed successfully."
        } catch {
            Write-Host "Failed to install subliminal via pip. Please install it manually."
            exit 1
        }
    }
}

function Invoke-SubtitlesForFile {
    param([string]$FilePath)
    $subtitlePath = [System.IO.Path]::ChangeExtension($FilePath, '.srt')
    if (Test-Path $subtitlePath) {
        Write-Host "Subtitle already exists for $FilePath. Skipping."
        return
    }
    Write-Host "Downloading subtitles for $FilePath..."
    subliminal download -l en "$FilePath"
}

function Invoke-SubtitlesForFolder {
    param([string]$FolderPath, [switch]$Recurse)
    $recurse = $Recurse.IsPresent
    if ($recurse) {
        $allFiles = Get-ChildItem -Path $FolderPath -File -Include $videoExtensions -Recurse | Select-Object -Unique -Property FullName
    } else {
        $allFiles = Get-ChildItem -Path (Join-Path $FolderPath '*') -File -Include $videoExtensions | Select-Object -Unique -Property FullName
    }
    if ($allFiles.Count -eq 0) {
        Write-Host "No video files found in $FolderPath"
        return
    }
    # Only include files that do not already have a .srt subtitle
    $filesWithoutSub = $allFiles | Where-Object { -not (Test-Path ([System.IO.Path]::ChangeExtension($_.FullName, '.srt'))) }
    $filesWithSub = $allFiles | Where-Object { Test-Path ([System.IO.Path]::ChangeExtension($_.FullName, '.srt')) }
    if ($filesWithSub.Count -gt 0) {
        Write-Host "The following video files are being ignored because they already have subtitles:"
        $filesWithSub | ForEach-Object { Write-Host $_.FullName }
    }
    if ($filesWithoutSub.Count -eq 0) {
        Write-Host "All video files already have subtitles."
        return
    }
    $batchSize = 5  # OpenSubtitles recommends 10 requests per 10 seconds for anonymous users (reduce to 5 for safety)
    $delaySeconds = 10
    for ($i = 0; $i -lt $filesWithoutSub.Count; $i += $batchSize) {
        $batch = $filesWithoutSub[$i..([math]::Min($i+$batchSize-1, $filesWithoutSub.Count-1))]
        $filePaths = $batch | ForEach-Object { '"' + $_.FullName + '"' }
        Write-Host "Downloading subtitles for $($batch.Count) files..."
        $argsString = $filePaths -join ' '
        $cmd = "subliminal download -l en $argsString"
        Invoke-Expression $cmd
        if ($i + $batchSize -lt $filesWithoutSub.Count) {
            Write-Host "Waiting $delaySeconds seconds before next batch to respect provider rate limits..."
            Start-Sleep -Seconds $delaySeconds
        }
    }
}

try {
    # Main script logic
    Install-Pip
    Install-Subliminal

    # Parse command line arguments
    if ($args.Count -gt 0 -and -not $Path) {
        $Path = $args[0]
    }

    if ($Path) {
        if (Test-Path $Path -PathType Leaf) {
            Invoke-SubtitlesForFile -FilePath $Path
            exit 0
        } elseif (Test-Path $Path -PathType Container) {
            Invoke-SubtitlesForFolder -FolderPath $Path -Recurse:$Recurse
            exit 0
        } else {
            Write-Host "Invalid path: $Path"
        }
    }

    Write-Host "Usage: subs.ps1 -Path <path to folder or video file> [-Recurse]"
    Write-Host "Supported video file extensions: $($videoExtensions -replace '\*\.', '.' -join ', ').`n"
    exit 1
} catch [System.Management.Automation.StopException] {
    Write-Host "Execution interrupted by user."
    exit 1
} catch {
    throw
}

