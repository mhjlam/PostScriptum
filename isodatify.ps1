<#
    isodatify.ps1 - Rename files and folders to use ISO date format (YYYY-MM-DD) for any detected date in the name.

    - Supports a wide range of date formats, including:
        * Numeric: 2024-05-19, 19-05-2024, 19.05.2024, 20240519, 19/05/2024, etc.
        * Month names: 19 May 2024, May 19 2024, 19-May-2024, etc. (full or abbreviated, any capitalization)
        * Ignores day names (e.g. Monday, Tue, etc.) if present at the start
    - Only the date part is replaced; any prefix/suffix is preserved.
    - Works for both files and folders.
    - Recurses over all items in a directory if a directory is given.
    - Usage: isodatify.ps1 <file|folder|directory>

    Examples:
        Rename all files/folders in the current directory: 
        isodatify .

        Rename a single file:
        isodatify "19-May-2024_report.txt"

        Rename a folder with a date and prefix:
        isodatify "backup_Monday 19.05.2024"

        Handles formats like: 
        "20240519", "19 May 2024", "2024-05-19", "19.05.2024", "May 19 2024", etc.
#>

# Converts file/folder names with various date formats to ISO format (YYYY-MM-DD)
param(
    [Parameter(Position=0, Mandatory=$true)]
    [string]$Path
)

if (-not $Path) {
    Write-Host "Usage: isodatify.ps1 <file|folder|directory>"
    exit 1
}

# Month and day names for regex patterns (case-insensitive)
$monthNames = @(
    'january','february','march','april','may','june','july','august','september','october','november','december',
    'jan','feb','mar','apr','may','jun','jul','aug','sep','sept','oct','nov','dec')
$dayNames = @(
    'monday','tuesday','wednesday','thursday','friday','saturday','sunday',
    'mon','tue','tues','wed','thu','thur','thurs','fri','sat','sun')

# Patterns for various date formats, supporting day/month names and separators
$patterns = @(
    # ISO and European numeric formats
    @{Pattern = '^(?i)(?:(' + ($dayNames -join '|') + ')[ ,._-]*)?(\d{4})[-.](\d{2})[-.](\d{2})$'; Format = 'yyyy-MM-dd'},
    @{Pattern = '^(?i)(?:(' + ($dayNames -join '|') + ')[ ,._-]*)?(\d{2})[-.](\d{2})[-.](\d{4})$'; Format = 'dd-MM-yyyy'},
    @{Pattern = '^(?i)(?:(' + ($dayNames -join '|') + ')[ ,._-]*)?(\d{2})[.](\d{2})[.](\d{4})$'; Format = 'dd.MM.yyyy'},
    @{Pattern = '^(?i)(?:(' + ($dayNames -join '|') + ')[ ,._-]*)?(\d{4})[.](\d{2})[.](\d{2})$'; Format = 'yyyy.MM.dd'},
    @{Pattern = '^(?i)(?:(' + ($dayNames -join '|') + ')[ ,._-]*)?(\d{8})$'; Format = 'yyyyMMdd'},
    # European and US with slashes
    @{Pattern = '^(?i)(?:(' + ($dayNames -join '|') + ')[ ,._-]*)?(\d{2})/(\d{2})/(\d{4})$'; Format = 'dd/MM/yyyy'},
    @{Pattern = '^(?i)(?:(' + ($dayNames -join '|') + ')[ ,._-]*)?(\d{4})/(\d{2})/(\d{2})$'; Format = 'yyyy/MM/dd'},
    # US-style numeric (MM-DD-YYYY, MM/DD/YYYY, MM.DD.YYYY)
    @{Pattern = '^(?i)(?:(' + ($dayNames -join '|') + ')[ ,._-]*)?(\d{2})[-/.](\d{2})[-/.](\d{4})$'; Format = 'MM-dd-yyyy'},
    # Month name formats (with and without comma)
    @{Pattern = '^(?i)(?:(' + ($dayNames -join '|') + ')[ ,._-]*)?(\d{1,2})[ ._-](' + ($monthNames -join '|') + ')[ ._-](\d{4})'; Format = 'd-MMM-yyyy'},
    @{Pattern = '^(?i)(?:(' + ($dayNames -join '|') + ')[ ,._-]*)?(' + ($monthNames -join '|') + ')[ ._-](\d{1,2})[ ,._-](\d{4})'; Format = 'MMM-d-yyyy'},
    # Month name, day, comma, year (e.g. May 19, 2024)
    @{Pattern = '^(?i)(?:(' + ($dayNames -join '|') + ')[ ,._-]*)?(' + ($monthNames -join '|') + ')[ ._-]?(\d{1,2}),[ ._-]?(\d{4})'; Format = 'MMM-d,yyyy'},
    # Day, month name, comma, year (e.g. 19 May, 2024)
    @{Pattern = '^(?i)(?:(' + ($dayNames -join '|') + ')[ ,._-]*)?(\d{1,2})[ ._-](' + ($monthNames -join '|') + '),[ ._-]?(\d{4})'; Format = 'd-MMM,yyyy'}
)

# Converts a single file/folder name to ISO date format if a date is detected
function Convert-NameToISO {
    param($item)
    $name = $item.Name
    $dt = $null
    $match = $null
    foreach ($pat in $patterns) {
        # Try to match the pattern
        if ($name -match $pat.Pattern) {
            $match = $matches[0]
            $isValid = $true
            $y = $null; $m = $null; $d = $null
            # Remove day name if present (case-insensitive)
            if ($pat.Pattern -like '*dayNames*') {
                $name = $name -replace "^(?i)($($dayNames -join '|'))[ ,._-]*", ''
            }
            # Extract year, month, day from match groups
            if ($pat.Format -eq 'yyyyMMdd') {
                $y = [int]$matches[1].Substring(0,4)
                $m = [int]$matches[1].Substring(4,2)
                $d = [int]$matches[1].Substring(6,2)
            } elseif ($pat.Format -eq 'yyyy-MM-dd' -or $pat.Format -eq 'yyyy.MM.dd' -or $pat.Format -eq 'yyyy/MM/dd') {
                $y = [int]$matches[2]; $m = [int]$matches[3]; $d = [int]$matches[4]
            } elseif ($pat.Format -eq 'dd-MM-yyyy' -or $pat.Format -eq 'dd.MM.yyyy' -or $pat.Format -eq 'dd/MM/yyyy') {
                $d = [int]$matches[2]; $m = [int]$matches[3]; $y = [int]$matches[4]
            } elseif ($pat.Format -eq 'MM-dd-yyyy') {
                $m = [int]$matches[2]; $d = [int]$matches[3]; $y = [int]$matches[4]
            } elseif ($pat.Format -eq 'd-MMM-yyyy') {
                $d = [int]$matches[2]; $m = $matches[3]; $y = [int]$matches[4]
            } elseif ($pat.Format -eq 'MMM-d-yyyy') {
                $m = $matches[2]; $d = [int]$matches[3]; $y = [int]$matches[4]
            } elseif ($pat.Format -eq 'MMM-d,yyyy') {
                $m = $matches[2]; $d = [int]$matches[3]; $y = [int]$matches[4]
            } elseif ($pat.Format -eq 'd-MMM,yyyy') {
                $d = [int]$matches[2]; $m = $matches[3]; $y = [int]$matches[4]
            } else {
                $isValid = $false
            }
            # Convert month name to number if needed (case-insensitive)
            if ($m -is [string]) {
                $mLower = $m.ToLower()
                $mNum = [array]::IndexOf($monthNames, $mLower) % 12 + 1
                $m = $mNum
            }
            # Validate date parts
            if ($isValid) {
                if ($y -lt 1900 -or $y -gt 2100) { $isValid = $false }
                if ($m -lt 1 -or $m -gt 12) { $isValid = $false }
                if ($d -lt 1 -or $d -gt 31) { $isValid = $false }
            }
            # Try to parse and rename if valid
            if ($isValid) {
                try {
                    $dt = [datetime]::new($y, $m, $d)
                } catch {}
                if ($dt) { break }
            }
        }
    }
    # If a valid date was found, replace only the date part in the name
    if ($dt -and $match) {
        $newName = $name -replace [regex]::Escape($match), $dt.ToString('yyyy-MM-dd')
        if ($newName -ne $name) {
            Rename-Item -Path $item.FullName -NewName $newName
        }
    }
}

# Main logic: process a directory, file, or folder
if (Test-Path $Path -PathType Container) {
    # Directory: process all files and folders inside
    Get-ChildItem -Path $Path | ForEach-Object { Convert-NameToISO $_ }
} elseif (Test-Path $Path -PathType Leaf) {
    # Single file or folder
    Convert-NameToISO (Get-Item $Path)
} else {
    Write-Host "Path not found: $Path"
    exit 1
}
