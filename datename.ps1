<#
    datename.ps1 - Rename file(s) to their creation or modified date (YYYY-MM-DD_HH-mm-ss), preserving extension.

    Usage:
        datename.ps1 <file|folder> [-Recurse] [-Modified]
    - If <file>: renames the file.
    - If <folder>: renames all files in the folder (optionally recursively).
    - -Recurse: process subfolders.
    - -Modified: use last modified date (default: creation date).

    Examples:
        datename.ps1 "myfile.txt"
        datename.ps1 "C:\MyFolder" -Recurse -Modified
#>

param(
    [Parameter(Position=0, Mandatory=$true)]
    [string]$Path,
    [switch]$Recurse,
    [switch]$Modified
)

function Get-DateString {
    param($item, $useModified)
    $dt = if ($useModified) { $item.LastWriteTime } else { $item.CreationTime }
    return $dt.ToString('yyyy-MM-dd_HH-mm-ss')
}

function Rename-ToDateName {
    param($item, $useModified)
    $dateStr = Get-DateString $item $useModified
    $ext = [System.IO.Path]::GetExtension($item.Name)
    $newName = "$dateStr$ext"
    if ($item.Name -eq $newName) { return }
    $newPath = Join-Path $item.DirectoryName $newName
    if (Test-Path $newPath) {
        Write-Host "Target exists: $newPath. Skipping $($item.FullName)"
        return
    }
    Write-Host "Renaming: $($item.Name) -> $newName"
    Rename-Item -Path $item.FullName -NewName $newName
}

function Show-Usage {
    Write-Host @"
Usage: datename.ps1 <file|folder> [-Recurse] [-Modified]
    <file|folder>   File or folder to process
    -Recurse       Process subfolders (if folder)
    -Modified      Use last modified date (default: creation date)
"@
}

if ($PSBoundParameters.Count -eq 0 -or -not $Path) {
    Show-Usage
    exit 1
}

if (-not (Test-Path $Path)) {
    Show-Usage
    Write-Host "Path not found: $Path"
    exit 1
}

if (Test-Path $Path -PathType Leaf) {
    Rename-ToDateName (Get-Item $Path) $Modified
} elseif (Test-Path $Path -PathType Container) {
    $opts = @{ Path = $Path; File = $true }
    if ($Recurse) { $opts['Recurse'] = $true }
    Get-ChildItem @opts | ForEach-Object { Rename-ToDateName $_ $Modified }
} else {
    Show-Usage
    Write-Host "Invalid path: $Path"
    exit 1
}
