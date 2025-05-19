<#
    fetch.ps1 - Recursively or non-recursively hard-resets, cleans, and pulls all git repos in a folder.
    Skips submodules at the top level. Handles both absolute and relative paths.

    Usage examples:
        fetch C:\Projects
        fetch . -Recurse
        fetch myfolder -Recurse
#>

# Parse command-line arguments and switches
param(
    [Parameter(Position=0, Mandatory=$false)]
    [string]$Path,
    [switch]$Recurse
)

# Prints usage instructions for the script
function Show-Usage {
    Write-Host "Usage: fetch <folder> [-Recurse]"
    Write-Host "  <folder>: Root folder to scan for git repos."
    Write-Host "  -Recurse : Scan all subfolders recursively."
}

# Prints the repo path and its remote URL in color
function Show-RepoPathAndRemote($repoPath, $root, $Path) {
    # Determine how to display the repo path: relative if possible, otherwise absolute
    $cwd = (Get-Location).ProviderPath
    $printPath = $repoPath
    if ($root -eq $cwd -or $root -eq '.') {
        $printPath = [System.IO.Path]::GetRelativePath($cwd, $repoPath)
    } elseif (-not ([System.IO.Path]::IsPathRooted($Path))) {
        # If input is a relative path (but not '.'), print the absolute path
        $printPath = (Resolve-Path $repoPath).Path
    }
    # Set up ANSI color codes for output
    $esc = [char]27
    $cyan = "${esc}[1;36m"
    $yellow = "${esc}[33m"
    $reset = "${esc}[0m"

    # Get the remote URL for the repo (origin)
    $remoteUrl = git -C $repoPath remote get-url origin 2>$null

    # Print the formatted repo path and remote URL
    Write-Host ("$cyan$printPath$reset ($yellow$remoteUrl$reset)")
}

# Shows the result of git reset --hard, formatting the HEAD message
function Show-ResetAndHead($repoPath) {
    # Run git reset --hard and capture output
    $resetOutput = git reset --hard 2>&1
    if ($resetOutput) {
        foreach ($line in $resetOutput) {
            # If output matches HEAD is now at <hash> <desc>, format it nicely
            if ($line -match '^HEAD is now at ([a-f0-9]+)(.*)$') {
                $hash = $matches[1]
                $rest = $matches[2].Trim()
                if ($rest) {
                    Write-Host ("HEAD is at ${hash}: $rest")
                } else {
                    Write-Host ("HEAD is at ${hash}:")
                }
            }

            # Print any other output lines as-is
            else {
                Write-Host $line
            }
        }
    }
}

# Visualizes the number of changes pulled as a colored bar
function Show-ChangesBar($additions, $deletions) {
    # Set up ANSI color codes for green (additions) and red (deletions)
    $esc = [char]27
    $green = "${esc}[32m"
    $red = "${esc}[31m"
    $reset = "${esc}[0m"
    $maxSymbols = 50
    $total = $additions + $deletions

    # Calculate how many + and - symbols to show, scaled to $maxSymbols
    if ($total -gt 0) {
        $plusCount = [Math]::Round($additions * $maxSymbols / $total)
        $minusCount = $maxSymbols - $plusCount
        if ($total -le $maxSymbols) {
            $plusCount = $additions
            $minusCount = $deletions
            $bar = $green + ('+' * $plusCount) + $reset + $red + ('-' * $minusCount) + $reset
        } else {
            $bar = $green + ('+' * $plusCount) + $reset + $red + ('-' * $minusCount) + $reset + '>'
        }

        # Print the summary bar with counts
        Write-Host ("Changes: $additions $bar $deletions")
    }
}

# For a single repo: reset, clean, pull, show changes, and update submodules
function Invoke-PullAndReport($repoPath, $root, $Path, $defaultBranch, $hasMain, $hasMaster) {
    # Print repo info and move into repo directory
    Show-RepoPathAndRemote $repoPath $root $Path
    Push-Location $repoPath

    # Hard reset and clean working tree
    Show-ResetAndHead $repoPath
    git clean -fdx
    $changesBefore = git rev-parse HEAD

    # Try to pull from the default, main, or master branch (in that order)
    if ($defaultBranch) {
        Write-Host "Resetting local changes and pulling origin/$defaultBranch..."
        git reset --hard "origin/$defaultBranch" >$null 2>&1
        $pullOutput = git pull origin $defaultBranch --quiet 2>&1
    } elseif ($hasMain) {
        Write-Host "Resetting local changes and pulling origin/main..."
        git reset --hard origin/main >$null 2>&1
        $pullOutput = git pull origin main --quiet 2>&1
    } elseif ($hasMaster) {
        Write-Host "Resetting local changes and pulling origin/master..."
        git reset --hard origin/master >$null 2>&1
        $pullOutput = git pull origin master --quiet 2>&1
    }
    $changesAfter = git rev-parse HEAD

    # If HEAD changed, show commit and diff summary
    if ($changesBefore -ne $changesAfter) {
        $commitsPulled = git rev-list --count $changesBefore..$changesAfter
        if ($commitsPulled -eq 0) {
            Write-Host ("HEAD changed from $changesBefore to $changesAfter (not a fast-forward or history rewritten)")
        } else {
            Write-Host ("$commitsPulled new commit(s) pulled.")
        }

        # Show a visual summary of code changes
        $diffNumStat = git diff --numstat $changesBefore $changesAfter
        $additions = 0; $deletions = 0
        foreach ($line in $diffNumStat) {
            $cols = $line -split '\t'
            if ($cols.Length -ge 2) {
                if ($cols[0] -match '^\d+$') { $additions += [int]$cols[0] }
                if ($cols[1] -match '^\d+$') { $deletions += [int]$cols[1] }
            }
        }
        if ($additions -gt 0 -or $deletions -gt 0) {
            Show-ChangesBar $additions $deletions
        }
    }

    # If nothing changed, print status
    elseif ($defaultBranch -or $hasMain -or $hasMaster) {
        if ($pullOutput) {
            Write-Host $pullOutput
        } else {
            Write-Host "Already up to date."
        }
    } else {
        Write-Host "No main, master, or default branch found in $repoPath."
    }

    # Always update submodules after pulling
    git submodule foreach --recursive "git reset --hard && git clean -fdx"
    git submodule update --init --recursive
    Pop-Location
    Write-Host ""
}

# Checks if a directory is a git repo (not a submodule), then pulls if appropriate
function Invoke-GitRepo($repoPath, $root, $Path, $suppressNotRepoMsg = $false) {
    # Check if .git exists, otherwise skip
    $gitDirPath = Join-Path $repoPath ".git"
    if (-not (Test-Path $gitDirPath)) {
        if (-not $suppressNotRepoMsg) {
            $repoPath = [System.IO.Path]::TrimEndingDirectorySeparator($repoPath)
            $repoPath = $repoPath + [System.IO.Path]::DirectorySeparatorChar
            Write-Host "$repoPath is not a git repository."
            Write-Host ""
        }
        return
    }

    # Skip submodules at the top level
    $gitDirContent = Get-Content $gitDirPath -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($gitDirContent -match '^gitdir:' -and $gitDirContent -match 'modules') {
        return  # skip submodules
    }

    # Get remote branches and default branch
    $branches = git -C $repoPath branch -r | ForEach-Object { $_.ToString().Trim() }
    $defaultBranch = git -C $repoPath symbolic-ref refs/remotes/origin/HEAD 2>$null
    if ($defaultBranch) {
        $defaultBranch = $defaultBranch -replace '^refs/remotes/origin/', ''
    }
    $hasMain = $branches | Where-Object { $_ -eq 'origin/main' }
    $hasMaster = $branches | Where-Object { $_ -eq 'origin/master' }
    
    # Only pull if there are remote branches
    if ($branches.Count -gt 0) {
        Invoke-PullAndReport $repoPath $root $Path $defaultBranch $hasMain $hasMaster
    } else {
        Write-Host "No remote branches found in $repoPath."
        Write-Host ""
    }
}

# Main entry: parse args, find repos, and process each
if ($args.Count -gt 2) {
    Show-Usage
    exit 1
}

if ($args.Count -ge 1 -and -not $Path) {
    $Path = $args[0]
}

if (-not $Path) {
    Show-Usage
    exit 1
}

$root = $Path
$originalCwd = Get-Location

try {
    if ($Recurse) {
        # Find all subdirectories that are git repos
        $allRepoDirs = Get-ChildItem -Path $root -Directory -Recurse | Where-Object {
            $gitDir = Join-Path $_.FullName ".git"
            Test-Path $gitDir
        }

        # Also include the root path itself if it is a git repo, while avoiding duplicates
        $rootRepoPath = (Resolve-Path $root).Path
        $allPaths = @($allRepoDirs | ForEach-Object { $_.FullName })
        if (-not ($allPaths -contains $rootRepoPath)) {
            $allPaths += $rootRepoPath
        }

        # If no repos found at all, show not-a-repo message for root
        if ($allPaths.Count -eq 0) {
            Invoke-GitRepo $rootRepoPath $root $Path $false
        } else {
            # Process each repo found, suppress not-a-repo message for root if others found
            foreach ($repoPath in $allPaths) {
                $suppress = ($repoPath -eq $rootRepoPath -and $allRepoDirs.Count -gt 0)
                Invoke-GitRepo $repoPath $root $Path $suppress
            }
        }
    } else {
        $repoPath = (Resolve-Path $root).Path
        Invoke-GitRepo $repoPath $root $Path $false
    }
}
finally {
    # Always restore the original working directory
    Set-Location $originalCwd
}
