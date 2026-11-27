<#
.SYNOPSIS
    Copy the main checkout's gitignored .env files into a git worktree.

.DESCRIPTION
    A worktree gets only tracked files. This copies every .env and .env.* file
    (but not *.example templates) from the main checkout's root and from each
    *.src export folder up to two levels down, into the same place in this
    worktree. Values that point into the main checkout are rewritten to point
    into the worktree, so ACCESS_VCS_DATABASE and similar paths name the
    worktree's own files. Existing files in the worktree are left alone unless
    -Force is given.

    With -CopyDatabase it also copies each project's built database and its
    vcs-index.idx, so the first export in the worktree is incremental, plus the
    files the database links to by a relative (rel:) path: Access, Excel, and
    text back ends, and VBA library references. A file that a .env value names
    inside the main checkout is copied to the rewritten path too. Projects are
    found by their export folders (a vcs-options.json beside a vbe-project.json,
    up to three levels down). A project whose ExportFolder option is absolute is
    skipped, because a worktree copy would export into the main checkout; so is
    a rel: path that leads outside the main checkout.

    A database that is open (a .laccdb or .ldb lock file beside it) is skipped
    unless -Force is also given, and even then the copy may be inconsistent.

    Cursor runs this from .cursor/worktrees.json with ROOT_WORKTREE_PATH set.
    For a worktree you created yourself, set ROOT_WORKTREE_PATH to the main
    checkout and run it from anywhere.

    It does not build anything: a build needs Access, and vcs_rebuild_addin
    installs machine-wide. See docs/worktrees.md.
#>
[CmdletBinding()]
param(
    [string]$MainCheckout = $env:ROOT_WORKTREE_PATH,
    [switch]$CopyDatabase,
    [switch]$Force
)

$ErrorActionPreference = 'Stop'

Add-Type -Namespace SetupWorktree -Name Native -MemberDefinition @'
[DllImport("kernel32.dll", CharSet = CharSet.Unicode)]
public static extern uint GetLongPathName(string shortPath, System.Text.StringBuilder longPath, uint size);
'@

# File listings return long names, so an 8.3 short path (C:\Users\EXAMPL~1) must not leak in.
function Get-LongPath([string]$path) {
    $full = (Resolve-Path -LiteralPath $path).Path
    $buffer = New-Object System.Text.StringBuilder 32768
    if ([SetupWorktree.Native]::GetLongPathName($full, $buffer, $buffer.Capacity) -gt 0) { $full = $buffer.ToString() }
    return $full.TrimEnd('\')
}

$worktree = Get-LongPath (Join-Path $PSScriptRoot '..')
if (-not $MainCheckout) {
    throw 'Set ROOT_WORKTREE_PATH (or pass -MainCheckout) to the main checkout folder.'
}
$main = Get-LongPath $MainCheckout
$mainGiven = (Resolve-Path -LiteralPath $MainCheckout).Path.TrimEnd('\')
if ($main -ieq $worktree) {
    Write-Host 'This is the main checkout; nothing to copy.'
    return
}

# The main checkout's path, long or as given, but not a sibling that starts with the
# same text (C:\Repos\project must not match C:\Repos\project-839).
$mainForms = @($main, $mainGiven) | Select-Object -Unique | ForEach-Object { [regex]::Escape($_) }
$mainPattern = '(?i)(?:' + ($mainForms -join '|') + ')(?=[\\/";''\s]|$)'

function Get-EnvFiles([string]$root) {
    $relativeFolders = @('')
    foreach ($dir in Get-ChildItem -LiteralPath $root -Directory -Force -ErrorAction SilentlyContinue) {
        if ($dir.Name -like '*.src') { $relativeFolders += $dir.Name }
        foreach ($sub in Get-ChildItem -LiteralPath $dir.FullName -Directory -Force -ErrorAction SilentlyContinue) {
            if ($sub.Name -like '*.src') { $relativeFolders += (Join-Path $dir.Name $sub.Name) }
        }
    }
    foreach ($relative in $relativeFolders) {
        $folder = if ($relative) { Join-Path $root $relative } else { $root }
        Get-ChildItem -LiteralPath $folder -File -Force -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -match '^\.env(\..+)?$' -and $_.Name -notlike '*.example' } |
            ForEach-Object {
                [pscustomobject]@{
                    File     = $_
                    Relative = if ($relative) { Join-Path $relative $_.Name } else { $_.Name }
                }
            }
    }
}

function Copy-EnvFile([System.IO.FileInfo]$source, [string]$target) {
    $bytes = [System.IO.File]::ReadAllBytes($source.FullName)
    $hasBom = $bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF
    $text = [System.Text.Encoding]::UTF8.GetString($bytes, [int]$hasBom * 3, $bytes.Length - [int]$hasBom * 3)

    $rewritten = 0
    $lines = $text -split '(?<=\n)'
    for ($i = 0; $i -lt $lines.Count; $i++) {
        $line = $lines[$i]
        if ($line -match '^\s*#' -or $line -notmatch '=') { continue }
        $updated = [regex]::Replace($line, $mainPattern, { param($m) $worktree })
        if ($updated -ne $line) {
            $lines[$i] = $updated
            $rewritten++
        }
    }

    $folder = Split-Path $target -Parent
    if (-not (Test-Path -LiteralPath $folder)) { New-Item -ItemType Directory -Path $folder | Out-Null }
    $encoding = New-Object System.Text.UTF8Encoding($hasBom)
    [System.IO.File]::WriteAllText($target, ($lines -join ''), $encoding)
    return $rewritten
}

$copied = 0
foreach ($entry in Get-EnvFiles $main) {
    $relative = $entry.Relative
    $target = Join-Path $worktree $relative
    if ((Test-Path -LiteralPath $target) -and -not $Force) {
        Write-Host "Kept existing $relative"
        continue
    }
    $rewritten = Copy-EnvFile $entry.File $target
    $copied++
    if ($rewritten) {
        Write-Host "Copied $relative ($rewritten value(s) now point into this worktree)"
    } else {
        Write-Host "Copied $relative"
    }
}

if ($copied -eq 0) {
    Write-Host 'No .env files to copy.'
}

if (-not $CopyDatabase) {
    Write-Host 'Not copied: built databases (build them from source, or rerun with -CopyDatabase) and the back-end or library files they link to.'
    return
}

function Get-RelativeToMain([string]$path) {
    $full = [System.IO.Path]::GetFullPath($path).TrimEnd('\')
    if (Test-Path -LiteralPath $full) { $full = Get-LongPath $full }
    if ($full.StartsWith($main + '\', [System.StringComparison]::OrdinalIgnoreCase)) {
        return $full.Substring($main.Length + 1)
    }
    return $null
}

function Read-Json([string]$path) {
    try { return Get-Content -LiteralPath $path -Raw -Encoding UTF8 | ConvertFrom-Json }
    catch { return $null }
}

# Copy one file from the main checkout to the same place in the worktree.
function Copy-FromMain([string]$source, [string]$what) {
    $relative = Get-RelativeToMain $source
    if (-not $relative) {
        Write-Warning "Skipped $what outside the main checkout: $source"
        return
    }
    if (-not (Test-Path -LiteralPath $source -PathType Leaf)) { return }
    $target = Join-Path $worktree $relative
    if ((Test-Path -LiteralPath $target) -and -not $Force) {
        Write-Host "Kept existing $relative"
        return
    }
    $folder = Split-Path $target -Parent
    if (-not (Test-Path -LiteralPath $folder)) { New-Item -ItemType Directory -Path $folder | Out-Null }
    Copy-Item -LiteralPath $source -Destination $target -Force
    Write-Host "Copied $relative ($what)"
}

function Test-DatabaseOpen([string]$path) {
    $base = [System.IO.Path]::ChangeExtension($path, $null).TrimEnd('.')
    return (Test-Path -LiteralPath "$base.laccdb") -or (Test-Path -LiteralPath "$base.ldb")
}

# A path the add-in stored relative to the database's folder, as rel:<path>.
function Resolve-RelPath([string]$value, [string]$dbFolder) {
    if ($value -notmatch '^rel:') { return $null }
    $relative = $value.Substring(4).TrimStart('\')
    if (-not $relative) { return $dbFolder }
    return [System.IO.Path]::GetFullPath((Join-Path $dbFolder $relative))
}

# Files a project links to by rel: paths, from its linked tables and references.
function Get-LinkedFiles([string]$src, [string]$dbFolder) {
    foreach ($file in Get-ChildItem -LiteralPath (Join-Path $src 'tbldefs') -Filter '*.json' -File -ErrorAction SilentlyContinue) {
        $items = (Read-Json $file.FullName).Items
        if (-not $items -or -not $items.Connect) { continue }
        if ($items.Connect -notmatch '(?i)(?:^|;)DATABASE=(rel:[^;]*)') { continue }
        $path = Resolve-RelPath $Matches[1] $dbFolder
        if (Test-Path -LiteralPath $path -PathType Container) {
            # Text links name the folder in the connection and the file in SourceTableName.
            if ($items.SourceTableName) {
                $name = $items.SourceTableName
                $candidate = Join-Path $path $name
                if (-not (Test-Path -LiteralPath $candidate)) { $candidate = Join-Path $path ($name -replace '#', '.') }
                [pscustomobject]@{ Path = $candidate; What = "linked by $($items.Name)" }
            }
        } else {
            [pscustomobject]@{ Path = $path; What = "linked by $($items.Name)" }
        }
    }
    $references = (Read-Json (Join-Path $src 'vbe-references.json')).Items
    if ($references) {
        foreach ($property in $references.PSObject.Properties) {
            $fullPath = $property.Value.FullPath
            if ($fullPath -match '^rel:') {
                [pscustomobject]@{ Path = (Resolve-RelPath $fullPath $dbFolder); What = "reference $($property.Name)" }
            }
        }
    }
}

$sourceFolders = Get-ChildItem -LiteralPath $main -Filter 'vbe-project.json' -File -Recurse -Depth 3 -Force -ErrorAction SilentlyContinue |
    Where-Object { $_.FullName -notmatch '\\(\.git|node_modules)\\' -and (Test-Path -LiteralPath (Join-Path $_.DirectoryName 'vcs-options.json')) } |
    ForEach-Object { $_.DirectoryName }

foreach ($src in $sourceFolders) {
    $srcRelative = Get-RelativeToMain $src
    $fileName = (Read-Json (Join-Path $src 'vbe-project.json')).Items.FileName
    if ($fileName -match '^rel:') { $fileName = $fileName.Substring(4) }
    $options = Read-Json (Join-Path $src 'vcs-options.json')
    $exportFolder = [string]$options.Options.ExportFolder
    if (-not $fileName) {
        Write-Warning "Skipped ${srcRelative}: vbe-project.json names no database file."
        continue
    }

    # Find the database from its export folder: blank is <database>.src beside it, and
    # a relative value (one leading \) is a folder path below the database's folder.
    if ($exportFolder -eq '') {
        $dbFolder = Split-Path $src -Parent
    } elseif ($exportFolder -match '^\\[^\\]') {
        $segments = @($exportFolder.Trim('\').Split('\') | Where-Object { $_ })
        if ($segments -contains '..') {
            Write-Warning "Skipped ${srcRelative}: its ExportFolder ($exportFolder) climbs above the database, so the database cannot be found from it. Use a blank ExportFolder or one below the database, such as \src\."
            continue
        }
        $dbFolder = $src
        foreach ($segment in $segments) { $dbFolder = Split-Path $dbFolder -Parent }
    } else {
        Write-Warning "Skipped ${srcRelative}: its ExportFolder ($exportFolder) is absolute, so a copy here would export into that folder, not this worktree. Use a blank ExportFolder or a relative one such as \src\."
        continue
    }

    $db = Join-Path $dbFolder $fileName
    if (-not (Test-Path -LiteralPath $db -PathType Leaf)) {
        Write-Host "No built database for $srcRelative (build it from source in the worktree)."
        continue
    }
    if (Test-DatabaseOpen $db) {
        if (-not $Force) {
            Write-Warning "Skipped $(Get-RelativeToMain $db): it is open (lock file beside it). Close it and rerun, or pass -Force to copy it anyway."
            continue
        }
        Write-Warning "Copying $(Get-RelativeToMain $db) while it is open; the copy may be inconsistent."
    }
    Copy-FromMain $db 'database'
    Copy-FromMain (Join-Path $src 'vcs-index.idx') 'change index'

    foreach ($linked in Get-LinkedFiles $src $dbFolder) {
        if (Test-DatabaseOpen $linked.Path) {
            Write-Warning "$(Split-Path $linked.Path -Leaf) is open; its copy may be inconsistent."
        }
        Copy-FromMain $linked.Path $linked.What
    }
}

# Files that .env values name inside the main checkout, now rewritten to the worktree.
$pathPattern = '(?i)(?:' + ($mainForms -join '|') + ')\\[^;"''\r\n]+'
foreach ($entry in Get-EnvFiles $main) {
    foreach ($line in Get-Content -LiteralPath $entry.File.FullName -Encoding UTF8) {
        if ($line -match '^\s*#' -or $line -notmatch '=') { continue }
        foreach ($m in [regex]::Matches($line, $pathPattern)) {
            $path = $m.Value.TrimEnd()
            if (Test-Path -LiteralPath $path -PathType Leaf) {
                Copy-FromMain $path "named in $($entry.Relative)"
            }
        }
    }
}
