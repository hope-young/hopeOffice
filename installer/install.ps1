# installer/install.ps1
#
# Install hope-Office into the current user's Office add-in
# manifest directory so Word/Excel/PowerPoint can sideload it
# from a fixed location.
#
# What this does:
#   1. Copies manifest.xml (and any sibling resource files) into
#      %LOCALAPPDATA%\Microsoft\Office\16.0\WEF\
#   2. (Optional) adds that directory to the Trust Center
#      "Trusted Add-in Catalogs" so Office actually loads the add-in.
#      That step requires modifying the registry and may need
#      Office to be closed first.
#
# Usage:
#   pwsh installer/install.ps1
#   pwsh installer/install.ps1 -TrustCatalog   # also trusts the dir
#   pwsh installer/install.ps1 -Uninstall
#
# Requirements:
#   - Windows + PowerShell 5+ (PowerShell 7+ recommended)
#   - manifest.xml at the project root
#   - The dev certs already installed via
#     `npx office-addin-dev-certs install` (HTTPS dev server)
#   - Or, for the production variant, a release build in dist/ and
#     a path the manifest can resolve.

[CmdletBinding()]
param(
    [switch]$Uninstall,
    [switch]$TrustCatalog,
    [switch]$DryRun,
    [switch]$Production,
    [string]$SourceDir = (Split-Path -Parent $PSScriptRoot),
    [string]$OfficeVersion = "16.0"
)

$ErrorActionPreference = "Stop"

# --- Constants ---------------------------------------------------------

# The per-user manifest directory. Office reads every *.xml in here
# on launch and surfaces matching ones under My Add-ins.
$wefDir = Join-Path $env:LOCALAPPDATA "Microsoft\Office\$OfficeVersion/WEF"
$addInName = "hopeOffice"
$targetDir = Join-Path $wefDir $addInName

# Registry path for the "Trusted Add-in Catalogs" group. The
# sub-key list contains the directories Office will trust without
# showing the security prompt.
$regKey = "HKCU:\Software\Microsoft\Office\$OfficeVersion\WEF\TrustedCatalogs\AllowList"

# --- Subcommands -------------------------------------------------------

if ($Uninstall) {
    Write-Host "Uninstalling $addInName from $wefDir..."
    if (Test-Path $targetDir) {
        if ($DryRun) {
            Write-Host "  [dry-run] would remove $targetDir"
        } else {
            Remove-Item -Recurse -Force $targetDir
            Write-Host "  removed $targetDir"
        }
    } else {
        Write-Host "  nothing to remove (directory absent)"
    }
    if (Test-Path $regKey) {
        $existing = Get-ItemProperty -Path $regKey -ErrorAction SilentlyContinue
        if ($existing) {
            if ($DryRun) {
                Write-Host "  [dry-run] would remove registry trust entry"
            } else {
                Remove-ItemProperty -Path $regKey -Name $addInName -ErrorAction SilentlyContinue
                Write-Host "  removed registry trust entry"
            }
        }
    }
    Write-Host "Done. Restart Word/Excel/PowerPoint for the change to take effect."
    exit 0
}

# --- Install -----------------------------------------------------------

if (-not (Test-Path $SourceDir)) {
    throw "Source directory not found: $SourceDir"
}
$manifestPath = Join-Path $SourceDir "manifest.xml"
if (-not (Test-Path $manifestPath)) {
    throw "manifest.xml not found at $manifestPath. Run from the project root, or pass -SourceDir."
}

# Make sure the WEF directory exists.
if (-not (Test-Path $wefDir)) {
    New-Item -ItemType Directory -Path $wefDir -Force | Out-Null
    Write-Host "Created $wefDir"
}

# Clean previous install (if any) so we don't leave stale files.
if (Test-Path $targetDir) {
    Write-Host "Removing previous install at $targetDir"
    if (-not $DryRun) {
        Remove-Item -Recurse -Force $targetDir
    }
}
if ($DryRun) {
    Write-Host "[dry-run] would create $targetDir"
} else {
    New-Item -ItemType Directory -Path $targetDir -Force | Out-Null
}

# Copy manifest.xml + any sibling resources. The manifest
# references files under SourceUrl (e.g.
# `https://localhost:3721/src/taskpane/index.html` in dev, or
# `/dist/taskpane/index.html` once we've moved SourceUrl to
# the built assets in a future release). Two install modes:
#
#   Default (dev):  copy `public/` + `src/`. Pair with
#                   `npm run dev`; dev certs must be installed
#                   via `npx office-addin-dev-certs install`.
#   -Production:    copy `public/` + `dist/`. Requires a prior
#                   `npm run build`. The user is responsible
#                   for serving `dist/` over HTTPS somewhere
#                   the manifest's SourceUrl can reach.
$manifestFiles = @(Get-ChildItem -Path $SourceDir -Filter "manifest.xml")
if ($DryRun) {
    Write-Host "[dry-run] would copy manifest.xml"
} else {
    Copy-Item -Path $manifestFiles -Destination $targetDir
}
$subs = if ($Production) { @("public", "dist") } else { @("public", "src") }
foreach ($sub in $subs) {
    $src = Join-Path $SourceDir $sub
    if (Test-Path $src) {
        if ($DryRun) {
            Write-Host "[dry-run] would copy -Recurse $src -> $targetDir\$sub"
        } else {
            Copy-Item -Recurse -Path $src -Destination (Join-Path $targetDir $sub)
        }
    } elseif ($Production -and $sub -eq "dist") {
        throw "Production install requested but `$SourceDir\dist` does not exist. Run 'npm run build' first."
    }
}
if ($DryRun) {
    Write-Host "[dry-run] would report install at $targetDir"
} else {
    Write-Host "Installed manifest + resources to $targetDir"
}

# --- Optional: register as a trusted catalog -----------------------
#
# The Trust Center's "Trusted Add-in Catalogs" list contains
# directories Office will load *without* showing the developer
# prompt. We append our install dir so production users get
# the add-in without a "do you trust this?" click.
if ($TrustCatalog) {
    if (-not (Test-Path $regKey)) {
        if ($DryRun) {
            Write-Host "[dry-run] would create $regKey"
        } else {
            New-Item -Path $regKey -Force | Out-Null
        }
    }
    if ($DryRun) {
        Write-Host "[dry-run] would set $regKey : $addInName = $targetDir"
    } else {
        Set-ItemProperty -Path $regKey -Name $addInName -Value $targetDir
        Write-Host "Added $targetDir to the Trusted Add-in Catalogs"
        Write-Host "  (registry: $regKey)"
    }
}

Write-Host ""
Write-Host "Done. Restart Word/Excel/PowerPoint for the change to take effect."
Write-Host "Sideload the add-in via Insert -> My Add-ins -> $addInName (or via the"
Write-Host ('manifest URL file:///' + $targetDir + '/manifest.xml if your Office build supports it).')
