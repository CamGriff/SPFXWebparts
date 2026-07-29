<#
.SYNOPSIS
    Sync a standalone SPFx project repo with its subtree folder in this monorepo.

.DESCRIPTION
    Assumes the standalone project repo lives as a sibling directory to this
    monorepo clone (e.g. ..\messageBanner next to this repo), and that the
    subtree folder name inside this repo matches that sibling directory name.

.PARAMETER Action
    "pull" brings changes from the standalone repo into this repo's subfolder.
    "push" sends changes from this repo's subfolder back out to the standalone repo.

.PARAMETER Name
    The project folder/prefix name, e.g. "messageBanner".

.PARAMETER Branch
    Branch to use on the standalone repo side. Defaults to "master".

.EXAMPLE
    .\scripts\subtree-sync.ps1 pull messageBanner
    .\scripts\subtree-sync.ps1 push scrolltotop-extension
#>
param(
    [Parameter(Mandatory = $true)]
    [ValidateSet("pull", "push")]
    [string]$Action,

    [Parameter(Mandatory = $true)]
    [string]$Name,

    [string]$Branch = "master"
)

$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot
$siblingPath = Join-Path (Split-Path -Parent $repoRoot) $Name

if (-not (Test-Path (Join-Path $siblingPath ".git"))) {
    Write-Error "No git repo found at '$siblingPath'. Expected the standalone '$Name' repo as a sibling of this monorepo clone."
    exit 1
}

Push-Location $repoRoot
try {
    if ($Action -eq "pull") {
        git -c core.longpaths=true subtree pull --prefix=$Name $siblingPath $Branch -m "Sync $Name from standalone repo"
    }
    else {
        git -c core.longpaths=true subtree push --prefix=$Name $siblingPath $Branch
    }
}
finally {
    Pop-Location
}
