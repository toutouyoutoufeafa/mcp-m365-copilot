<#
.SYNOPSIS
    Désinstalle complètement le MCP M365 Copilot Client

.DESCRIPTION
    Supprime :
    - Processus MCP et tunnel Cloudflare
    - Répertoire d'installation
    - Fichier .env
    - Cache npm/node_modules (optionnel)

.PARAMETER InstallDir
    Répertoire d'installation (défaut: $env:USERPROFILE\mcp-m365-copilot)

.PARAMETER Force
    Ne demande pas de confirmation

.PARAMETER KeepNodeModules
    Garde node_modules (pour réinstallation rapide)
#>

param(
    [string]$InstallDir = "$env:USERPROFILE\mcp-m365-copilot",
    [switch]$Force,
    [switch]$KeepNodeModules
)

$ErrorActionPreference = "Continue"

function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Write-Host "[$timestamp] [$Level] $Message"
}

Write-Log "=== Désinstallation MCP M365 Copilot Client ==="
Write-Log "Répertoire: $InstallDir"

if (-not $Force) {
    $confirm = Read-Host "Confirmer la suppression de $InstallDir ? (o/N)"
    if ($confirm -ne "o" -and $confirm -ne "O") {
        Write-Log "Annulé"
        exit 0
    }
}

# Arrêter processus MCP
Write-Log "Arrêt des processus MCP..."
Get-Process -Name "node" -ErrorAction SilentlyContinue | Where-Object { $_.Path -like "*mcp-m365-copilot*" } | Stop-Process -Force -ErrorAction SilentlyContinue
Get-Process -Name "cloudflared" -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue

# Supprimer répertoire
if (Test-Path $InstallDir) {
    if (-not $KeepNodeModules) {
        Write-Log "Suppression complète: $InstallDir"
        Remove-Item -Path $InstallDir -Recurse -Force -ErrorAction SilentlyContinue
    } else {
        Write-Log "Suppression (sauf node_modules): $InstallDir"
        Get-ChildItem $InstallDir -Exclude "node_modules" | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
    }
    Write-Log "Répertoire supprimé"
} else {
    Write-Log "Répertoire déjà absent"
}

# Nettoyer cache npm (optionnel)
# Write-Log "Nettoyage cache npm..."
# npm cache clean --force -ErrorAction SilentlyContinue

Write-Log "=== Désinstallation terminée ==="