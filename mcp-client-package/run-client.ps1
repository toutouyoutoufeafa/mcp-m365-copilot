<#
.SYNOPSIS
    Installe, configure et lance le MCP M365 Copilot pour un client

.DESCRIPTION
    Ce script :
    1. Vérifie/installe Node.js (via winget)
    2. Clone le repo (si pas déjà fait)
    3. Installe les dépendances (npm ci)
    4. Build TypeScript (npm run build)
    5. Crée le fichier .env depuis config.template.env
    6. Lance le serveur MCP + tunnel Cloudflare

.PARAMETER ClientId
    Client ID de l'App Registration (fourni par create-app-reg.ps1)

.PARAMETER TenantId
    Tenant ID du client

.PARAMETER PurviewAccountName
    Nom du compte Purview (optionnel, active l'agent Purview)

.PARAMETER RepoUrl
    URL du repo Git privé (défaut: variable d'env MCP_REPO_URL)

.PARAMETER InstallDir
    Répertoire d'installation (défaut: $env:USERPROFILE\mcp-m365-copilot)

.PARAMETER SkipBuild
    Ignore le build (si déjà fait)

.PARAMETER NoTunnel
    Lance seulement le MCP sans tunnel Cloudflare

.EXAMPLE
    .\run-client.ps1 -ClientId "xxx" -TenantId "yyy"

.EXAMPLE
    .\run-client.ps1 -ClientId "xxx" -TenantId "yyy" -PurviewAccountName "mon-purview"
#>

param(
    [Parameter(Mandatory=$true)]
    [string]$ClientId,

    [Parameter(Mandatory=$true)]
    [string]$TenantId,

    [string]$PurviewAccountName = "",

    [string]$RepoUrl = $env:MCP_REPO_URL,

    [string]$InstallDir = "$env:USERPROFILE\mcp-m365-copilot",

    [switch]$SkipBuild,

    [switch]$NoTunnel
)

$ErrorActionPreference = "Stop"

function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Write-Host "[$timestamp] [$Level] $Message"
}

function Check-Command {
    param([string]$Command)
    try {
        Get-Command $Command -ErrorAction Stop | Out-Null
        return $true
    } catch {
        return $false
    }
}

function Install-NodeJS {
    if (Check-Command "node") {
        $version = node --version
        Write-Log "Node.js déjà installé: $version"
        return
    }
    
    Write-Log "Installation de Node.js via winget..."
    if (Check-Command "winget") {
        winget install --id OpenJS.NodeJS.LTS --source winget --accept-source-agreements --accept-package-agreements
    } elseif (Check-Command "choco") {
        choco install nodejs-lts -y
    } else {
        throw "Ni winget ni chocolatey trouvé. Installez Node.js manuellement: https://nodejs.org"
    }
    
    # Rafraîchir PATH
    $env:PATH = [System.Environment]::GetEnvironmentVariable("Path","Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path","User")
    Write-Log "Node.js installé: $(node --version)"
}

function Install-Cloudflared {
    $cloudflaredPath = "$InstallDir\cloudflared.exe"
    if (Test-Path $cloudflaredPath) {
        Write-Log "cloudflared déjà présent"
        return $cloudflaredPath
    }
    
    Write-Log "Téléchargement de cloudflared..."
    $url = "https://github.com/cloudflare/cloudflared/releases/latest/download/cloudflared-windows-amd64.exe"
    try {
        Invoke-WebRequest -Uri $url -OutFile $cloudflaredPath -UseBasicParsing
        Write-Log "cloudflared téléchargé: $cloudflaredPath"
    } catch {
        Write-Log "Échec téléchargement cloudflared: $_" "WARN"
        Write-Log "Téléchargez manuellement: https://github.com/cloudflare/cloudflared/releases" "WARN"
    }
    return $cloudflaredPath
}

function Setup-Repo {
    if (Test-Path "$InstallDir\package.json") {
        Write-Log "Repo déjà présent: $InstallDir"
        return
    }
    
    if (-not $RepoUrl) {
        throw "RepoUrl requis (paramètre ou variable MCP_REPO_URL). Exemple: https://github.com/tonorg/mcp-m365-copilot.git"
    }
    
    Write-Log "Clonage du repo: $RepoUrl"
    git clone $RepoUrl $InstallDir
    Write-Log "Repo cloné dans: $InstallDir"
}

function Create-EnvFile {
    $templatePath = "$InstallDir\config.template.env"
    $envPath = "$InstallDir\.env"
    
    if (-not (Test-Path $templatePath)) {
        throw "config.template.env introuvable dans $InstallDir"
    }
    
    $content = Get-Content $templatePath -Raw
    $content = $content -replace 'ENTRA_CLIENT_ID=".*"', "ENTRA_CLIENT_ID=`"$ClientId`""
    $content = $content -replace 'ENTRA_TENANT_ID=".*"', "ENTRA_TENANT_ID=`"$TenantId`""
    
    if ($PurviewAccountName) {
        $content = $content -replace '# PURVIEW_ACCOUNT_NAME=.*', "PURVIEW_ACCOUNT_NAME=`"$PurviewAccountName`""
        $content = $content -replace '# PURVIEW_SCOPES_ENABLED=.*', 'PURVIEW_SCOPES_ENABLED="true"'
    }
    
    Set-Content -Path $envPath -Value $content -Encoding UTF8
    Write-Log "Fichier .env créé: $envPath"
}

function Install-Dependencies {
    Write-Log "Installation des dépendances npm..."
    Set-Location $InstallDir
    npm ci
    Write-Log "Dépendances installées"
}

function Build-Project {
    if ($SkipBuild) {
        Write-Log "Build ignoré (--SkipBuild)"
        return
    }
    Write-Log "Build TypeScript..."
    Set-Location $InstallDir
    npm run build
    Write-Log "Build terminé"
}

function Start-MCP {
    Write-Log "Démarrage du serveur MCP sur le port $env:PORT (défaut 3000)..."
    Set-Location $InstallDir
    
    $env:ENTRA_CLIENT_ID = $ClientId
    $env:ENTRA_TENANT_ID = $TenantId
    if ($PurviewAccountName) { $env:PURVIEW_ACCOUNT_NAME = $PurviewAccountName }
    
    # Lancer le serveur MCP en arrière-plan
    $mcpProcess = Start-Process -FilePath "node" -ArgumentList "dist/index.js" -PassThru -WindowStyle Hidden
    Write-Log "MCP Server PID: $($mcpProcess.Id)"
    
    # Attendre que le serveur démarre
    Start-Sleep -Seconds 3
    
    return $mcpProcess
}

function Start-Tunnel {
    param($mcpProcess)
    
    if ($NoTunnel) {
        Write-Log "Tunnel désactivé (--NoTunnel). MCP accessible sur localhost:3000"
        return
    }
    
    $cloudflared = Install-Cloudflared
    if (-not (Test-Path $cloudflared)) {
        Write-Log "cloudflared non trouvé, tunnel ignoré" "WARN"
        return
    }
    
    Write-Log "Démarrage du tunnel Cloudflare..."
    $tunnelProcess = Start-Process -FilePath $cloudflared -ArgumentList "tunnel", "--url", "http://localhost:3000" -PassThru -WindowStyle Hidden
    Write-Log "Tunnel Cloudflare PID: $($tunnelProcess.Id)"
    
    # Attendre et récupérer l'URL
    Start-Sleep -Seconds 5
    
    # Essayer de récupérer l'URL depuis les logs
    try {
        $output = & $cloudflared tunnel --url http://localhost:3000 2>&1 | Select-String "trycloudflare.com"
        if ($output) {
            $url = $output.Matches.Value | Select-Object -First 1
            Write-Log "============================================"
            Write-Log "TUNNEL ACTIF: $url"
            Write-Log "============================================"
            Write-Log "Configurez ce Custom Connector dans Copilot Studio:"
            Write-Log "  URL: $url"
            Write-Log "  OpenAPI: $url/openapi.json"
            Write-Log "  Auth: OAuth2 OBO avec Client ID: $ClientId"
            Write-Log "============================================"
        }
    } catch {
        Write-Log "Tunnel démarré (URL visible dans la console cloudflared)"
    }
    
    return $tunnelProcess
}

function Wait-ForExit {
    param($mcpProcess, $tunnelProcess)
    
    Write-Log "Appuyez sur Ctrl+C pour arrêter..."
    try {
        while ($true) {
            Start-Sleep -Seconds 10
            if ($mcpProcess.HasExited) { throw "MCP process exited" }
            if ($tunnelProcess -and $tunnelProcess.HasExited) { throw "Tunnel process exited" }
        }
    } catch {
        Write-Log "Arrêt détecté..."
    } finally {
        if ($mcpProcess -and -not $mcpProcess.HasExited) { Stop-Process -Id $mcpProcess.Id -Force; Write-Log "MCP arrêté" }
        if ($tunnelProcess -and -not $tunnelProcess.HasExited) { Stop-Process -Id $tunnelProcess.Id -Force; Write-Log "Tunnel arrêté" }
    }
}

# ===== MAIN =====
Write-Log "=== MCP M365 Copilot Client Launcher ==="
Write-Log "Client ID: $ClientId"
Write-Log "Tenant ID: $TenantId"
if ($PurviewAccountName) { Write-Log "Purview Account: $PurviewAccountName" }

try {
    Install-NodeJS
    Setup-Repo
    Create-EnvFile
    Install-Dependencies
    Build-Project
    
    $mcp = Start-MCP
    $tunnel = $null
    if (-not $NoTunnel) { $tunnel = Start-Tunnel $mcp }
    
    Wait-ForExit $mcp $tunnel
} catch {
    Write-Log "ERREUR: $_" "ERROR"
    exit 1
}