<#
.SYNOPSIS
    Crée l'App Registration Azure AD pour le client avec les bonnes permissions

.DESCRIPTION
    Ce script utilise Microsoft Graph PowerShell pour créer :
    1. App Registration (Public client, redirect URI http://localhost:8400)
    2. Permissions déléguées (base + optionnelles Purview/Intune/Security/Autopilot/Defender/M365Admin/Exchange)
    3. Client Secret (valide 2 ans)
    4. Admin Consent (si exécuté par admin)

.REQUIREMENTS
    - Module Microsoft Graph PowerShell: Install-Module Microsoft.Graph -Scope CurrentUser
    - Connexion: Connect-MgGraph -Scopes "Application.ReadWrite.All", "AppRoleAssignment.ReadWrite.All"

.PARAMETER TenantId
    Tenant ID du client (où créer l'app)

.PARAMETER AppName
    Nom de l'application (défaut: "mcp-m365-copilot-client")

.PARAMETER EnablePurview
    Ajoute les scopes Purview (DataPolicies, Catalog, Accounts)

.PARAMETER EnableIntune
    Ajoute les scopes Intune (Devices, Configuration)

.PARAMETER EnableSecurity
    Ajoute les scopes Security (Events, Alerts)

.PARAMETER EnableAutopilot
    Ajoute les scopes Autopilot (Devices, ServiceConfig, Configuration)

.PARAMETER EnableDefender
    Ajoute les scopes Defender (Alerts, Devices, Vulnerabilities, ThreatIndicators)

.PARAMETER EnableM365Admin
    Ajoute les scopes M365 Admin (Organization, Domains, ServiceHealth, Reports, Directory, Roles, Policies)

.PARAMETER EnableExchange
    Ajoute les scopes Exchange (Mail, MailboxSettings, Calendars, Groups)

.PARAMETER OutputFile
    Fichier de sortie pour les credentials (défaut: .\app-registration-output.json)

.EXAMPLE
    Connect-MgGraph -Scopes "Application.ReadWrite.All"
    .\create-app-reg.ps1 -TenantId "xxx" -EnablePurview

.EXAMPLE
    .\create-app-reg.ps1 -TenantId "xxx" -EnablePurview -EnableIntune -EnableAutopilot -EnableDefender -EnableM365Admin -EnableExchange -OutputFile ".\client-creds.json"
#>

param(
    [Parameter(Mandatory=$true)]
    [string]$TenantId,

    [string]$AppName = "mcp-m365-copilot-client",

    [switch]$EnablePurview,

    [switch]$EnableIntune,

    [switch]$EnableSecurity,

    [switch]$EnableAutopilot,

    [switch]$EnableDefender,

    [switch]$EnableM365Admin,

    [switch]$EnableExchange,

    [string]$OutputFile = ".\app-registration-output.json"
)

$ErrorActionPreference = "Stop"

function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Write-Host "[$timestamp] [$Level] $Message"
}

# Vérifier connexion Graph
$context = Get-MgContext
if (-not $context) {
    throw "Non connecté à Microsoft Graph. Exécutez: Connect-MgGraph -Scopes 'Application.ReadWrite.All', 'AppRoleAssignment.ReadWrite.All'"
}
Write-Log "Connecté en tant que: $($context.Account)"

# Scopes de base (toujours)
$baseScopes = @(
    "User.Read",
    "Files.Read",
    "Team.ReadBasic.All",
    "Channel.ReadBasic.All"
)

# Scopes optionnels par catégorie
$optionalScopes = @{}

if ($EnablePurview) {
    $optionalScopes["Purview"] = @(
        "PurviewDataPolicies.Read.All",
        "PurviewCatalog.Read.All",
        "PurviewAccounts.Read.All"
    )
}
if ($EnableIntune) {
    $optionalScopes["Intune"] = @(
        "DeviceManagementManagedDevices.Read.All",
        "DeviceManagementConfiguration.Read.All"
    )
}
if ($EnableSecurity) {
    $optionalScopes["Security"] = @(
        "SecurityEvents.Read.All",
        "Alert.Read.All"
    )
}
if ($EnableAutopilot) {
    $optionalScopes["Autopilot"] = @(
        "DeviceManagementManagedDevices.Read.All",
        "DeviceManagementManagedDevices.ReadWrite.All",
        "DeviceManagementServiceConfig.Read.All",
        "DeviceManagementServiceConfig.ReadWrite.All",
        "DeviceManagementConfiguration.Read.All",
        "DeviceManagementConfiguration.ReadWrite.All"
    )
}
if ($EnableDefender) {
    $optionalScopes["Defender"] = @(
        "SecurityEvents.Read.All",
        "Alert.Read.All",
        "Alert.ReadWrite.All",
        "ThreatIndicators.Read.All",
        "ThreatIndicators.ReadWrite.OwnedBy",
        "Vulnerability.Read.All",
        "Machine.Read.All",
        "Machine.ReadWrite.All"
    )
}
if ($EnableM365Admin) {
    $optionalScopes["M365Admin"] = @(
        "Organization.Read.All",
        "Organization.ReadWrite.All",
        "Domain.Read.All",
        "Domain.ReadWrite.All",
        "ServiceHealth.Read.All",
        "Reports.Read.All",
        "Directory.Read.All",
        "Directory.ReadWrite.All",
        "Policy.Read.All",
        "Policy.ReadWrite.All",
        "RoleManagement.Read.All",
        "RoleManagement.ReadWrite.Directory"
    )
}
if ($EnableExchange) {
    $optionalScopes["Exchange"] = @(
        "Mail.Read",
        "Mail.ReadWrite",
        "Mail.Send",
        "MailboxSettings.Read",
        "MailboxSettings.ReadWrite",
        "Calendars.Read",
        "Calendars.ReadWrite",
        "Contacts.Read",
        "Contacts.ReadWrite",
        "Group.Read.All",
        "Group.ReadWrite.All"
    )
}

$allScopes = $baseScopes
foreach ($category in $optionalScopes.Keys) {
    $allScopes += $optionalScopes[$category]
}

Write-Log "Scopes à configurer:"
$baseScopes | ForEach-Object { Write-Log "  [BASE] $_" }
$optionalScopes.GetEnumerator() | ForEach-Object {
    $_.Value | ForEach-Object { Write-Log "  [$($_.Key)] $_" }
}

# Créer l'application
Write-Log "Création de l'App Registration: $AppName"
$app = New-MgApplication -DisplayName $AppName `
    -SignInAudience "AzureADMyOrg" `
    -PublicClient @{ RedirectUris = @("http://localhost:8400") } `
    -RequiredResourceAccess @()

$appId = $app.AppId
$objectId = $app.Id
Write-Log "Application créée: $appId (ObjectId: $objectId)"

# Attendre la réplication
Start-Sleep -Seconds 5

# Récupérer les Service Principals
Write-Log "Récupération des Service Principals..."

$graphSp = Get-MgServicePrincipal -Filter "appId eq '00000003-0000-0000-c000-000000000000'" -ErrorAction SilentlyContinue
if (-not $graphSp) {
    throw "Microsoft Graph Service Principal introuvable"
}

# Purview SP
$purviewSp = $null
if ($EnablePurview) {
    $purviewSp = Get-MgServicePrincipal -Filter "appId eq '05cb130f-1f67-4122-a582-3e16c3a29b7f'" -ErrorAction SilentlyContinue
    if (-not $purviewSp) {
        Write-Log "Purview Service Principal introuvable - tentative avec nom" "WARN"
        $purviewSp = Get-MgServicePrincipal -Filter "displayName eq 'Microsoft Purview'" -ErrorAction SilentlyContinue
    }
}

# Intune SP (same as Graph for device management)
$intuneSp = $graphSp

# Security/Defender SP
$securitySp = $null
if ($EnableSecurity -or $EnableDefender) {
    $securitySp = Get-MgServicePrincipal -Filter "appId eq '00000002-0000-0000-c000-000000000000'" -ErrorAction SilentlyContinue
}

# M365 Admin - uses Graph SP
$m365adminSp = $graphSp

# Exchange - uses Graph SP
$exchangeSp = $graphSp

# Construire RequiredResourceAccess
$requiredResourceAccess = @()

# Microsoft Graph (base + Intune + M365Admin + Exchange + Autopilot)
$graphAccess = @{
    resourceAppId = "00000003-0000-0000-c000-000000000000"
    resourceAccess = @()
}
$baseScopes | ForEach-Object {
    $scope = $graphSp.Oauth2PermissionScopes | Where-Object { $_.Value -eq $_ }
    if ($scope) {
        $graphAccess.resourceAccess += @{ id = $scope.Id; type = "Scope" }
    } else {
        Write-Log "Scope Graph non trouvé: $_" "WARN"
    }
}

# Add optional scopes that use Graph
$graphOptionalCategories = @("Intune", "M365Admin", "Exchange", "Autopilot")
foreach ($category in $graphOptionalCategories) {
    if ($optionalScopes.ContainsKey($category)) {
        $optionalScopes[$category] | ForEach-Object {
            $scope = $graphSp.Oauth2PermissionScopes | Where-Object { $_.Value -eq $_ }
            if ($scope) {
                $graphAccess.resourceAccess += @{ id = $scope.Id; type = "Scope" }
            } else {
                Write-Log "Scope Graph non trouvé: $_" "WARN"
            }
        }
    }
}
$requiredResourceAccess += $graphAccess

# Purview
if ($EnablePurview -and $purviewSp) {
    $purviewAccess = @{
        resourceAppId = $purviewSp.AppId
        resourceAccess = @()
    }
    $optionalScopes["Purview"] | ForEach-Object {
        $scope = $purviewSp.Oauth2PermissionScopes | Where-Object { $_.Value -eq $_ }
        if ($scope) {
            $purviewAccess.resourceAccess += @{ id = $scope.Id; type = "Scope" }
        } else {
            Write-Log "Scope Purview non trouvé: $_" "WARN"
        }
    }
    $requiredResourceAccess += $purviewAccess
} elseif ($EnablePurview) {
    Write-Log "Purview SP non trouvé - scopes Purview non ajoutés" "WARN"
}

# Security/Defender
if (($EnableSecurity -or $EnableDefender) -and $securitySp) {
    $securityAccess = @{
        resourceAppId = $securitySp.AppId
        resourceAccess = @()
    }
    if ($EnableSecurity) {
        $optionalScopes["Security"] | ForEach-Object {
            $scope = $securitySp.Oauth2PermissionScopes | Where-Object { $_.Value -eq $_ }
            if ($scope) {
                $securityAccess.resourceAccess += @{ id = $scope.Id; type = "Scope" }
            } else {
                Write-Log "Scope Security non trouvé: $_" "WARN"
            }
        }
    }
    if ($EnableDefender) {
        $optionalScopes["Defender"] | ForEach-Object {
            $scope = $securitySp.Oauth2PermissionScopes | Where-Object { $_.Value -eq $_ }
            if ($scope) {
                $securityAccess.resourceAccess += @{ id = $scope.Id; type = "Scope" }
            } else {
                Write-Log "Scope Defender non trouvé: $_" "WARN"
            }
        }
    }
    $requiredResourceAccess += $securityAccess
} elseif ($EnableSecurity -or $EnableDefender) {
    Write-Log "Security/Defender SP non trouvé - scopes non ajoutés" "WARN"
}

# Mettre à jour l'application avec les permissions
Write-Log "Configuration des permissions..."
Update-MgApplication -ApplicationId $objectId -RequiredResourceAccess $requiredResourceAccess
Write-Log "Permissions configurées"

# Créer le Client Secret (valide 2 ans)
Write-Log "Création du Client Secret..."
$secret = Add-MgApplicationPassword -ApplicationId $objectId `
    -PasswordCredential @{ DisplayName = "MCP Client Secret"; EndDateTime = (Get-Date).AddYears(2) }
$clientSecret = $secret.SecretText
Write-Log "Client Secret créé (expire: $($secret.EndDateTime))"

# Admin Consent
Write-Log "Attribution du consentement admin..."
$requiredResourceAccess | ForEach-Object {
    $resourceAppId = $_.resourceAppId
    $_.resourceAccess | ForEach-Object {
        try {
            New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $objectId `
                -ResourceId (Get-MgServicePrincipal -Filter "appId eq '$resourceAppId'").Id `
                -AppRoleId $_.id -PrincipalId $objectId -ErrorAction Stop | Out-Null
        } catch {
            # Ignorer erreurs (déjà assigné, etc.)
        }
    }
}
Write-Log "Consentement admin appliqué (si permissions supportées)"

# Service Principal pour l'app
$sp = Get-MgServicePrincipal -Filter "appId eq '$appId'" -ErrorAction SilentlyContinue
if (-not $sp) {
    Write-Log "Création du Service Principal..."
    $sp = New-MgServicePrincipal -AppId $appId
}

# Construire la sortie scopes
$scopeOutput = @{
    base = $baseScopes
}
foreach ($category in $optionalScopes.Keys) {
    $scopeOutput[$category.ToLower()] = $optionalScopes[$category]
}

# Sortie
$output = @{
    clientId = $appId
    tenantId = $TenantId
    clientSecret = $clientSecret
    secretExpires = $secret.EndDateTime
    appName = $AppName
    scopes = $scopeOutput
    redirectUri = "http://localhost:8400"
    createdAt = (Get-Date).ToString("o")
}

$output | ConvertTo-Json -Depth 5 | Set-Content -Path $OutputFile -Encoding UTF8
Write-Log "============================================"
Write-Log "APP REGISTRATION CRÉÉE AVEC SUCCÈS"
Write-Log "============================================"
Write-Log "Client ID: $appId"
Write-Log "Tenant ID: $TenantId"
Write-Log "Client Secret: $clientSecret"
Write-Log "Secret expire: $($secret.EndDateTime)"
Write-Log "Fichier de sortie: $OutputFile"
Write-Log "============================================"
Write-Log "AJOUTEZ CES VALEURS DANS config.template.env :"
Write-Log '  ENTRA_CLIENT_ID="' + $appId + '"'
Write-Log '  ENTRA_TENANT_ID="' + $TenantId + '"'
Write-Log "============================================"
Write-Log "IMPORTANT: Sauvegardez le Client Secret maintenant - il ne sera plus affiché !"