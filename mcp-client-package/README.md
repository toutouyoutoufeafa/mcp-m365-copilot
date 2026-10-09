# MCP M365 Copilot - Package Client

Package prêt-à-déployer pour héberger le serveur MCP M365 Copilot **chez le client** (zéro coût, zéro infra MSP).

## 🎯 Architecture

```
┌─────────────────┐     HTTPS      ┌──────────────────┐     HTTPS      ┌─────────────────┐
│  Copilot Studio │◄──────────────►│  Client MCP      │◄──────────────►│  Microsoft 365  │
│  (MSP Tenant)   │  Custom Connector │  (Client Tenant)│  Graph API     │  (Client Data)  │
└─────────────────┘                 └──────────────────┘                └─────────────────┘
                                           │
                                    ┌──────┴──────┐
                                    │ cloudflared │ (gratuit, sortant seulement)
                                    └─────────────┘
```

## 📦 Contenu du package

```
mcp-client-package/
├── run-client.ps1          # Script principal (install + build + run + tunnel)
├── create-app-reg.ps1      # Crée l'App Registration Azure AD + permissions
├── config.template.env     # Template de configuration
├── uninstall.ps1           # Désinstallation propre
├── cloudflared/            # Binaire cloudflared (téléchargé auto si absent)
└── README.md               # Ce fichier
```

## 🚀 Déploiement en 3 étapes (10 min)

### Prérequis client
- Windows 10/11 ou Windows Server 2019+
- Accès admin local (pour installer Node.js via winget)
- Accès Azure AD (pour créer l'App Registration) - **Admin Global requis**
- Sortie Internet HTTPS (pour cloudflared tunnel)

---

### Étape 1 : Créer l'App Registration (Admin Global client)

```powershell
# 1. Installer Microsoft Graph PowerShell (si pas déjà fait)
Install-Module Microsoft.Graph -Scope CurrentUser

# 2. Se connecter au tenant CLIENT
Connect-MgGraph -Scopes "Application.ReadWrite.All", "AppRoleAssignment.ReadWrite.All"

# 3. Lancer le script (adapte les flags selon besoins)
.\create-app-reg.ps1 `
    -TenantId "TENANT_ID_DU_CLIENT" `
    -EnablePurview `
    -EnableIntune `
    -OutputFile ".\client-creds.json"
```

**Sortie** : `client-creds.json` avec `clientId`, `tenantId`, `clientSecret`

> ⚠️ **Sauvegardez le `clientSecret` immédiatement** - il n'est affiché qu'une fois.

---

### Étape 2 : Configurer le .env

```powershell
# Copier le template
Copy-Item config.template.env .env

# Éditer .env avec les valeurs de client-creds.json
notepad .env
```

Contenu minimal `.env` :
```env
ENTRA_CLIENT_ID="xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
ENTRA_TENANT_ID="xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"

# Optionnel - Purview
# PURVIEW_ACCOUNT_NAME="mon-purview-account"
```

---

### Étape 3 : Lancer le client

```powershell
# Avec les valeurs du client-creds.json
.\run-client.ps1 `
    -ClientId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" `
    -TenantId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" `
    -PurviewAccountName "mon-purview-account"  # Optionnel
```

**Résultat** :
```
============================================
TUNNEL ACTIF: https://abc123.trycloudflare.com
============================================
Configurez ce Custom Connector dans Copilot Studio:
  URL: https://abc123.trycloudflare.com
  OpenAPI: https://abc123.trycloudflare.com/openapi.json
  Auth: OAuth2 OBO avec Client ID: xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx
============================================
```

---

## 🔧 Configuration Copilot Studio (Côté MSP)

1. Aller sur **https://copilotstudio.microsoft.com** (tenant MSP)
2. **Create** → **Agent** → **New agent**
3. **Configure** → **Add knowledge** → **Custom connector**
4. **New custom connector** :
   - **OpenAPI spec URL** : `https://abc123.trycloudflare.com/openapi.json`
   - **Authentication** : **OAuth 2.0** → **On-Behalf-Of (OBO)**
     - Client ID : `xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx` (du client)
     - Client Secret : `yyyyyyyyyyyyyyyyyyyy` (du client)
     - Token URL : `https://login.microsoftonline.com/TENANT_ID_CLIENT/oauth2/v2.0/token`
     - Scope : `api://xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx/.default`
5. **Save** → **Test connection** → **Add actions** (sélectionner les tools souhaités)
6. **Publish** → **Publish to M365 Copilot**

---

## 🛠️ Tools disponibles par agent

### Dev Agent (toujours)
| Tool | Description |
|------|-------------|
| `dev_read_file` | Lecture fichier workspace local |
| `dev_write_file` | Écriture fichier (preview + approval) |
| `dev_list_directory` | Liste répertoire |
| `dev_generate_powershell` | Génère script PowerShell (Graph, Intune, Teams, etc.) |
| `dev_generate_python` | Génère script Python (Graph) |
| `dev_generate_terraform` | Génère config Terraform |

### Knowledge Agent (toujours)
| Tool | Description |
|------|-------------|
| `engineering_knowledge_search` | Recherche dans Standards/, Runbooks/, Templates/, Architecture/ |
| `engineering_knowledge_read` | Lit un fichier de la allowlist |
| `engineering_knowledge_list` | Liste les fichiers allowlistés |

### Collaboration Agent (toujours)
| Tool | Description |
|------|-------------|
| `collab_preview_teams_message` | Prévoit message Teams (approval requise) |
| `collab_preview_teams_channel` | Prévoit création channel (approval) |
| `collab_preview_sharepoint_file` | Prévoit upload SharePoint (approval) |
| `collab_list_teams` | Liste teams (lecture seule) |
| `collab_list_channels` | Liste channels (lecture seule) |
| `collab_approve_plan` | Approuve un plan preview |
| `collab_execute_plan` | Exécute plan approuvé (hash vérifié) |

### Purview Agent (si `PURVIEW_ACCOUNT_NAME` configuré)
| Tool | Description |
|------|-------------|
| `purview_list_accounts` | Liste comptes Purview accessibles |
| `purview_list_data_sources` | Liste sources de données enregistrées |
| `purview_list_assets` | Recherche/assets dans Data Catalog |
| `purview_get_asset` | Détail asset (schéma, classification, lineage) |
| `purview_get_lineage` | Ligne de données (upstream/downstream) |
| `purview_list_scans` | Liste scans d'une source |
| `purview_get_scan_status` | Statut détaillé d'un scan |
| `purview_list_classifications` | Règles de classification / labels |
| `purview_list_policies` | Politiques d'accès, rétention, qualité |
| `purview_search_assets` | Recherche avancée avec filtres |

---

## 🔐 Scopes & Permissions

### Base (toujours requis)
| Permission | Type | Usage |
|------------|------|-------|
| `User.Read` | Delegated | Profil utilisateur |
| `Files.Read` | Delegated | OneDrive/SharePoint lecture |
| `Team.ReadBasic.All` | Delegated | Liste teams |
| `Channel.ReadBasic.All` | Delegated | Liste channels |

### Purview (optionnel `-EnablePurview`)
| Permission | Type | Usage |
|------------|------|-------|
| `PurviewDataPolicies.Read.All` | Delegated | Politiques données |
| `PurviewCatalog.Read.All` | Delegated | Catalogue de données |
| `PurviewAccounts.Read.All` | Delegated | Comptes Purview |

### Intune (optionnel `-EnableIntune`)
| Permission | Type | Usage |
|------------|------|-------|
| `DeviceManagementManagedDevices.Read.All` | Delegated | Appareils gérés |
| `DeviceManagementConfiguration.Read.All` | Delegated | Config Intune |

### Security (optionnel `-EnableSecurity`)
| Permission | Type | Usage |
|------------|------|-------|
| `SecurityEvents.Read.All` | Delegated | Événements sécurité |
| `Alert.Read.All` | Delegated | Alertes Defender |

---

## 🔄 Maintenance

### Mise à jour
```powershell
cd C:\mcp-m365-copilot
git pull
npm ci
npm run build
# Relancer run-client.ps1
```

### Rotation Client Secret (tous les 2 ans)
1. Azure Portal → App Registration → Certificates & secrets → New client secret
2. Mettre à jour `.env` → Relancer `run-client.ps1`

### Logs & Debug
```powershell
# Logs MCP (stdout/stderr visible dans la console)
# Logs audit JSON : correlationId, userIdHash, tool, action, result, durationMs

# Niveau de log
$env:LOG_LEVEL="debug"  # debug, info, warn, error
```

---

## 🗑️ Désinstallation

```powershell
.\uninstall.ps1 -Force
```

---

## ❓ Dépannage

| Problème | Solution |
|----------|----------|
| `AADSTS50011: Reply URL mismatch` | Vérifier `http://localhost:8400` dans App Registration > Authentication |
| `AADSTS65001: Consent required` | Admin Consent manquant → Azure Portal > API permissions > Grant admin consent |
| `cloudflared: command not found` | Script le télécharge auto ; sinon manuel: https://github.com/cloudflare/cloudflared/releases |
| `Purview: 403 Forbidden` | Vérifier scopes Purview + Admin Consent + accès au compte Purview |
| `Tunnel ne démarre pas` | Port 3000 libre ? Firewall sortie HTTPS (443) autorisé ? |
| `Node.js non trouvé` | `winget install OpenJS.NodeJS.LTS` ou installez manuellement |

---

## 📋 Checklist livraison client

- [ ] App Registration créée (`create-app-reg.ps1`)
- [ ] `client-creds.json` remis au client (Client ID, Tenant ID, **Secret**)
- [ ] `.env` configuré sur machine client
- [ ] `run-client.ps1` testé (tunnel actif, URL `trycloudflare.com`)
- [ ] Custom Connector créé dans Copilot Studio MSP
- [ ] Test outil `purview_list_assets` (si Purview) ou `dev_read_file`
- [ ] Documentation remise (ce README + procédure rotation secret)

---

## 🔒 Sécurité & Conformité

- **Données jamais chez le MSP** : MCP + tunnel tournent **chez le client**
- **Auth OBO** : Token utilisateur MSP → token pour Graph client (pas de secret partagé côté MSP)
- **Scopes minimaux** : Juste ce qui est nécessaire, pas de `*.All` admin
- **Approval workflow** : Toute écriture Teams/SharePoint = preview + hash + approval humaine
- **Audit local** : Logs JSON structurés sans contenu sensible
- **RGPD** : Données personnelles restent dans le tenant client

---

## 📞 Support

- **Repo source** : `https://github.com/tonorg/mcp-m365-copilot` (privé)
- **Issues** : GitHub Issues du repo
- **Doc technique** : `/docs` dans le repo principal