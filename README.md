# Microsoft Entra Scripts

Collection de scripts PowerShell pour la gestion des délégations Exchange Online dans Microsoft Entra.

## 📋 Description

Ce dépôt contient des scripts PowerShell optimisés pour :
- **Recherche et analyse** des délégations Exchange Online
- **Ajout de délégations** sur les boîtes aux lettres
- **Génération de rapports** formatés par service

## 🚀 Scripts disponibles

### 👥 `script/export_utilisateurs_macos.ps1` (Recommandé pour macOS)
Script pour exporter la liste complète des utilisateurs Microsoft Graph avec leurs informations détaillées.

**Fonctionnalités :**
- Récupération automatique de tous les utilisateurs Microsoft Graph
- Extraction des informations : nom, prénom, département, poste, email
- Génération de rapport texte formaté par département
- Export CSV pour analyses complémentaires
- Statistiques détaillées et regroupement par service
- Gestion des comptes actifs/désactivés
- **Exclusion automatique des boîtes partagées** (par défaut)
- **Optimisé pour macOS avec Microsoft.Graph**

**Utilisation :**
```bash
# Export de base (utilisateurs actifs uniquement, boîtes partagées exclues)
pwsh ./script/export_utilisateurs_macos.ps1

# Export complet avec comptes désactivés
pwsh ./script/export_utilisateurs_macos.ps1 -IncludeDisabled

# Export avec boîtes partagées incluses
pwsh ./script/export_utilisateurs_macos.ps1 -IncludeSharedMailboxes

# Export vers un répertoire spécifique
pwsh ./script/export_utilisateurs_macos.ps1 -OutputPath "/chemin/vers/dossier"
```

### 👥 `script/export_utilisateurs_complet.ps1` (Version AzureAD)
Script pour exporter la liste complète des utilisateurs Azure AD avec leurs informations détaillées.

**Fonctionnalités :**
- Récupération automatique de tous les utilisateurs Azure AD
- Extraction des informations : nom, prénom, département, poste, email
- Génération de rapport texte formaté par département
- Export CSV pour analyses complémentaires
- Statistiques détaillées et regroupement par service
- Gestion des comptes actifs/désactivés

**Utilisation :**
```bash
# Export de base (utilisateurs actifs uniquement)
pwsh ./script/export_utilisateurs_complet.ps1

# Export complet avec comptes désactivés
pwsh ./script/export_utilisateurs_complet.ps1 -IncludeDisabled

# Export vers un répertoire spécifique
pwsh ./script/export_utilisateurs_complet.ps1 -OutputPath "/chemin/vers/dossier"
```

### 📊 `script/rapport_delegations_complet.ps1`
Script principal pour analyser toutes les délégations et générer un rapport complet.

**Fonctionnalités :**
- Recherche automatique sur toutes les boîtes aux lettres
- Barre de progression en temps réel
- Génération de rapport formaté par service
- Export CSV pour analyses complémentaires

**Utilisation :**
```bash
pwsh ./script/rapport_delegations_complet.ps1
```

### ➕ `script/ajouter_delegation.ps1`
Script interactif pour ajouter des délégations sur des boîtes aux lettres.

**Fonctionnalités :**
- Interface utilisateur guidée
- Validation des adresses email
- Application automatique de toutes les permissions
- Vérification post-application

**Utilisation :**
```bash
pwsh ./script/ajouter_delegation.ps1
```

### 🔧 `script/installer_modules_macos.ps1` (Recommandé pour macOS)
Script d'installation automatique des modules PowerShell requis pour macOS.

**Fonctionnalités :**
- Installation automatique du module Microsoft.Graph
- Vérification de la version PowerShell
- Configuration complète de l'environnement macOS
- Détection automatique de l'environnement macOS

**Utilisation :**
```bash
pwsh ./script/installer_modules_macos.ps1
```

### 🔧 `script/installer_modules.ps1` (Version AzureAD)
Script d'installation automatique des modules PowerShell requis.

**Fonctionnalités :**
- Installation automatique du module AzureAD
- Vérification de la version PowerShell
- Configuration complète de l'environnement

**Utilisation :**
```bash
pwsh ./script/installer_modules.ps1
```

## 📁 Fichiers générés

### Scripts de délégations :
- **`Rapport_Delegations_Formate.txt`** - Rapport formaté par service avec liens mailto
- **`Delegations_Possedees_Report.csv`** - Données complètes au format CSV

### Script d'export utilisateurs :
- **`Utilisateurs_Graph_YYYYMMDD_HHMMSS.txt`** - Rapport formaté par département avec statistiques (Microsoft Graph)
- **`Utilisateurs_Graph_YYYYMMDD_HHMMSS.csv`** - Données complètes au format CSV (Microsoft Graph)
- **`Utilisateurs_AzureAD_YYYYMMDD_HHMMSS.txt`** - Rapport formaté par département avec statistiques (AzureAD)
- **`Utilisateurs_AzureAD_YYYYMMDD_HHMMSS.csv`** - Données complètes au format CSV (AzureAD)

## ⚙️ Prérequis

### Sur macOS :
1. **PowerShell Core** :
   ```bash
   brew install --cask powershell
   ```

2. **Module ExchangeOnlineManagement** :
   ```powershell
   Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser
   ```

3. **Compte administrateur** Microsoft Entra/Microsoft 365

## 🔧 Configuration

**IMPORTANT :** Avant d'utiliser les scripts, modifiez la configuration dans les fichiers :

### Dans `script/rapport_delegations_complet.ps1` :
```powershell
# Liste des utilisateurs dont on cherche les délégations
$TargetUsers = @(
    "utilisateur1@votre-domaine.com",
    "utilisateur2@votre-domaine.com",
    "utilisateur3@votre-domaine.com"
)

# Configuration des services
$Services = @{
    "Votre Service" = @{
        "Responsable" = "Nom du Responsable"
        "Email" = "responsable@votre-domaine.com"
        "Utilisateurs" = @("utilisateur1@votre-domaine.com")
    }
}
```

## 📊 Types de délégations

- **Full Access** - Accès complet à la boîte aux lettres
- **Send As** - Permission d'envoyer des emails au nom de l'utilisateur
- **Send on Behalf** - Permission d'envoyer des emails de la part de l'utilisateur

## 🛡️ Sécurité

- **Opérations en lecture seule** pour l'analyse
- **Validation des entrées** utilisateur
- **Gestion d'erreurs** complète
- **Déconnexion automatique** d'Exchange Online

## 📝 Exemples de rapports

### Rapport des délégations :
```markdown
## Votre Service

Template de base: **Nom du Responsable**

- [boite1@votre-domaine.com](mailto:boite1@votre-domaine.com)
- [boite2@votre-domaine.com](mailto:boite2@votre-domaine.com) (Nom d'affichage)
- [boite3@votre-domaine.com](mailto:boite3@votre-domaine.com) (Autre boîte)
```

### Rapport des utilisateurs :
Voir le fichier `Exemple_Rapport_Utilisateurs.txt` pour un exemple complet du format de rapport généré par le script d'export des utilisateurs.

## 🤝 Contribution

Pour contribuer à ce projet :
1. Fork le dépôt
2. Créer une branche feature
3. Commiter les changements
4. Pousser vers la branche
5. Ouvrir une Pull Request

## 📄 Licence

Ce projet est sous licence MIT. Voir le fichier `LICENSE` pour plus de détails.

## 🔗 Liens utiles

- [Documentation Exchange Online PowerShell](https://docs.microsoft.com/en-us/powershell/exchange/exchange-online-powershell)
- [Module ExchangeOnlineManagement](https://www.powershellgallery.com/packages/ExchangeOnlineManagement)
- [PowerShell Core sur macOS](https://docs.microsoft.com/en-us/powershell/scripting/install/installing-powershell-core-on-macos)
