# Microsoft Entra Scripts

Collection de scripts PowerShell pour la gestion des délégations Exchange Online dans Microsoft Entra.

## 📋 Description

Ce dépôt contient des scripts PowerShell optimisés pour :
- **Recherche et analyse** des délégations Exchange Online
- **Ajout de délégations** sur les boîtes aux lettres
- **Génération de rapports** formatés par service

## 🚀 Scripts disponibles

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

## 📁 Fichiers générés

- **`Rapport_Delegations_Formate.txt`** - Rapport formaté par service avec liens mailto
- **`Delegations_Possedees_Report.csv`** - Données complètes au format CSV

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

## 📝 Exemple de rapport

```markdown
## Votre Service

Template de base: **Nom du Responsable**

- [boite1@votre-domaine.com](mailto:boite1@votre-domaine.com)
- [boite2@votre-domaine.com](mailto:boite2@votre-domaine.com) (Nom d'affichage)
- [boite3@votre-domaine.com](mailto:boite3@votre-domaine.com) (Autre boîte)
```

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
