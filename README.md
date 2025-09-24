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

Les scripts sont configurés pour rechercher les délégations des utilisateurs suivants :
- sophie.runtz@lde.fr
- celine.risch@lde.fr
- sarah.merah@lde.fr
- maxime.klein@lde.fr
- monia.belebbed@lde.fr
- elodie.urban@lde.fr
- elisabeth.laux@lde.fr

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
## Numérique

Template de base: **Céline Risch**

- [numerique@lde.fr](mailto:numerique@lde.fr)
- [support@lde.fr](mailto:support@lde.fr) (Support)
- [archives.techniques@lde.fr](mailto:archives.techniques@lde.fr) (Archives techniques)
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
