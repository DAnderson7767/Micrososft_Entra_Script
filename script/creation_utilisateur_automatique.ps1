#!/usr/bin/env pwsh

<#
.SYNOPSIS
    Script de création automatique d'utilisateur avec délégations des boîtes partagées
    
.DESCRIPTION
    Ce script permet de créer un nouvel utilisateur dans Microsoft 365 et de configurer
    automatiquement les délégations des boîtes partagées selon le service/département.
    
    Fonctionnalités :
    - Interface interactive pour saisie des informations utilisateur
    - Génération automatique des adresses email (prenom.nom@domain)
    - Support des domaines lde.fr et poplab.education
    - Configuration automatique des délégations par service
    - Gestion des prénoms composés (prenom-compose@domain)
    
.NOTES
    Prérequis :
    1. PowerShell Core : brew install --cask powershell
    2. Modules requis :
       - Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser
       - Install-Module -Name Microsoft.Graph.Users -Scope CurrentUser
       - Install-Module -Name Microsoft.Graph.Identity.DirectoryManagement -Scope CurrentUser
    3. Exécuter avec : pwsh ./script/creation_utilisateur_automatique.ps1
    
.AUTHOR
    Script optimisé pour macOS
#>

# Configuration globale
$Script:Config = @{
    Domains = @("lde.fr", "poplab.education")
    DefaultPassword = "Lde2025@@"
    OutputPath = "/Users/davidchiche/Desktop/Microsoft Entra/Rapport_Creation_Utilisateur.txt"
}

# Mapping des délégations par service/département
$Script:DelegationMapping = @{
    "Numérique" = @{
        "SharedMailboxes" = @(
            "aide@poplab.education",
            "aura@lde.fr",
            "bonjour@poplab.education",
            "commandes_numeriques@lde.fr",
            "contact@poplab.education",
            "evenements@lde.fr",
            "fablab@lde.fr",
            "grandest@lde.fr",
            "ile-de-france@lde.fr",
            "lareunion@lde.fr",
            "lucie@poplab.education",
            "notifications@poplab.education",
            "num_rc_gpt@lde.fr",
            "numerique@lde.fr",
            "occitanie@lde.fr",
            "signalementCatalogue@lde.fr",
            "support@lde.fr"
        )
        "Permissions" = @("FullAccess", "SendAs")
    }
    "Relations Clients" = @{
        "SharedMailboxes" = @(
            "abonnements@lde.fr",
            "accueil@lde.fr",
            "achats@lde.fr",
            "aide@poplab.education",
            "archives.techniques@lde.fr",
            "commandes_numeriques@lde.fr",
            "etiquettes@lde.fr",
            "grandest@lde.fr",
            "ile-de-france@lde.fr",
            "num_rc_gpt@lde.fr",
            "numerique@lde.fr",
            "parents@lde.fr",
            "support@lde.fr"
        )
        "Permissions" = @("FullAccess", "SendAs")
    }
    "Export" = @{
        "SharedMailboxes" = @(
            "commerce_export_gpt@lde.fr",
            "parents@lde.fr",
            "zzz_expeditions@lde.fr"
        )
        "Permissions" = @("FullAccess", "SendAs")
    }
    "Commerce" = @{
        "SharedMailboxes" = @(
            "commerce_export_gpt@lde.fr",
            "cristalweb@lde.fr",
            "marketing@lde.fr",
            "panier@lde.fr"
        )
        "Permissions" = @("FullAccess", "SendAs")
    }
    "Chef de Projet" = @{
        "SharedMailboxes" = @(
            "perplexity@lde.fr",
            "serveurlocal@lde.fr",
            "si_gpt@lde.fr",
            "aide@poplab.education",
            "fablab@lde.fr",
            "notifications@poplab.education",
            "archives.techniques@lde.fr",
            "support@lde.fr"
        )
        "Permissions" = @("FullAccess", "SendAs")
    }
    "Comptabilité" = @{
        "SharedMailboxes" = @(
            "comptabilite@lde.fr",
            "comptarh_gpt@lde.fr"
        )
        "Permissions" = @("FullAccess", "SendAs")
    }
    "Marketing" = @{
        "SharedMailboxes" = @(
            "figma@lde.fr",
            "marketing_gpt@lde.fr",
            "marketing@lde.fr",
            "aide@poplab.education",
            "bonjour@poplab.education",
            "contact@poplab.education",
            "ile-de-france@lde.fr",
            "notifications@poplab.education",
            "lucie@poplab.education",
            "accueil@lde.fr"
        )
        "Permissions" = @("FullAccess", "SendAs")
    }
}

# ===== FONCTIONS UTILITAIRES =====

function Write-ColorOutput {
    param(
        [string]$Message,
        [string]$Color = "White"
    )
    Write-Host $Message -ForegroundColor $Color
}

function Test-ModuleAvailability {
    param([string[]]$ModuleNames)
    
    $MissingModules = @()
    foreach ($Module in $ModuleNames) {
        if (-not (Get-Module -ListAvailable -Name $Module)) {
            $MissingModules += $Module
        }
    }

    if ($MissingModules.Count -gt 0) {
        Write-ColorOutput "Modules manquants : $($MissingModules -join ', ')" "Red"
        Write-ColorOutput "Installez-les avec :" "Yellow"
        foreach ($Module in $MissingModules) {
            Write-ColorOutput "Install-Module -Name $Module -Scope CurrentUser" "Yellow"
        }
        return $false
    }
    return $true
}

function Connect-ToServices {
    try {
        Write-ColorOutput "Connexion à Exchange Online..." "Green"
        Connect-ExchangeOnline -ErrorAction Stop
        
        Write-ColorOutput "Connexion à Microsoft Graph..." "Green"
        Connect-MgGraph -Scopes "User.ReadWrite.All", "Directory.ReadWrite.All" -ErrorAction Stop
        
        return $true
    }
    catch {
        Write-ColorOutput "Erreur de connexion : $($_.Exception.Message)" "Red"
        return $false
    }
}

function Disconnect-FromServices {
    try {
        Write-ColorOutput "Déconnexion des services..." "Cyan"
        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
        Disconnect-MgGraph -ErrorAction SilentlyContinue
        Write-ColorOutput "Déconnexion terminée." "Green"
    }
    catch {
        Write-ColorOutput "Erreur lors de la déconnexion : $($_.Exception.Message)" "Yellow"
    }
}

# ===== FONCTIONS DE VALIDATION ET GÉNÉRATION D'EMAIL =====

function Normalize-Name {
    param([string]$Name)
    
    if ([string]::IsNullOrWhiteSpace($Name)) {
        return ""
    }
    
    # Supprimer les espaces multiples et normaliser
    $normalized = $Name.Trim() -replace "\s+", " "
    
    # Supprimer les caractères spéciaux et accents
    $normalized = $normalized -replace "[àáâãäå]", "a"
    $normalized = $normalized -replace "[èéêë]", "e"
    $normalized = $normalized -replace "[ìíîï]", "i"
    $normalized = $normalized -replace "[òóôõö]", "o"
    $normalized = $normalized -replace "[ùúûü]", "u"
    $normalized = $normalized -replace "[ýÿ]", "y"
    $normalized = $normalized -replace "[ç]", "c"
    $normalized = $normalized -replace "[ñ]", "n"
    $normalized = $normalized -replace "[ÀÁÂÃÄÅ]", "A"
    $normalized = $normalized -replace "[ÈÉÊË]", "E"
    $normalized = $normalized -replace "[ÌÍÎÏ]", "I"
    $normalized = $normalized -replace "[ÒÓÔÕÖ]", "O"
    $normalized = $normalized -replace "[ÙÚÛÜ]", "U"
    $normalized = $normalized -replace "[ÝŸ]", "Y"
    $normalized = $normalized -replace "[Ç]", "C"
    $normalized = $normalized -replace "[Ñ]", "N"
    
    # Supprimer les caractères non alphanumériques sauf les espaces et tirets
    $normalized = $normalized -replace "[^a-zA-Z0-9\s\-]", ""
    
    return $normalized.ToLower()
}

function Generate-EmailAddress {
    param(
        [string]$FirstName,
        [string]$LastName,
        [string]$Domain
    )
    
    $normalizedFirst = Normalize-Name -Name $FirstName
    $normalizedLast = Normalize-Name -Name $LastName
    
    if ([string]::IsNullOrWhiteSpace($normalizedFirst) -or [string]::IsNullOrWhiteSpace($normalizedLast)) {
        return $null
    }
    
    # Gérer les prénoms composés (remplacer les espaces par des tirets)
    $normalizedFirst = $normalizedFirst -replace "\s+", "-"
    
    # Générer l'adresse email
    $email = "$normalizedFirst.$normalizedLast@$Domain"
    
    return $email
}

function Test-EmailAvailability {
    param(
        [string]$EmailAddress
    )
    
    try {
        $User = Get-MgUser -Filter "userPrincipalName eq '$EmailAddress'" -ErrorAction SilentlyContinue
        return $User -eq $null
    }
    catch {
        return $false
    }
}

# ===== INTERFACE UTILISATEUR =====

function Get-UserInformation {
    Write-ColorOutput "`n=== CRÉATION D'UN NOUVEL UTILISATEUR ===" "Yellow"
    Write-ColorOutput "Veuillez saisir les informations du nouvel utilisateur :" "Cyan"
    
    # Saisie du prénom
    do {
        $FirstName = Read-Host "Prénom"
        if ([string]::IsNullOrWhiteSpace($FirstName)) {
            Write-ColorOutput "Le prénom ne peut pas être vide." "Red"
        }
    } while ([string]::IsNullOrWhiteSpace($FirstName))
    
    # Saisie du nom
    do {
        $LastName = Read-Host "Nom"
        if ([string]::IsNullOrWhiteSpace($LastName)) {
            Write-ColorOutput "Le nom ne peut pas être vide." "Red"
        }
    } while ([string]::IsNullOrWhiteSpace($LastName))
    
    # Sélection du domaine
    Write-ColorOutput "`nSélection du domaine :" "Yellow"
    for ($i = 0; $i -lt $Script:Config.Domains.Count; $i++) {
        Write-ColorOutput "  [$($i+1)] $($Script:Config.Domains[$i])" "White"
    }
    
    do {
        $domainChoice = Read-Host "Votre choix (1-$($Script:Config.Domains.Count))"
    } while ($domainChoice -notin @(1..$Script:Config.Domains.Count))
    
    $SelectedDomain = $Script:Config.Domains[$domainChoice - 1]
    
    # Génération de l'adresse email
    $EmailAddress = Generate-EmailAddress -FirstName $FirstName -LastName $LastName -Domain $SelectedDomain
    
    if (-not $EmailAddress) {
        Write-ColorOutput "Erreur lors de la génération de l'adresse email." "Red"
        return $null
    }
    
    Write-ColorOutput "`nAdresse email générée : $EmailAddress" "Green"
    
    # Vérification de la disponibilité
    Write-ColorOutput "Vérification de la disponibilité de l'adresse email..." "Cyan"
    $IsAvailable = Test-EmailAvailability -EmailAddress $EmailAddress
    
    if (-not $IsAvailable) {
        Write-ColorOutput "ATTENTION: L'adresse email $EmailAddress existe déjà !" "Red"
        $continue = Read-Host "Voulez-vous continuer malgré tout ? (o/N)"
        if ($continue -notin @("o", "O", "oui", "OUI")) {
            return $null
        }
    } else {
        Write-ColorOutput "✓ L'adresse email est disponible." "Green"
    }
    
    # Saisie du département/service
    Write-ColorOutput "`nDépartements/services disponibles :" "Yellow"
    $Departments = $Script:DelegationMapping.Keys | Sort-Object
    for ($i = 0; $i -lt $Departments.Count; $i++) {
        Write-ColorOutput "  [$($i+1)] $($Departments[$i])" "White"
    }
    
    do {
        $deptChoice = Read-Host "Sélectionnez le département/service (1-$($Departments.Count))"
    } while ($deptChoice -notin @(1..$Departments.Count))
    
    $SelectedDepartment = $Departments[$deptChoice - 1]
    
    # Saisie du poste (optionnel)
    $JobTitle = Read-Host "Poste (optionnel)"
    
    # Saisie du mot de passe temporaire
    Write-ColorOutput "`nMot de passe temporaire (laisser vide pour utiliser le mot de passe par défaut) :" "Yellow"
    $TempPasswordInput = Read-Host "Mot de passe" -AsSecureString
    
    # Vérifier si un mot de passe a été saisi
    $TempPasswordPlain = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($TempPasswordInput))
    
    if ([string]::IsNullOrWhiteSpace($TempPasswordPlain)) {
        $TempPassword = $Script:Config.DefaultPassword
        Write-ColorOutput "Utilisation du mot de passe par défaut." "Yellow"
    } else {
        $TempPassword = $TempPasswordPlain
        Write-ColorOutput "Utilisation du mot de passe saisi." "Green"
    }
    
    return @{
        FirstName = $FirstName
        LastName = $LastName
        EmailAddress = $EmailAddress
        Domain = $SelectedDomain
        Department = $SelectedDepartment
        JobTitle = $JobTitle
        TempPassword = $TempPassword
    }
}

function Show-UserSummary {
    param([hashtable]$UserInfo)
    
    Write-ColorOutput "`n=== RÉCAPITULATIF ===" "Yellow"
    Write-ColorOutput "Prénom : $($UserInfo.FirstName)" "White"
    Write-ColorOutput "Nom : $($UserInfo.LastName)" "White"
    Write-ColorOutput "Adresse email : $($UserInfo.EmailAddress)" "White"
    Write-ColorOutput "Domaine : $($UserInfo.Domain)" "White"
    Write-ColorOutput "Département : $($UserInfo.Department)" "White"
    Write-ColorOutput "Poste : $($UserInfo.JobTitle)" "White"
    
    # Afficher les délégations qui seront configurées
    $Delegations = $Script:DelegationMapping[$UserInfo.Department]
    if ($Delegations) {
        Write-ColorOutput "`nDélégations qui seront configurées :" "Cyan"
        foreach ($Mailbox in $Delegations.SharedMailboxes) {
            Write-ColorOutput "  - $Mailbox ($($Delegations.Permissions -join ', '))" "White"
        }
    }
    
    Write-ColorOutput ""
    $confirm = Read-Host "Confirmez-vous la création de cet utilisateur ? (o/N)"
    return $confirm -in @("o", "O", "oui", "OUI")
}

# ===== FONCTIONS DE CRÉATION D'UTILISATEUR =====

function New-M365User {
    param([hashtable]$UserInfo)
    
    try {
        Write-ColorOutput "`nCréation de l'utilisateur dans Microsoft 365..." "Green"
        
        # Paramètres de création d'utilisateur
        $UserParams = @{
            UserPrincipalName = $UserInfo.EmailAddress
            DisplayName = "$($UserInfo.FirstName) $($UserInfo.LastName)"
            GivenName = $UserInfo.FirstName
            Surname = $UserInfo.LastName
            MailNickname = $UserInfo.EmailAddress.Split('@')[0]
            PasswordProfile = @{
                Password = $UserInfo.TempPassword
                ForceChangePasswordNextSignIn = $true
            }
            AccountEnabled = $true
            Department = $UserInfo.Department
        }
        
        if (-not [string]::IsNullOrWhiteSpace($UserInfo.JobTitle)) {
            $UserParams.JobTitle = $UserInfo.JobTitle
        }
        
        # Créer l'utilisateur
        $NewUser = New-MgUser @UserParams -ErrorAction Stop
        
        Write-ColorOutput "✓ Utilisateur créé avec succès : $($NewUser.UserPrincipalName)" "Green"
        Write-ColorOutput "  ID : $($NewUser.Id)" "Cyan"
        
        return $NewUser
    }
    catch {
        Write-ColorOutput "Erreur lors de la création de l'utilisateur : $($_.Exception.Message)" "Red"
        return $null
    }
}

function Wait-ForUserSync {
    param(
        [string]$UserEmail,
        [int]$MaxRetries = 10,
        [int]$DelaySeconds = 30
    )
    
    Write-ColorOutput "`nAttente de la synchronisation de l'utilisateur avec Exchange Online..." "Yellow"
    Write-ColorOutput "Cela peut prendre quelques minutes. Veuillez patienter..." "Cyan"
    
    for ($i = 1; $i -le $MaxRetries; $i++) {
        Write-ColorOutput "Tentative $i/$MaxRetries - Vérification de la disponibilité de $UserEmail..." "Cyan"
        
        try {
            $User = Get-Mailbox -Identity $UserEmail -ErrorAction SilentlyContinue
            if ($User) {
                Write-ColorOutput "✓ Utilisateur synchronisé avec Exchange Online !" "Green"
                return $true
            }
        }
        catch {
            # Ignorer les erreurs de recherche
        }
        
        if ($i -lt $MaxRetries) {
            Write-ColorOutput "Utilisateur pas encore disponible. Attente de $DelaySeconds secondes..." "Yellow"
            Start-Sleep -Seconds $DelaySeconds
        }
    }
    
    Write-ColorOutput "⚠ L'utilisateur n'est pas encore synchronisé après $MaxRetries tentatives." "Yellow"
    Write-ColorOutput "Les délégations seront configurées plus tard via un script séparé." "Yellow"
    return $false
}

function Check-UserLicense {
    param([string]$UserEmail)
    
    try {
        Write-ColorOutput "`nVérification de la licence Exchange Online de l'utilisateur..." "Cyan"
        
        # Vérifier si l'utilisateur a une boîte aux lettres
        $Mailbox = Get-Mailbox -Identity $UserEmail -ErrorAction SilentlyContinue
        if ($Mailbox) {
            Write-ColorOutput "✓ L'utilisateur a une boîte aux lettres Exchange Online." "Green"
            return $true
        }
        
        # Si pas de boîte aux lettres, vérifier les licences dans Azure AD
        $User = Get-MgUser -Filter "userPrincipalName eq '$UserEmail'" -Property "AssignedLicenses" -ErrorAction SilentlyContinue
        if ($User -and $User.AssignedLicenses -and $User.AssignedLicenses.Count -gt 0) {
            Write-ColorOutput "⚠ L'utilisateur a des licences assignées mais pas encore de boîte aux lettres." "Yellow"
            Write-ColorOutput "La synchronisation peut prendre quelques minutes." "Yellow"
            return "pending"
        } else {
            Write-ColorOutput "❌ L'utilisateur n'a aucune licence assignée !" "Red"
            Write-ColorOutput "`n=== ACTION REQUISE ===" "Red"
            Write-ColorOutput "Un administrateur doit assigner une licence Exchange Online à l'utilisateur :" "Yellow"
            Write-ColorOutput "  - $UserEmail" "White"
            Write-ColorOutput "`nLicences recommandées :" "Cyan"
            Write-ColorOutput "  - Microsoft 365 Business Premium" "White"
            Write-ColorOutput "  - Microsoft 365 E3/E5" "White"
            Write-ColorOutput "  - Exchange Online Plan 1/2" "White"
            Write-ColorOutput "`nUne fois la licence assignée, appuyez sur Entrée pour continuer..." "Green"
            Read-Host
            
            # Vérifier à nouveau après l'action de l'admin
            Write-ColorOutput "Vérification après assignation de licence..." "Cyan"
            
            # Vérifier si une licence a été assignée
            $UserAfter = Get-MgUser -Filter "userPrincipalName eq '$UserEmail'" -Property "AssignedLicenses" -ErrorAction SilentlyContinue
            if ($UserAfter -and $UserAfter.AssignedLicenses -and $UserAfter.AssignedLicenses.Count -gt 0) {
                Write-ColorOutput "✓ Licence assignée détectée !" "Green"
                return "retry"
            } else {
                Write-ColorOutput "⚠ Aucune licence détectée. Vérifiez que la licence a bien été assignée." "Yellow"
                Write-ColorOutput "Appuyez sur Entrée pour réessayer ou Ctrl+C pour annuler..." "Yellow"
                Read-Host
                return "retry"
            }
        }
    }
    catch {
        Write-ColorOutput "Erreur lors de la vérification de la licence : $($_.Exception.Message)" "Red"
        return $false
    }
}

function Set-MailboxDelegations {
    param(
        [string]$UserEmail,
        [string]$Department
    )
    
    try {
        Write-ColorOutput "`nConfiguration des délégations des boîtes partagées..." "Green"
        
        # Vérifier d'abord si l'utilisateur a une licence
        $LicenseStatus = Check-UserLicense -UserEmail $UserEmail
        
        if ($LicenseStatus -eq "retry") {
            # L'admin a assigné une licence, attendre la synchronisation
            if (-not (Wait-ForUserSync -UserEmail $UserEmail)) {
                Write-ColorOutput "`nLes délégations ne peuvent pas être configurées maintenant." "Red"
                Write-ColorOutput "L'utilisateur sera configuré automatiquement dans quelques minutes." "Yellow"
                Write-ColorOutput "Vous pouvez relancer le script de délégations plus tard." "Cyan"
                return @("Délégations reportées - utilisateur pas encore synchronisé")
            }
        } elseif ($LicenseStatus -eq "pending") {
            # L'utilisateur a une licence mais pas encore de boîte aux lettres
            if (-not (Wait-ForUserSync -UserEmail $UserEmail)) {
                Write-ColorOutput "`nLes délégations ne peuvent pas être configurées maintenant." "Red"
                Write-ColorOutput "L'utilisateur sera configuré automatiquement dans quelques minutes." "Yellow"
                Write-ColorOutput "Vous pouvez relancer le script de délégations plus tard." "Cyan"
                return @("Délégations reportées - utilisateur pas encore synchronisé")
            }
        } elseif ($LicenseStatus -eq $false) {
            Write-ColorOutput "`nImpossible de configurer les délégations sans licence Exchange Online." "Red"
            return @("Erreur - aucune licence Exchange Online")
        }
        
        $Delegations = $Script:DelegationMapping[$Department]
        if (-not $Delegations) {
            Write-ColorOutput "Aucune délégation configurée pour le département : $Department" "Yellow"
            return @()
        }
        
        $Results = @()
        
        foreach ($Mailbox in $Delegations.SharedMailboxes) {
            Write-ColorOutput "Configuration de $Mailbox..." "Cyan"
            
            try {
                # Vérifier que la boîte aux lettres existe
                $MailboxExists = Get-Mailbox -Identity $Mailbox -ErrorAction SilentlyContinue
                if (-not $MailboxExists) {
                    Write-ColorOutput "  ⚠ Boîte aux lettres $Mailbox introuvable, ignorée." "Yellow"
                    continue
                }
                
                # Configurer les permissions selon le type
                foreach ($Permission in $Delegations.Permissions) {
                    switch ($Permission) {
                        "FullAccess" {
                            try {
                                Add-MailboxPermission -Identity $Mailbox -User $UserEmail -AccessRights FullAccess -InheritanceType All -ErrorAction Stop
                                Write-ColorOutput "  ✓ Accès complet accordé" "Green"
                                $Results += "FullAccess sur $Mailbox"
                            }
                            catch {
                                Write-ColorOutput "  ⚠ Erreur accès complet : $($_.Exception.Message)" "Yellow"
                            }
                        }
                        "SendAs" {
                            try {
                                Add-RecipientPermission -Identity $Mailbox -Trustee $UserEmail -AccessRights SendAs -Confirm:$false -ErrorAction Stop
                                Write-ColorOutput "  ✓ Permission Send As accordée" "Green"
                                $Results += "SendAs sur $Mailbox"
                            }
                            catch {
                                Write-ColorOutput "  ⚠ Erreur Send As : $($_.Exception.Message)" "Yellow"
                            }
                        }
                        "SendOnBehalf" {
                            try {
                                Set-Mailbox -Identity $Mailbox -GrantSendOnBehalfTo @{Add=$UserEmail} -ErrorAction Stop
                                Write-ColorOutput "  ✓ Permission Send On Behalf accordée" "Green"
                                $Results += "SendOnBehalf sur $Mailbox"
                            }
                            catch {
                                Write-ColorOutput "  ⚠ Erreur Send On Behalf : $($_.Exception.Message)" "Yellow"
                            }
                        }
                    }
                }
            }
            catch {
                Write-ColorOutput "  ⚠ Erreur lors de la configuration de $Mailbox : $($_.Exception.Message)" "Yellow"
            }
        }
        
        return $Results
    }
    catch {
        Write-ColorOutput "Erreur lors de la configuration des délégations : $($_.Exception.Message)" "Red"
        return @()
    }
}

# ===== GÉNÉRATION DE RAPPORT =====

function Generate-CreationReport {
    param(
        [hashtable]$UserInfo,
        [object]$CreatedUser,
        [array]$Delegations
    )
    
    try {
        Write-ColorOutput "`nGénération du rapport de création..." "Cyan"
        
        $ReportContent = @()
        $ReportContent += "# Rapport de Création d'Utilisateur"
        $ReportContent += ""
        $ReportContent += "Généré le : $(Get-Date -Format 'dd/MM/yyyy à HH:mm')"
        $ReportContent += ""
        $ReportContent += "=== INFORMATIONS UTILISATEUR ==="
        $ReportContent += "Prénom : $($UserInfo.FirstName)"
        $ReportContent += "Nom : $($UserInfo.LastName)"
        $ReportContent += "Adresse email : $($UserInfo.EmailAddress)"
        $ReportContent += "Domaine : $($UserInfo.Domain)"
        $ReportContent += "Département : $($UserInfo.Department)"
        $ReportContent += "Poste : $($UserInfo.JobTitle)"
        $ReportContent += ""
        
        if ($CreatedUser) {
            $ReportContent += "=== CRÉATION RÉUSSIE ==="
            $ReportContent += "ID utilisateur : $($CreatedUser.Id)"
            $ReportContent += "Date de création : $(Get-Date -Format 'dd/MM/yyyy à HH:mm')"
            $ReportContent += ""
        } else {
            $ReportContent += "=== ÉCHEC DE LA CRÉATION ==="
            $ReportContent += "L'utilisateur n'a pas pu être créé."
            $ReportContent += ""
        }
        
        if ($Delegations.Count -gt 0) {
            $ReportContent += "=== DÉLÉGATIONS CONFIGURÉES ==="
            foreach ($Delegation in $Delegations) {
                $ReportContent += "- $Delegation"
            }
            $ReportContent += ""
        }
        
        # Écrire le rapport
        $ReportContent | Out-File -FilePath $Script:Config.OutputPath -Encoding UTF8
        Write-ColorOutput "Rapport généré : $($Script:Config.OutputPath)" "Green"
        
        return $ReportContent
    }
    catch {
        Write-ColorOutput "Erreur lors de la génération du rapport : $($_.Exception.Message)" "Red"
        return @()
    }
}

# ===== SCRIPT PRINCIPAL =====

function Main {
    try {
        # Vérification des modules
        if (-not (Test-ModuleAvailability -ModuleNames @("ExchangeOnlineManagement", "Microsoft.Graph.Users", "Microsoft.Graph.Identity.DirectoryManagement"))) {
            exit 1
        }
        
        # Import des modules
        Import-Module ExchangeOnlineManagement -ErrorAction Stop
        Import-Module Microsoft.Graph.Users -ErrorAction Stop
        Import-Module Microsoft.Graph.Identity.DirectoryManagement -ErrorAction Stop
        
        # Connexion aux services
        if (-not (Connect-ToServices)) {
            exit 1
        }
        
        # Saisie des informations utilisateur
        $UserInfo = Get-UserInformation
        if (-not $UserInfo) {
            Write-ColorOutput "Création d'utilisateur annulée." "Yellow"
            Disconnect-FromServices
            exit 0
        }
        
        # Affichage du récapitulatif et confirmation
        if (-not (Show-UserSummary -UserInfo $UserInfo)) {
            Write-ColorOutput "Création d'utilisateur annulée." "Yellow"
            Disconnect-FromServices
            exit 0
        }
        
        # Création de l'utilisateur
        $CreatedUser = New-M365User -UserInfo $UserInfo
        if (-not $CreatedUser) {
            Write-ColorOutput "Échec de la création de l'utilisateur." "Red"
            Disconnect-FromServices
            exit 1
        }
        
        # Configuration des délégations
        $Delegations = Set-MailboxDelegations -UserEmail $UserInfo.EmailAddress -Department $UserInfo.Department
        
        # Génération du rapport
        $ReportContent = Generate-CreationReport -UserInfo $UserInfo -CreatedUser $CreatedUser -Delegations $Delegations
        
        # Résumé final
        Write-ColorOutput "`n=== CRÉATION TERMINÉE ===" "Green"
        Write-ColorOutput "Utilisateur créé : $($UserInfo.EmailAddress)" "Green"
        Write-ColorOutput "Délégations configurées : $($Delegations.Count)" "Green"
        Write-ColorOutput "Rapport généré : $($Script:Config.OutputPath)" "Green"
        
        Write-ColorOutput "`nIMPORTANT : L'utilisateur devra changer son mot de passe lors de sa première connexion." "Yellow"
        
    }
    catch {
        Write-ColorOutput "Erreur critique : $($_.Exception.Message)" "Red"
        Write-ColorOutput "Stack trace : $($_.ScriptStackTrace)" "Red"
    }
    finally {
        # Déconnexion propre
        Disconnect-FromServices
    }
}

# Exécution du script principal
Main
