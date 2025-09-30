#!/usr/bin/env pwsh

<#
.SYNOPSIS
    Script pour configurer les délégations d'un utilisateur existant
    
.DESCRIPTION
    Ce script permet de configurer les délégations des boîtes partagées
    pour un utilisateur existant selon son département/service.
    
.NOTES
    Prérequis :
    1. PowerShell Core : brew install --cask powershell
    2. Modules requis :
       - Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser
       - Install-Module -Name Microsoft.Graph.Users -Scope CurrentUser
    3. Exécuter avec : pwsh ./script/configurer_delegations_utilisateur.ps1
    
.AUTHOR
    Script optimisé pour macOS
#>

# Configuration globale
$Script:Config = @{
    OutputPath = "/Users/davidchiche/Desktop/Microsoft Entra/Rapport_Configuration_Delegations.txt"
}

# Mapping des délégations par service/département (identique au script principal)
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
        Connect-MgGraph -Scopes "User.Read.All", "Directory.Read.All" -ErrorAction Stop
        
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

# ===== INTERFACE UTILISATEUR =====

function Get-UserInformation {
    Write-ColorOutput "`n=== CONFIGURATION DES DÉLÉGATIONS D'UN UTILISATEUR ===" "Yellow"
    Write-ColorOutput "Veuillez saisir les informations de l'utilisateur :" "Cyan"
    
    # Saisie de l'email utilisateur
    do {
        $UserEmail = Read-Host "Adresse email de l'utilisateur"
        if ([string]::IsNullOrWhiteSpace($UserEmail)) {
            Write-ColorOutput "L'adresse email ne peut pas être vide." "Red"
        }
    } while ([string]::IsNullOrWhiteSpace($UserEmail))
    
    # Vérification que l'utilisateur existe
    try {
        $User = Get-MgUser -Filter "userPrincipalName eq '$UserEmail'" -ErrorAction Stop
        if (-not $User) {
            Write-ColorOutput "Utilisateur $UserEmail introuvable dans Azure AD." "Red"
            return $null
        }
        
        Write-ColorOutput "✓ Utilisateur trouvé : $($User.DisplayName)" "Green"
        Write-ColorOutput "  Département actuel : $($User.Department)" "Cyan"
    }
    catch {
        Write-ColorOutput "Erreur lors de la recherche de l'utilisateur : $($_.Exception.Message)" "Red"
        return $null
    }
    
    # Sélection du département/service
    Write-ColorOutput "`nDépartements/services disponibles :" "Yellow"
    $Departments = $Script:DelegationMapping.Keys | Sort-Object
    for ($i = 0; $i -lt $Departments.Count; $i++) {
        Write-ColorOutput "  [$($i+1)] $($Departments[$i])" "White"
    }
    
    do {
        $deptChoice = Read-Host "Sélectionnez le département/service (1-$($Departments.Count))"
    } while ($deptChoice -notin @(1..$Departments.Count))
    
    $SelectedDepartment = $Departments[$deptChoice - 1]
    
    return @{
        UserEmail = $UserEmail
        Department = $SelectedDepartment
        UserInfo = $User
    }
}

function Show-UserSummary {
    param([hashtable]$UserInfo)
    
    Write-ColorOutput "`n=== RÉCAPITULATIF ===" "Yellow"
    Write-ColorOutput "Utilisateur : $($UserInfo.UserEmail)" "White"
    Write-ColorOutput "Nom d'affichage : $($UserInfo.UserInfo.DisplayName)" "White"
    Write-ColorOutput "Département sélectionné : $($UserInfo.Department)" "White"
    
    # Afficher les délégations qui seront configurées
    $Delegations = $Script:DelegationMapping[$UserInfo.Department]
    if ($Delegations) {
        Write-ColorOutput "`nDélégations qui seront configurées :" "Cyan"
        foreach ($Mailbox in $Delegations.SharedMailboxes) {
            Write-ColorOutput "  - $Mailbox ($($Delegations.Permissions -join ', '))" "White"
        }
    }
    
    Write-ColorOutput ""
    $confirm = Read-Host "Confirmez-vous la configuration des délégations ? (o/N)"
    return $confirm -in @("o", "O", "oui", "OUI")
}

# ===== FONCTIONS DE CONFIGURATION DES DÉLÉGATIONS =====

function Set-MailboxDelegations {
    param(
        [string]$UserEmail,
        [string]$Department
    )
    
    try {
        Write-ColorOutput "`nConfiguration des délégations des boîtes partagées..." "Green"
        
        # Vérifier que l'utilisateur existe dans Exchange Online
        try {
            $Mailbox = Get-Mailbox -Identity $UserEmail -ErrorAction Stop
            Write-ColorOutput "✓ Utilisateur trouvé dans Exchange Online" "Green"
        }
        catch {
            Write-ColorOutput "⚠ Utilisateur $UserEmail introuvable dans Exchange Online." "Red"
            Write-ColorOutput "Vérifiez que l'utilisateur a bien une boîte aux lettres." "Yellow"
            return @("Erreur - utilisateur introuvable dans Exchange Online")
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

function Generate-ConfigurationReport {
    param(
        [hashtable]$UserInfo,
        [array]$Delegations
    )
    
    try {
        Write-ColorOutput "`nGénération du rapport de configuration..." "Cyan"
        
        $ReportContent = @()
        $ReportContent += "# Rapport de Configuration des Délégations"
        $ReportContent += ""
        $ReportContent += "Généré le : $(Get-Date -Format 'dd/MM/yyyy à HH:mm')"
        $ReportContent += ""
        $ReportContent += "=== INFORMATIONS UTILISATEUR ==="
        $ReportContent += "Email : $($UserInfo.UserEmail)"
        $ReportContent += "Nom d'affichage : $($UserInfo.UserInfo.DisplayName)"
        $ReportContent += "Département : $($UserInfo.Department)"
        $ReportContent += ""
        
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
        if (-not (Test-ModuleAvailability -ModuleNames @("ExchangeOnlineManagement", "Microsoft.Graph.Users"))) {
            exit 1
        }
        
        # Import des modules
        Import-Module ExchangeOnlineManagement -ErrorAction Stop
        Import-Module Microsoft.Graph.Users -ErrorAction Stop
        
        # Connexion aux services
        if (-not (Connect-ToServices)) {
            exit 1
        }
        
        # Saisie des informations utilisateur
        $UserInfo = Get-UserInformation
        if (-not $UserInfo) {
            Write-ColorOutput "Configuration des délégations annulée." "Yellow"
            Disconnect-FromServices
            exit 0
        }
        
        # Affichage du récapitulatif et confirmation
        if (-not (Show-UserSummary -UserInfo $UserInfo)) {
            Write-ColorOutput "Configuration des délégations annulée." "Yellow"
            Disconnect-FromServices
            exit 0
        }
        
        # Configuration des délégations
        $Delegations = Set-MailboxDelegations -UserEmail $UserInfo.UserEmail -Department $UserInfo.Department
        
        # Génération du rapport
        $ReportContent = Generate-ConfigurationReport -UserInfo $UserInfo -Delegations $Delegations
        
        # Résumé final
        Write-ColorOutput "`n=== CONFIGURATION TERMINÉE ===" "Green"
        Write-ColorOutput "Utilisateur : $($UserInfo.UserEmail)" "Green"
        Write-ColorOutput "Délégations configurées : $($Delegations.Count)" "Green"
        Write-ColorOutput "Rapport généré : $($Script:Config.OutputPath)" "Green"
        
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
