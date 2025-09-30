#!/usr/bin/env pwsh

<#
.SYNOPSIS
    Script simplifié de recherche et génération de rapport des délégations
    
.DESCRIPTION
    Ce script recherche les délégations possédées par des utilisateurs spécifiques
    et génère un rapport formaté organisé par utilisateur.
    
.NOTES
    Prérequis sur macOS :
    1. PowerShell Core : brew install --cask powershell
    2. Modules requis :
       - Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser
       - Install-Module -Name Microsoft.Graph.Users -Scope CurrentUser
    3. Exécuter avec : pwsh ./script/rapport_delegations_simple.ps1
#>

# Vérifier la présence des modules requis
$RequiredModules = @("ExchangeOnlineManagement", "Microsoft.Graph.Users")
$MissingModules = @()

foreach ($Module in $RequiredModules) {
    if (-not (Get-Module -ListAvailable -Name $Module)) {
        $MissingModules += $Module
    }
}

if ($MissingModules.Count -gt 0) {
    Write-Error "Modules manquants : $($MissingModules -join ', ')"
    Write-Host "Installez-les avec :" -ForegroundColor Yellow
    foreach ($Module in $MissingModules) {
        Write-Host "Install-Module -Name $Module -Scope CurrentUser" -ForegroundColor Yellow
    }
    exit 1
}

# Importer les modules
Import-Module ExchangeOnlineManagement
Import-Module Microsoft.Graph.Users

# Connexion aux services Microsoft
Write-Host "Connexion à Exchange Online..." -ForegroundColor Green
Connect-ExchangeOnline

Write-Host "Connexion à Microsoft Graph..." -ForegroundColor Green
Connect-MgGraph -Scopes "User.Read.All"

# Liste par défaut des utilisateurs
$defaultUsers = @(
    "celine.rish@lde.fr",
    "tom.wolff@lde.fr", 
    "tony.aussel@lde.fr",
    "stephanie.tiratel@lde.fr",
    "sarah.merah@lde.fr",
    "sophie.runtz@lde.fr",
    "monia.belebbed@lde.fr",
    "marine.bernauer@lde.fr",
    "david.weil@lde.fr"
)

# Menu principal simplifié
Write-Host "`n=== CONFIGURATION DES UTILISATEURS CIBLES ===" -ForegroundColor Yellow
Write-Host "Choisissez le mode de recherche des délégations :" -ForegroundColor Cyan
Write-Host ""
Write-Host "Options de recherche :" -ForegroundColor Yellow
Write-Host "  [1] Utiliser la liste par défaut" -ForegroundColor White
Write-Host "  [2] Recherche par département" -ForegroundColor White
Write-Host "  [3] Saisir des utilisateurs manuellement" -ForegroundColor White
Write-Host "  [4] Annuler" -ForegroundColor White

do {
    $choice = Read-Host "Votre choix (1-4)"
} while ($choice -notin @("1", "2", "3", "4"))

$TargetUsers = @()

switch ($choice) {
    "1" {
        $TargetUsers = $defaultUsers
        Write-Host "`nUtilisation de la liste par défaut :" -ForegroundColor Green
        $TargetUsers | ForEach-Object { Write-Host "  - $_" -ForegroundColor Cyan }
    }
    "2" {
        Write-Host "`n=== RECHERCHE PAR DÉPARTEMENT ===" -ForegroundColor Yellow
        
        # Récupérer les départements disponibles
        Write-Host "Récupération des départements disponibles..." -ForegroundColor Cyan
        try {
            $Users = Get-MgUser -All -Filter "userType eq 'Member' and accountEnabled eq true" -Property "UserPrincipalName", "DisplayName", "Department"
            $Departments = $Users | 
                Where-Object { $_.Department -and $_.Department.Trim() -ne "" } | 
                Select-Object -ExpandProperty Department | 
                Sort-Object | 
                Get-Unique
            
            if ($Departments.Count -eq 0) {
                Write-Warning "Aucun département trouvé dans Azure AD."
                Write-Host "Retour au menu principal..." -ForegroundColor Yellow
                exit 1
            }
            
            Write-Host "`nDépartements disponibles :" -ForegroundColor Green
            for ($i = 0; $i -lt $Departments.Count; $i++) {
                Write-Host "  [$($i+1)] $($Departments[$i])" -ForegroundColor White
            }
            
            Write-Host "  [$(($Departments.Count + 1))] Saisir un département manuellement" -ForegroundColor White
            Write-Host "  [$(($Departments.Count + 2))] Annuler" -ForegroundColor White
            
            do {
                $deptChoice = Read-Host "Votre choix (1-$(($Departments.Count + 2)))"
            } while ($deptChoice -notin @(1..($Departments.Count + 2)))
            
            if ($deptChoice -le $Departments.Count) {
                $SelectedDepartment = $Departments[$deptChoice - 1]
                Write-Host "Département sélectionné : $SelectedDepartment" -ForegroundColor Green
                
                # Rechercher les utilisateurs du département
                $FilteredUsers = $Users | Where-Object { 
                    $_.Department -and 
                    $_.Department.Trim() -ne "" -and 
                    $_.Department -like "*$SelectedDepartment*" 
                }
                
                if ($FilteredUsers.Count -eq 0) {
                    Write-Warning "Aucun utilisateur trouvé pour le département : $SelectedDepartment"
                    exit 1
                }
                
                $TargetUsers = $FilteredUsers | ForEach-Object { $_.UserPrincipalName }
                Write-Host "$($FilteredUsers.Count) utilisateur(s) trouvé(s) dans le département $SelectedDepartment" -ForegroundColor Green
            }
            elseif ($deptChoice -eq ($Departments.Count + 1)) {
                $ManualDepartment = Read-Host "Saisissez le nom du département"
                if ($ManualDepartment.Trim() -ne "") {
                    Write-Host "Département saisi : $ManualDepartment" -ForegroundColor Green
                    
                    $FilteredUsers = $Users | Where-Object { 
                        $_.Department -and 
                        $_.Department.Trim() -ne "" -and 
                        $_.Department -like "*$($ManualDepartment.Trim())*" 
                    }
                    
                    if ($FilteredUsers.Count -eq 0) {
                        Write-Warning "Aucun utilisateur trouvé pour le département : $ManualDepartment"
                        exit 1
                    }
                    
                    $TargetUsers = $FilteredUsers | ForEach-Object { $_.UserPrincipalName }
                    Write-Host "$($FilteredUsers.Count) utilisateur(s) trouvé(s) dans le département $ManualDepartment" -ForegroundColor Green
                } else {
                    Write-Error "Nom de département vide."
                    exit 1
                }
            }
            else {
                Write-Host "Opération annulée." -ForegroundColor Yellow
                exit 1
            }
        }
        catch {
            Write-Error "Erreur lors de la récupération des départements : $($_.Exception.Message)"
            exit 1
        }
    }
    "3" {
        Write-Host "`nSaisie des utilisateurs :" -ForegroundColor Green
        Write-Host "Tapez les adresses email (une par ligne), puis 'FIN' pour terminer :" -ForegroundColor Cyan
        
        $emails = @()
        $lineNumber = 1
        
        do {
            $userInput = Read-Host "Email $lineNumber"
            if ($userInput -ne "FIN" -and $userInput.Trim() -ne "") {
                $emails += $userInput.Trim()
                $lineNumber++
            }
        } while ($userInput -ne "FIN")
        
        # Validation des emails
        $validEmails = @()
        foreach ($email in $emails) {
            if ($email -match "^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$") {
                $validEmails += $email
            } else {
                Write-Warning "Email invalide ignoré : $email"
            }
        }
        
        if ($validEmails.Count -eq 0) {
            Write-Error "Aucun email valide saisi."
            exit 1
        }
        
        $TargetUsers = $validEmails
        Write-Host "`nUtilisateurs saisis :" -ForegroundColor Green
        $TargetUsers | ForEach-Object { Write-Host "  - $_" -ForegroundColor Cyan }
    }
    "4" {
        Write-Host "Opération annulée par l'utilisateur." -ForegroundColor Yellow
        exit 0
    }
}

# Vérifier si on a des utilisateurs valides
if ($TargetUsers.Count -eq 0) {
    Write-Host "`nAucun utilisateur sélectionné. Arrêt du script." -ForegroundColor Red
    exit 1
}

Write-Host "`nUtilisateurs configurés : $($TargetUsers.Count)" -ForegroundColor Green
$TargetUsers | ForEach-Object { Write-Host "  - $_" -ForegroundColor Cyan }

Write-Host "`nRecherche des délégations possedees par $($TargetUsers.Count) utilisateurs..." -ForegroundColor Green
Write-Host "ATTENTION: Cette operation peut prendre plusieurs minutes selon la taille de votre organisation." -ForegroundColor Yellow

# Récupérer toutes les boîtes aux lettres (utilisateurs et partagées)
Write-Host "`nRecuperation de toutes les boites aux lettres..." -ForegroundColor Cyan
$AllMailboxes = Get-Mailbox -ResultSize Unlimited | Where-Object { 
    $_.RecipientTypeDetails -in @('UserMailbox', 'SharedMailbox', 'RoomMailbox', 'EquipmentMailbox') 
}

Write-Host "$($AllMailboxes.Count) boites aux lettres trouvees" -ForegroundColor Green

# Initialiser le tableau des résultats
$Results = @()
$ProcessedCount = 0

# Fonction pour afficher une barre de progression
function Show-Progress {
    param($Current, $Total, $Activity = "Traitement")
    
    $PercentComplete = [math]::Round(($Current / $Total) * 100, 1)
    $ProgressBar = ""
    $BarLength = 50
    $FilledLength = [math]::Round(($Current / $Total) * $BarLength)
    
    for ($i = 0; $i -lt $FilledLength; $i++) { $ProgressBar += "█" }
    for ($i = $FilledLength; $i -lt $BarLength; $i++) { $ProgressBar += "░" }
    
    Write-Progress -Activity $Activity -Status "Boite $Current sur $Total ($PercentComplete%)" -PercentComplete $PercentComplete
    Write-Host "`r[$ProgressBar] $PercentComplete% ($Current/$Total)" -NoNewline -ForegroundColor Green
}

# Parcourir chaque boîte aux lettres
foreach ($Mailbox in $AllMailboxes) {
    $ProcessedCount++
    Show-Progress -Current $ProcessedCount -Total $AllMailboxes.Count -Activity "Analyse des delegations"
    
    try {
        # Récupérer les permissions Full Access (comparaison insensible à la casse)
        $FullAccessPerms = Get-MailboxPermission -Identity $Mailbox.PrimarySmtpAddress -ErrorAction SilentlyContinue | 
            Where-Object { 
                ($TargetUsers | ForEach-Object { $_.ToLower() }) -contains $_.User.ToString().ToLower() -and 
                $_.IsInherited -eq $false 
            }

        # Récupérer les permissions Send As (comparaison insensible à la casse)
        $SendAsPerms = Get-RecipientPermission -Identity $Mailbox.PrimarySmtpAddress -ErrorAction SilentlyContinue |
            Where-Object { ($TargetUsers | ForEach-Object { $_.ToLower() }) -contains $_.Trustee.ToString().ToLower() }

        # Récupérer les permissions Send on Behalf (comparaison insensible à la casse)
        $SendOnBehalfUsers = $Mailbox.GrantSendOnBehalfTo | Where-Object { 
            ($TargetUsers | ForEach-Object { $_.ToLower() }) -contains $_.ToString().ToLower() 
        }

        # Si des délégations sont trouvées, ajouter aux résultats
        if ($FullAccessPerms -or $SendAsPerms -or $SendOnBehalfUsers) {
            $Results += [PSCustomObject]@{
                BoiteAuxLettres    = $Mailbox.PrimarySmtpAddress
                TypeBoite          = $Mailbox.RecipientTypeDetails
                NomAffichage       = $Mailbox.DisplayName
                FullAccess         = if ($FullAccessPerms) { ($FullAccessPerms.User -join ", ") } else { "" }
                SendAs             = if ($SendAsPerms) { ($SendAsPerms.Trustee -join ", ") } else { "" }
                SendOnBehalf       = if ($SendOnBehalfUsers) { ($SendOnBehalfUsers -join ", ") } else { "" }
            }
        }
    }
    catch {
        Write-Warning "Erreur lors du traitement de $($Mailbox.PrimarySmtpAddress) : $($_.Exception.Message)"
    }
}

Write-Host "`nRecherche terminee !" -ForegroundColor Green
Write-Progress -Activity "Analyse des delegations" -Completed

# ===== GÉNÉRATION DU RAPPORT FORMATÉ =====

Write-Host "`nGeneration du rapport formate..." -ForegroundColor Cyan

# Initialiser le contenu du rapport
$ReportContent = @()
$ReportContent += "# Rapport des Délégations par Collaborateur"
$ReportContent += ""
$ReportContent += "Généré le : $(Get-Date -Format 'dd/MM/yyyy à HH:mm')"
$ReportContent += ""

# Créer un dictionnaire pour organiser les délégations par utilisateur
$UserDelegations = @{}

# Parcourir tous les résultats pour organiser par utilisateur
foreach ($Delegation in $Results) {
    # Extraire tous les utilisateurs qui ont des délégations sur cette boîte
    $UsersWithAccess = @()
    
    # Ajouter les utilisateurs avec accès complet
    if ($Delegation.FullAccess) {
        $UsersWithAccess += $Delegation.FullAccess -split ", " | ForEach-Object { $_.Trim() }
    }
    
    # Ajouter les utilisateurs avec Send As
    if ($Delegation.SendAs) {
        $UsersWithAccess += $Delegation.SendAs -split ", " | ForEach-Object { $_.Trim() }
    }
    
    # Ajouter les utilisateurs avec Send On Behalf
    if ($Delegation.SendOnBehalf) {
        $UsersWithAccess += $Delegation.SendOnBehalf -split ", " | ForEach-Object { $_.Trim() }
    }
    
    # Organiser par utilisateur
    foreach ($User in $UsersWithAccess) {
        if ($User -and $User.Trim() -ne "") {
            if (-not $UserDelegations.ContainsKey($User)) {
                $UserDelegations[$User] = @()
            }
            $UserDelegations[$User] += $Delegation
        }
    }
}

# Générer le rapport par utilisateur
foreach ($User in ($UserDelegations.Keys | Sort-Object)) {
    $UserDelegations[$User] = $UserDelegations[$User] | Sort-Object BoiteAuxLettres | Get-Unique -AsString
    
    $ReportContent += "$User :"
    foreach ($Delegation in $UserDelegations[$User]) {
        $ReportContent += "  - $($Delegation.BoiteAuxLettres) ($($Delegation.TypeBoite))"
    }
    $ReportContent += ""
}

# Chemins des fichiers de sortie
$OutputPath = "/Users/davidchiche/Desktop/Microsoft Entra/Rapport_Delegations_Formate.txt"
$CsvPath = "/Users/davidchiche/Desktop/Microsoft Entra/Delegations_Possedees_Report.csv"

# Écrire le rapport formaté
$ReportContent | Out-File -FilePath $OutputPath -Encoding UTF8

# Écrire le rapport CSV (optionnel, pour référence)
$Results | Export-Csv -Path $CsvPath -NoTypeInformation -Encoding UTF8

# Afficher les résultats
if ($Results.Count -gt 0) {
    Write-Host "`n=== RESUME ===" -ForegroundColor Yellow
    Write-Host "$($Results.Count) boite(s) aux lettres avec des delegations trouvees" -ForegroundColor Green
    
    # Résumé par utilisateur
    Write-Host "`n=== RESUME PAR UTILISATEUR ===" -ForegroundColor Yellow
    foreach ($User in $TargetUsers) {
        $UserDelegations = $Results | Where-Object { 
            $_.FullAccess.ToLower() -like "*$($User.ToLower())*" -or 
            $_.SendAs.ToLower() -like "*$($User.ToLower())*" -or 
            $_.SendOnBehalf.ToLower() -like "*$($User.ToLower())*" 
        }
        
        if ($UserDelegations.Count -gt 0) {
            Write-Host "`n$User a des delegations sur $($UserDelegations.Count) boite(s)" -ForegroundColor Cyan
        } else {
            Write-Host "`n$User : Aucune delegation trouvee" -ForegroundColor Gray
        }
    }
    
    Write-Host "`nRapport formate genere : $OutputPath" -ForegroundColor Green
    Write-Host "Rapport CSV genere : $CsvPath" -ForegroundColor Green
    
} else {
    Write-Host "`nAucune delegation trouvee pour les utilisateurs specifies." -ForegroundColor Red
}

# Déconnexion propre
Write-Host "`nDeconnexion des services..." -ForegroundColor Cyan
Disconnect-ExchangeOnline -Confirm:$false
Disconnect-MgGraph
Write-Host "Deconnexion terminee." -ForegroundColor Green

