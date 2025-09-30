#!/usr/bin/env pwsh

<#
.SYNOPSIS
    Script complet de recherche et génération de rapport des délégations avec support des départements
    
.DESCRIPTION
    Ce script :
    1. Recherche en lecture seule toutes les délégations possédées par les utilisateurs spécifiés
    2. Supporte la recherche par utilisateurs spécifiques ou par département
    3. Génère un rapport formaté organisé par utilisateur avec informations de département
    4. Affiche le type de boîte aux lettres pour chaque délégation
    
.NOTES
    Prérequis sur macOS :
    1. PowerShell Core : brew install --cask powershell
    2. Modules requis :
       - Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser
       - Install-Module -Name Microsoft.Graph.Users -Scope CurrentUser
    3. Exécuter avec : pwsh ./script/rapport_delegations_complet_v2.ps1
    
    Fonctionnalités :
    - Recherche par utilisateurs spécifiques (mode classique)
    - Recherche par département (nouveau)
    - Affichage des informations de département dans le rapport
    - Interface utilisateur améliorée avec menus interactifs
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
    Write-Host "`nNote: Sur macOS, utilisez Microsoft.Graph au lieu d'AzureAD" -ForegroundColor Cyan
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

# Fonction pour récupérer les utilisateurs par département
function Get-UsersByDepartment {
    param([string]$Department)
    
    Write-Host "Recherche des utilisateurs du département : $Department" -ForegroundColor Cyan
    
    try {
        # Récupérer tous les utilisateurs avec Microsoft Graph en spécifiant les propriétés nécessaires
        $Users = Get-MgUser -All -Filter "userType eq 'Member' and accountEnabled eq true" -Property "UserPrincipalName", "DisplayName", "Department", "JobTitle", "OfficeLocation"
        
        # Filtrer localement pour une recherche plus flexible (insensible à la casse)
        $FilteredUsers = $Users | Where-Object { 
            $_.Department -and 
            $_.Department.Trim() -ne "" -and 
            $_.Department -like "*$Department*" 
        }
        
        if ($FilteredUsers.Count -eq 0) {
            Write-Warning "Aucun utilisateur trouvé pour le département : $Department"
            Write-Host "Vérifiez l'orthographe ou essayez une recherche partielle." -ForegroundColor Yellow
            
            # Afficher quelques exemples de départements disponibles pour aider
            $AvailableDepts = $Users | Where-Object { $_.Department -and $_.Department.Trim() -ne "" } | 
                Select-Object -ExpandProperty Department | 
                Sort-Object | 
                Get-Unique | 
                Select-Object -First 10
            
            if ($AvailableDepts.Count -gt 0) {
                Write-Host "`nExemples de départements disponibles :" -ForegroundColor Cyan
                foreach ($Dept in $AvailableDepts) {
                    Write-Host "  - '$Dept'" -ForegroundColor White
                }
            }
            
            return @()
        }
        
        # Extraire les adresses email
        $UserEmails = $FilteredUsers | ForEach-Object { $_.UserPrincipalName }
        
        Write-Host "$($FilteredUsers.Count) utilisateur(s) trouvé(s) dans le département $Department" -ForegroundColor Green
        
        # Afficher les utilisateurs trouvés pour confirmation
        Write-Host "`nUtilisateurs trouvés :" -ForegroundColor Cyan
        foreach ($User in $FilteredUsers) {
            Write-Host "  - $($User.UserPrincipalName) (Département: $($User.Department))" -ForegroundColor White
        }
        
        return $UserEmails
        
    } catch {
        Write-Error "Erreur lors de la récupération des utilisateurs du département $Department : $($_.Exception.Message)"
        return @()
    }
}

# Fonction pour lister tous les départements disponibles
function Get-AvailableDepartments {
    try {
        Write-Host "Récupération des utilisateurs depuis Microsoft Graph..." -ForegroundColor Cyan
        
        # Récupérer tous les utilisateurs avec Microsoft Graph en spécifiant les propriétés nécessaires
        $Users = Get-MgUser -All -Filter "userType eq 'Member' and accountEnabled eq true" -Property "UserPrincipalName", "DisplayName", "Department", "JobTitle", "OfficeLocation"
        
        Write-Host "$($Users.Count) utilisateur(s) trouvé(s)" -ForegroundColor Green
        
        # Diagnostic : vérifier combien d'utilisateurs ont un département défini
        $UsersWithDepartment = $Users | Where-Object { 
            $_.Department -and 
            $_.Department.Trim() -ne "" -and 
            $null -ne $_.Department 
        }
        Write-Host "$($UsersWithDepartment.Count) utilisateur(s) avec département défini" -ForegroundColor Yellow
        
        if ($UsersWithDepartment.Count -eq 0) {
            Write-Warning "Aucun utilisateur n'a de département défini dans Azure AD."
            Write-Host "`nExemples d'utilisateurs pour diagnostic :" -ForegroundColor Yellow
            
            # Afficher quelques exemples d'utilisateurs pour diagnostic
            $SampleUsers = $Users | Select-Object -First 5 | Select-Object UserPrincipalName, DisplayName, Department, JobTitle, OfficeLocation
            $SampleUsers | Format-Table -AutoSize
            
            return @()
        }
        
        # Extraire les départements uniques de manière plus robuste
        $Departments = @()
        foreach ($User in $UsersWithDepartment) {
            if ($User.Department -and $User.Department.Trim() -ne "") {
                $Departments += $User.Department.Trim()
            }
        }
        
        # Supprimer les doublons et trier
        $Departments = $Departments | Sort-Object | Get-Unique
        
        Write-Host "$($Departments.Count) département(s) unique(s) trouvé(s)" -ForegroundColor Green
        
        # Afficher les départements trouvés pour diagnostic
        if ($Departments.Count -gt 0) {
            Write-Host "`nDépartements trouvés :" -ForegroundColor Cyan
            foreach ($Dept in $Departments) {
                Write-Host "  - '$Dept'" -ForegroundColor White
            }
        }
        
        return $Departments
        
    } catch {
        Write-Error "Erreur lors de la récupération des départements : $($_.Exception.Message)"
        Write-Host "Vérifiez que vous avez les permissions nécessaires pour lire les informations utilisateur." -ForegroundColor Yellow
        return @()
    }
}

# Fonction principale pour sélectionner les utilisateurs
function Select-Users {
    Write-Host "`n=== CONFIGURATION DES UTILISATEURS CIBLES ===" -ForegroundColor Yellow
    Write-Host "Choisissez le mode de recherche des délégations :" -ForegroundColor Cyan
    Write-Host ""
    
    # Liste par défaut
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
    
    Write-Host "Options de recherche :" -ForegroundColor Yellow
    Write-Host "  [1] Recherche par utilisateurs spécifiques" -ForegroundColor White
    Write-Host "  [2] Recherche par département" -ForegroundColor White
    Write-Host "  [3] Annuler" -ForegroundColor White
    
    do {
        $searchMode = Read-Host "Votre choix (1-3)"
    } while ($searchMode -notin @("1", "2", "3"))
    
    switch ($searchMode) {
        "1" {
            Write-Host "`n=== RECHERCHE PAR UTILISATEURS SPÉCIFIQUES ===" -ForegroundColor Yellow
            Write-Host "Liste actuelle des utilisateurs :" -ForegroundColor Green
            for ($i = 0; $i -lt $defaultUsers.Count; $i++) {
                Write-Host "  [$($i+1)] $($defaultUsers[$i])" -ForegroundColor White
            }
            
            Write-Host ""
            Write-Host "Options :" -ForegroundColor Yellow
            Write-Host "  [1] Utiliser la liste actuelle" -ForegroundColor White
            Write-Host "  [2] Modifier la liste" -ForegroundColor White
            Write-Host "  [3] Annuler" -ForegroundColor White
            
            do {
                $choice = Read-Host "Votre choix (1-3)"
            } while ($choice -notin @("1", "2", "3"))
            
            switch ($choice) {
                "1" {
                    return $defaultUsers
                }
                "2" {
                    Write-Host "`nSaisie des nouveaux utilisateurs :" -ForegroundColor Green
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
                        Write-Error "Aucun email valide saisi. Utilisation de la liste par défaut."
                        return $defaultUsers
                    }
                    
                    return $validEmails
                }
                "3" {
                    Write-Host "Opération annulée." -ForegroundColor Yellow
                    return @()
                }
            }
        }
        "2" {
            Write-Host "`n=== RECHERCHE PAR DÉPARTEMENT ===" -ForegroundColor Yellow
            
            # Récupérer la liste des départements disponibles
            Write-Host "Récupération des départements disponibles..." -ForegroundColor Cyan
            $AvailableDepartments = Get-AvailableDepartments
            
            if ($AvailableDepartments.Count -eq 0) {
                Write-Error "Aucun département trouvé dans Azure AD."
                Write-Host "`nAlternatives disponibles :" -ForegroundColor Yellow
                Write-Host "  [1] Saisir un département manuellement" -ForegroundColor White
                Write-Host "  [2] Annuler" -ForegroundColor White
                
                do {
                    $altChoice = Read-Host "Votre choix (1-2)"
                } while ($altChoice -notin @("1", "2"))
                
                switch ($altChoice) {
                    "1" {
                        $ManualDepartment = Read-Host "Saisissez le nom du département"
                        if ($ManualDepartment.Trim() -ne "") {
                            Write-Host "Département saisi : $ManualDepartment" -ForegroundColor Green
                            return Get-UsersByDepartment -Department $ManualDepartment.Trim()
                        } else {
                            Write-Error "Nom de département vide."
                            return @()
                        }
                    }
                    "2" {
                        Write-Host "Opération annulée." -ForegroundColor Yellow
                        return @()
                    }
                }
            }
            
            if ($AvailableDepartments.Count -gt 0) {
                Write-Host "`nDépartements disponibles :" -ForegroundColor Green
                for ($i = 0; $i -lt $AvailableDepartments.Count; $i++) {
                    Write-Host "  [$($i+1)] $($AvailableDepartments[$i])" -ForegroundColor White
                }
            }
            
            Write-Host ""
            Write-Host "Options :" -ForegroundColor Yellow
            if ($AvailableDepartments.Count -gt 0) {
                Write-Host "  [1-$(($AvailableDepartments.Count))] Sélectionner un département" -ForegroundColor White
                Write-Host "  [$(($AvailableDepartments.Count + 1))] Saisir un département manuellement" -ForegroundColor White
                Write-Host "  [$(($AvailableDepartments.Count + 2))] Annuler" -ForegroundColor White
                
                do {
                    $choice = Read-Host "Votre choix (1-$(($AvailableDepartments.Count + 2)))"
                } while ($choice -notin @(1..($AvailableDepartments.Count + 2)))
                
                if ($choice -le $AvailableDepartments.Count) {
                    # Sélection d'un département de la liste
                    $SelectedDepartment = $AvailableDepartments[$choice - 1]
                    Write-Host "Département sélectionné : $SelectedDepartment" -ForegroundColor Green
                    return Get-UsersByDepartment -Department $SelectedDepartment
                }
                elseif ($choice -eq ($AvailableDepartments.Count + 1)) {
                    # Saisie manuelle d'un département
                    $ManualDepartment = Read-Host "Saisissez le nom du département"
                    if ($ManualDepartment.Trim() -ne "") {
                        Write-Host "Département saisi : $ManualDepartment" -ForegroundColor Green
                        return Get-UsersByDepartment -Department $ManualDepartment.Trim()
                    } else {
                        Write-Error "Nom de département vide."
                        return @()
                    }
                }
                else {
                    # Annulation
                    Write-Host "Opération annulée." -ForegroundColor Yellow
                    return @()
                }
            }
        }
        "3" {
            Write-Host "Opération annulée par l'utilisateur." -ForegroundColor Yellow
            return @()
        }
    }
}

# Fonction pour récupérer les informations de département d'un utilisateur
function Get-UserDepartment {
    param([string]$UserEmail)
    
    try {
        $User = Get-MgUser -Filter "userPrincipalName eq '$UserEmail'" -ErrorAction SilentlyContinue
        if ($User) {
            return $User.Department
        }
    } catch {
        # Ignorer les erreurs silencieusement
    }
    return "Non défini"
}

# ===== SCRIPT PRINCIPAL =====

# Récupérer les utilisateurs
$TargetUsers = Select-Users

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

# Générer le rapport par utilisateur avec informations de département
foreach ($User in ($UserDelegations.Keys | Sort-Object)) {
    $UserDelegations[$User] = $UserDelegations[$User] | Sort-Object BoiteAuxLettres | Get-Unique -AsString
    
    # Récupérer le département de l'utilisateur
    $UserDepartment = Get-UserDepartment -UserEmail $User
    
    $ReportContent += "$User (Département: $UserDepartment) :"
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

