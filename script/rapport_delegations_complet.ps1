#!/usr/bin/env pwsh

<#
.SYNOPSIS
    Script complet de recherche et génération de rapport des délégations
    
.DESCRIPTION
    Ce script :
    1. Recherche en lecture seule toutes les délégations possédées par les utilisateurs spécifiés
    2. Génère directement un rapport formaté organisé par service avec des liens mailto
    
.NOTES
    Prérequis sur macOS :
    1. PowerShell Core : brew install --cask powershell
    2. Module : Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser
    3. Exécuter avec : pwsh ./script/rapport_delegations_complet.ps1
#>

# Vérifier la présence du module
if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
    Write-Error "Module ExchangeOnlineManagement non trouvé. Installez-le avec :"
    Write-Host "Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser" -ForegroundColor Yellow
    exit 1
}

# Importer le module
Import-Module ExchangeOnlineManagement

# Connexion à Exchange Online
Write-Host "Connexion à Exchange Online..." -ForegroundColor Green
Connect-ExchangeOnline

# Interface en ligne de commande pour saisir les utilisateurs cibles (compatible macOS)
function Get-UserEmails {
    Write-Host "`n=== CONFIGURATION DES UTILISATEURS CIBLES ===" -ForegroundColor Yellow
    Write-Host "Saisissez les adresses email des utilisateurs dont vous voulez analyser les délégations." -ForegroundColor Cyan
    Write-Host "Une adresse par ligne. Tapez 'FIN' sur une ligne vide pour terminer." -ForegroundColor Cyan
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
                $input = Read-Host "Email $lineNumber"
                if ($input -ne "FIN" -and $input.Trim() -ne "") {
                    $emails += $input.Trim()
                    $lineNumber++
                }
            } while ($input -ne "FIN")
            
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
            Write-Host "Opération annulée par l'utilisateur." -ForegroundColor Yellow
            exit 0
        }
    }
}

# Récupérer les utilisateurs
$TargetUsers = Get-UserEmails

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
        $ReportContent += $Delegation.BoiteAuxLettres
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
Disconnect-ExchangeOnline -Confirm:$false
Write-Host "`nDeconnexion terminee." -ForegroundColor Green
