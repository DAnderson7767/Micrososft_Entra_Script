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

# Liste des utilisateurs dont on cherche les délégations
$TargetUsers = @(
    "sophie.runtz@lde.fr",
    "celine.risch@lde.fr",
    "sarah.merah@lde.fr",
    "maxime.klein@lde.fr",
    "monia.belebbed@lde.fr",
    "elodie.urban@lde.fr",
    "elisabeth.laux@lde.fr"
)

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

# Définir les services et leurs responsables principaux
$Services = @{
    "Numérique" = @{
        "Responsable" = "Céline Risch"
        "Email" = "celine.risch@lde.fr"
        "Utilisateurs" = @("celine.risch@lde.fr", "elodie.urban@lde.fr")
    }
    "Marketing & Communication" = @{
        "Responsable" = "Sarah Merah"
        "Email" = "sarah.merah@lde.fr" 
        "Utilisateurs" = @("sarah.merah@lde.fr", "maxime.klein@lde.fr")
    }
    "Comptabilité & RH" = @{
        "Responsable" = "Monia Belebbed"
        "Email" = "monia.belebbed@lde.fr"
        "Utilisateurs" = @("monia.belebbed@lde.fr", "Elisabeth.laux@lde.fr")
    }
    "Commercial & Export" = @{
        "Responsable" = "Sophie Runtz"
        "Email" = "sophie.runtz@lde.fr"
        "Utilisateurs" = @("sophie.runtz@lde.fr")
    }
    "PopLab" = @{
        "Responsable" = "Maxime Klein"
        "Email" = "maxime.klein@lde.fr"
        "Utilisateurs" = @("maxime.klein@lde.fr")
    }
}

# Fonction pour formater une adresse email avec son nom d'affichage
function Format-EmailEntry {
    param($Email, $DisplayName)
    
    if ($DisplayName -and $DisplayName -ne $Email -and $DisplayName -ne "") {
        return "- [$Email](mailto:$Email) ($DisplayName)"
    } else {
        return "- [$Email](mailto:$Email)"
    }
}

# Initialiser le contenu du rapport
$ReportContent = @()
$ReportContent += "# Rapport des Délégations par Service"
$ReportContent += ""
$ReportContent += "Généré le : $(Get-Date -Format 'dd/MM/yyyy à HH:mm')"
$ReportContent += ""

# Traiter chaque service
foreach ($ServiceName in $Services.Keys) {
    $Service = $Services[$ServiceName]
    $ServiceUsers = $Service.Utilisateurs
    
    # Récupérer toutes les délégations pour les utilisateurs de ce service
    $ServiceDelegations = @()
    
    foreach ($Delegation in $Results) {
        $HasAccess = $false
        
        # Vérifier si un utilisateur du service a accès (comparaison insensible à la casse)
        foreach ($User in $ServiceUsers) {
            if ($Delegation.FullAccess.ToLower() -like "*$($User.ToLower())*" -or 
                $Delegation.SendAs.ToLower() -like "*$($User.ToLower())*" -or 
                $Delegation.SendOnBehalf.ToLower() -like "*$($User.ToLower())*") {
                $HasAccess = $true
                break
            }
        }
        
        if ($HasAccess) {
            $ServiceDelegations += $Delegation
        }
    }
    
    # Générer la section du service si des délégations existent
    if ($ServiceDelegations.Count -gt 0) {
        $ReportContent += "## $ServiceName"
        $ReportContent += ""
        $ReportContent += "Template de base: **$($Service.Responsable)**"
        $ReportContent += ""
        
        # Trier les délégations par ordre alphabétique
        $SortedDelegations = $ServiceDelegations | Sort-Object BoiteAuxLettres
        
        foreach ($Delegation in $SortedDelegations) {
            $FormattedEntry = Format-EmailEntry -Email $Delegation.BoiteAuxLettres -DisplayName $Delegation.NomAffichage
            $ReportContent += $FormattedEntry
        }
        
        $ReportContent += ""
        $ReportContent += "---"
        $ReportContent += ""
    }
}

# Ajouter une section avec toutes les délégations non catégorisées (au cas où)
$AllCategorizedEmails = @()
foreach ($Service in $Services.Values) {
    foreach ($Delegation in $Results) {
        foreach ($User in $Service.Utilisateurs) {
            if ($Delegation.FullAccess.ToLower() -like "*$($User.ToLower())*" -or 
                $Delegation.SendAs.ToLower() -like "*$($User.ToLower())*" -or 
                $Delegation.SendOnBehalf.ToLower() -like "*$($User.ToLower())*") {
                $AllCategorizedEmails += $Delegation.BoiteAuxLettres
            }
        }
    }
}

$UncategorizedDelegations = $Results | Where-Object { $_.BoiteAuxLettres -notin $AllCategorizedEmails }

if ($UncategorizedDelegations.Count -gt 0) {
    $ReportContent += "## Autres Délégations"
    $ReportContent += ""
    $ReportContent += "Délégations non catégorisées :"
    $ReportContent += ""
    
    foreach ($Delegation in ($UncategorizedDelegations | Sort-Object BoiteAuxLettres)) {
        $FormattedEntry = Format-EmailEntry -Email $Delegation.BoiteAuxLettres -DisplayName $Delegation.NomAffichage
        $ReportContent += $FormattedEntry
    }
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
