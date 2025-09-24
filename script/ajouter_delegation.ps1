#!/usr/bin/env pwsh

<#
.SYNOPSIS
    Script interactif pour ajouter des délégations Exchange Online
    
.DESCRIPTION
    Ce script demande à l'utilisateur :
    1. L'adresse email de l'utilisateur qui recevra la délégation
    2. L'adresse email de la boîte aux lettres cible
    3. Le type de délégation (Full Access, Send As, Send on Behalf)
    Puis applique la délégation demandée
    
.NOTES
    Prérequis :
    1. PowerShell Core : brew install --cask powershell
    2. Module : Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser
    3. Compte administrateur Exchange Online
    4. Exécuter avec : pwsh ./script/ajouter_delegation.ps1
#>

# Vérifier la présence du module
if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
    Write-Error "Module ExchangeOnlineManagement non trouve. Installez-le avec :"
    Write-Host "Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser" -ForegroundColor Yellow
    exit 1
}

# Importer le module
Import-Module ExchangeOnlineManagement

# Connexion à Exchange Online
Write-Host "Connexion a Exchange Online..." -ForegroundColor Green
Connect-ExchangeOnline

Write-Host "`n=== AJOUT DE DELEGATION EXCHANGE ONLINE ===" -ForegroundColor Cyan
Write-Host "Ce script vous permet d'ajouter des delegations sur des boites aux lettres." -ForegroundColor White
Write-Host ""

# Fonction pour valider une adresse email (plus permissive)
function Test-EmailAddress {
    param($Email)
    # Validation simple : doit contenir @ et au moins un point après @
    if ($Email -match '^[^@]+@[^@]+\.[^@]+$') {
        return $true
    }
    return $false
}

# Fonction pour vérifier si une boîte aux lettres existe
function Test-MailboxExists {
    param($Email)
    try {
        $Mailbox = Get-Mailbox -Identity $Email -ErrorAction Stop
        return $true
    }
    catch {
        return $false
    }
}

# Fonction pour vérifier si un utilisateur existe
function Test-UserExists {
    param($Email)
    try {
        $User = Get-Recipient -Identity $Email -ErrorAction Stop
        return $true
    }
    catch {
        return $false
    }
}

# Demander l'adresse email de l'utilisateur qui recevra la délégation
do {
    $UserEmail = Read-Host "Entrez l'adresse email de l'utilisateur qui recevra la delegation"
    
    if (-not (Test-EmailAddress -Email $UserEmail)) {
        Write-Host "ERREUR: Format d'adresse email invalide." -ForegroundColor Red
        continue
    }
    
    if (-not (Test-UserExists -Email $UserEmail)) {
        Write-Host "ERREUR: L'utilisateur '$UserEmail' n'existe pas dans l'organisation." -ForegroundColor Red
        continue
    }
    
    break
} while ($true)

Write-Host "Utilisateur valide: $UserEmail" -ForegroundColor Green

# Demander l'adresse email de la boîte aux lettres cible
do {
    $TargetMailbox = Read-Host "Entrez l'adresse email de la boite aux lettres cible"
    
    if (-not (Test-EmailAddress -Email $TargetMailbox)) {
        Write-Host "ERREUR: Format d'adresse email invalide." -ForegroundColor Red
        continue
    }
    
    if (-not (Test-MailboxExists -Email $TargetMailbox)) {
        Write-Host "ERREUR: La boite aux lettres '$TargetMailbox' n'existe pas." -ForegroundColor Red
        continue
    }
    
    break
} while ($true)

Write-Host "Boite aux lettres valide: $TargetMailbox" -ForegroundColor Green

# Appliquer automatiquement toutes les permissions
$DelegationType = "All"
$DelegationName = "Toutes les permissions (Full Access + Send As + Send on Behalf)"
Write-Host "`nType de delegation: $DelegationName" -ForegroundColor Green

# Afficher un résumé avant application
Write-Host "`n=== RESUME DE LA DELEGATION ===" -ForegroundColor Cyan
Write-Host "Utilisateur: $UserEmail" -ForegroundColor White
Write-Host "Boite aux lettres cible: $TargetMailbox" -ForegroundColor White
Write-Host "Type de delegation: $DelegationName" -ForegroundColor White
Write-Host "`nApplication automatique de toutes les permissions..." -ForegroundColor Yellow

# Appliquer les délégations
Write-Host "`nApplication de la delegation..." -ForegroundColor Cyan

$SuccessCount = 0
$ErrorCount = 0

try {
    switch ($DelegationType) {
        "FullAccess" {
            Write-Host "Ajout de la permission Full Access..." -ForegroundColor Yellow
            Add-MailboxPermission -Identity $TargetMailbox -User $UserEmail -AccessRights FullAccess -InheritanceType All
            $SuccessCount++
            Write-Host "SUCCES: Permission Full Access ajoutee." -ForegroundColor Green
        }
        
        "SendAs" {
            Write-Host "Ajout de la permission Send As..." -ForegroundColor Yellow
            Add-RecipientPermission -Identity $TargetMailbox -Trustee $UserEmail -AccessRights SendAs -Confirm:$false
            $SuccessCount++
            Write-Host "SUCCES: Permission Send As ajoutee." -ForegroundColor Green
        }
        
        "SendOnBehalf" {
            Write-Host "Ajout de la permission Send on Behalf..." -ForegroundColor Yellow
            Set-Mailbox -Identity $TargetMailbox -GrantSendOnBehalfTo @{Add=$UserEmail}
            $SuccessCount++
            Write-Host "SUCCES: Permission Send on Behalf ajoutee." -ForegroundColor Green
        }
        
        "All" {
            Write-Host "Ajout de toutes les permissions..." -ForegroundColor Yellow
            
            # Full Access
            try {
                Add-MailboxPermission -Identity $TargetMailbox -User $UserEmail -AccessRights FullAccess -InheritanceType All
                Write-Host "SUCCES: Permission Full Access ajoutee." -ForegroundColor Green
                $SuccessCount++
            }
            catch {
                Write-Host "ERREUR: Impossible d'ajouter Full Access - $($_.Exception.Message)" -ForegroundColor Red
                $ErrorCount++
            }
            
            # Send As
            try {
                Add-RecipientPermission -Identity $TargetMailbox -Trustee $UserEmail -AccessRights SendAs -Confirm:$false
                Write-Host "SUCCES: Permission Send As ajoutee." -ForegroundColor Green
                $SuccessCount++
            }
            catch {
                Write-Host "ERREUR: Impossible d'ajouter Send As - $($_.Exception.Message)" -ForegroundColor Red
                $ErrorCount++
            }
            
            # Send on Behalf
            try {
                Set-Mailbox -Identity $TargetMailbox -GrantSendOnBehalfTo @{Add=$UserEmail}
                Write-Host "SUCCES: Permission Send on Behalf ajoutee." -ForegroundColor Green
                $SuccessCount++
            }
            catch {
                Write-Host "ERREUR: Impossible d'ajouter Send on Behalf - $($_.Exception.Message)" -ForegroundColor Red
                $ErrorCount++
            }
        }
    }
}
catch {
    Write-Host "ERREUR: Impossible d'appliquer la delegation - $($_.Exception.Message)" -ForegroundColor Red
    $ErrorCount++
}

# Résumé final
Write-Host "`n=== RESUME DE L'OPERATION ===" -ForegroundColor Cyan
if ($SuccessCount -gt 0) {
    Write-Host "SUCCES: $SuccessCount permission(s) ajoutee(s) avec succes." -ForegroundColor Green
}
if ($ErrorCount -gt 0) {
    Write-Host "ERREURS: $ErrorCount erreur(s) rencontree(s)." -ForegroundColor Red
}

# Vérifier les délégations actuelles
Write-Host "`n=== VERIFICATION DES DELEGATIONS ACTUELLES ===" -ForegroundColor Cyan
Write-Host "Delegations actuelles pour $TargetMailbox :" -ForegroundColor White

try {
    # Full Access
    $FullAccess = Get-MailboxPermission -Identity $TargetMailbox | Where-Object { 
        $_.User -like "*$UserEmail*" -and $_.IsInherited -eq $false 
    }
    if ($FullAccess) {
        Write-Host "- Full Access: OUI" -ForegroundColor Green
    } else {
        Write-Host "- Full Access: NON" -ForegroundColor Gray
    }
    
    # Send As
    $SendAs = Get-RecipientPermission -Identity $TargetMailbox | Where-Object { 
        $_.Trustee -like "*$UserEmail*" 
    }
    if ($SendAs) {
        Write-Host "- Send As: OUI" -ForegroundColor Green
    } else {
        Write-Host "- Send As: NON" -ForegroundColor Gray
    }
    
    # Send on Behalf
    $MailboxInfo = Get-Mailbox -Identity $TargetMailbox
    $SendOnBehalf = $MailboxInfo.GrantSendOnBehalfTo | Where-Object { $_ -like "*$UserEmail*" }
    if ($SendOnBehalf) {
        Write-Host "- Send on Behalf: OUI" -ForegroundColor Green
    } else {
        Write-Host "- Send on Behalf: NON" -ForegroundColor Gray
    }
}
catch {
    Write-Host "ERREUR: Impossible de verifier les delegations actuelles." -ForegroundColor Red
}

# Déconnexion propre
Disconnect-ExchangeOnline -Confirm:$false
Write-Host "`nDeconnexion terminee." -ForegroundColor Green
Write-Host "`nScript termine. Fermeture automatique dans 3 secondes..." -ForegroundColor Yellow
Start-Sleep -Seconds 1
