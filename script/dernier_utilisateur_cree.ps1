#Requires -Version 7.0

<#
.SYNOPSIS
    Script pour récupérer le dernier compte créé sur Microsoft Entra ID.
    Version optimisée pour macOS.

.DESCRIPTION
    Ce script se connecte à Microsoft Graph et récupère le dernier utilisateur créé
    avec ses informations : nom, prénom, département et date de création.
    
    Format de sortie : Nom Prenom Departement date de création

.PARAMETER ShowDetails
    Afficher des détails supplémentaires sur l'utilisateur (par défaut : false)

.EXAMPLE
    .\dernier_utilisateur_cree.ps1

.EXAMPLE
    .\dernier_utilisateur_cree.ps1 -ShowDetails

.NOTES
    Auteur: Assistant IA
    Version: 1.0 - macOS Compatible
    Date: $(Get-Date -Format "yyyy-MM-dd")
    
    Prérequis:
    - Module Microsoft.Graph
    - Connexion administrateur Microsoft 365
    - PowerShell Core 7.0+ (recommandé sur macOS)
#>

param(
    [switch]$ShowDetails = $false
)

function Test-MicrosoftGraphConnection {
    try {
        # Vérifier si on est connecté à Microsoft Graph
        $context = Get-MgContext -ErrorAction Stop
        return ($null -ne $context)
    }
    catch {
        return $false
    }
}

function Connect-ToMicrosoftGraph {
    try {
        # Se connecter à Microsoft Graph avec les permissions nécessaires
        Connect-MgGraph -Scopes "User.Read.All", "Directory.Read.All" -NoWelcome
        return $true
    }
    catch {
        return $false
    }
}

function Get-LastCreatedUser {
    try {
        # Récupérer les utilisateurs triés par date de création (le plus récent en premier)
        $users = Get-MgUser -All -Property @(
            "id", "displayName", "givenName", "surname", "userPrincipalName",
            "department", "jobTitle", "accountEnabled", "createdDateTime",
            "userType", "mail"
        ) | Where-Object {
            # Filtrer pour ne garder que les utilisateurs membres (pas les invités)
            $_.UserType -eq "Member" -and 
            $_.UserPrincipalName -notlike "*#EXT#*" -and
            $null -ne $_.CreatedDateTime
        } | Sort-Object CreatedDateTime -Descending | Select-Object -First 1
        
        return $users
    }
    catch {
        return $null
    }
}

function Format-UserOutput {
    param(
        [object]$User,
        [bool]$ShowDetails
    )
    
    # Extraire les informations
    $firstName = if ($User.GivenName) { $User.GivenName.Trim() } else { "Non renseigné" }
    $lastName = if ($User.Surname) { $User.Surname.Trim() } else { "Non renseigné" }
    $department = if ($User.Department) { $User.Department.Trim() } else { "Non renseigné" }
    $createdDate = if ($User.CreatedDateTime) { 
        $User.CreatedDateTime.ToString("dd/MM/yyyy") 
    } else { 
        "Date inconnue" 
    }
    
    # Format de sortie demandé : Nom Prenom Departement date de création
    $output = "$lastName $firstName $department $createdDate"
    
    if ($ShowDetails) {
        Write-Host ""
        Write-Host "DÉTAILS DE L'UTILISATEUR"
        Write-Host ("=" * 50)
        Write-Host "Nom complet: $($User.DisplayName)"
        Write-Host "Email: $($User.Mail)"
        Write-Host "Poste: $(if ($User.JobTitle) { $User.JobTitle } else { 'Non renseigné' })"
        Write-Host "Statut: $(if ($User.AccountEnabled) { 'Actif' } else { 'Désactivé' })"
        Write-Host "Date de création: $createdDate"
        Write-Host ("=" * 50)
        Write-Host ""
    }
    
    return $output
}

# Script principal
function Main {
    # Importer le module Microsoft.Graph
    try {
        Import-Module Microsoft.Graph.Users -Force
    }
    catch {
        exit 1
    }
    
    # Vérifier la connexion Microsoft Graph
    if (-not (Test-MicrosoftGraphConnection)) {
        if (-not (Connect-ToMicrosoftGraph)) {
            exit 1
        }
    }
    
    # Récupérer le dernier utilisateur créé
    $lastUser = Get-LastCreatedUser
    
    if (-not $lastUser) {
        exit 1
    }
    
    # Formater et afficher le résultat
    $formattedOutput = Format-UserOutput -User $lastUser -ShowDetails $ShowDetails
    
    # Afficher le résultat dans le format demandé
    Write-Host $formattedOutput
    
    # Déconnexion
    try {
        Disconnect-MgGraph | Out-Null
    }
    catch {
        # Ignorer les erreurs de déconnexion
    }
}

# Gestion des erreurs globales
trap {
    # Tentative de déconnexion en cas d'erreur
    try {
        Disconnect-MgGraph | Out-Null
    }
    catch {
        # Ignorer les erreurs de déconnexion
    }
    
    exit 1
}

# Exécution du script principal
Main
