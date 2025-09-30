#Requires -Version 7.0

<#
.SYNOPSIS
    Script pour exporter la liste complète des utilisateurs Microsoft Graph avec leurs informations détaillées.
    Version optimisée pour macOS.

.DESCRIPTION
    Ce script se connecte à Microsoft Graph et récupère tous les utilisateurs avec leurs informations :
    - Nom et prénom
    - Département
    - Poste/Titre
    - Email
    - Statut du compte
    - Date de création

    Les données sont exportées dans un fichier texte formaté et un fichier CSV.

.PARAMETER OutputPath
    Chemin de sortie pour les fichiers générés (par défaut : répertoire courant)

.PARAMETER IncludeDisabled
    Inclure les comptes désactivés dans l'export (par défaut : false)

.PARAMETER IncludeSharedMailboxes
    Inclure les boîtes partagées dans l'export (par défaut : false - les boîtes partagées sont exclues)

.EXAMPLE
    .\export_utilisateurs_macos.ps1

.EXAMPLE
    .\export_utilisateurs_macos.ps1 -IncludeDisabled

.EXAMPLE
    .\export_utilisateurs_macos.ps1 -IncludeSharedMailboxes

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
    [string]$OutputPath = ".",
    [switch]$IncludeDisabled = $false,
    [switch]$IncludeSharedMailboxes = $false
)

# Configuration des couleurs pour les messages
$Colors = @{
    Success = "Green"
    Warning = "Yellow"
    Error = "Red"
    Info = "Cyan"
    Header = "Magenta"
}

function Write-ColorOutput {
    param(
        [string]$Message,
        [string]$Color = "White"
    )
    Write-Host $Message -ForegroundColor $Color
}

function Test-MicrosoftGraphConnection {
    try {
        # Vérifier si on est connecté à Microsoft Graph
        $context = Get-MgContext -ErrorAction Stop
        return ($context -ne $null)
    }
    catch {
        return $false
    }
}

function Connect-ToMicrosoftGraph {
    Write-ColorOutput "🔐 Connexion à Microsoft Graph..." $Colors.Info
    
    try {
        # Se connecter à Microsoft Graph avec les permissions nécessaires
        Connect-MgGraph -Scopes "User.Read.All", "Directory.Read.All" -NoWelcome
        Write-ColorOutput "✅ Connexion réussie à Microsoft Graph" $Colors.Success
        return $true
    }
    catch {
        Write-ColorOutput "❌ Erreur de connexion à Microsoft Graph : $($_.Exception.Message)" $Colors.Error
        Write-ColorOutput "💡 Assurez-vous d'avoir installé le module Microsoft.Graph et d'avoir les permissions nécessaires" $Colors.Warning
        return $false
    }
}

function Get-AllUsers {
    param(
        [bool]$IncludeDisabledUsers
    )
    
    Write-ColorOutput "📊 Récupération de la liste des utilisateurs..." $Colors.Info
    
    try {
        # Récupérer tous les utilisateurs avec Microsoft Graph
        $users = @()
        $pageSize = 999
        $skipToken = $null
        
        do {
            $params = @{
                All = $true
                PageSize = $pageSize
                Property = @(
                    "id", "displayName", "givenName", "surname", "userPrincipalName",
                    "department", "jobTitle", "accountEnabled", "createdDateTime",
                    "userType", "mail"
                )
            }
            
            if ($skipToken) {
                $params.SkipToken = $skipToken
            }
            
            $result = Get-MgUser @params
            
            if ($result) {
                $users += $result
                $skipToken = $result | Select-Object -Last 1 | ForEach-Object { $_.AdditionalProperties.'@odata.nextLink' }
            }
        } while ($skipToken)
        
        # Filtrer les utilisateurs selon les critères
        if ($IncludeSharedMailboxes) {
            # Inclure tous les utilisateurs (y compris les boîtes partagées)
            $filteredUsers = $users | Where-Object {
                $_.UserType -eq "Member" -and 
                $_.UserPrincipalName -notlike "*#EXT#*" -and
                ($IncludeDisabledUsers -or $_.AccountEnabled -eq $true)
            }
            Write-ColorOutput "📋 Mode : Inclure les boîtes partagées" $Colors.Info
        } else {
            # Exclure les boîtes partagées (mode par défaut)
            $filteredUsers = $users | Where-Object {
                $_.UserType -eq "Member" -and 
                $_.UserPrincipalName -notlike "*#EXT#*" -and
                $_.UserPrincipalName -notlike "*_*" -and  # Exclure les comptes avec underscore (souvent des boîtes partagées)
                $_.DisplayName -notlike "*partagé*" -and  # Exclure les comptes avec "partagé" dans le nom
                $_.DisplayName -notlike "*shared*" -and   # Exclure les comptes avec "shared" dans le nom
                $_.DisplayName -notlike "*mailbox*" -and  # Exclure les comptes avec "mailbox" dans le nom
                $_.DisplayName -notlike "*boite*" -and    # Exclure les comptes avec "boite" dans le nom
                $_.DisplayName -notlike "*box*" -and      # Exclure les comptes avec "box" dans le nom
                $_.JobTitle -notlike "*partagé*" -and     # Exclure les postes avec "partagé"
                $_.JobTitle -notlike "*shared*" -and      # Exclure les postes avec "shared"
                $_.JobTitle -notlike "*mailbox*" -and     # Exclure les postes avec "mailbox"
                $_.Department -notlike "*partagé*" -and   # Exclure les départements avec "partagé"
                $_.Department -notlike "*shared*" -and    # Exclure les départements avec "shared"
                $_.Department -notlike "*mailbox*" -and   # Exclure les départements avec "mailbox"
                ($IncludeDisabledUsers -or $_.AccountEnabled -eq $true)
            }
            Write-ColorOutput "📋 Mode : Exclure les boîtes partagées" $Colors.Info
        }
        
        Write-ColorOutput "✅ $($filteredUsers.Count) utilisateurs récupérés" $Colors.Success
        return $filteredUsers
    }
    catch {
        Write-ColorOutput "❌ Erreur lors de la récupération des utilisateurs : $($_.Exception.Message)" $Colors.Error
        return @()
    }
}

function Format-UserData {
    param(
        [array]$Users
    )
    
    Write-ColorOutput "🔄 Formatage des données utilisateurs..." $Colors.Info
    
    $formattedUsers = @()
    $totalUsers = $Users.Count
    $currentUser = 0
    
    foreach ($user in $Users) {
        $currentUser++
        $progress = [math]::Round(($currentUser / $totalUsers) * 100, 1)
        Write-Progress -Activity "Formatage des données" -Status "Utilisateur $currentUser sur $totalUsers" -PercentComplete $progress
        
        # Extraire les informations
        $firstName = if ($user.GivenName) { $user.GivenName.Trim() } else { "Non renseigné" }
        $lastName = if ($user.Surname) { $user.Surname.Trim() } else { "Non renseigné" }
        $department = if ($user.Department) { $user.Department.Trim() } else { "Non renseigné" }
        $jobTitle = if ($user.JobTitle) { $user.JobTitle.Trim() } else { "Non renseigné" }
        $email = if ($user.Mail) { $user.Mail.Trim() } elseif ($user.UserPrincipalName) { $user.UserPrincipalName.Trim() } else { "Non renseigné" }
        $displayName = if ($user.DisplayName) { $user.DisplayName.Trim() } else { "Non renseigné" }
        $accountStatus = if ($user.AccountEnabled) { "Actif" } else { "Désactivé" }
        $createdDate = if ($user.CreatedDateTime) { 
            $user.CreatedDateTime.ToString("dd/MM/yyyy") 
        } else { 
            "Date inconnue" 
        }
        
        $formattedUsers += [PSCustomObject]@{
            Nom = $lastName
            Prenom = $firstName
            NomComplet = $displayName
            Email = $email
            Departement = $department
            Poste = $jobTitle
            Statut = $accountStatus
            DateCreation = $createdDate
        }
    }
    
    Write-Progress -Activity "Formatage des données" -Completed
    return $formattedUsers
}

function Export-ToTextFile {
    param(
        [array]$Users,
        [string]$FilePath
    )
    
    Write-ColorOutput "📝 Génération du fichier texte..." $Colors.Info
    
    try {
        $content = @()
        $content += "=" * 80
        $content += "RAPPORT DES UTILISATEURS MICROSOFT GRAPH"
        $content += "Généré le : $(Get-Date -Format 'dd/MM/yyyy à HH:mm:ss')"
        $content += "Total d'utilisateurs : $($Users.Count)"
        $content += "=" * 80
        $content += ""
        
        # Grouper par département
        $usersByDepartment = $Users | Group-Object -Property Departement | Sort-Object Name
        
        foreach ($department in $usersByDepartment) {
            $deptName = if ($department.Name -eq "Non renseigné") { "Département non renseigné" } else { $department.Name }
            
            $content += "📁 $deptName ($($department.Count) utilisateur(s))"
            $content += "-" * 60
            
            foreach ($user in $department.Group | Sort-Object Nom, Prenom) {
                $content += "👤 $($user.Prenom) $($user.Nom)"
                $content += "   📧 Email: $($user.Email)"
                $content += "   💼 Poste: $($user.Poste)"
                $content += "   📊 Statut: $($user.Statut)"
                $content += ""
            }
            $content += ""
        }
        
        # Statistiques
        $content += "=" * 80
        $content += "STATISTIQUES"
        $content += "=" * 80
        $content += "Total d'utilisateurs: $($Users.Count)"
        $content += "Utilisateurs actifs: $(($Users | Where-Object { $_.Statut -eq 'Actif' }).Count)"
        $content += "Utilisateurs désactivés: $(($Users | Where-Object { $_.Statut -eq 'Désactivé' }).Count)"
        $content += "Départements: $($usersByDepartment.Count)"
        $content += ""
        
        # Top 5 des départements
        $content += "Top 5 des départements:"
        $topDepartments = $usersByDepartment | Sort-Object Count -Descending | Select-Object -First 5
        foreach ($dept in $topDepartments) {
            $deptName = if ($dept.Name -eq "Non renseigné") { "Non renseigné" } else { $dept.Name }
            $content += "  • $deptName : $($dept.Count) utilisateur(s)"
        }
        
        # Écrire le fichier
        $content | Out-File -FilePath $FilePath -Encoding UTF8
        Write-ColorOutput "✅ Fichier texte généré : $FilePath" $Colors.Success
    }
    catch {
        Write-ColorOutput "❌ Erreur lors de la génération du fichier texte : $($_.Exception.Message)" $Colors.Error
    }
}

function Export-ToCSVFile {
    param(
        [array]$Users,
        [string]$FilePath
    )
    
    Write-ColorOutput "📊 Génération du fichier CSV..." $Colors.Info
    
    try {
        $Users | Export-Csv -Path $FilePath -NoTypeInformation -Encoding UTF8
        Write-ColorOutput "✅ Fichier CSV généré : $FilePath" $Colors.Success
    }
    catch {
        Write-ColorOutput "❌ Erreur lors de la génération du fichier CSV : $($_.Exception.Message)" $Colors.Error
    }
}

# Script principal
function Main {
    Write-ColorOutput "🚀 DÉMARRAGE DU SCRIPT D'EXPORT DES UTILISATEURS (macOS)" $Colors.Header
    Write-ColorOutput ("=" * 60) $Colors.Header
    
    # Importer le module Microsoft.Graph
    try {
        Import-Module Microsoft.Graph.Users -Force
        Write-ColorOutput "✅ Module Microsoft.Graph importé avec succès" $Colors.Success
    }
    catch {
        Write-ColorOutput "❌ Erreur lors de l'import du module Microsoft.Graph : $($_.Exception.Message)" $Colors.Error
        Write-ColorOutput "💡 Exécutez d'abord : pwsh -Command 'Install-Module Microsoft.Graph -Scope CurrentUser'" $Colors.Warning
        exit 1
    }
    
    # Vérifier la connexion Microsoft Graph
    if (-not (Test-MicrosoftGraphConnection)) {
        if (-not (Connect-ToMicrosoftGraph)) {
            Write-ColorOutput "❌ Impossible de se connecter à Microsoft Graph. Arrêt du script." $Colors.Error
            exit 1
        }
    } else {
        Write-ColorOutput "✅ Déjà connecté à Microsoft Graph" $Colors.Success
    }
    
    # Récupérer les utilisateurs
    $users = Get-AllUsers -IncludeDisabledUsers $IncludeDisabled
    
    if ($users.Count -eq 0) {
        Write-ColorOutput "❌ Aucun utilisateur trouvé. Arrêt du script." $Colors.Error
        exit 1
    }
    
    # Formater les données
    $formattedUsers = Format-UserData -Users $users
    
    # Générer les noms de fichiers
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $textFile = Join-Path $OutputPath "Utilisateurs_Graph_$timestamp.txt"
    $csvFile = Join-Path $OutputPath "Utilisateurs_Graph_$timestamp.csv"
    
    # Exporter les données
    Export-ToTextFile -Users $formattedUsers -FilePath $textFile
    Export-ToCSVFile -Users $formattedUsers -FilePath $csvFile
    
    # Résumé final
    Write-ColorOutput "" $Colors.Info
    Write-ColorOutput ("=" * 60) $Colors.Header
    Write-ColorOutput "✅ EXPORT TERMINÉ AVEC SUCCÈS" $Colors.Success
    Write-ColorOutput ("=" * 60) $Colors.Header
    Write-ColorOutput "📊 Utilisateurs exportés : $($formattedUsers.Count)" $Colors.Info
    Write-ColorOutput "📝 Fichier texte : $textFile" $Colors.Info
    Write-ColorOutput "📊 Fichier CSV : $csvFile" $Colors.Info
    Write-ColorOutput "" $Colors.Info
    
    # Déconnexion
    try {
        Disconnect-MgGraph | Out-Null
        Write-ColorOutput "🔐 Déconnexion de Microsoft Graph réussie" $Colors.Success
    }
    catch {
        Write-ColorOutput "⚠️ Erreur lors de la déconnexion de Microsoft Graph" $Colors.Warning
    }
}

# Gestion des erreurs globales
trap {
    Write-ColorOutput "❌ Erreur inattendue : $($_.Exception.Message)" $Colors.Error
    Write-ColorOutput "📍 Ligne : $($_.InvocationInfo.ScriptLineNumber)" $Colors.Error
    
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
