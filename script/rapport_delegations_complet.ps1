#!/usr/bin/env pwsh

<#
.SYNOPSIS
    Script optimisé de recherche et génération de rapport des délégations pour macOS
    
.DESCRIPTION
    Ce script recherche toutes les délégations possédées par les utilisateurs spécifiés
    et génère un rapport formaté organisé par utilisateur avec informations de département.
    
    Fonctionnalités :
    - Recherche par utilisateurs spécifiques ou par département
    - Interface utilisateur interactive
    - Gestion d'erreurs robuste
    - Optimisé pour macOS et PowerShell Core
    - Rapport formaté avec informations de département
    
.NOTES
    Prérequis sur macOS :
    1. PowerShell Core : brew install --cask powershell
    2. Modules requis :
       - Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser
       - Install-Module -Name Microsoft.Graph.Users -Scope CurrentUser
    3. Exécuter avec : pwsh ./script/rapport_delegations_complet.ps1
    
.AUTHOR
    Script optimisé pour macOS
#>

# Configuration globale
$Script:Config = @{
    OutputPath = "/Users/davidchiche/Desktop/Microsoft Entra/Rapport_Delegations_Formate.txt"
    CsvPath = "/Users/davidchiche/Desktop/Microsoft Entra/Delegations_Possedees_Report.csv"
    DefaultUsers = @(
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
        Write-ColorOutput "`nNote: Sur macOS, utilisez Microsoft.Graph au lieu d'AzureAD" "Cyan"
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

# ===== FONCTIONS DE NORMALISATION =====

function Normalize-Department {
    param([string]$Department)
    
    if ([string]::IsNullOrWhiteSpace($Department)) {
        return "Non défini"
    }
    
    # Nettoyer la chaîne
    $normalized = $Department.Trim()
    
    # Vérifier si c'est un email (contient @)
    if ($normalized -match "@") {
        return "DONNÉES INCORRECTES (Email)"
    }
    
    # Normaliser la casse et les espaces
    $normalized = $normalized -replace "\s+", " "  # Remplacer les espaces multiples par un seul
    $normalized = $normalized.ToLower()
    
    # Règles de normalisation spécifiques
    $normalizationRules = @{
        "chef de projet" = "Chef de Projet"
        "chefferie de projet" = "Chef de Projet"
        "relation client" = "Relations Clients"
        "relation clients" = "Relations Clients"
        "relations client" = "Relations Clients"
        "commercial et marketing" = "Commercial et Marketing"
        "développement informatique" = "Développement Informatique"
        "ressources humaines" = "Ressources Humaines"
        "comptabilité" = "Comptabilité"
        "direction" = "Direction"
        "production" = "Production"
        "technique" = "Technique"
        "marketing" = "Marketing"
        "commerce" = "Commerce"
        "export" = "Export"
        "informatique" = "Informatique"
        "numérique" = "Numérique"
        "lde" = "LDE"
    }
    
    # Appliquer les règles de normalisation
    if ($normalizationRules.ContainsKey($normalized)) {
        return $normalizationRules[$normalized]
    }
    
    # Si pas de règle spécifique, capitaliser la première lettre de chaque mot
    return (Get-Culture).TextInfo.ToTitleCase($normalized)
}

function Get-NormalizedDepartments {
    param([array]$AllUsers)
    
    try {
        $UsersWithDepartment = $AllUsers | Where-Object { 
            $_.Department -and 
            $_.Department.Trim() -ne "" -and 
            $null -ne $_.Department 
        }
        
        Write-ColorOutput "$($UsersWithDepartment.Count) utilisateur(s) avec département défini" "Yellow"
        
        if ($UsersWithDepartment.Count -eq 0) {
            Write-ColorOutput "Aucun utilisateur n'a de département défini dans Azure AD." "Yellow"
            return @()
        }
        
        # Normaliser tous les départements et collecter les statistiques
        $NormalizedDepartments = @{}
        $DepartmentStats = @{}
        $IncorrectData = @()
        
        foreach ($User in $UsersWithDepartment) {
            if ($User.Department -and $User.Department.Trim() -ne "") {
                $OriginalDept = $User.Department.Trim()
                $NormalizedDept = Normalize-Department -Department $OriginalDept
                
                # Identifier les données incorrectes
                if ($NormalizedDept -like "DONNÉES INCORRECTES*") {
                    $IncorrectData += @{
                        User = $User.UserPrincipalName
                        OriginalDepartment = $OriginalDept
                        Issue = "Email dans le champ département"
                    }
                    continue
                }
                
                # Compter les occurrences
                if (-not $DepartmentStats.ContainsKey($NormalizedDept)) {
                    $DepartmentStats[$NormalizedDept] = @{
                        Count = 0
                        Originals = @()
                        Users = @()
                    }
                }
                $DepartmentStats[$NormalizedDept].Count++
                if ($DepartmentStats[$NormalizedDept].Originals -notcontains $OriginalDept) {
                    $DepartmentStats[$NormalizedDept].Originals += $OriginalDept
                }
                $DepartmentStats[$NormalizedDept].Users += $User.UserPrincipalName
            }
        }
        
        # Afficher les données incorrectes
        if ($IncorrectData.Count -gt 0) {
            Write-ColorOutput "`n=== DONNÉES INCORRECTES À CORRIGER ===" "Red"
            foreach ($Issue in $IncorrectData) {
                Write-ColorOutput "ERREUR: $($Issue.User) : '$($Issue.OriginalDepartment)' - $($Issue.Issue)" "Red"
            }
            Write-ColorOutput ""
        }
        
        # Afficher les statistiques de normalisation
        Write-ColorOutput "=== NORMALISATION DES DÉPARTEMENTS ===" "Cyan"
        foreach ($Dept in ($DepartmentStats.Keys | Sort-Object)) {
            $Stats = $DepartmentStats[$Dept]
            Write-ColorOutput "`nDépartement normalisé: $Dept ($($Stats.Count) utilisateur(s))" "Green"
            
            if ($Stats.Originals.Count -gt 1) {
                Write-ColorOutput "   Normalisé depuis : $($Stats.Originals -join ', ')" "Yellow"
                Write-ColorOutput "   Utilisateurs concernés :" "Yellow"
                foreach ($User in $Stats.Users) {
                    Write-ColorOutput "     - $User" "White"
                }
            }
        }
        
        $Departments = $DepartmentStats.Keys | Sort-Object
        Write-ColorOutput "`n$($Departments.Count) département(s) unique(s) après normalisation" "Green"
        
        return $Departments
        
    }
    catch {
        Write-ColorOutput "Erreur lors de la normalisation des départements : $($_.Exception.Message)" "Red"
        return @()
    }
}

# ===== FONCTIONS DE RECHERCHE D'UTILISATEURS =====

function Get-AllUsers {
    try {
        Write-ColorOutput "Récupération des utilisateurs depuis Microsoft Graph..." "Cyan"
        $Users = Get-MgUser -All -Filter "userType eq 'Member' and accountEnabled eq true" -Property "UserPrincipalName", "DisplayName", "Department", "JobTitle", "OfficeLocation" -ErrorAction Stop
        Write-ColorOutput "$($Users.Count) utilisateur(s) trouvé(s)" "Green"
        return $Users
    }
    catch {
        Write-ColorOutput "Erreur lors de la récupération des utilisateurs : $($_.Exception.Message)" "Red"
            return @()
        }
}

function Get-UsersByDepartment {
    param(
        [string]$Department,
        [array]$AllUsers
    )
    
    Write-ColorOutput "Recherche des utilisateurs du département : $Department" "Cyan"
    
    try {
        # Normaliser le département recherché
        $NormalizedSearchDept = Normalize-Department -Department $Department
        
        # Filtrer les utilisateurs en comparant avec les départements normalisés
        $FilteredUsers = $AllUsers | Where-Object { 
            if ($_.Department -and $_.Department.Trim() -ne "") {
                $UserNormalizedDept = Normalize-Department -Department $_.Department
                return $UserNormalizedDept -eq $NormalizedSearchDept
            }
            return $false
        }
        
        if ($FilteredUsers.Count -eq 0) {
            Write-ColorOutput "Aucun utilisateur trouvé pour le département : $Department" "Yellow"
            Write-ColorOutput "Département normalisé recherché : $NormalizedSearchDept" "Yellow"
            
            # Afficher des exemples de départements disponibles (originaux)
            $AvailableDepts = $AllUsers | Where-Object { $_.Department -and $_.Department.Trim() -ne "" } | 
                Select-Object -ExpandProperty Department | 
            Sort-Object | 
                Get-Unique | 
                Select-Object -First 10
            
            if ($AvailableDepts.Count -gt 0) {
                Write-ColorOutput "`nExemples de départements disponibles (originaux) :" "Cyan"
                foreach ($Dept in $AvailableDepts) {
                    Write-ColorOutput "  - '$Dept'" "White"
                }
            }
            
            return @()
        }
        
        $UserEmails = $FilteredUsers | ForEach-Object { $_.UserPrincipalName }
        Write-ColorOutput "$($FilteredUsers.Count) utilisateur(s) trouvé(s) dans le département $Department" "Green"
        Write-ColorOutput "Département normalisé : $NormalizedSearchDept" "Green"
        
        # Afficher les utilisateurs trouvés avec leurs départements originaux
        Write-ColorOutput "`nUtilisateurs trouvés :" "Cyan"
        foreach ($User in $FilteredUsers) {
            $OriginalDept = $User.Department
            $NormalizedDept = Normalize-Department -Department $User.Department
            Write-ColorOutput "  - $($User.UserPrincipalName) (Original: '$OriginalDept' → Normalisé: '$NormalizedDept')" "White"
        }
        
        return $UserEmails
        
    }
    catch {
        Write-ColorOutput "Erreur lors de la recherche par département : $($_.Exception.Message)" "Red"
        return @()
    }
}

function Get-AvailableDepartments {
    param([array]$AllUsers)
    
    try {
        $UsersWithDepartment = $AllUsers | Where-Object { 
            $_.Department -and 
            $_.Department.Trim() -ne "" -and 
            $null -ne $_.Department 
        }
        
        Write-ColorOutput "$($UsersWithDepartment.Count) utilisateur(s) avec département défini" "Yellow"
        
        if ($UsersWithDepartment.Count -eq 0) {
            Write-ColorOutput "Aucun utilisateur n'a de département défini dans Azure AD." "Yellow"
            return @()
        }
        
        # Extraire TOUS les départements originaux (sans normalisation pour l'affichage)
        $Departments = @()
        foreach ($User in $UsersWithDepartment) {
            if ($User.Department -and $User.Department.Trim() -ne "") {
                $Departments += $User.Department.Trim()
            }
        }
        
        $Departments = $Departments | Sort-Object | Get-Unique
        Write-ColorOutput "$($Departments.Count) département(s) unique(s) trouvé(s)" "Green"
        
        # Afficher les départements trouvés pour diagnostic
        if ($Departments.Count -gt 0) {
            Write-ColorOutput "`nDépartements trouvés (ORIGINAUX - pour correction) :" "Cyan"
            foreach ($Dept in $Departments) {
                Write-ColorOutput "  - '$Dept'" "White"
            }
        }
        
        return $Departments
        
    }
    catch {
        Write-ColorOutput "Erreur lors de la récupération des départements : $($_.Exception.Message)" "Red"
        return @()
    }
}

# ===== INTERFACE UTILISATEUR =====

function Show-MainMenu {
    Write-ColorOutput "`n=== CONFIGURATION DES UTILISATEURS CIBLES ===" "Yellow"
    Write-ColorOutput "Choisissez le mode de recherche des délégations :" "Cyan"
    Write-ColorOutput ""
    
    Write-ColorOutput "Options de recherche :" "Yellow"
    Write-ColorOutput "  [1] Recherche par utilisateurs spécifiques" "White"
    Write-ColorOutput "  [2] Recherche par département" "White"
    Write-ColorOutput "  [3] Afficher le rapport de normalisation des départements" "White"
    Write-ColorOutput "  [4] Afficher TOUS les départements (pour correction)" "White"
    Write-ColorOutput "  [5] Export utilisateurs avec licences actives" "White"
    Write-ColorOutput "  [6] Annuler" "White"
    
    do {
        $choice = Read-Host "Votre choix (1-6)"
    } while ($choice -notin @("1", "2", "3", "4", "5", "6"))
    
    return $choice
}

function Get-UserEmailsFromSpecificUsers {
    Write-ColorOutput "`n=== RECHERCHE PAR UTILISATEURS SPÉCIFIQUES ===" "Yellow"
    Write-ColorOutput "Liste actuelle des utilisateurs :" "Green"
    
    for ($i = 0; $i -lt $Script:Config.DefaultUsers.Count; $i++) {
        Write-ColorOutput "  [$($i+1)] $($Script:Config.DefaultUsers[$i])" "White"
    }
    
    Write-ColorOutput ""
    Write-ColorOutput "Options :" "Yellow"
    Write-ColorOutput "  [1] Utiliser la liste actuelle" "White"
    Write-ColorOutput "  [2] Modifier la liste" "White"
    Write-ColorOutput "  [3] Retour au menu principal" "White"
            
            do {
                $choice = Read-Host "Votre choix (1-3)"
            } while ($choice -notin @("1", "2", "3"))
            
            switch ($choice) {
        "1" { return $Script:Config.DefaultUsers }
        "2" { return Get-CustomUserList }
        "3" { return @() }
    }
}

function Get-CustomUserList {
    Write-ColorOutput "`nSaisie des nouveaux utilisateurs :" "Green"
    Write-ColorOutput "Tapez les adresses email (une par ligne), puis 'FIN' pour terminer :" "Cyan"
                    
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
            Write-ColorOutput "Email invalide ignoré : $email" "Yellow"
                        }
                    }
                    
                    if ($validEmails.Count -eq 0) {
        Write-ColorOutput "Aucun email valide saisi. Utilisation de la liste par défaut." "Red"
        return $Script:Config.DefaultUsers
                    }
                    
                    return $validEmails
                }

function Get-UserEmailsFromDepartment {
    param([array]$AllUsers)
    
    Write-ColorOutput "`n=== RECHERCHE PAR DÉPARTEMENT ===" "Yellow"
    
    $AvailableDepartments = Get-AvailableDepartments -AllUsers $AllUsers
    
    if ($AvailableDepartments.Count -eq 0) {
        Write-ColorOutput "Aucun département trouvé dans Azure AD." "Red"
                        return @()
            }
            
    Write-ColorOutput "`nDépartements disponibles :" "Green"
                for ($i = 0; $i -lt $AvailableDepartments.Count; $i++) {
        Write-ColorOutput "  [$($i+1)] $($AvailableDepartments[$i])" "White"
    }
    
    Write-ColorOutput ""
    Write-ColorOutput "Options :" "Yellow"
    Write-ColorOutput "  [1-$(($AvailableDepartments.Count))] Sélectionner un département" "White"
    Write-ColorOutput "  [$(($AvailableDepartments.Count + 1))] Saisir un département manuellement" "White"
    Write-ColorOutput "  [$(($AvailableDepartments.Count + 2))] Retour au menu principal" "White"
                
                do {
                    $choice = Read-Host "Votre choix (1-$(($AvailableDepartments.Count + 2)))"
                } while ($choice -notin @(1..($AvailableDepartments.Count + 2)))
            
            if ($choice -le $AvailableDepartments.Count) {
                $SelectedDepartment = $AvailableDepartments[$choice - 1]
        Write-ColorOutput "Département sélectionné : $SelectedDepartment" "Green"
        return Get-UsersByDepartment -Department $SelectedDepartment -AllUsers $AllUsers
            }
            elseif ($choice -eq ($AvailableDepartments.Count + 1)) {
                $ManualDepartment = Read-Host "Saisissez le nom du département"
                if ($ManualDepartment.Trim() -ne "") {
            Write-ColorOutput "Département saisi : $ManualDepartment" "Green"
            return Get-UsersByDepartment -Department $ManualDepartment.Trim() -AllUsers $AllUsers
                } else {
            Write-ColorOutput "Nom de département vide." "Red"
                    return @()
                }
            }
            else {
                return @()
            }
        }

function Show-AllDepartmentsForCorrection {
    param([array]$AllUsers)
    
    Write-ColorOutput "`n=== TOUS LES DÉPARTEMENTS (POUR CORRECTION) ===" "Yellow"
    
    try {
        $UsersWithDepartment = $AllUsers | Where-Object { 
            $_.Department -and 
            $_.Department.Trim() -ne "" -and 
            $null -ne $_.Department 
        }
        
        # Grouper par département original
        $DepartmentGroups = $UsersWithDepartment | Group-Object -Property Department | Sort-Object Name
        
        Write-ColorOutput "`n$($DepartmentGroups.Count) département(s) unique(s) trouvé(s) :" "Green"
        Write-ColorOutput ""
        
        # Initialiser le contenu du rapport TXT
        $ReportContent = @()
        $ReportContent += "# Rapport de Vérification des Départements"
        $ReportContent += ""
        $ReportContent += "Généré le : $(Get-Date -Format 'dd/MM/yyyy à HH:mm')"
        $ReportContent += ""
        $ReportContent += "=== TOUS LES DÉPARTEMENTS (POUR CORRECTION) ==="
        $ReportContent += ""
        $ReportContent += "$($DepartmentGroups.Count) département(s) unique(s) trouvé(s) :"
        $ReportContent += ""
        
        foreach ($Group in $DepartmentGroups) {
            $DeptName = $Group.Name
            $UserCount = $Group.Count
            
            # Identifier les problèmes
            $Issues = @()
            if ($DeptName -match "@") {
                $Issues += "ERREUR: EMAIL DANS LE CHAMP DÉPARTEMENT"
            }
            
            # Vérifier la normalisation
            $Normalized = Normalize-Department -Department $DeptName
            if ($Normalized -ne $DeptName -and $Normalized -notlike "DONNÉES INCORRECTES*") {
                $Issues += "ATTENTION: SERA NORMALISÉ VERS: '$Normalized'"
            }
            
            Write-ColorOutput "Département: '$DeptName' ($UserCount utilisateur(s))" "Cyan"
            $ReportContent += "Département: '$DeptName' ($UserCount utilisateur(s))"
            
            if ($Issues.Count -gt 0) {
                foreach ($Issue in $Issues) {
                    Write-ColorOutput "   $Issue" "Red"
                    $ReportContent += "   $Issue"
                }
            }
            
            # Afficher les utilisateurs
            Write-ColorOutput "   Utilisateurs :" "Yellow"
            $ReportContent += "   Utilisateurs :"
            foreach ($User in $Group.Group) {
                Write-ColorOutput "     - $($User.UserPrincipalName)" "White"
                $ReportContent += "     - $($User.UserPrincipalName)"
            }
            Write-ColorOutput ""
            $ReportContent += ""
        }
        
        Write-ColorOutput "`n=== RÉSUMÉ DES PROBLÈMES À CORRIGER ===" "Red"
        $ReportContent += ""
        $ReportContent += "=== RÉSUMÉ DES PROBLÈMES À CORRIGER ==="
        $ReportContent += ""
        
        # Compter les problèmes
        $EmailIssues = $DepartmentGroups | Where-Object { $_.Name -match "@" }
        $NormalizationIssues = $DepartmentGroups | Where-Object { 
            $Normalized = Normalize-Department -Department $_.Name
            $Normalized -ne $_.Name -and $Normalized -notlike "DONNÉES INCORRECTES*"
        }
        
        if ($EmailIssues.Count -gt 0) {
            Write-ColorOutput "ERREUR: $($EmailIssues.Count) département(s) avec des EMAILS à corriger :" "Red"
            $ReportContent += "ERREUR: $($EmailIssues.Count) département(s) avec des EMAILS à corriger :"
            foreach ($Issue in $EmailIssues) {
                Write-ColorOutput "   - '$($Issue.Name)' ($($Issue.Count) utilisateur(s))" "Red"
                $ReportContent += "   - '$($Issue.Name)' ($($Issue.Count) utilisateur(s))"
            }
        }
        
        if ($NormalizationIssues.Count -gt 0) {
            Write-ColorOutput "`nATTENTION: $($NormalizationIssues.Count) département(s) avec des variantes d'orthographe :" "Yellow"
            $ReportContent += ""
            $ReportContent += "ATTENTION: $($NormalizationIssues.Count) département(s) avec des variantes d'orthographe :"
            foreach ($Issue in $NormalizationIssues) {
                $Normalized = Normalize-Department -Department $Issue.Name
                Write-ColorOutput "   - '$($Issue.Name)' → '$Normalized' ($($Issue.Count) utilisateur(s))" "Yellow"
                $ReportContent += "   - '$($Issue.Name)' → '$Normalized' ($($Issue.Count) utilisateur(s))"
            }
        }
        
        if ($EmailIssues.Count -eq 0 -and $NormalizationIssues.Count -eq 0) {
            Write-ColorOutput "Aucun problème détecté !" "Green"
            $ReportContent += "Aucun problème détecté !"
        }
        
        # Exporter le rapport TXT
        $DepartmentReportPath = "/Users/davidchiche/Desktop/Microsoft Entra/Rapport_Verification_Departements.txt"
        $ReportContent | Out-File -FilePath $DepartmentReportPath -Encoding UTF8
        
        Write-ColorOutput "`nRapport de vérification des départements exporté : $DepartmentReportPath" "Green"
        
    }
    catch {
        Write-ColorOutput "Erreur lors de l'affichage des départements : $($_.Exception.Message)" "Red"
    }
}

function Export-UsersWithActiveLicenses {
    param([array]$AllUsers)
    
    Write-ColorOutput "`n=== EXPORT UTILISATEURS AVEC LICENCES ACTIVES ===" "Yellow"
    
    try {
        Write-ColorOutput "Récupération des informations de licences..." "Cyan"
        
        # Récupérer les utilisateurs avec leurs licences
        $UsersWithLicenses = @()
$ProcessedCount = 0
        $ErrorCount = 0
        
        foreach ($User in $AllUsers) {
            $ProcessedCount++
            Write-Progress -Activity "Vérification des licences" -Status "Utilisateur $ProcessedCount sur $($AllUsers.Count)" -PercentComplete (($ProcessedCount / $AllUsers.Count) * 100)
            
            # Vérifier que l'utilisateur a un UserPrincipalName valide
            if (-not $User.UserPrincipalName -or $User.UserPrincipalName.Trim() -eq "") {
                continue
            }
            
            try {
                # Récupérer les licences de l'utilisateur en utilisant l'UserPrincipalName
                $UserLicenses = Get-MgUserLicenseDetail -UserId $User.UserPrincipalName -ErrorAction SilentlyContinue
                
                if ($UserLicenses -and $UserLicenses.Count -gt 0) {
                    # Vérifier si l'utilisateur a au moins une licence active
                    $HasActiveLicense = $false
                    $LicenseDetails = @()
                    
                    foreach ($License in $UserLicenses) {
                        if ($License.ServicePlans) {
                            foreach ($ServicePlan in $License.ServicePlans) {
                                if ($ServicePlan.ProvisioningStatus -eq "Success") {
                                    $HasActiveLicense = $true
                                    $LicenseDetails += "$($License.SkuPartNumber): $($ServicePlan.ServicePlanName)"
                                }
                            }
                        }
                    }
                    
                    if ($HasActiveLicense) {
                        $UsersWithLicenses += [PSCustomObject]@{
                            UserPrincipalName = $User.UserPrincipalName
                            DisplayName = $User.DisplayName
                            Department = $User.Department
                            JobTitle = $User.JobTitle
                            OfficeLocation = $User.OfficeLocation
                            LicenseCount = $UserLicenses.Count
                            ActiveLicenses = ($LicenseDetails -join "; ")
                        }
                    }
                }
            }
            catch {
                # Compter les erreurs silencieusement
                $ErrorCount++
            }
        }
        
        Write-Progress -Activity "Vérification des licences" -Completed
        
        Write-ColorOutput "`n$($UsersWithLicenses.Count) utilisateur(s) avec licences actives trouvé(s)" "Green"
        if ($ErrorCount -gt 0) {
            Write-ColorOutput "$ErrorCount utilisateur(s) avec erreurs de vérification des licences (comptes sans licences ou permissions insuffisantes)" "Yellow"
        }
        
        if ($UsersWithLicenses.Count -gt 0) {
            # Générer le rapport
            $ReportContent = @()
            $ReportContent += "# Rapport des Utilisateurs avec Licences Actives"
            $ReportContent += ""
            $ReportContent += "Généré le : $(Get-Date -Format 'dd/MM/yyyy à HH:mm')"
            $ReportContent += "Total utilisateurs avec licences actives : $($UsersWithLicenses.Count)"
            $ReportContent += ""
            $ReportContent += "=== DÉTAIL PAR UTILISATEUR ==="
            $ReportContent += ""
            
            foreach ($User in $UsersWithLicenses) {
                $ReportContent += "Utilisateur : $($User.UserPrincipalName)"
                $ReportContent += "Nom d'affichage : $($User.DisplayName)"
                $ReportContent += "Département : $($User.Department)"
                $ReportContent += "Poste : $($User.JobTitle)"
                $ReportContent += "Bureau : $($User.OfficeLocation)"
                $ReportContent += "Nombre de licences : $($User.LicenseCount)"
                $ReportContent += "Licences actives : $($User.ActiveLicenses)"
                $ReportContent += ""
            }
            
            # Statistiques par département
            $DepartmentStats = $UsersWithLicenses | Group-Object -Property Department | Sort-Object Count -Descending
            
            $ReportContent += "=== STATISTIQUES PAR DÉPARTEMENT ==="
            $ReportContent += ""
            foreach ($Dept in $DepartmentStats) {
                $ReportContent += "$($Dept.Name) : $($Dept.Count) utilisateur(s)"
            }
            
            # Exporter le rapport TXT
            $LicenseReportPath = "/Users/davidchiche/Desktop/Microsoft Entra/Rapport_Utilisateurs_Licences_Actives.txt"
            $ReportContent | Out-File -FilePath $LicenseReportPath -Encoding UTF8
            
            # Exporter en CSV
            $CsvPath = "/Users/davidchiche/Desktop/Microsoft Entra/Utilisateurs_Licences_Actives.csv"
            $UsersWithLicenses | Export-Csv -Path $CsvPath -NoTypeInformation -Encoding UTF8
            
            Write-ColorOutput "`nRapport TXT généré : $LicenseReportPath" "Green"
            Write-ColorOutput "Rapport CSV généré : $CsvPath" "Green"
            
            # Afficher un résumé à l'écran
            Write-ColorOutput "`n=== RÉSUMÉ ===" "Yellow"
            Write-ColorOutput "Total utilisateurs avec licences actives : $($UsersWithLicenses.Count)" "Green"
            Write-ColorOutput "`nTop 5 des départements :" "Cyan"
            foreach ($Dept in ($DepartmentStats | Select-Object -First 5)) {
                Write-ColorOutput "  - $($Dept.Name) : $($Dept.Count) utilisateur(s)" "White"
            }
        } else {
            Write-ColorOutput "Aucun utilisateur avec licence active trouvé." "Yellow"
        }
        
    }
    catch {
        Write-ColorOutput "Erreur lors de l'export des utilisateurs avec licences : $($_.Exception.Message)" "Red"
    }
}

function Get-TargetUsers {
    param([array]$AllUsers)
    
    $searchMode = Show-MainMenu
    
    switch ($searchMode) {
        "1" { return Get-UserEmailsFromSpecificUsers }
        "2" { return Get-UserEmailsFromDepartment -AllUsers $AllUsers }
        "3" { 
            Write-ColorOutput "`n=== RAPPORT DE NORMALISATION DES DÉPARTEMENTS ===" "Yellow"
            Get-NormalizedDepartments -AllUsers $AllUsers | Out-Null
            Write-ColorOutput "`nAppuyez sur Entrée pour continuer..." "Cyan"
            Read-Host
            return Get-TargetUsers -AllUsers $AllUsers  # Retour au menu principal
        }
        "4" { 
            Show-AllDepartmentsForCorrection -AllUsers $AllUsers
            Write-ColorOutput "`nAppuyez sur Entrée pour continuer..." "Cyan"
            Read-Host
            return Get-TargetUsers -AllUsers $AllUsers  # Retour au menu principal
        }
        "5" { 
            Export-UsersWithActiveLicenses -AllUsers $AllUsers
            Write-ColorOutput "`nAppuyez sur Entrée pour continuer..." "Cyan"
            Read-Host
            return Get-TargetUsers -AllUsers $AllUsers  # Retour au menu principal
        }
        "6" { 
            Write-ColorOutput "Opération annulée par l'utilisateur." "Yellow"
            exit 0
        }
    }
}

# ===== FONCTIONS DE RECHERCHE DE DÉLÉGATIONS =====

function Get-AllMailboxes {
    try {
        Write-ColorOutput "`nRécupération de toutes les boîtes aux lettres..." "Cyan"
        $AllMailboxes = Get-Mailbox -ResultSize Unlimited -ErrorAction Stop | Where-Object { 
            $_.RecipientTypeDetails -in @('UserMailbox', 'SharedMailbox', 'RoomMailbox', 'EquipmentMailbox') 
        }
        Write-ColorOutput "$($AllMailboxes.Count) boîtes aux lettres trouvées" "Green"
        return $AllMailboxes
    }
    catch {
        Write-ColorOutput "Erreur lors de la récupération des boîtes aux lettres : $($_.Exception.Message)" "Red"
        return @()
    }
}

function Show-Progress {
    param(
        [int]$Current,
        [int]$Total,
        [string]$Activity = "Traitement"
    )
    
    $PercentComplete = [math]::Round(($Current / $Total) * 100, 1)
    $ProgressBar = ""
    $BarLength = 50
    $FilledLength = [math]::Round(($Current / $Total) * $BarLength)
    
    for ($i = 0; $i -lt $FilledLength; $i++) { $ProgressBar += "█" }
    for ($i = $FilledLength; $i -lt $BarLength; $i++) { $ProgressBar += "░" }
    
    Write-Progress -Activity $Activity -Status "Boîte $Current sur $Total ($PercentComplete%)" -PercentComplete $PercentComplete
    Write-Host "`r[$ProgressBar] $PercentComplete% ($Current/$Total)" -NoNewline -ForegroundColor Green
}

function Find-Delegations {
    param(
        [array]$TargetUsers,
        [array]$AllMailboxes
    )
    
    Write-ColorOutput "`nRecherche des délégations possédées par $($TargetUsers.Count) utilisateurs..." "Green"
    Write-ColorOutput "ATTENTION: Cette opération peut prendre plusieurs minutes selon la taille de votre organisation." "Yellow"
    
    $Results = @()
    $ProcessedCount = 0
    
    # Créer un tableau des utilisateurs en minuscules pour comparaison
    $TargetUsersLower = $TargetUsers | ForEach-Object { $_.ToLower() }
    
foreach ($Mailbox in $AllMailboxes) {
    $ProcessedCount++
        Show-Progress -Current $ProcessedCount -Total $AllMailboxes.Count -Activity "Analyse des délégations"
    
    try {
            # Récupérer les permissions Full Access
        $FullAccessPerms = Get-MailboxPermission -Identity $Mailbox.PrimarySmtpAddress -ErrorAction SilentlyContinue | 
            Where-Object { 
                    $TargetUsersLower -contains $_.User.ToString().ToLower() -and 
                $_.IsInherited -eq $false 
            }

            # Récupérer les permissions Send As
        $SendAsPerms = Get-RecipientPermission -Identity $Mailbox.PrimarySmtpAddress -ErrorAction SilentlyContinue |
                Where-Object { $TargetUsersLower -contains $_.Trustee.ToString().ToLower() }

            # Récupérer les permissions Send on Behalf
        $SendOnBehalfUsers = $Mailbox.GrantSendOnBehalfTo | Where-Object { 
                $TargetUsersLower -contains $_.ToString().ToLower() 
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
            Write-ColorOutput "`nErreur lors du traitement de $($Mailbox.PrimarySmtpAddress) : $($_.Exception.Message)" "Yellow"
        }
    }
    
    Write-ColorOutput "`nRecherche terminée !" "Green"
    Write-Progress -Activity "Analyse des délégations" -Completed
    
    return $Results
}

# ===== GÉNÉRATION DE RAPPORT =====

function Get-UserDepartment {
    param([string]$UserEmail)
    
    try {
        $User = Get-MgUser -Filter "userPrincipalName eq '$UserEmail'" -ErrorAction SilentlyContinue
        if ($User) {
            return $User.Department
        }
    }
    catch {
        # Ignorer les erreurs silencieusement
    }
    return "Non défini"
}

function Generate-Report {
    param(
        [array]$Results,
        [array]$TargetUsers
    )
    
    Write-ColorOutput "`nGénération du rapport formaté..." "Cyan"

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

# Écrire le rapport formaté
    $ReportContent | Out-File -FilePath $Script:Config.OutputPath -Encoding UTF8
    
    # Écrire le rapport CSV
    $Results | Export-Csv -Path $Script:Config.CsvPath -NoTypeInformation -Encoding UTF8
    
    return $ReportContent
}

function Show-Summary {
    param(
        [array]$Results,
        [array]$TargetUsers
    )
    
if ($Results.Count -gt 0) {
        Write-ColorOutput "`n=== RÉSUMÉ ===" "Yellow"
        Write-ColorOutput "$($Results.Count) boîte(s) aux lettres avec des délégations trouvées" "Green"
    
    # Résumé par utilisateur
        Write-ColorOutput "`n=== RÉSUMÉ PAR UTILISATEUR ===" "Yellow"
    foreach ($User in $TargetUsers) {
        $UserDelegations = $Results | Where-Object { 
            $_.FullAccess.ToLower() -like "*$($User.ToLower())*" -or 
            $_.SendAs.ToLower() -like "*$($User.ToLower())*" -or 
            $_.SendOnBehalf.ToLower() -like "*$($User.ToLower())*" 
        }
        
        if ($UserDelegations.Count -gt 0) {
                Write-ColorOutput "`n$User a des délégations sur $($UserDelegations.Count) boîte(s)" "Cyan"
        } else {
                Write-ColorOutput "`n$User : Aucune délégation trouvée" "Gray"
        }
    }
    
        Write-ColorOutput "`nRapport formaté généré : $($Script:Config.OutputPath)" "Green"
        Write-ColorOutput "Rapport CSV généré : $($Script:Config.CsvPath)" "Green"
    
} else {
        Write-ColorOutput "`nAucune délégation trouvée pour les utilisateurs spécifiés." "Red"
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
        
        # Récupération de tous les utilisateurs (une seule fois)
        $AllUsers = Get-AllUsers
        if ($AllUsers.Count -eq 0) {
            Write-ColorOutput "Aucun utilisateur trouvé. Arrêt du script." "Red"
            Disconnect-FromServices
            exit 1
        }
        
        # Sélection des utilisateurs cibles
        $TargetUsers = Get-TargetUsers -AllUsers $AllUsers
        
        # Vérification des utilisateurs sélectionnés
        if ($TargetUsers.Count -eq 0) {
            Write-ColorOutput "`nAucun utilisateur sélectionné. Arrêt du script." "Red"
            Disconnect-FromServices
            exit 1
        }
        
        Write-ColorOutput "`nUtilisateurs configurés : $($TargetUsers.Count)" "Green"
        $TargetUsers | ForEach-Object { Write-ColorOutput "  - $_" "Cyan" }
        
        # Récupération des boîtes aux lettres
        $AllMailboxes = Get-AllMailboxes
        if ($AllMailboxes.Count -eq 0) {
            Write-ColorOutput "Aucune boîte aux lettres trouvée. Arrêt du script." "Red"
            Disconnect-FromServices
            exit 1
        }
        
        # Recherche des délégations
        $Results = Find-Delegations -TargetUsers $TargetUsers -AllMailboxes $AllMailboxes
        
        # Génération du rapport
        $ReportContent = Generate-Report -Results $Results -TargetUsers $TargetUsers
        
        # Affichage du résumé
        Show-Summary -Results $Results -TargetUsers $TargetUsers
        
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