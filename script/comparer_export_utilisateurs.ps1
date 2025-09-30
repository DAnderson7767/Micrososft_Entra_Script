#Requires -Version 7.0

<#
.SYNOPSIS
    Script pour comparer les exports avec et sans boîtes partagées.

.DESCRIPTION
    Ce script exécute deux exports :
    1. Un export sans boîtes partagées (mode par défaut)
    2. Un export avec boîtes partagées incluses
    
    Il affiche ensuite les statistiques de comparaison.

.EXAMPLE
    .\comparer_export_utilisateurs.ps1
#>

# Configuration des couleurs
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

function Get-UserCountFromCSV {
    param(
        [string]$CsvPath
    )
    
    if (Test-Path $CsvPath) {
        $csv = Import-Csv $CsvPath
        return $csv.Count
    }
    return 0
}

function Show-Comparison {
    param(
        [string]$WithoutSharedPath,
        [string]$WithSharedPath
    )
    
    $countWithout = Get-UserCountFromCSV $WithoutSharedPath
    $countWith = Get-UserCountFromCSV $WithSharedPath
    $sharedCount = $countWith - $countWithout
    
    Write-ColorOutput "" $Colors.Info
    Write-ColorOutput "📊 COMPARAISON DES EXPORTS" $Colors.Header
    Write-ColorOutput ("=" * 40) $Colors.Header
    Write-ColorOutput "👥 Utilisateurs sans boîtes partagées : $countWithout" $Colors.Info
    Write-ColorOutput "📦 Utilisateurs avec boîtes partagées : $countWith" $Colors.Info
    Write-ColorOutput "🔍 Boîtes partagées détectées : $sharedCount" $Colors.Warning
    Write-ColorOutput "" $Colors.Info
    
    if ($sharedCount -gt 0) {
        Write-ColorOutput "✅ Les boîtes partagées ont été correctement exclues du premier export" $Colors.Success
    } else {
        Write-ColorOutput "ℹ️ Aucune boîte partagée détectée dans votre environnement" $Colors.Info
    }
}

# Script principal
Write-ColorOutput "🔍 COMPARAISON DES EXPORTS UTILISATEURS" $Colors.Header
Write-ColorOutput ("=" * 50) $Colors.Header

# Vérifier que le script principal existe
$scriptPath = Join-Path $PSScriptRoot "export_utilisateurs_macos.ps1"

if (-not (Test-Path $scriptPath)) {
    Write-ColorOutput "❌ Script principal non trouvé : $scriptPath" $Colors.Error
    exit 1
}

Write-ColorOutput "✅ Script principal trouvé" $Colors.Success

# Export 1 : Sans boîtes partagées
Write-ColorOutput "" $Colors.Info
Write-ColorOutput "📊 Export 1 : Sans boîtes partagées..." $Colors.Info
try {
    & $scriptPath -OutputPath "." | Out-Null
    $withoutSharedFile = Get-ChildItem -Path "." -Name "Utilisateurs_Graph_*.csv" | Sort-Object LastWriteTime -Descending | Select-Object -First 1
    Write-ColorOutput "✅ Export 1 terminé : $withoutSharedFile" $Colors.Success
}
catch {
    Write-ColorOutput "❌ Erreur lors de l'export 1 : $($_.Exception.Message)" $Colors.Error
    exit 1
}

# Export 2 : Avec boîtes partagées
Write-ColorOutput "" $Colors.Info
Write-ColorOutput "📦 Export 2 : Avec boîtes partagées..." $Colors.Info
try {
    & $scriptPath -OutputPath "." -IncludeSharedMailboxes | Out-Null
    $withSharedFile = Get-ChildItem -Path "." -Name "Utilisateurs_Graph_*.csv" | Sort-Object LastWriteTime -Descending | Select-Object -First 1
    Write-ColorOutput "✅ Export 2 terminé : $withSharedFile" $Colors.Success
}
catch {
    Write-ColorOutput "❌ Erreur lors de l'export 2 : $($_.Exception.Message)" $Colors.Error
    exit 1
}

# Comparaison
Show-Comparison -WithoutSharedPath $withoutSharedFile -WithSharedPath $withSharedFile

Write-ColorOutput "" $Colors.Info
Write-ColorOutput "🎯 Fichiers générés :" $Colors.Info
Write-ColorOutput "   📄 Sans boîtes partagées : $withoutSharedFile" $Colors.Info
Write-ColorOutput "   📄 Avec boîtes partagées : $withSharedFile" $Colors.Info
Write-ColorOutput "" $Colors.Info
Write-ColorOutput "✅ Comparaison terminée avec succès !" $Colors.Success
