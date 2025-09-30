#Requires -Version 7.0

<#
.SYNOPSIS
    Script d'installation des modules PowerShell nécessaires pour macOS.

.DESCRIPTION
    Ce script installe automatiquement les modules PowerShell requis pour utiliser
    les scripts d'export des utilisateurs Microsoft Graph sur macOS.

.EXAMPLE
    .\installer_modules_macos.ps1
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

function Install-RequiredModule {
    param(
        [string]$ModuleName,
        [string]$Description
    )
    
    Write-ColorOutput "🔍 Vérification du module $ModuleName..." $Colors.Info
    
    if (Get-Module -ListAvailable -Name $ModuleName) {
        Write-ColorOutput "✅ Module $ModuleName déjà installé" $Colors.Success
        return $true
    }
    
    Write-ColorOutput "📦 Installation du module $ModuleName ($Description)..." $Colors.Info
    
    try {
        Install-Module -Name $ModuleName -Scope CurrentUser -Force -AllowClobber
        Write-ColorOutput "✅ Module $ModuleName installé avec succès" $Colors.Success
        return $true
    }
    catch {
        Write-ColorOutput "❌ Erreur lors de l'installation du module $ModuleName : $($_.Exception.Message)" $Colors.Error
        return $false
    }
}

function Test-PowerShellVersion {
    $currentVersion = $PSVersionTable.PSVersion
    $requiredVersion = [Version]"7.0"
    
    Write-ColorOutput "🔍 Vérification de la version PowerShell..." $Colors.Info
    Write-ColorOutput "Version actuelle : $($currentVersion.ToString())" $Colors.Info
    
    if ($currentVersion -ge $requiredVersion) {
        Write-ColorOutput "✅ Version PowerShell compatible" $Colors.Success
        return $true
    } else {
        Write-ColorOutput "❌ Version PowerShell incompatible. Version requise : 7.0+" $Colors.Error
        Write-ColorOutput "💡 Installez PowerShell Core 7.0 ou supérieur avec : brew install --cask powershell" $Colors.Warning
        return $false
    }
}

function Test-macOSEnvironment {
    Write-ColorOutput "🍎 Vérification de l'environnement macOS..." $Colors.Info
    
    if ($IsMacOS -or $PSVersionTable.Platform -eq "Unix") {
        Write-ColorOutput "✅ Environnement macOS détecté" $Colors.Success
        return $true
    } else {
        Write-ColorOutput "⚠️ Environnement non-macOS détecté" $Colors.Warning
        Write-ColorOutput "💡 Ce script est optimisé pour macOS" $Colors.Info
        return $true  # On continue quand même
    }
}

# Script principal
Write-ColorOutput "🍎 INSTALLATION DES MODULES POWERSHELL POUR MACOS" $Colors.Header
Write-ColorOutput ("=" * 50) $Colors.Header

# Vérifier l'environnement
Test-macOSEnvironment

# Vérifier la version PowerShell
if (-not (Test-PowerShellVersion)) {
    Write-ColorOutput "❌ Installation interrompue" $Colors.Error
    exit 1
}

# Modules à installer pour macOS
$modules = @(
    @{
        Name = "Microsoft.Graph"
        Description = "Module Microsoft Graph (recommandé pour macOS)"
    },
    @{
        Name = "Microsoft.Graph.Users"
        Description = "Module Microsoft Graph Users"
    }
)

$allInstalled = $true

foreach ($module in $modules) {
    if (-not (Install-RequiredModule -ModuleName $module.Name -Description $module.Description)) {
        $allInstalled = $false
    }
}

Write-ColorOutput "" $Colors.Info
Write-ColorOutput ("=" * 50) $Colors.Header

if ($allInstalled) {
    Write-ColorOutput "✅ TOUS LES MODULES INSTALLÉS AVEC SUCCÈS" $Colors.Success
    Write-ColorOutput "" $Colors.Info
    Write-ColorOutput "🎯 Vous pouvez maintenant utiliser le script d'export des utilisateurs :" $Colors.Info
    Write-ColorOutput "   .\export_utilisateurs_macos.ps1" $Colors.Info
    Write-ColorOutput "" $Colors.Info
    Write-ColorOutput "💡 Note : Sur macOS, Microsoft.Graph est recommandé au lieu d'AzureAD" $Colors.Info
} else {
    Write-ColorOutput "❌ CERTAINS MODULES N'ONT PAS PU ÊTRE INSTALLÉS" $Colors.Error
    Write-ColorOutput "💡 Vérifiez les erreurs ci-dessus et réessayez" $Colors.Warning
    Write-ColorOutput "💡 Sur macOS, assurez-vous d'avoir PowerShell Core 7.0+" $Colors.Warning
}

Write-ColorOutput "" $Colors.Info
