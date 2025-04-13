param(
    [string]$ScriptPath = "",
    [string]$TaskName = "SurveillanceLicencesQuotidienne",
    [string]$Time = "09:00",
    [switch]$Force
)

# Vérifier l'élévation des privilèges
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Warning "Ce script nécessite des droits administrateur. Veuillez relancer en tant qu'administrateur."
    exit 1
}

# Vérifier si le script existe
if (-not (Test-Path -Path $ScriptPath -PathType Leaf)) {
    Write-Error "Le script principal n'a pas été trouvé à l'emplacement spécifié: $ScriptPath"
    exit 1
}

# Vérifier le format de l'heure
if ($Time -notmatch '^\d{2}:\d{2}$') {
    Write-Error "Format de l'heure invalide. Utilisez HH:mm (ex: 09:00 ou 16:30)"
    exit 1
}

# Vérifier si la tâche existe déjà
$existingTask = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue

if ($existingTask) {
    if (-not $Force) {
        $confirmation = Read-Host "La tâche '$TaskName' existe déjà. Voulez-vous la mettre à jour ? (O/N)"
        if ($confirmation -notin 'O','o') {
            Write-Host "Annulation de l'opération."
            exit
        }
    }
    try {
        Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction Stop
        Write-Host "Ancienne tâche supprimée avec succès."
    }
    catch {
        Write-Error "Échec de la suppression de la tâche existante: $_"
        exit 1
    }
}

try {
    # Créer l'action
    $action = New-ScheduledTaskAction `
        -Execute "PowerShell.exe" `
        -Argument "-NoProfile -ExecutionPolicy Bypass -File `"$ScriptPath`""

    # Créer le déclencheur quotidien
    $trigger = New-ScheduledTaskTrigger `
        -Daily `
        -At $Time

    # Configurer les paramètres
    $settings = New-ScheduledTaskSettingsSet `
        -StartWhenAvailable `
        -DontStopOnIdleEnd `
        -AllowStartIfOnBatteries `
        -DontStopIfGoingOnBatteries `
        -RunOnlyIfNetworkAvailable `
        -RestartCount 3 `
        -RestartInterval (New-TimeSpan -Minutes 30)

    # Enregistrer la tâche
    Register-ScheduledTask `
        -TaskName $TaskName `
        -Action $action `
        -Trigger $trigger `
        -Settings $settings `
        -RunLevel Highest `
        -Description "Exécution quotidienne du script de surveillance des licences à $Time" `
        -ErrorAction Stop

    Write-Host "`nTâche planifiée configurée avec succès :" -ForegroundColor Green
    Write-Host "- Nom : $TaskName" -ForegroundColor Cyan
    Write-Host "- Exécution quotidienne à : $Time" -ForegroundColor Cyan
    Write-Host "- Script : $ScriptPath" -ForegroundColor Cyan
    Write-Host "- Compte : $env:USERDOMAIN\$env:USERNAME" -ForegroundColor Cyan
    Write-Host "- Niveau d'exécution : Élevé" -ForegroundColor Cyan
    
    # Vérification optionnelle
    $createdTask = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
    if ($createdTask) {
        Write-Host "`nVérification : La tâche a bien été créée dans le Planificateur de tâches." -ForegroundColor Green
    }
}
catch {
    Write-Error "Erreur lors de la création de la tâche : $_"
    exit 1
}
