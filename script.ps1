<#
.SYNOPSIS
    Enhanced License Monitoring System with simplified email security
.DESCRIPTION
    This script monitors software license expiration dates, sends alerts, and generates reports.
    Uses plain text credentials for SMTP authentication.
.NOTES
    Version: 4.2
    Author: billel-eren
    Last Updated: 19/04/2025
#>

# --- Configuration ---
param (
    [string]$BasePath = "",
    [string]$ExcelFilePath = "$BasePath\SoftwareLicenses.xlsx",
    [string]$LogPath = "$BasePath\LicenseMonitor.log",
    [string]$ReportPath = "$BasePath\LicenseReport.html",
    [string]$BackupFolder = "$BasePath\Backups",
    
    # Email Configuration
    [string]$FromAddress = "",
    [string]$ToAddress = "",
    [string[]]$CCAddresses = @(),
    [string]$SMTPServer = "",
    [int]$SMTPPort = 587,
    [string]$SMTPUsername = "",  # Plain text username
    [string]$SMTPPassword = "",  # Plain text password
    
    # Alert Thresholds
    [int]$WarningDays = 90,
    [int]$CriticalDays = 30,
    [int]$ExpiredFrequency = 7,
    [int]$WarningFrequency = 14,
    [int]$NormalFrequency = 30,
    
    # Features
    [bool]$EnableReporting = $true,
    [bool]$EnableLogging = $true,
    [bool]$TrackRenewalCosts = $true,
    [bool]$AutoBackupExcel = $true,
    
    # Company Information
    [string]$CompanyName = "Your Company",
    [string]$CompanyLogo = "$BasePath\logo.png"
)

# --- Constants ---
$EmailRetryAttempts = 3
$EmailRetryDelay = 5  # seconds (will double with each attempt)

# Column Mapping (with empty value handling)
$ColumnMap = @{
    Software = "Logicielle"
    ExpirationDate = "date-experation-license"
    LicenseType = "Type"
    Owner = "Proprietaire"
    LastCheck = "Derniere-verification"
    LastEmail = "Dernier-email"
    RenewalCost = "Cout-renouvellement"
}

# Language Settings (French)
$Language = @{
    Strings = @{
        "Report-Title" = "Rapport d'Expiration des Licences"
        "Licenses" = "Licences"
        "Summary" = "Résumé"
        "Priority" = "Priorité"
        "Software" = "Logiciel"
        "ExpirationDate" = "Date d'expiration"
        "DaysRemaining" = "Jours restants"
        "LicenseType" = "Type de licence"
        "Owner" = "Propriétaire"
        "RenewalCost" = "Coût de renouvellement"
        "Status" = "Statut"
        "Low" = "Faible"
        "Medium" = "Moyenne"
        "High" = "Élevée"
        "Critical" = "Critique"
        "OK" = "OK"
        "Warning" = "Attention"
        "Expired" = "EXPIRÉ"
        "GeneratedBy" = "Généré automatiquement par le script de surveillance des licences"
    }
}

# --- Backup Functions ---
function Backup-ExcelFile {
    param (
        [string]$SourcePath,
        [string]$BackupFolder,
        [bool]$Force = $false
    )
    
    try {
        # Check if backup is needed (once per month)
        $BackupTrackerFile = Join-Path $BackupFolder "backup_tracker.txt"
        $CurrentMonth = (Get-Date).ToString("yyyy-MM")
        
        $LastBackupMonth = $null
        if (Test-Path $BackupTrackerFile) {
            $LastBackupMonth = Get-Content $BackupTrackerFile -ErrorAction SilentlyContinue
        }
        
        # Skip backup if already done this month and not forced
        if (-not $Force -and $LastBackupMonth -eq $CurrentMonth) {
            Write-Log "La sauvegarde a déjà été effectuée ce mois-ci ($CurrentMonth). Ignoré." "INFO"
            return
        }
        
        # Create backup directory if it doesn't exist
        if (-not (Test-Path $BackupFolder)) {
            New-Item -ItemType Directory -Path $BackupFolder -Force | Out-Null
            Write-Log "Répertoire de sauvegarde créé: $BackupFolder" "INFO"
        }
        
        # Create monthly backup subfolder
        $MonthlyBackupFolder = Join-Path $BackupFolder $CurrentMonth
        if (-not (Test-Path $MonthlyBackupFolder)) {
            New-Item -ItemType Directory -Path $MonthlyBackupFolder -Force | Out-Null
        }
        
        # Generate backup filename with timestamp
        $BackupFileName = "SoftwareLicenses_$(Get-Date -Format 'yyyy-MM-dd_HHmmss').xlsx"
        $BackupPath = Join-Path $MonthlyBackupFolder $BackupFileName
        
        # Perform the backup
        Copy-Item -Path $SourcePath -Destination $BackupPath -Force
        Write-Log "Fichier Excel sauvegardé vers: $BackupPath" "INFO"
        
        # Update the backup tracker
        $CurrentMonth | Out-File -FilePath $BackupTrackerFile -Force
        Write-Log "Sauvegarde mensuelle effectuée pour le mois $CurrentMonth" "INFO"
    }
    catch {
        Write-Log "Erreur lors de la sauvegarde du fichier Excel: $($_.Exception.Message)" "ERROR"
    }
}

function Create-AlertEmailBody {
    param (
        [PSCustomObject]$Alert,
        [int]$CriticalDays,
        [int]$WarningDays,
        [string]$CompanyName = "your-company"
    )
    
    # Determine status and colors
    if ($Alert.DaysUntilExpiration -lt 0) {
        $StatusText = "EXPIRÉE"
        $StatusColor = "#E53E3E" # Red
        $HeaderBg = "#E53E3E"
        $AccentColor = "#C53030"
    }
    elseif ($Alert.DaysUntilExpiration -le $CriticalDays) {
        $StatusText = "CRITIQUE"
        $StatusColor = "#E53E3E" # Red
        $HeaderBg = "#E53E3E"
        $AccentColor = "#C53030"
    }
    else {
        $StatusText = "AVERTISSEMENT"
        $StatusColor = "#ED8936" # Orange
        $HeaderBg = "#ED8936"
        $AccentColor = "#DD6B20"
    }
    
    $ExpirationDateFormatted = $Alert.ExpirationDate.ToString("dd/MM/yyyy")
    $GeneratedDate = Get-Date -Format "dd/MM/yyyy HH:mm"
    
    $HTMLBody = @"
<!DOCTYPE html>
<html lang="fr">
<head>
    <meta charset="UTF-8">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Alerte d'expiration de licence</title>
    <style>
        /* Inline all styles to ensure email works without internet */
        body {
            font-family: Arial, 'Helvetica Neue', Helvetica, sans-serif;
            line-height: 1.5;
            color: #2D3748;
            margin: 0;
            padding: 0;
            background-color: #F7FAFC;
        }
        
        .container {
            max-width: 650px;
            margin: 20px auto;
            background-color: #FFFFFF;
            border-radius: 10px;
            box-shadow: 0 4px 20px rgba(0, 0, 0, 0.08);
            overflow: hidden;
        }
        
        .header {
            background-color: $HeaderBg;
            color: white;
            padding: 25px 30px;
            position: relative;
        }
        
        .header h1 {
            margin: 0;
            font-size: 22px;
            font-weight: 600;
            letter-spacing: 0.015em;
        }
        
        .status-badge {
            display: inline-block;
            padding: 5px 12px;
            background-color: rgba(255, 255, 255, 0.25);
            color: white;
            border-radius: 20px;
            font-size: 13px;
            font-weight: 600;
            margin-left: 10px;
            letter-spacing: 0.03em;
            text-transform: uppercase;
            vertical-align: middle;
        }
        
        .content {
            padding: 30px;
        }
        
        .info-card {
            background-color: #F7FAFC;
            border-left: 4px solid $StatusColor;
            margin-bottom: 25px;
            padding: 20px;
            border-radius: 6px;
            box-shadow: 0 2px 5px rgba(0, 0, 0, 0.04);
        }
        
        .software-name {
            font-size: 20px;
            font-weight: 600;
            margin-bottom: 10px;
            color: #2D3748;
        }
        
        .days-counter {
            font-size: 22px;
            font-weight: 700;
            color: $StatusColor;
            text-align: center;
            margin: 25px 0;
            padding: 20px;
            background-color: rgba(0, 0, 0, 0.02);
            border-radius: 8px;
            border: 1px solid #E2E8F0;
        }
        
        .details-table {
            width: 100%;
            border-collapse: separate;
            border-spacing: 0;
            margin: 20px 0 25px;
            border-radius: 8px;
            overflow: hidden;
            box-shadow: 0 2px 5px rgba(0, 0, 0, 0.04);
        }
        
        .details-table th {
            text-align: left;
            padding: 12px 16px;
            background-color: #F1F5F9;
            font-weight: 600;
            font-size: 14px;
            color: #718096;
            border-bottom: 1px solid #E2E8F0;
        }
        
        .details-table td {
            padding: 12px 16px;
            border-bottom: 1px solid #E2E8F0;
            font-size: 15px;
            background-color: #FFFFFF;
        }
        
        .details-table tr:last-child td {
            border-bottom: none;
        }
        
        h3 {
            color: #2D3748;
            font-size: 18px;
            margin-top: 25px;
            margin-bottom: 15px;
        }
        
        .info-section {
            background-color: #F7FAFC;
            padding: 18px;
            border-radius: 8px;
            margin-top: 20px;
            border: 1px solid #E2E8F0;
        }
        
        .info-section h4 {
            margin-top: 0;
            margin-bottom: 10px;
            font-weight: 600;
            color: #2D3748;
            font-size: 16px;
        }
        
        .footer {
            margin-top: 30px;
            padding: 20px 30px;
            border-top: 1px solid #E2E8F0;
            font-size: 13px;
            color: #718096;
            text-align: center;
            background-color: #FAFBFC;
        }
        
        p {
            margin: 0 0 15px;
        }
        
        strong {
            font-weight: 600;
        }
        
        /* Basic responsiveness for mobile clients */
        @media screen and (max-width: 600px) {
            .container {
                margin: 10px;
                width: auto;
            }
            
            .content {
                padding: 20px;
            }
            
            .header {
                padding: 20px;
            }
            
            .header h1 {
                font-size: 18px;
            }
            
            .status-badge {
                display: block;
                margin: 10px 0 0;
                width: fit-content;
            }
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>Alerte d'expiration de licence <span class="status-badge">$StatusText</span></h1>
        </div>
        
        <div class="content">
            <div class="info-card">
                <div class="software-name">$($Alert.Software)</div>
                <p>
"@

    if ($Alert.DaysUntilExpiration -lt 0) {
        $HTMLBody += "Cette licence a <strong>expiré</strong> depuis <strong>$([Math]::Abs($Alert.DaysUntilExpiration))</strong> jours.<br>Date d'expiration: <strong>$ExpirationDateFormatted</strong>"
    }
    else {
        $HTMLBody += "Cette licence va expirer dans <strong>$($Alert.DaysUntilExpiration)</strong> jours.<br>Date d'expiration: <strong>$ExpirationDateFormatted</strong>"
    }

    $HTMLBody += @"
                </p>
            </div>
            
            <div class="days-counter">
                $(if ($Alert.DaysUntilExpiration -lt 0) { "EXPIRÉE DEPUIS $([Math]::Abs($Alert.DaysUntilExpiration)) JOURS" } else { "EXPIRE DANS $($Alert.DaysUntilExpiration) JOURS" })
            </div>
            
            <h3>Détails de la licence</h3>
            <table class="details-table">
                <tr><th>Logiciel</th><td>$($Alert.Software)</td></tr>
"@

    if ($Alert.LicenseType) {
        $HTMLBody += "<tr><th>Type de licence</th><td>$($Alert.LicenseType)</td></tr>"
    }

    if ($Alert.Owner) {
        $HTMLBody += "<tr><th>Propriétaire</th><td>$($Alert.Owner)</td></tr>"
    }

    if ($Alert.RenewalCost) {
        $HTMLBody += "<tr><th>Coût de renouvellement</th><td>$($Alert.RenewalCost) €</td></tr>"
    }

    $HTMLBody += @"
                <tr><th>Date d'expiration</th><td>$ExpirationDateFormatted</td></tr>
            </table>
            
            <div class="info-section">
                <h4>Action requise</h4>
                <p>Veuillez prendre les mesures appropriées pour renouveler cette licence dès que possible afin d'éviter toute interruption de service.</p>
                <p>Pour plus d'informations, veuillez contacter votre équipe informatique ou le gestionnaire des licences.</p>
            </div>
        </div>
        
        <div class="footer">
            <p>Ce message a été généré automatiquement par le système de surveillance des licences$(if ($CompanyName) { " de $CompanyName" }).<br>
            Généré le $GeneratedDate</p>
        </div>
    </div>
</body>
</html>
"@
    
    return $HTMLBody
}

function Send-EmailWithRetry {
    param (
        [hashtable]$MailParams,
        [int]$MaxAttempts = 3,
        [int]$InitialDelay = 5
    )
    
    $attempt = 1
    $delay = $InitialDelay
    
    while ($attempt -le $MaxAttempts) {
        try {
            # Remove SSL requirement and use plain authentication
            $MailParams.Remove("UseSsl")
            Send-MailMessage @MailParams -ErrorAction Stop
            Write-Log "Email envoyé avec succès (tentative $attempt)" "INFO"
            return $true
        }
        catch [System.Net.Mail.SmtpException] {
            if ($attempt -ge $MaxAttempts) {
                Write-Log "Échec d'envoi de l'email après $MaxAttempts tentatives. Erreur: $($_.Exception.Message)" "ERROR"
                return $false
            }
            
            Write-Log "Tentative d'envoi d'email $attempt échouée. Nouvelle tentative dans $delay secondes. Erreur: $($_.Exception.Message)" "WARNING"
            Start-Sleep -Seconds $delay
            $attempt++
            $delay *= 2
        }
        catch {
            Write-Log "Erreur inattendue lors de l'envoi de l'email: $($_.Exception.Message)" "ERROR"
            return $false
        }
    }
    
    return $false
}

# --- Input Validation ---
function Test-EmailAddress {
    param([string]$Email)
    try {
        $null = [mailaddress]$Email
        return $true
    }
    catch {
        return $false
    }
}

function Validate-Inputs {
    # Validate paths
    $invalidPaths = @()
    if (-not (Test-Path -Path $BasePath -IsValid)) { $invalidPaths += $BasePath }
    if (-not (Test-Path -Path $BackupFolder -IsValid)) { $invalidPaths += $BackupFolder }
    
    if ($invalidPaths.Count -gt 0) {
        throw "Invalid path(s) detected: $($invalidPaths -join ', ')"
    }
    
    # Validate email addresses
    $invalidEmails = @()
    if (-not (Test-EmailAddress $FromAddress)) { $invalidEmails += "FromAddress" }
    if (-not (Test-EmailAddress $ToAddress)) { $invalidEmails += "ToAddress" }
    foreach ($cc in $CCAddresses) {
        if (-not (Test-EmailAddress $cc)) { $invalidEmails += "CCAddress: $cc" }
    }
    
    if ($invalidEmails.Count -gt 0) {
        throw "Invalid email address(es) detected: $($invalidEmails -join ', ')"
    }
    
    # Validate Excel file
    if (-not (Test-Path $ExcelFilePath -PathType Leaf)) {
        throw "Excel file not found at: $ExcelFilePath"
    }
    
    if ((Get-Item $ExcelFilePath).Extension -notin '.xlsx', '.xls') {
        throw "Invalid file type. Expected Excel file (.xlsx or .xls)"
    }
}

# --- Enhanced Logging Function ---
function Write-Log {
    param (
        [string]$Message,
        [string]$Level = "INFO",
        [string]$LogPath = $script:LogPath
    )
    
    if ($EnableLogging) {
        $TimeStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        $LogEntry = "[$TimeStamp] [$Level] $Message"
        
        try {
            $LogEntry | Out-File -FilePath $LogPath -Append -ErrorAction Stop
        }
        catch {
            Write-Host $LogEntry -ForegroundColor Red
        }
        
        # Color-coded console output
        $Color = @{
            "INFO" = "White"
            "WARNING" = "Yellow"
            "ERROR" = "Red"
            "DEBUG" = "Gray"
        }[$Level]
        
        Write-Host $LogEntry -ForegroundColor $Color
    }
}

# --- Excel Helper Functions ---
function Get-ColumnIndex {
    param (
        [object]$Worksheet,
        [string]$HeaderName,
        [int]$DefaultColumn = 0
    )
    
    $HeaderRow = 1
    $UsedColumnsCount = $Worksheet.UsedRange.Columns.Count
    
    for ($Col = 1; $Col -le $UsedColumnsCount; $Col++) {
        $CellValue = $Worksheet.Cells.Item($HeaderRow, $Col).Value2
        if ($CellValue -eq $HeaderName) {
            return $Col
        }
    }
    
    return $DefaultColumn
}

function Get-CellValue {
    param (
        [object]$Worksheet,
        [int]$Row,
        [int]$Column
    )
    
    if ($Column -gt 0) {
        $value = $Worksheet.Cells.Item($Row, $Column).Value2
        if ($null -eq $value -or $value -eq "") {
            return $null
        }
        return $value
    }
    return $null
}

# --- Report Generation Functions --
function Generate-HTMLReport {
    param (
        [array]$LicenseData,
        [string]$ReportPath = $script:ReportPath,
        [string]$CompanyLogo = $script:CompanyLogo,
        [string]$CompanyName = $script:CompanyName
    )

    try {
        # Calculate summary statistics with proper null handling
        if (-not $LicenseData -or $LicenseData.Count -eq 0) {
            Write-Log "Aucune donnée de licence à analyser" "WARNING"
            $expiredCount = $criticalCount = $warningCount = $totalCost = 0
        }
        else {
            $expiredCount = @($LicenseData | Where-Object { $_.DaysUntilExpiration -lt 0 }).Count
            $criticalCount = @($LicenseData | Where-Object { 
                $_.DaysUntilExpiration -ge 0 -and $_.DaysUntilExpiration -le $CriticalDays 
            }).Count
            $warningCount = @($LicenseData | Where-Object { 
                $_.DaysUntilExpiration -gt $CriticalDays -and $_.DaysUntilExpiration -le $WarningDays 
            }).Count
        
            $validCosts = @($LicenseData | Where-Object { $null -ne $_.RenewalCost -and $_.RenewalCost -ne "" })
            $totalCost = if ($validCosts.Count -gt 0) { 
                ($validCosts | Measure-Object -Property RenewalCost -Sum).Sum 
            } else { 0 }
        }
        
        # Modern color palette with direct values for offline use
        $PrimaryColor = "#4F46E5"       # Indigo
        $SecondaryColor = "#4338CA"     # Indigo darker
        $SuccessColor = "#10B981"       # Emerald
        $WarningColor = "#F59E0B"       # Amber
        $CriticalColor = "#EF4444"      # Red
        $ExpiredColor = "#B91C1C"       # Dark red
        $TextColor = "#1F2937"          # Gray 800
        $TextMuted = "#6B7280"          # Gray 500
        $BackgroundColor = "#F9FAFB"    # Gray 50
        $CardBackground = "#FFFFFF"     # White
        $BorderColor = "#E5E7EB"        # Gray 200
        
        $HTML = @"
<!DOCTYPE html>
<html lang="fr">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>$($Language.Strings.'Report-Title') - $CompanyName</title>
    <style>
        /* Reset and base styles */
        *, *::before, *::after { 
            margin: 0; 
            padding: 0; 
            box-sizing: border-box; 
        }
        
        body { 
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif; 
            line-height: 1.6; 
            color: $TextColor; 
            background-color: $BackgroundColor; 
            padding: 0; 
            margin: 0;
            -webkit-font-smoothing: antialiased;
            -moz-osx-font-smoothing: grayscale;
        }
        
        /* Container & main layout */
        .container { 
            max-width: 1200px; 
            margin: 2.5rem auto; 
            padding: 0; 
            background: $CardBackground; 
            box-shadow: 0 10px 25px rgba(0,0,0,0.1); 
            border-radius: 16px; 
            overflow: hidden;
        }
        
        /* Header section */
        .header-banner {
            background: linear-gradient(135deg, $PrimaryColor, $SecondaryColor);
            padding: 2rem;
            color: white;
            position: relative;
        }
        
        header { 
            display: flex; 
            justify-content: space-between; 
            align-items: center; 
            padding: 2rem;
        }
        
        .logo-container { 
            display: flex; 
            align-items: center; 
            gap: 1.2rem;
        }
        
        .logo { 
            width: 66px; 
            height: 66px; 
            background-color: white; 
            color: $SecondaryColor; 
            display: flex; 
            align-items: center; 
            justify-content: center; 
            border-radius: 14px; 
            font-weight: bold; 
            font-size: 1.8rem;
            overflow: hidden;
            box-shadow: 0 4px 12px rgba(0,0,0,0.08);
        }
        
        .report-title h1 { 
            font-size: 2rem; 
            margin: 0; 
            font-weight: 700;
            color: white;
        }
        
        .report-title p { 
            color: rgba(255,255,255,0.9);
            font-size: 1.1rem; 
            margin-top: 0.4rem;
        }
        
        .date-time { 
            font-size: 0.95rem; 
            color: white;
            background: rgba(255,255,255,0.2);
            padding: 0.7rem 1.2rem;
            border-radius: 12px;
            backdrop-filter: blur(5px);
            font-weight: 500;
        }
        
        /* Dashboard stats */
        .dashboard-container {
            padding: 0 2rem;
        }
        
        .dashboard { 
            display: grid; 
            grid-template-columns: repeat(auto-fit, minmax(240px, 1fr)); 
            gap: 1.8rem; 
            margin: -3rem 0 2rem 0;
            position: relative;
            z-index: 10;
        }
        
        .stat-card { 
            background: $CardBackground; 
            padding: 1.8rem 1.5rem; 
            border-radius: 14px; 
            box-shadow: 0 10px 20px rgba(0,0,0,0.08); 
            text-align: center; 
            transition: all 0.3s ease;
            border-top: 5px solid transparent;
            display: flex;
            flex-direction: column;
            justify-content: center;
            gap: 0.7rem;
        }
        
        .stat-card:hover { 
            transform: translateY(-5px); 
            box-shadow: 0 15px 30px rgba(0,0,0,0.12);
        }
        
        .stat-card h3 { 
            font-size: 2.5rem; 
            font-weight: 700;
            line-height: 1.2;
        }
        
        .stat-card p { 
            color: $TextMuted; 
            font-size: 1rem; 
            font-weight: 500;
        }
        
        .stat-expired { border-top-color: $ExpiredColor; }
        .stat-expired h3 { color: $ExpiredColor; }
        
        .stat-critical { border-top-color: $CriticalColor; }
        .stat-critical h3 { color: $CriticalColor; }
        
        .stat-warning { border-top-color: $WarningColor; }
        .stat-warning h3 { color: $WarningColor; }
        
        .stat-cost { border-top-color: $SecondaryColor; }
        .stat-cost h3 { color: $SecondaryColor; }
        
        /* Content sections */
        .content-section { 
            margin: 0 2rem 2.5rem; 
            background: $CardBackground; 
            padding: 2rem; 
            border-radius: 14px; 
            box-shadow: 0 5px 15px rgba(0,0,0,0.05);
            border: 1px solid $BorderColor;
        }
        
        .section-title { 
            color: $SecondaryColor; 
            margin-bottom: 1.5rem; 
            padding-bottom: 0.8rem; 
            border-bottom: 3px solid $PrimaryColor; 
            font-size: 1.4rem; 
            font-weight: 600;
            position: relative;
        }
        
        .section-title::after {
            content: '';
            position: absolute;
            bottom: -3px;
            left: 0;
            width: 60px;
            height: 3px;
            background-color: $SecondaryColor;
        }
        
        /* Tables */
        table { 
            width: 100%; 
            border-collapse: separate; 
            border-spacing: 0;
            margin-top: 1rem; 
            font-size: 0.92rem; 
            border-radius: 8px;
            overflow: hidden;
        }
        
        th { 
            background-color: #F3F4F6; 
            color: $SecondaryColor; 
            font-weight: 600; 
            text-align: left; 
            padding: 14px 16px; 
            border-bottom: 2px solid $BorderColor;
            position: sticky;
            top: 0;
        }
        
        td { 
            padding: 14px 16px; 
            border-bottom: 1px solid $BorderColor; 
        }
        
        tr:last-child td { 
            border-bottom: none; 
        }
        
        tbody tr { 
            transition: background-color 0.2s ease; 
        }
        
        tbody tr:hover { 
            background-color: rgba(79, 70, 229, 0.05); 
        }
        
        /* Status badges */
        .status-badge { 
            padding: 6px 12px; 
            border-radius: 20px; 
            font-size: 0.8rem; 
            font-weight: 600; 
            display: inline-block; 
            text-align: center; 
            min-width: 100px;
            letter-spacing: 0.5px;
        }
        
        .status-expired { 
            background-color: rgba(185, 28, 28, 0.1); 
            color: $ExpiredColor; 
        }
        
        .status-critical { 
            background-color: rgba(239, 68, 68, 0.1); 
            color: $CriticalColor; 
        }
        
        .status-warning { 
            background-color: rgba(245, 158, 11, 0.15); 
            color: $WarningColor; 
        }
        
        .status-ok { 
            background-color: rgba(16, 185, 129, 0.1); 
            color: $SuccessColor; 
        }
        
        /* Priority indicators */
        .priority-high { 
            font-weight: bold; 
            color: $CriticalColor;
            padding: 4px 10px;
            background-color: rgba(239, 68, 68, 0.1);
            border-radius: 6px;
            display: inline-block;
        }
        
        .priority-medium { 
            color: $WarningColor;
            padding: 4px 10px;
            background-color: rgba(245, 158, 11, 0.15);
            border-radius: 6px;
            display: inline-block;
        }
        
        .priority-low { 
            color: $SuccessColor;
            padding: 4px 10px;
            background-color: rgba(16, 185, 129, 0.1);
            border-radius: 6px;
            display: inline-block;
        }
        
        /* Days remaining */
        .days-remaining { 
            font-weight: 600;
            padding: 5px 10px;
            border-radius: 6px;
            display: inline-block;
            text-align: center;
            min-width: 40px;
        }
        
        .days-negative { 
            color: white;
            background-color: $ExpiredColor; 
        }
        
        .days-critical { 
            color: white;
            background-color: $CriticalColor; 
        }
        
        .days-warning { 
            color: $TextColor;
            background-color: $WarningColor; 
        }
        
        .days-ok { 
            color: white;
            background-color: $SuccessColor; 
        }
        
        /* Summary list */
        .summary-list {
            list-style: none;
            padding: 0;
        }
        
        .summary-list li {
            padding: 14px 16px;
            border-bottom: 1px solid $BorderColor;
            display: flex;
            justify-content: space-between;
            align-items: center;
            border-radius: 8px;
            margin-bottom: 8px;
            background-color: #F9FAFB;
        }
        
        .summary-list li:last-child {
            margin-bottom: 0;
        }
        
        .summary-list li strong {
            color: $SecondaryColor;
            font-weight: 600;
        }
        
        .summary-value {
            font-weight: 600;
            font-size: 1.1rem;
        }
        
        /* Footer */
        footer { 
            text-align: center; 
            color: $TextMuted; 
            font-size: 0.9rem; 
            padding: 2rem; 
            border-top: 1px solid $BorderColor; 
            background-color: #F9FAFB;
        }
        
        /* Action buttons */
        .action-button {
            display: inline-block;
            background-color: $PrimaryColor;
            color: white;
            padding: 10px 20px;
            border-radius: 10px;
            text-decoration: none;
            font-weight: 600;
            transition: all 0.2s ease;
            margin-top: 1.5rem;
            border: none;
            cursor: pointer;
            font-size: 0.95rem;
            box-shadow: 0 4px 6px rgba(79, 70, 229, 0.2);
        }
        
        .action-button:hover {
            background-color: $SecondaryColor;
            transform: translateY(-2px);
            box-shadow: 0 6px 10px rgba(79, 70, 229, 0.3);
        }
        
        /* Print styles */
        @media print {
            body { 
                background: white; 
                color: black;
            }
            
            .container { 
                box-shadow: none; 
                max-width: 100%; 
                margin: 0;
                padding: 0;
            }
            
            .header-banner {
                background: #f3f4f6 !important;
                -webkit-print-color-adjust: exact;
                print-color-adjust: exact;
                color-adjust: exact;
            }
            
            .report-title h1, 
            .date-time, 
            .report-title p {
                color: $TextColor !important;
            }
            
            .content-section, 
            .stat-card { 
                box-shadow: none !important; 
                border: 1px solid #ddd !important; 
                break-inside: avoid;
            }
            
            .dashboard {
                margin-top: 1rem;
            }
            
            .stat-card {
                box-shadow: none !important;
            }
            
            .no-print { 
                display: none !important; 
            }
            
            .action-button { 
                display: none !important; 
            }
            
            table {
                page-break-inside: auto;
            }
            
            tr {
                page-break-inside: avoid;
                page-break-after: auto;
            }
        }
        
        /* Responsive design */
        @media (max-width: 768px) {
            .container {
                margin: 1rem;
                border-radius: 12px;
            }
            
            .header-banner,
            header {
                padding: 1.5rem;
            }
            
            header {
                flex-direction: column;
                gap: 1.5rem;
                align-items: flex-start;
            }
            
            .date-time {
                text-align: left;
                width: 100%;
            }
            
            .dashboard-container {
                padding: 0 1.5rem;
            }
            
            .dashboard {
                grid-template-columns: 1fr;
                margin-top: -1.5rem;
            }
            
            .content-section {
                margin: 0 1.5rem 2rem;
                padding: 1.5rem;
            }
            
            table {
                font-size: 0.85rem;
            }
            
            td, th {
                padding: 12px 10px;
            }
            
            .status-badge {
                min-width: auto;
                padding: 4px 8px;
            }
            
            .summary-list li {
                flex-direction: column;
                align-items: flex-start;
                gap: 0.5rem;
            }
            
            .summary-value {
                font-size: 1rem;
            }
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header-banner">
            <header>
                <div class="logo-container">
                    <div class="logo">$(if (Test-Path $CompanyLogo) { "<img src='$CompanyLogo' alt='Logo' width='66' height='66'>" } else { $CompanyName.Substring(0, 1) })</div>
                    <div class="report-title">
                        <h1>$($Language.Strings.'Report-Title')</h1>
                        <p>$CompanyName</p>
                    </div>
                </div>
                <div class="date-time">
                    $(Get-Date -Format "dddd, dd MMMM yyyy HH:mm")
                </div>
            </header>
        </div>

        <div class="dashboard-container">
            <div class="dashboard">
                <div class="stat-card stat-expired">
                    <h3>$expiredCount</h3>
                    <p>$($Language.Strings.'Expired')</p>
                </div>
                <div class="stat-card stat-critical">
                    <h3>$criticalCount</h3>
                    <p>$($Language.Strings.'Critical') (≤ $CriticalDays jours)</p>
                </div>
                <div class="stat-card stat-warning">
                    <h3>$warningCount</h3>
                    <p>$($Language.Strings.'Warning') (≤ $WarningDays jours)</p>
                </div>
                <div class="stat-card stat-cost">
                    <h3>$(if ($totalCost) { "$totalCost €" } else { "N/A" })</h3>
                    <p>$($Language.Strings.'RenewalCost')</p>
                </div>
            </div>
        </div>

        <div class="content-section">
            <h2 class="section-title">$($Language.Strings.'Licenses')</h2>
            <div style="overflow-x: auto;">
                <table>
                    <thead>
                        <tr>
                            <th>$($Language.Strings.'Priority')</th>
                            <th>$($Language.Strings.'Software')</th>
                            <th>$($Language.Strings.'ExpirationDate')</th>
                            <th>$($Language.Strings.'DaysRemaining')</th>
                            <th>$($Language.Strings.'LicenseType')</th>
                            <th>$($Language.Strings.'Owner')</th>
                            <th>$($Language.Strings.'RenewalCost')</th>
                            <th>$($Language.Strings.'Status')</th>
                        </tr>
                    </thead>
                    <tbody>
"@

        foreach ($License in $LicenseData) {
            if ($License.DaysUntilExpiration -lt 0) {
                $StatusClass = "status-expired"
                $DaysClass = "days-negative"
                $PriorityClass = "priority-high"
                $Status = $Language.Strings.'Expired'
                $Priority = $Language.Strings.'Critical'
            }
            elseif ($License.DaysUntilExpiration -le $CriticalDays) {
                $StatusClass = "status-critical"
                $DaysClass = "days-critical"
                $PriorityClass = "priority-high"
                $Status = $Language.Strings.'Critical'
                $Priority = $Language.Strings.'High'
            }
            elseif ($License.DaysUntilExpiration -le $WarningDays) {
                $StatusClass = "status-warning"
                $DaysClass = "days-warning"
                $PriorityClass = "priority-medium"
                $Status = $Language.Strings.'Warning'
                $Priority = $Language.Strings.'Medium'
            }
            else {
                $StatusClass = "status-ok"
                $DaysClass = "days-ok"
                $PriorityClass = "priority-low"
                $Status = $Language.Strings.'OK'
                $Priority = $Language.Strings.'Low'
            }
            
            $HTML += @"
                        <tr>
                            <td><span class="$PriorityClass">$Priority</span></td>
                            <td><strong>$($License.Software)</strong></td>
                            <td>$($License.ExpirationDate.ToString("dd/MM/yyyy"))</td>
                            <td><span class="days-remaining $DaysClass">$($License.DaysUntilExpiration)</span></td>
                            <td>$($License.LicenseType)</td>
                            <td>$($License.Owner)</td>
                            <td>$(if ($License.RenewalCost) { "$($License.RenewalCost) €" } else { "N/A" })</td>
                            <td><span class="status-badge $StatusClass">$Status</span></td>
                        </tr>
"@
        }
        
        $HTML += @"
                    </tbody>
                </table>
            </div>
            <div class="no-print" style="text-align: right;">
                <button class="action-button" onclick="window.print(); return false;">Imprimer ce rapport</button>
            </div>
        </div>
        
        <div class="content-section">
            <h2 class="section-title">$($Language.Strings.'Summary')</h2>
            <ul class="summary-list">
                <li>
                    <strong>$($Language.Strings.'Expired'):</strong>
                    <span class="summary-value $(if ($expiredCount -gt 0) { 'priority-high' })">$expiredCount licences</span>
                </li>
                <li>
                    <strong>$($Language.Strings.'Critical'):</strong>
                    <span class="summary-value $(if ($criticalCount -gt 0) { 'priority-high' })">$criticalCount licences (expirent dans $CriticalDays jours ou moins)</span>
                </li>
                <li>
                    <strong>$($Language.Strings.'Warning'):</strong>
                    <span class="summary-value $(if ($warningCount -gt 0) { 'priority-medium' })">$warningCount licences (expirent dans $WarningDays jours ou moins)</span>
                </li>
                <li>
                    <strong>$($Language.Strings.'RenewalCost'):</strong>
                    <span class="summary-value">$(if ($totalCost) { "$totalCost €" } else { "N/A" })</span>
                </li>
            </ul>
        </div>
        
        <footer>
            <p>$($Language.Strings.'GeneratedBy') $(Get-Date -Format "yyyy-MM-dd HH:mm") | Rapport généré automatiquement</p>
        </footer>
    </div>
</body>
</html>
"@

        $HTML | Out-File -FilePath $ReportPath -Force -Encoding UTF8
        Write-Log "Rapport HTML généré à: $ReportPath"
        return $true
    }
    catch {
        Write-Log "Erreur lors de la génération du rapport: $($_.Exception.Message)" "ERROR"
        return $false
    }
}

# --- Main Processing ---
try {
    # Validate all inputs before proceeding
    Validate-Inputs
    
    # Create directories if they don't exist
    if (-not (Test-Path $BasePath)) { 
        New-Item -ItemType Directory -Path $BasePath -Force | Out-Null 
        Write-Log "Répertoire de base créé: $BasePath" "INFO"
    }
    
    if (-not (Test-Path $BackupFolder) -and $AutoBackupExcel) { 
        New-Item -ItemType Directory -Path $BackupFolder -Force | Out-Null 
        Write-Log "Répertoire de sauvegarde créé: $BackupFolder" "INFO"
    }

    Write-Log "Démarrage du script de surveillance des licences" "INFO"
    
    # Backup Excel file if enabled (once per month)
    if ($AutoBackupExcel) {
        Backup-ExcelFile -SourcePath $ExcelFilePath -BackupFolder $BackupFolder
    }
    
    # Initialize Excel
    $Excel = $null
    $Workbook = $null
    $Worksheet = $null
    
    try {
        $Excel = New-Object -ComObject Excel.Application
        $Excel.Visible = $false
        $Excel.DisplayAlerts = $false

        # Open the workbook
        $Workbook = $Excel.Workbooks.Open($ExcelFilePath)
        $Worksheet = $Workbook.Sheets.Item(1)

        # Detect column indexes
        $ColumnIndexes = @{
            Software = Get-ColumnIndex -Worksheet $Worksheet -HeaderName $ColumnMap.Software -DefaultColumn 1
            ExpirationDate = Get-ColumnIndex -Worksheet $Worksheet -HeaderName $ColumnMap.ExpirationDate -DefaultColumn 2
            LicenseType = Get-ColumnIndex -Worksheet $Worksheet -HeaderName $ColumnMap.LicenseType
            Owner = Get-ColumnIndex -Worksheet $Worksheet -HeaderName $ColumnMap.Owner
            LastCheck = Get-ColumnIndex -Worksheet $Worksheet -HeaderName $ColumnMap.LastCheck
            LastEmail = Get-ColumnIndex -Worksheet $Worksheet -HeaderName $ColumnMap.LastEmail 
            RenewalCost = Get-ColumnIndex -Worksheet $Worksheet -HeaderName $ColumnMap.RenewalCost
        }

        # Get data range
        $LastRow = $Worksheet.UsedRange.Rows.Count
        $LicenseData = @()
        $AlertsToSend = @()
        
        Write-Log "Traitement de $($LastRow-1) enregistrements de licence" "INFO"

        # Process each license
        for ($Row = 2; $Row -le $LastRow; $Row++) {
            try {
                $Software = Get-CellValue -Worksheet $Worksheet -Row $Row -Column $ColumnIndexes.Software
                $ExpirationDateString = Get-CellValue -Worksheet $Worksheet -Row $Row -Column $ColumnIndexes.ExpirationDate

                # Skip empty rows
                if ([string]::IsNullOrEmpty($Software) -or [string]::IsNullOrEmpty($ExpirationDateString)) {
                    continue
                }

                # Get additional data (handling empty values)
                $LicenseInfo = @{
                    Software = $Software
                    Row = $Row
                    LicenseType = Get-CellValue -Worksheet $Worksheet -Row $Row -Column $ColumnIndexes.LicenseType
                    Owner = Get-CellValue -Worksheet $Worksheet -Row $Row -Column $ColumnIndexes.Owner
                    RenewalCost = Get-CellValue -Worksheet $Worksheet -Row $Row -Column $ColumnIndexes.RenewalCost
                }

                # Parse expiration date
                try {
                    $LicenseInfo["ExpirationDate"] = [datetime]::ParseExact($ExpirationDateString, "dd/MM/yyyy", [System.Globalization.CultureInfo]::InvariantCulture)
                    $LicenseInfo["DaysUntilExpiration"] = ($LicenseInfo.ExpirationDate - (Get-Date)).Days
                }
                catch {
                    Write-Log "Format de date invalide pour '$Software' dans la ligne $Row. Format attendu: DD/MM/YYYY. Ignoré." "WARNING"
                    continue
                }

                # Get last email date if available
                $LastEmailDate = $null
                if ($ColumnIndexes.LastEmail -gt 0) {
                    $LastEmailString = Get-CellValue -Worksheet $Worksheet -Row $Row -Column $ColumnIndexes.LastEmail
                    if (-not [string]::IsNullOrEmpty($LastEmailString)) {
                        try {
                            $LastEmailDate = [datetime]::ParseExact($LastEmailString, "dd/MM/yyyy", [System.Globalization.CultureInfo]::InvariantCulture)
                        }
                        catch {
                            Write-Log "Format de date de dernier email invalide pour '$Software'. Traité comme aucun email précédent." "WARNING"
                        }
                    }
                }

                # Determine alert priority
                if ($LicenseInfo.DaysUntilExpiration -lt 0) {
                    $Priority = "Critical"
                    $ShouldAlert = (-not $LastEmailDate) -or ((Get-Date) - $LastEmailDate).Days -ge $ExpiredFrequency
                }
                elseif ($LicenseInfo.DaysUntilExpiration -le $CriticalDays) {
                    $Priority = "Critical"
                    $ShouldAlert = (-not $LastEmailDate) -or ((Get-Date) - $LastEmailDate).Days -ge $ExpiredFrequency
                }
                elseif ($LicenseInfo.DaysUntilExpiration -le $WarningDays) {
                    $Priority = "Warning"
                    $ShouldAlert = (-not $LastEmailDate) -or ((Get-Date) - $LastEmailDate).Days -ge $WarningFrequency
                }
                else {
                    $Priority = "Normal"
                    $ShouldAlert = $false
                }

                $LicenseInfo["Priority"] = $Priority
                $LicenseInfo["ShouldAlert"] = $ShouldAlert

                # Update last check date
                if ($ColumnIndexes.LastCheck -gt 0) {
                    $cell = $Worksheet.Cells.Item($Row, $ColumnIndexes.LastCheck)
                    $cell.NumberFormat = "@"  # Forcer le format Texte
                    $cell.Value2 = (Get-Date).ToString("dd/MM/yyyy")
                }

                $LicenseData += [PSCustomObject]$LicenseInfo
                
                if ($LicenseInfo.ShouldAlert) {
                    $AlertsToSend += [PSCustomObject]$LicenseInfo
                }
            }
            catch {
                Write-Log "Erreur lors du traitement de la ligne $Row : $($_.Exception.GetType().Name): $($_.Exception.Message)" "ERROR"
            }
        }

        # Save workbook after all updates
        $Workbook.Save()

        # Process alerts
        foreach ($Alert in $AlertsToSend) {
            try {
                # Build email subject 
                if ($Alert.DaysUntilExpiration -lt 0) {
                    $Subject = "[EXPIRÉ] Alerte d'expiration de licence - $($Alert.Software)"
                }
                elseif ($Alert.DaysUntilExpiration -le $CriticalDays) {
                    $Subject = "[CRITIQUE] Alerte d'expiration de licence - $($Alert.Software)"
                }
                else {
                    $Subject = "[AVERTISSEMENT] Alerte d'expiration de licence - $($Alert.Software)"
                }
        
                # Create HTML email body
                $HTMLBody = Create-AlertEmailBody -Alert $Alert -CriticalDays $CriticalDays -WarningDays $WarningDays -CompanyName $CompanyName
                
                # Prepare email parameters with encoding specified
                $MailParams = @{
                    From = $FromAddress
                    To = $ToAddress
                    Subject = $Subject
                    Body = $HTMLBody
                    BodyAsHtml = $true
                    SmtpServer = $SMTPServer
                    Port = $SMTPPort
                    Encoding = [System.Text.Encoding]::UTF8
                }
        
                # Add credentials if provided
                if (-not [string]::IsNullOrEmpty($SMTPUsername) -and -not [string]::IsNullOrEmpty($SMTPPassword)) {
                    $MailParams["Credential"] = New-Object System.Management.Automation.PSCredential ($SMTPUsername, (ConvertTo-SecureString $SMTPPassword -AsPlainText -Force))
                }
        
                if ($CCAddresses.Count -gt 0) {
                    $MailParams["Cc"] = $CCAddresses -join ","
                }
        
                # Send email with retry logic
                $emailSent = Send-EmailWithRetry -MailParams $MailParams -MaxAttempts $EmailRetryAttempts -InitialDelay $EmailRetryDelay
        
                if ($emailSent) {
                    Write-Log "Email envoyé pour '$($Alert.Software)'" "INFO"
                    
                    # Update last email date if sent successfully
                    if ($ColumnIndexes.LastEmail -gt 0) {
                        $Worksheet.Cells.Item($Alert.Row, $ColumnIndexes.LastEmail).NumberFormat = "@"
                        $Worksheet.Cells.Item($Alert.Row, $ColumnIndexes.LastEmail).Value2 = (Get-Date).ToString("dd/MM/yyyy")
                        Write-Log "Date du dernier email mise à jour pour '$($Alert.Software)' à la ligne $($Alert.Row)" "INFO"
                        $Workbook.Save()
                    }
                }
            }
            catch {
                Write-Log "Erreur lors de l'envoi de l'email pour '$($Alert.Software)': $($_.Exception.GetType().Name): $($_.Exception.Message)" "ERROR"
            }
        }
       
        # Generate report if enabled
        if ($EnableReporting -and $LicenseData.Count -gt 0) {
            $reportGenerated = Generate-HTMLReport -LicenseData $LicenseData
            if ($reportGenerated) {
                Write-Log "Rapport généré avec succès" "INFO"
            }
        }
        
        Write-Log "Surveillance des licences terminée avec succès" "INFO"
    }
    catch [System.Runtime.InteropServices.COMException] {
        Write-Log "Erreur Excel COM: $($_.Exception.Message)" "ERROR"
        throw
    }
    finally {
        # Clean up Excel
        if ($Workbook) { 
            try { $Workbook.Close($false) } catch { Write-Log "Error closing workbook: $($_.Exception.Message)" "ERROR" }
        }
        if ($Excel) { 
            try { $Excel.Quit() } catch { Write-Log "Error quitting Excel: $($_.Exception.Message)" "ERROR" }
        }
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Worksheet) | Out-Null
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Workbook) | Out-Null
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Excel) | Out-Null
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
}
catch [System.IO.IOException] {
    Write-Log "Erreur d'E/S: $($_.Exception.Message)" "ERROR"
    exit 1
}
catch [System.Net.Mail.SmtpException] {
    Write-Log "Erreur d'envoi d'email: $($_.Exception.Message)" "ERROR"
    exit 1
}
catch {
    Write-Log "Erreur critique: $($_.Exception.GetType().Name): $($_.Exception.Message)" "ERROR"
    exit 1
}
