# Check PowerShell version
$PSVersionTable.PSVersion

# Check if Excel is installed
try {
    $excel = New-Object -ComObject Excel.Application -ErrorAction Stop
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    Write-Host "Excel is available." -ForegroundColor Green
} catch {
    Write-Warning "Excel not found. Install Excel or use 'ImportExcel' module."
}

# Check .NET version
[System.Environment]::Version
