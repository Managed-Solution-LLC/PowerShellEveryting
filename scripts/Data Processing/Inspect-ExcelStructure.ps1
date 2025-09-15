# Excel Structure Inspector
# This script examines the structure of your Excel file to identify column names and data

param(
    [Parameter(Mandatory=$true)]
    [string]$ExcelFilePath,
    
    [Parameter(Mandatory=$false)]
    [string]$WorksheetName = "Sheet1"
)

# Check if ImportExcel module is available
if (-not (Get-Module -Name ImportExcel -ListAvailable)) {
    Write-Host "ImportExcel module not found. Installing..." -ForegroundColor Yellow
    Install-Module -Name ImportExcel -Force -AllowClobber
    Write-Host "ImportExcel module installed successfully." -ForegroundColor Green
}

Import-Module ImportExcel -Force

try {
    Write-Host "Inspecting Excel file: $ExcelFilePath" -ForegroundColor Cyan
    Write-Host "Worksheet: $WorksheetName" -ForegroundColor Cyan
    
    # Try to get worksheet names first
    try {
        $worksheets = Get-ExcelSheetInfo -Path $ExcelFilePath
        Write-Host "`nAvailable Worksheets:" -ForegroundColor Yellow
        $worksheets | ForEach-Object { Write-Host "  - $($_.Name)" -ForegroundColor White }
    }
    catch {
        Write-Host "Could not get worksheet info, proceeding with specified worksheet..." -ForegroundColor Yellow
    }
    
    # Import just the first few rows to examine structure
    Write-Host "`nImporting first 5 rows..." -ForegroundColor Cyan
    $sampleData = Import-Excel -Path $ExcelFilePath -WorksheetName $WorksheetName -EndRow 5
    
    if ($sampleData) {
        Write-Host "Found $($sampleData.Count) sample rows" -ForegroundColor Green
        
        # Get column names/headers
        $headers = $sampleData[0].PSObject.Properties.Name
        Write-Host "`nColumn Headers Found:" -ForegroundColor Yellow
        for ($i = 0; $i -lt $headers.Count; $i++) {
            $columnLetter = [char](65 + $i)  # Convert to A, B, C, D, E, etc.
            Write-Host "  Column $columnLetter ($($i+1)): '$($headers[$i])'" -ForegroundColor White
        }
        
        # Show first few rows of data
        Write-Host "`nFirst 3 rows of data:" -ForegroundColor Magenta
        $sampleData | Select-Object -First 3 | Format-Table -AutoSize
        
        # Check if there's data that looks like your status data
        Write-Host "`nLooking for columns that might contain status data..." -ForegroundColor Cyan
        foreach ($header in $headers) {
            $columnLetter = [char](65 + [array]::IndexOf($headers, $header))
            $sampleValue = $sampleData[0].$header
            if ($sampleValue -and $sampleValue.ToString().Contains("Status") -or $sampleValue.ToString().Contains("HARDWARE") -or $sampleValue.ToString().Contains("Details")) {
                Write-Host "  Column $columnLetter ($header) might contain status data!" -ForegroundColor Green
                Write-Host "    Sample: $($sampleValue.ToString().Substring(0, [Math]::Min(100, $sampleValue.ToString().Length)))..." -ForegroundColor Gray
            }
        }
        
        # Try to find the 5th column (E) specifically
        if ($headers.Count -ge 5) {
            $columnE = $headers[4]  # 5th column (index 4)
            Write-Host "`nColumn E is named: '$columnE'" -ForegroundColor Yellow
            Write-Host "Sample data from Column E:" -ForegroundColor Yellow
            for ($i = 0; $i -lt [Math]::Min(3, $sampleData.Count); $i++) {
                $value = $sampleData[$i].$columnE
                if ($value) {
                    Write-Host "Row $($i+1): $($value.ToString().Substring(0, [Math]::Min(150, $value.ToString().Length)))..." -ForegroundColor Gray
                } else {
                    Write-Host "Row $($i+1): (empty)" -ForegroundColor Gray
                }
            }
        } else {
            Write-Host "`nThis Excel file has fewer than 5 columns. Column E does not exist." -ForegroundColor Red
        }
    } else {
        Write-Host "No data found in the Excel file." -ForegroundColor Red
    }
    
} catch {
    Write-Error "An error occurred: $($_.Exception.Message)"
    Write-Host "Stack trace: $($_.ScriptStackTrace)" -ForegroundColor Red
}
