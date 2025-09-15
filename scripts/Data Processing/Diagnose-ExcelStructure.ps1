# Diagnose Excel File Structure
# This script will help us understand the structure of your Excel file

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
}

Import-Module ImportExcel -Force

try {
    Write-Host "Analyzing Excel file: $ExcelFilePath" -ForegroundColor Cyan
    
    # Get worksheet names
    Write-Host "`nAvailable worksheets:" -ForegroundColor Yellow
    $worksheets = Get-ExcelSheetInfo -Path $ExcelFilePath
    $worksheets | ForEach-Object { Write-Host "  - $($_.Name)" -ForegroundColor White }
    
    # Import first few rows to analyze structure
    Write-Host "`nImporting first 5 rows from worksheet '$WorksheetName'..." -ForegroundColor Cyan
    $sampleData = Import-Excel -Path $ExcelFilePath -WorksheetName $WorksheetName -StartRow 1 -EndRow 5
    
    if ($sampleData) {
        Write-Host "`nColumn Headers Found:" -ForegroundColor Yellow
        $headers = $sampleData[0].PSObject.Properties.Name
        for ($i = 0; $i -lt $headers.Count; $i++) {
            $columnLetter = [char](65 + $i)  # Convert to A, B, C, D, E, etc.
            Write-Host "  Column $columnLetter ($($i+1)): $($headers[$i])" -ForegroundColor White
        }
        
        Write-Host "`nSample data from first row:" -ForegroundColor Yellow
        $firstRow = $sampleData[0]
        foreach ($property in $firstRow.PSObject.Properties) {
            $value = if ($property.Value) { $property.Value.ToString().Substring(0, [Math]::Min(100, $property.Value.ToString().Length)) } else { "(empty)" }
            Write-Host "  $($property.Name): $value" -ForegroundColor Gray
        }
        
        # Check if there's data that looks like your status information
        Write-Host "`nLooking for columns that might contain status data..." -ForegroundColor Yellow
        foreach ($property in $firstRow.PSObject.Properties) {
            if ($property.Value -and $property.Value.ToString().Contains("Status") -or 
                $property.Value.ToString().Contains("HARDWARE") -or 
                $property.Value.ToString().Contains("----")) {
                Write-Host "  ** Potential status column found: $($property.Name)" -ForegroundColor Green
                Write-Host "     Sample content: $($property.Value.ToString().Substring(0, [Math]::Min(200, $property.Value.ToString().Length)))" -ForegroundColor Gray
            }
        }
        
        # Show how many total rows
        $allData = Import-Excel -Path $ExcelFilePath -WorksheetName $WorksheetName
        Write-Host "`nTotal rows in worksheet: $($allData.Count)" -ForegroundColor Cyan
        
    } else {
        Write-Host "No data found in the specified worksheet." -ForegroundColor Red
    }
    
} catch {
    Write-Error "Error analyzing Excel file: $($_.Exception.Message)"
}
