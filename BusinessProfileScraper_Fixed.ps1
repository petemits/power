# ===================================================================
# Business Profile Web Search - SIMPLE WORKING VERSION
# ===================================================================
# Minimal script with error handling that actually works
# ===================================================================

# 1. Setup and Configuration
# -------------------------------------------------------------------
Write-Host "Starting Business Profile Scraper..." -ForegroundColor Cyan

# Configuration
$SearchQuery = "software development company New York"
$OutputExcelPath = "C:\Temp\BusinessProfiles.xlsx"

# Create output directory if needed
$outputDir = Split-Path -Path $OutputExcelPath -Parent
if (-not (Test-Path -Path $outputDir)) {
    New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
}

# 2. Initialize Chrome Browser
# -------------------------------------------------------------------
Write-Host "Starting Chrome browser..." -ForegroundColor Cyan

try {
    # Import Selenium module
    Import-Module -Name Selenium -ErrorAction Stop
    
    # Start Chrome - will use ChromeDriver from PATH
    $Driver = Start-SeChrome -Headless
    Write-Host "Browser started successfully" -ForegroundColor Green
}
catch {
    Write-Error "Failed to start browser: $_"
    Write-Host "Make sure:" -ForegroundColor Yellow
    Write-Host "1. Chrome is installed" -ForegroundColor Yellow
    Write-Host "2. ChromeDriver is in PATH (run 'chromedriver --version' to check)" -ForegroundColor Yellow
    Write-Host "3. Selenium module is installed (run 'Install-Module -Name Selenium -Force')" -ForegroundColor Yellow
    exit
}

# 3. Perform Google Search
# -------------------------------------------------------------------
Write-Host "Searching for: '$SearchQuery'..." -ForegroundColor Cyan

try {
    $SearchUrl = "https://www.google.com/search?q=$([System.Web.HttpUtility]::UrlEncode($SearchQuery))"
    Enter-SeUrl -Url $SearchUrl -Driver $Driver
    
    # Wait for page to load
    Start-Sleep -Seconds 3
    
    Write-Host "Page loaded successfully" -ForegroundColor Green
}
catch {
    Write-Error "Failed to load search page: $_"
    Stop-SeDriver -Driver $Driver
    exit
}

# 4. Extract Search Results
# -------------------------------------------------------------------
Write-Host "Extracting search results..." -ForegroundColor Cyan

$BusinessProfiles = @()

try {
    # Get search results - try different selectors
    $results = Get-SeElement -Driver $Driver -By CssSelector -Value "div.g" -ErrorAction SilentlyContinue
    
    if (-not $results) {
        $results = Get-SeElement -Driver $Driver -By CssSelector -Value "div[data-snf]" -ErrorAction SilentlyContinue
    }
    
    if (-not $results) {
        $results = Get-SeElement -Driver $Driver -By CssSelector -Value "div.tF2Cxc" -ErrorAction SilentlyContinue
    }
    
    if ($results) {
        Write-Host "Found $($results.Count) results" -ForegroundColor Green
        
        $count = 0
        foreach ($result in $results) {
            $count++
            
            try {
                # Extract title
                $title = "N/A"
                $titleElement = Get-SeElement -Element $result -By CssSelector -Value "h3" -ErrorAction SilentlyContinue
                if ($titleElement) {
                    $title = $titleElement.Text.Trim()
                }
                
                # Extract URL
                $url = "N/A"
                $linkElement = Get-SeElement -Element $result -By CssSelector -Value "a[href]" -ErrorAction SilentlyContinue
                if ($linkElement) {
                    $url = $linkElement.GetAttribute("href")
                }
                
                # Extract description
                $desc = "N/A"
                $descElement = Get-SeElement -Element $result -By CssSelector -Value "div[data-sncf]" -ErrorAction SilentlyContinue
                if (-not $descElement) {
                    $descElement = Get-SeElement -Element $result -By CssSelector -Value "div.VwiC3b" -ErrorAction SilentlyContinue
                }
                if ($descElement) {
                    $desc = $descElement.Text.Trim()
                }
                
                # Create profile object
                $profile = [PSCustomObject]@{
                    Rank = $count
                    Title = $title
                    URL = $url
                    Description = $desc
                    Date = Get-Date -Format "yyyy-MM-dd"
                }
                
                $BusinessProfiles += $profile
                
                Write-Host "  Extracted: $title" -ForegroundColor Gray
            }
            catch {
                Write-Host "  Error extracting result $count" -ForegroundColor Red
            }
        }
    }
    else {
        Write-Host "No results found. Google's layout may have changed." -ForegroundColor Yellow
    }
}
catch {
    Write-Error "Error during extraction: $_"
}

# 5. Save Results to Excel
# -------------------------------------------------------------------
if ($BusinessProfiles.Count -gt 0) {
    Write-Host "`nSaving $($BusinessProfiles.Count) profiles to Excel..." -ForegroundColor Cyan
    
    try {
        # Try to use ImportExcel module
        if (Get-Module -ListAvailable -Name ImportExcel) {
            Import-Module -Name ImportExcel -ErrorAction SilentlyContinue
            $BusinessProfiles | Export-Excel -Path $OutputExcelPath -WorksheetName "Profiles" -AutoSize
            
            Write-Host "✅ Excel file saved: $OutputExcelPath" -ForegroundColor Green
        }
        else {
            # Fallback to CSV
            $csvPath = $OutputExcelPath -replace '\.xlsx$', '.csv'
            $BusinessProfiles | Export-Csv -Path $csvPath -NoTypeInformation
            
            Write-Host "⚠ ImportExcel not found. Saved as CSV: $csvPath" -ForegroundColor Yellow
            Write-Host "   Install ImportExcel: 'Install-Module -Name ImportExcel -Force'" -ForegroundColor Yellow
        }
        
        # Show preview
        Write-Host "`nPreview of results:" -ForegroundColor Cyan
        $BusinessProfiles | Select-Object -First 3 | Format-Table -Property Rank, Title, URL -AutoSize
    }
    catch {
        Write-Error "Failed to save results: $_"
    }
}
else {
    Write-Host "No data to save." -ForegroundColor Yellow
}

# 6. Cleanup
# -------------------------------------------------------------------
Write-Host "`nCleaning up..." -ForegroundColor Cyan

try {
    if ($Driver) {
        Stop-SeDriver -Driver $Driver
        Write-Host "Browser closed." -ForegroundColor Gray
    }
}
catch {
    Write-Host "Error during cleanup." -ForegroundColor Red
}

Write-Host "`nScript completed." -ForegroundColor Green