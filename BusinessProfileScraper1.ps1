# ===================================================================
# COMPLETE BUSINESS PROFILE SCRAPER WITH EMBEDDED DRIVER
# ===================================================================
# This script includes its own ChromeDriver and doesn't rely on PATH
# ===================================================================

# 1. CONFIGURATION - SET YOUR PATHS HERE
# -------------------------------------------------------------------
$ProjectRoot = "C:\Users\user\power"  # Change to your project folder
$ChromePath = "C:\Program Files\Google\Chrome\Application\chrome.exe"
$SearchQuery = "software development company New York"
$OutputExcelPath = "$ProjectRoot\BusinessProfiles_$(Get-Date -Format 'yyyyMMdd_HHmmss').xlsx"

# 2. SETUP CHROMEDRIVER IN PROJECT FOLDER
# -------------------------------------------------------------------
Write-Host "Setting up ChromeDriver in project folder..." -ForegroundColor Cyan

# Define driver paths
$ChromeDriverExe = "$ProjectRoot\chromedriver.exe"
$ChromeDriverUrl = "https://chromedriver.storage.googleapis.com/143.0.7499.169/chromedriver_win32.zip"

# Check if we have the correct driver (version 143)
$hasCorrectDriver = $false
if (Test-Path $ChromeDriverExe) {
    try {
        $driverInfo = & $ChromeDriverExe --version 2>&1
        if ($driverInfo -match "143\.0\.7499\.169") {
            $hasCorrectDriver = $true
            Write-Host "✓ ChromeDriver 143.0.7499.169 found in project folder" -ForegroundColor Green
        } else {
            Write-Host "✗ Wrong ChromeDriver version found: $driverInfo" -ForegroundColor Yellow
            Remove-Item $ChromeDriverExe -Force -ErrorAction SilentlyContinue
        }
    } catch {
        Write-Host "✗ Existing ChromeDriver is corrupt" -ForegroundColor Yellow
        Remove-Item $ChromeDriverExe -Force -ErrorAction SilentlyContinue
    }
}

# Download ChromeDriver 143 if needed
if (-not $hasCorrectDriver) {
    Write-Host "Downloading ChromeDriver 143.0.7499.169..." -ForegroundColor Yellow
    try {
        # Create temp folder
        $tempDir = "$env:TEMP\chromedriver_$(Get-Random)"
        New-Item -ItemType Directory -Path $tempDir -Force | Out-Null
        
        # Download and extract
        $zipPath = "$tempDir\chromedriver.zip"
        Invoke-WebRequest -Uri $ChromeDriverUrl -OutFile $zipPath
        
        # Extract using .NET (no external dependencies)
        Add-Type -AssemblyName System.IO.Compression.FileSystem
        [System.IO.Compression.ZipFile]::ExtractToDirectory($zipPath, $tempDir)
        
        # Move to project folder
        $extractedDriver = Get-ChildItem -Path $tempDir -Filter "chromedriver.exe" -Recurse | Select-Object -First 1
        if ($extractedDriver) {
            Copy-Item -Path $extractedDriver.FullName -Destination $ChromeDriverExe -Force
            Write-Host "✓ ChromeDriver 143 downloaded to project folder" -ForegroundColor Green
        } else {
            throw "ChromeDriver not found in downloaded archive"
        }
        
        # Cleanup
        Remove-Item -Path $tempDir -Recurse -Force -ErrorAction SilentlyContinue
    } catch {
        Write-Error "Failed to download ChromeDriver: $_"
        Write-Host "Please manually download ChromeDriver 143 from:" -ForegroundColor Yellow
        Write-Host "https://chromedriver.storage.googleapis.com/index.html?path=143.0.7499.169/" -ForegroundColor Cyan
        Write-Host "And save it as: $ChromeDriverExe" -ForegroundColor Cyan
        exit
    }
}

# 3. VERIFY CHROME INSTALLATION
# -------------------------------------------------------------------
Write-Host "Verifying Chrome installation..." -ForegroundColor Cyan
if (-not (Test-Path $ChromePath)) {
    Write-Error "Chrome not found at: $ChromePath"
    Write-Host "Please update the `$ChromePath` variable in the script" -ForegroundColor Yellow
    exit
}

# Get Chrome version
try {
    $chromeVersion = (Get-Item $ChromePath).VersionInfo.FileVersion
    Write-Host "✓ Chrome version: $chromeVersion" -ForegroundColor Green
    
    # Verify ChromeDriver compatibility
    if ($chromeVersion -notmatch "^143\.") {
        Write-Warning "Chrome version $chromeVersion may not be fully compatible with ChromeDriver 143"
        Write-Host "Consider updating Chrome or using matching ChromeDriver" -ForegroundColor Yellow
    }
} catch {
    Write-Warning "Could not determine Chrome version"
}

# 4. INITIALIZE SELENIUM WITH PROJECT DRIVER
# -------------------------------------------------------------------
Write-Host "Initializing Selenium with project ChromeDriver..." -ForegroundColor Cyan

# Load required assemblies
try {
    # First, try to load Selenium module
    Import-Module -Name Selenium -ErrorAction Stop
    
    # Set environment to use our project ChromeDriver
    $env:PATH = "$ProjectRoot;$env:PATH"
    
    # Start Chrome using our specific driver
    $ChromeOptions = New-SeChromeOptions
    $ChromeOptions.AddArgument("--headless")
    $ChromeOptions.AddArgument("--no-sandbox")
    $ChromeOptions.AddArgument("--disable-dev-shm-usage")
    
    # Create ChromeDriver service pointing to our executable
    $DriverService = [OpenQA.Selenium.Chrome.ChromeDriverService]::CreateDefaultService($ProjectRoot, "chromedriver.exe")
    $DriverService.HideCommandPromptWindow = $true
    
    # Create driver instance
    $Driver = New-Object OpenQA.Selenium.Chrome.ChromeDriver($DriverService, $ChromeOptions)
    
    Write-Host "✓ Browser started with project ChromeDriver" -ForegroundColor Green
} catch {
    Write-Error "Failed to initialize browser: $_"
    Write-Host "Troubleshooting:" -ForegroundColor Yellow
    Write-Host "1. Ensure Selenium module is installed: Install-Module -Name Selenium -Force" -ForegroundColor Yellow
    Write-Host "2. Check ChromeDriver.exe exists at: $ChromeDriverExe" -ForegroundColor Yellow
    Write-Host "3. Run PowerShell as Administrator if permission issues" -ForegroundColor Yellow
    exit
}

# 5. PERFORM WEB SEARCH
# -------------------------------------------------------------------
Write-Host "Searching for: '$SearchQuery'..." -ForegroundColor Cyan

try {
    $SearchUrl = "https://www.google.com/search?q=$([System.Web.HttpUtility]::UrlEncode($SearchQuery))"
    $Driver.Navigate().GoToUrl($SearchUrl)
    
    # Wait for page to load
    Start-Sleep -Seconds 3
    
    Write-Host "✓ Search page loaded" -ForegroundColor Green
} catch {
    Write-Error "Failed to load search page: $_"
    $Driver.Quit()
    exit
}

# 6. EXTRACT BUSINESS PROFILES
# -------------------------------------------------------------------
Write-Host "Extracting business profiles..." -ForegroundColor Cyan

$BusinessProfiles = New-Object System.Collections.Generic.List[PSCustomObject]

try {
    # Wait for results to appear
    $wait = New-Object OpenQA.Selenium.Support.UI.WebDriverWait($Driver, [System.TimeSpan]::FromSeconds(10))
    
    # Try multiple selectors for search results
    $selectors = @(
        "//div[@class='g']",
        "//div[contains(@class, 'MjjYud')]",
        "//div[@data-snf]",
        "//div[@class='tF2Cxc']"
    )
    
    $resultsFound = $false
    foreach ($xpath in $selectors) {
        try {
            $elements = $Driver.FindElements([OpenQA.Selenium.By]::XPath($xpath))
            if ($elements.Count -gt 0) {
                Write-Host "Found $($elements.Count) results using XPath: $xpath" -ForegroundColor Green
                $resultsFound = $true
                
                $resultCount = 0
                foreach ($element in $elements) {
                    $resultCount++
                    if ($resultCount -gt 10) { break } # Limit to 10 results
                    
                    try {
                        # Extract data
                        $title = "N/A"
                        $url = "N/A"
                        $description = "N/A"
                        
                        # Try to get title
                        try {
                            $titleElements = $element.FindElements([OpenQA.Selenium.By]::TagName("h3"))
                            if ($titleElements.Count -gt 0) {
                                $title = $titleElements[0].Text
                            }
                        } catch { }
                        
                        # Try to get URL
                        try {
                            $linkElements = $element.FindElements([OpenQA.Selenium.By]::TagName("a"))
                            if ($linkElements.Count -gt 0) {
                                $url = $linkElements[0].GetAttribute("href")
                            }
                        } catch { }
                        
                        # Try to get description
                        try {
                            $descSelectors = @(
                                [OpenQA.Selenium.By]::ClassName("VwiC3b"),
                                [OpenQA.Selenium.By]::ClassName("MUxGbd"),
                                [OpenQA.Selenium.By]::XPath(".//div[contains(@class, 'VwiC3b')]")
                            )
                            
                            foreach ($descBy in $descSelectors) {
                                try {
                                    $descElements = $element.FindElements($descBy)
                                    if ($descElements.Count -gt 0 -and $descElements[0].Text.Trim().Length -gt 0) {
                                        $description = $descElements[0].Text.Trim()
                                        break
                                    }
                                } catch { }
                            }
                        } catch { }
                        
                        # Only add if we have valid data
                        if ($title -ne "N/A" -or $url -ne "N/A") {
                            $profile = [PSCustomObject]@{
                                Rank = $BusinessProfiles.Count + 1
                                Title = $title
                                URL = $url
                                Description = $description
                                SearchQuery = $SearchQuery
                                DateScraped = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                            }
                            
                            $BusinessProfiles.Add($profile)
                            Write-Host "  ✓ Extracted: $title" -ForegroundColor Gray
                        }
                    } catch {
                        Write-Host "  ⚠ Error processing result $resultCount" -ForegroundColor Yellow
                    }
                }
                break
            }
        } catch { }
    }
    
    if (-not $resultsFound) {
        Write-Warning "No search results found. The page structure may have changed."
    }
} catch {
    Write-Error "Error during extraction: $_"
}

Write-Host "`nExtraction complete: $($BusinessProfiles.Count) profiles found" -ForegroundColor Green

# 7. SAVE TO EXCEL
# -------------------------------------------------------------------
if ($BusinessProfiles.Count -gt 0) {
    Write-Host "Saving to Excel..." -ForegroundColor Cyan
    
    try {
        # Check for ImportExcel module
        $importExcelModule = Get-Module -ListAvailable -Name ImportExcel
        if (-not $importExcelModule) {
            Write-Host "Installing ImportExcel module..." -ForegroundColor Yellow
            Install-Module -Name ImportExcel -Force -Scope CurrentUser
        }
        
        # Export to Excel
        $BusinessProfiles | Export-Excel -Path $OutputExcelPath `
            -WorksheetName "Business Profiles" `
            -TableName "SearchResults" `
            -AutoSize `
            -BoldTopRow `
            -FreezeTopRow `
            -ClearSheet
        
        Write-Host "`n✅ SUCCESS: Excel file created!" -ForegroundColor Green
        Write-Host "   Location: $OutputExcelPath" -ForegroundColor Yellow
        
        # Show preview
        Write-Host "`nPreview of results:" -ForegroundColor Cyan
        Write-Host ("-" * 80) -ForegroundColor DarkGray
        $BusinessProfiles | Select-Object -First 3 | Format-Table -Property Rank, Title, URL -AutoSize | Out-Host
        Write-Host ("-" * 80) -ForegroundColor DarkGray
        
        # Offer to open the file
        $openFile = Read-Host "`nOpen Excel file? (Y/N)"
        if ($openFile -eq "Y" -or $openFile -eq "y") {
            Invoke-Item $OutputExcelPath
        }
    } catch {
        Write-Error "Failed to save to Excel: $_"
        
        # Fallback to CSV
        try {
            $csvPath = $OutputExcelPath -replace '\.xlsx$', '.csv'
            $BusinessProfiles | Export-Csv -Path $csvPath -NoTypeInformation
            Write-Host "Saved as CSV: $csvPath" -ForegroundColor Yellow
        } catch {
            Write-Error "Could not save data in any format"
        }
    }
} else {
    Write-Host "No data to save" -ForegroundColor Yellow
}

# 8. CLEANUP
# -------------------------------------------------------------------
Write-Host "`nCleaning up..." -ForegroundColor Cyan
try {
    if ($Driver -ne $null) {
        $Driver.Quit()
        Write-Host "Browser closed" -ForegroundColor Gray
    }
} catch {
    Write-Warning "Error during cleanup"
}

Write-Host "`n" + ("=" * 60) -ForegroundColor DarkGray
Write-Host "SCRIPT COMPLETED" -ForegroundColor Green
Write-Host ("=" * 60) -ForegroundColor DarkGray