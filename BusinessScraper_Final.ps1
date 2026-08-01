# BUSINESS PROFILE SCRAPER - CLEAN VERSION
Write-Host "Starting scraper..."

# Configuration
$ProjectFolder = "C:\Users\user\power"
$ChromePath = "C:\Program Files\Google\Chrome\Application\chrome.exe"
$SearchQuery = "software development company"
$OutputFile = "$ProjectFolder\Results_$(Get-Date -Format 'yyyyMMdd').xlsx"

# 1. Setup ChromeDriver
Write-Host "Setting up ChromeDriver..."
$DriverPath = "$ProjectFolder\chromedriver.exe"

# Check if ChromeDriver exists
if (-not (Test-Path $DriverPath)) {
    Write-Host "ChromeDriver not found. Downloading..."
    try {
        $url = "https://chromedriver.storage.googleapis.com/143.0.7499.169/chromedriver_win32.zip"
        $tempZip = "$env:TEMP\chromedriver.zip"
        
        # Download
        Invoke-WebRequest -Uri $url -OutFile $tempZip
        
        # Extract
        Add-Type -AssemblyName System.IO.Compression.FileSystem
        [System.IO.Compression.ZipFile]::ExtractToDirectory($tempZip, $ProjectFolder)
        
        Write-Host "ChromeDriver downloaded"
    } catch {
        Write-Host "Download failed. Please download manually from:"
        Write-Host "https://chromedriver.storage.googleapis.com/index.html?path=143.0.7499.169/"
        Write-Host "Save as: $DriverPath"
        exit
    }
}

# 2. Verify Chrome
Write-Host "Checking Chrome..."
if (-not (Test-Path $ChromePath)) {
    Write-Host "Chrome not found at: $ChromePath"
    exit
}

# 3. Start Browser
Write-Host "Starting browser..."
try {
    Import-Module Selenium
    
    # Create service using our ChromeDriver
    $service = [OpenQA.Selenium.Chrome.ChromeDriverService]::CreateDefaultService($ProjectFolder, "chromedriver.exe")
    
    # Create options
    $options = New-SeChromeOptions
    $options.AddArgument("--headless")
    $options.AddArgument("--no-sandbox")
    
    # Create driver
    $Driver = New-Object OpenQA.Selenium.Chrome.ChromeDriver($service, $options)
    
    Write-Host "Browser started"
} catch {
    Write-Host "Failed to start browser: $_"
    Write-Host "Make sure Selenium module is installed: Install-Module -Name Selenium -Force"
    exit
}

# 4. Search Google
Write-Host "Searching Google..."
try {
    $searchUrl = "https://www.google.com/search?q=$([System.Web.HttpUtility]::UrlEncode($SearchQuery))"
    $Driver.Navigate().GoToUrl($searchUrl)
    Start-Sleep -Seconds 3
    Write-Host "Search page loaded"
} catch {
    Write-Host "Search failed: $_"
    $Driver.Quit()
    exit
}

# 5. Extract Results
Write-Host "Extracting results..."
$Results = @()

try {
    # Get search results
    $elements = $Driver.FindElements([OpenQA.Selenium.By]::CssSelector("div.g"))
    
    if ($elements.Count -eq 0) {
        $elements = $Driver.FindElements([OpenQA.Selenium.By]::CssSelector("div[class*='MjjYud']"))
    }
    
    if ($elements.Count -gt 0) {
        Write-Host "Found $($elements.Count) results"
        
        $count = 0
        foreach ($element in $elements) {
            $count++
            if ($count -gt 10) { break }
            
            $title = "N/A"
            $url = "N/A"
            
            # Get title
            $titleEl = $element.FindElements([OpenQA.Selenium.By]::TagName("h3"))
            if ($titleEl.Count -gt 0) {
                $title = $titleEl[0].Text
            }
            
            # Get URL
            $linkEl = $element.FindElements([OpenQA.Selenium.By]::TagName("a"))
            if ($linkEl.Count -gt 0) {
                $url = $linkEl[0].GetAttribute("href")
            }
            
            # Add to results if we have data
            if ($title -ne "N/A" -or $url -ne "N/A") {
                $result = [PSCustomObject]@{
                    Number = $count
                    Title = $title
                    URL = $url
                    Date = Get-Date -Format "yyyy-MM-dd"
                }
                $Results += $result
            }
        }
    } else {
        Write-Host "No results found"
    }
} catch {
    Write-Host "Error extracting: $_"
}

Write-Host "Extracted $($Results.Count) profiles"

# 6. Save Results
if ($Results.Count -gt 0) {
    Write-Host "Saving results..."
    try {
        # Try Excel first
        if (Get-Module -ListAvailable -Name ImportExcel) {
            $Results | Export-Excel -Path $OutputFile -WorksheetName "Results" -AutoSize
            Write-Host "Saved to Excel: $OutputFile"
        } else {
            # Fallback to CSV
            $csvFile = $OutputFile -replace '.xlsx$', '.csv'
            $Results | Export-Csv -Path $csvFile -NoTypeInformation
            Write-Host "Saved to CSV: $csvFile"
        }
        
        # Show preview
        Write-Host ""
        Write-Host "First 3 results:"
        $Results | Select-Object -First 3 | Format-Table -AutoSize
        
    } catch {
        Write-Host "Save failed: $_"
    }
} else {
    Write-Host "No data to save"
}

# 7. Cleanup
Write-Host "Cleaning up..."
try {
    if ($Driver) {
        $Driver.Quit()
    }
} catch {
    Write-Host "Cleanup error"
}

Write-Host "Script completed"