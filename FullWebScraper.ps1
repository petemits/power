<#
    COMPLETE CLICKABLE WEB SCRAPER
    Clicks search results and extracts business data
#>

# Clear screen
Clear-Host
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host "   CLICKABLE WEB SCRAPER v1.0" -ForegroundColor Cyan
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host ""

# Setup paths
$ScriptFolder = "C:\Users\user\power\"
$ChromeDriverExe = Join-Path $ScriptFolder "chromedriver.exe"
$WebDriverDll = Join-Path $ScriptFolder "WebDriver.dll"

# Check required files
if (-not (Test-Path $ChromeDriverExe)) {
    Write-Host "ERROR: ChromeDriver not found!" -ForegroundColor Red
    Write-Host "Place chromedriver.exe in: $ScriptFolder" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "Press any key to exit..." -ForegroundColor Gray
    $null = [System.Console]::ReadKey($true)
    exit
}

if (-not (Test-Path $WebDriverDll)) {
    Write-Host "ERROR: WebDriver.dll not found!" -ForegroundColor Red
    Write-Host "Place WebDriver.dll in: $ScriptFolder" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "Press any key to exit..." -ForegroundColor Gray
    $null = [System.Console]::ReadKey($true)
    exit
}

# Force using our ChromeDriver
$env:PATH = "$ScriptFolder;$env:PATH"

# Load WebDriver
try {
    Add-Type -Path $WebDriverDll
    Write-Host "✓ WebDriver.dll loaded" -ForegroundColor Green
} catch {
    Write-Host "ERROR loading WebDriver.dll: $($_.Exception.Message)" -ForegroundColor Red
    exit
}

# Load ImportExcel module
try {
    Import-Module ImportExcel -ErrorAction Stop
    Write-Host "✓ ImportExcel module loaded" -ForegroundColor Green
} catch {
    Write-Host "Installing ImportExcel module..." -ForegroundColor Yellow
    Install-Module -Name ImportExcel -Force -Scope CurrentUser
    Import-Module ImportExcel
    Write-Host "✓ ImportExcel installed" -ForegroundColor Green
}

# Get search term
$searchTerm = Read-Host "`nEnter search term (or press Enter for 'business services')"
if ([string]::IsNullOrWhiteSpace($searchTerm)) {
    $searchTerm = "business services"
    Write-Host "Using: 'business services'" -ForegroundColor Yellow
}

Write-Host "`nStarting Chrome..." -ForegroundColor Green

# Configure Chrome
$options = New-Object OpenQA.Selenium.Chrome.ChromeOptions
$options.AddArgument("--disable-blink-features=AutomationControlled")
$options.AddArgument("--start-maximized")

try {
    # Start Chrome
    $service = [OpenQA.Selenium.Chrome.ChromeDriverService]::CreateDefaultService($ScriptFolder)
    $service.HideCommandPromptWindow = $true
    $driver = New-Object OpenQA.Selenium.Chrome.ChromeDriver($service, $options)
    
    Write-Host "✓ Chrome started" -ForegroundColor Green
    
    # Navigate to Google
    Write-Host "`nNavigating to Google..." -ForegroundColor Green
    $driver.Navigate().GoToUrl("https://www.google.com")
    Start-Sleep -Seconds 2
    
    # Perform search
    Write-Host "Searching for: '$searchTerm'" -ForegroundColor Green
    $searchBox = $driver.FindElement([OpenQA.Selenium.By]::Name("q"))
    $searchBox.SendKeys($searchTerm)
    $searchBox.SendKeys([OpenQA.Selenium.Keys]::Enter)
    Start-Sleep -Seconds 3
    
    # Get search result links
    Write-Host "Finding search results..." -ForegroundColor Green
    $links = @()
    
    # Try multiple selectors for Google results
    $selectors = @(
        "div.g a",
        "a h3",
        "div.tF2Cxc a",
        "div.yuRUbf a"
    )
    
    foreach ($selector in $selectors) {
        $elements = $driver.FindElements([OpenQA.Selenium.By]::CssSelector($selector))
        if ($elements.Count -gt 0) {
            Write-Host "Found $($elements.Count) elements with: $selector" -ForegroundColor Gray
            
            foreach ($element in $elements) {
                try {
                    $href = $element.GetAttribute("href")
                    $text = $element.Text
                    
                    if ($href -and $href.StartsWith("http") -and -not $href.Contains("google.com") -and $text.Trim().Length -gt 0) {
                        if ($links.URL -notcontains $href) {
                            $links += [PSCustomObject]@{
                                URL = $href
                                Title = $text
                            }
                        }
                    }
                } catch {
                    continue
                }
            }
        }
        
        if ($links.Count -ge 10) {
            break
        }
    }
    
    if ($links.Count -eq 0) {
        Write-Host "ERROR: No links found!" -ForegroundColor Red
        Write-Host "Taking screenshot for debugging..." -ForegroundColor Yellow
        $screenshotPath = Join-Path $ScriptFolder "debug.png"
        $driver.GetScreenshot().SaveAsFile($screenshotPath)
        Write-Host "Screenshot saved to: $screenshotPath" -ForegroundColor Cyan
        $driver.Quit()
        Write-Host "`nPress any key to exit..." -ForegroundColor Gray
        $null = [System.Console]::ReadKey($true)
        exit
    }
    
    Write-Host "✓ Found $($links.Count) links" -ForegroundColor Green
    
    # Process first 5 links
    $results = @()
    $maxToProcess = [Math]::Min(5, $links.Count)
    
    Write-Host "`nProcessing first $maxToProcess links..." -ForegroundColor Cyan
    
    for ($i = 0; $i -lt $maxToProcess; $i++) {
        $link = $links[$i]
        Write-Host "`n[$($i+1)/$maxToProcess] Processing: $($link.Title)" -ForegroundColor Cyan
        
        $profileData = Extract-PageData -Driver $driver -Url $link.URL
        
        $results += [PSCustomObject]@{
            Number = $i + 1
            SearchTitle = $link.Title
            URL = $link.URL
            PageTitle = $profileData.PageTitle
            Email = $profileData.Email
            Phone = $profileData.Phone
            SocialMedia = $profileData.SocialMedia
            HasAddress = $profileData.HasAddress
            WordCount = $profileData.WordCount
            ScrapeTime = $profileData.ScrapeTime
        }
        
        # Show what was found
        $foundItems = @()
        if ($profileData.Email) { $foundItems += "email" }
        if ($profileData.Phone) { $foundItems += "phone" }
        if ($profileData.SocialMedia) { $foundItems += "social" }
        
        if ($foundItems.Count -gt 0) {
            Write-Host "  Found: $($foundItems -join ', ')" -ForegroundColor Green
        } else {
            Write-Host "  No contact info found" -ForegroundColor Yellow
        }
    }
    
    # Save to Excel
    Write-Host "`nSaving to Excel..." -ForegroundColor Green
    
    $timestamp = Get-Date -Format "yyyyMMdd_HHmm"
    $excelFile = "$env:USERPROFILE\Desktop\BusinessData_$timestamp.xlsx"
    
    try {
        $results | Export-Excel -Path $excelFile -WorksheetName "Results" -AutoSize -BoldTopRow -FreezeTopRow
        Write-Host "✓ Saved to: $excelFile" -ForegroundColor Green
        
        # Show summary
        Write-Host "`n==========================================" -ForegroundColor Cyan
        Write-Host "           SUMMARY" -ForegroundColor Cyan
        Write-Host "==========================================" -ForegroundColor Cyan
        
        $totalEmails = ($results | Where-Object { $_.Email -ne "" }).Count
        $totalPhones = ($results | Where-Object { $_.Phone -ne "" }).Count
        $totalSocial = ($results | Where-Object { $_.SocialMedia -ne "" }).Count
        
        Write-Host "Pages processed: $($results.Count)" -ForegroundColor Gray
        Write-Host "Emails found: $totalEmails" -ForegroundColor Gray
        Write-Host "Phones found: $totalPhones" -ForegroundColor Gray
        Write-Host "Social media found: $totalSocial" -ForegroundColor Gray
        Write-Host ""
        
        # Ask to open file
        $openFile = Read-Host "Open Excel file? (Y/N)"
        if ($openFile -eq "Y" -or $openFile -eq "y") {
            Start-Process $excelFile
            Write-Host "Opening Excel..." -ForegroundColor Green
        }
        
    } catch {
        Write-Host "ERROR saving Excel: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "Displaying data below:" -ForegroundColor Yellow
        $results | Format-Table -AutoSize
    }
    
} catch {
    Write-Host "`nMAIN ERROR: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Script failed. Check ChromeDriver version (should be 143.x)" -ForegroundColor Yellow
} finally {
    if ($driver -ne $null) {
        Write-Host "`nClosing browser..." -ForegroundColor Gray
        $driver.Quit()
    }
}

Write-Host "`nScript finished." -ForegroundColor Gray

# Function to extract data from a page
function Extract-PageData {
    param(
        [object]$Driver,
        [string]$Url
    )
    
    $result = [PSCustomObject]@{
        PageTitle = ""
        Email = ""
        Phone = ""
        SocialMedia = ""
        HasAddress = "No"
        WordCount = 0
        ScrapeTime = 0
    }
    
    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    
    try {
        # Save current window
        $mainWindow = $Driver.CurrentWindowHandle
        
        # Open URL in new tab
        $Driver.ExecuteScript("window.open(arguments[0]);", $Url)
        Start-Sleep -Seconds 2
        
        # Switch to new tab
        $allWindows = $Driver.WindowHandles
        $newWindow = $allWindows[-1]
        $Driver.SwitchTo().Window($newWindow)
        Start-Sleep -Seconds 2
        
        # Get page title
        $result.PageTitle = $Driver.Title
        
        # Get all text from page
        $bodyText = $Driver.FindElement([OpenQA.Selenium.By]::TagName("body")).Text
        $result.WordCount = ($bodyText -split '\s+').Count
        
        # Look for email addresses
        $emailPattern = "[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}"
        $emailMatches = [regex]::Matches($bodyText, $emailPattern)
        if ($emailMatches.Count -gt 0) {
            $result.Email = $emailMatches[0].Value
        }
        
        # Look for phone numbers
        $phonePattern = "\(?\d{3}\)?[-.\s]?\d{3}[-.\s]?\d{4}"
        $phoneMatches = [regex]::Matches($bodyText, $phonePattern)
        if ($phoneMatches.Count -gt 0) {
            $result.Phone = $phoneMatches[0].Value
        }
        
        # Check for address keywords
        $addressWords = @("street", "avenue", "road", "boulevard", "drive", "lane", "suite", "floor")
        foreach ($word in $addressWords) {
            if ($bodyText.ToLower().Contains($word)) {
                $result.HasAddress = "Yes"
                break
            }
        }
        
        # Look for social media links
        $socialSites = @("facebook.com", "twitter.com", "linkedin.com", "instagram.com")
        $socialLinks = @()
        
        $allLinks = $Driver.FindElements([OpenQA.Selenium.By]::TagName("a"))
        foreach ($link in $allLinks) {
            $href = $link.GetAttribute("href")
            if ($href) {
                foreach ($site in $socialSites) {
                    if ($href.ToLower().Contains($site)) {
                        $socialLinks += $site
                        break
                    }
                }
            }
        }
        
        if ($socialLinks.Count -gt 0) {
            $result.SocialMedia = $socialLinks -join ", "
        }
        
        # Close the tab and switch back
        $Driver.Close()
        $Driver.SwitchTo().Window($mainWindow)
        
    } catch {
        Write-Host "  Warning: Error scraping page - $($_.Exception.Message)" -ForegroundColor Yellow
        
        # Try to return to main window
        try {
            if ($Driver.WindowHandles.Count -gt 1) {
                $Driver.Close()
                $Driver.SwitchTo().Window($Driver.WindowHandles[0])
            }
        } catch {
            # Ignore cleanup errors
        }
    }
    
    $stopwatch.Stop()
    $result.ScrapeTime = [Math]::Round($stopwatch.Elapsed.TotalSeconds, 1)
    
    return $result
}