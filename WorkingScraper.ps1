<#
    WORKING CLICKABLE WEB SCRAPER
    Searches Google, clicks results, extracts data to Excel
#>

Clear-Host
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host "   WORKING WEB SCRAPER v1.0" -ForegroundColor Cyan
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host ""

$ScriptFolder = "C:\Users\user\power\"
$ChromeDriverExe = Join-Path $ScriptFolder "chromedriver.exe"
$WebDriverDll = Join-Path $ScriptFolder "WebDriver.dll"

# Check files exist
if (-not (Test-Path $ChromeDriverExe)) {
    Write-Host "ERROR: ChromeDriver not found!" -ForegroundColor Red
    Write-Host "Place chromedriver.exe in: $ScriptFolder" -ForegroundColor Yellow
    pause
    exit
}

if (-not (Test-Path $WebDriverDll)) {
    Write-Host "ERROR: WebDriver.dll not found!" -ForegroundColor Red
    Write-Host "Place WebDriver.dll in: $ScriptFolder" -ForegroundColor Yellow
    pause
    exit
}

# Force using our ChromeDriver
$env:PATH = "$ScriptFolder;$env:PATH"

# Load WebDriver
try {
    Add-Type -Path $WebDriverDll
    Write-Host "✓ WebDriver.dll loaded" -ForegroundColor Green
} catch {
    Write-Host "ERROR loading WebDriver.dll" -ForegroundColor Red
    pause
    exit
}

# Load ImportExcel
try {
    Import-Module ImportExcel
    Write-Host "✓ ImportExcel loaded" -ForegroundColor Green
} catch {
    Write-Host "Installing ImportExcel..." -ForegroundColor Yellow
    Install-Module -Name ImportExcel -Force -Scope CurrentUser
    Import-Module ImportExcel
    Write-Host "✓ ImportExcel installed" -ForegroundColor Green
}

# Get search term
$searchTerm = Read-Host "`nEnter search term (or press Enter for 'business')"
if ($searchTerm -eq "") {
    $searchTerm = "business"
    Write-Host "Using: 'business'" -ForegroundColor Yellow
}

Write-Host "`nStarting Chrome..." -ForegroundColor Green

# Configure Chrome
$options = New-Object OpenQA.Selenium.Chrome.ChromeOptions
$options.AddArgument("--disable-blink-features=AutomationControlled")
$options.AddArgument("--start-maximized")

$driver = $null

try {
    # Start Chrome
    $service = [OpenQA.Selenium.Chrome.ChromeDriverService]::CreateDefaultService($ScriptFolder)
    $service.HideCommandPromptWindow = $true
    $driver = New-Object OpenQA.Selenium.Chrome.ChromeDriver($service, $options)
    
    Write-Host "✓ Chrome started" -ForegroundColor Green
    
    # Go to Google
    Write-Host "`nNavigating to Google..." -ForegroundColor Green
    $driver.Navigate().GoToUrl("https://www.google.com")
    Start-Sleep -Seconds 2
    
    # Search
    Write-Host "Searching: '$searchTerm'" -ForegroundColor Green
    $searchBox = $driver.FindElement([OpenQA.Selenium.By]::Name("q"))
    $searchBox.SendKeys($searchTerm)
    $searchBox.SendKeys([OpenQA.Selenium.Keys]::Enter)
    Start-Sleep -Seconds 3
    
    # Find links
    Write-Host "Finding links..." -ForegroundColor Green
    $links = @()
    
    # Try different selectors
    $selectors = @("div.g a", "a h3", "div.tF2Cxc a")
    
    foreach ($selector in $selectors) {
        $elements = $driver.FindElements([OpenQA.Selenium.By]::CssSelector($selector))
        if ($elements.Count -gt 0) {
            Write-Host "Found $($elements.Count) with: $selector" -ForegroundColor Gray
            
            foreach ($element in $elements) {
                try {
                    $href = $element.GetAttribute("href")
                    $text = $element.Text
                    
                    if ($href -and $href.StartsWith("http") -and -not $href.Contains("google.com") -and $text.Length -gt 0) {
                        # Check if not already added
                        $alreadyExists = $false
                        foreach ($existing in $links) {
                            if ($existing.URL -eq $href) {
                                $alreadyExists = $true
                                break
                            }
                        }
                        
                        if (-not $alreadyExists) {
                            $links += [PSCustomObject]@{
                                URL = $href
                                Title = $text
                            }
                        }
                    }
                } catch {
                    # Skip this element
                }
            }
        }
        
        if ($links.Count -ge 10) { break }
    }
    
    if ($links.Count -eq 0) {
        Write-Host "ERROR: No links found!" -ForegroundColor Red
        $driver.Quit()
        pause
        exit
    }
    
    Write-Host "✓ Found $($links.Count) links" -ForegroundColor Green
    
    # Process up to 5 links
    $results = @()
    $maxToProcess = [Math]::Min(5, $links.Count)
    
    Write-Host "`nProcessing $maxToProcess links..." -ForegroundColor Cyan
    
    for ($i = 0; $i -lt $maxToProcess; $i++) {
        $link = $links[$i]
        Write-Host "`n[$($i+1)/$maxToProcess] $($link.Title)" -ForegroundColor Cyan
        
        $data = Get-PageData -Driver $driver -Url $link.URL
        
        $results += [PSCustomObject]@{
            Number = $i + 1
            SearchTitle = $link.Title
            URL = $link.URL
            PageTitle = $data.PageTitle
            Email = $data.Email
            Phone = $data.Phone
            Social = $data.Social
            HasAddress = $data.HasAddress
            Words = $data.Words
            TimeSec = $data.TimeSec
        }
        
        # Show what was found
        $found = @()
        if ($data.Email) { $found += "email" }
        if ($data.Phone) { $found += "phone" }
        if ($data.Social) { $found += "social" }
        
        if ($found.Count -gt 0) {
            Write-Host "  Found: $($found -join ', ')" -ForegroundColor Green
        } else {
            Write-Host "  No contact info" -ForegroundColor Gray
        }
    }
    
    # Save to Excel
    Write-Host "`nSaving to Excel..." -ForegroundColor Green
    
    $timestamp = Get-Date -Format "yyyyMMdd_HHmm"
    $excelFile = "$env:USERPROFILE\Desktop\ScrapedData_$timestamp.xlsx"
    
    $results | Export-Excel -Path $excelFile -WorksheetName "Results" -AutoSize -BoldTopRow
    
    Write-Host "✓ Saved to: $excelFile" -ForegroundColor Green
    
    # Show summary
    Write-Host "`n==========================================" -ForegroundColor Cyan
    Write-Host "           RESULTS SUMMARY" -ForegroundColor Cyan
    Write-Host "==========================================" -ForegroundColor Cyan
    
    $emailCount = 0
    $phoneCount = 0
    $socialCount = 0
    
    foreach ($result in $results) {
        if ($result.Email) { $emailCount++ }
        if ($result.Phone) { $phoneCount++ }
        if ($result.Social) { $socialCount++ }
    }
    
    Write-Host "Pages processed: $($results.Count)" -ForegroundColor Gray
    Write-Host "Emails found: $emailCount" -ForegroundColor Gray
    Write-Host "Phones found: $phoneCount" -ForegroundColor Gray
    Write-Host "Social links found: $socialCount" -ForegroundColor Gray
    Write-Host ""
    
    # Ask to open
    $openFile = Read-Host "Open Excel file? (Y/N)"
    if ($openFile -eq "Y" -or $openFile -eq "y") {
        Start-Process $excelFile
        Write-Host "Opening Excel..." -ForegroundColor Green
    }
    
} catch {
    Write-Host "ERROR: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Make sure ChromeDriver is version 143.x" -ForegroundColor Yellow
} finally {
    if ($driver -ne $null) {
        Write-Host "`nClosing browser..." -ForegroundColor Gray
        $driver.Quit()
    }
}

Write-Host "`nDone." -ForegroundColor Gray
pause

# Function to get data from a page
function Get-PageData {
    param(
        [object]$Driver,
        [string]$Url
    )
    
    $result = [PSCustomObject]@{
        PageTitle = ""
        Email = ""
        Phone = ""
        Social = ""
        HasAddress = "No"
        Words = 0
        TimeSec = 0
    }
    
    $startTime = Get-Date
    
    try {
        # Save main window
        $mainWindow = $Driver.CurrentWindowHandle
        
        # Open new tab
        $Driver.ExecuteScript("window.open('" + $Url + "');")
        Start-Sleep -Seconds 2
        
        # Switch to new tab
        $windows = $Driver.WindowHandles
        $Driver.SwitchTo().Window($windows[1])
        Start-Sleep -Seconds 2
        
        # Get page info
        $result.PageTitle = $Driver.Title
        
        $bodyText = $Driver.FindElement([OpenQA.Selenium.By]::TagName("body")).Text
        $result.Words = ($bodyText -split '\s+').Count
        
        # Find email
        if ($bodyText -match "[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}") {
            $result.Email = $Matches[0]
        }
        
        # Find phone
        if ($bodyText -match "\(?\d{3}\)?[-.\s]?\d{3}[-.\s]?\d{4}") {
            $result.Phone = $Matches[0]
        }
        
        # Check for address words
        $addressWords = @("street", "avenue", "road", "boulevard", "address", "location")
        foreach ($word in $addressWords) {
            if ($bodyText.ToLower().Contains($word)) {
                $result.HasAddress = "Yes"
                break
            }
        }
        
        # Check for social links
        $socialSites = @("facebook.com", "twitter.com", "linkedin.com", "instagram.com")
        $foundSocial = @()
        
        $allLinks = $Driver.FindElements([OpenQA.Selenium.By]::TagName("a"))
        foreach ($link in $allLinks) {
            $href = $link.GetAttribute("href")
            if ($href) {
                foreach ($site in $socialSites) {
                    if ($href.ToLower().Contains($site)) {
                        $foundSocial += $site
                        break
                    }
                }
            }
        }
        
        if ($foundSocial.Count -gt 0) {
            $result.Social = $foundSocial -join ", "
        }
        
        # Close tab and return
        $Driver.Close()
        $Driver.SwitchTo().Window($mainWindow)
        
    } catch {
        Write-Host "  Warning: Could not scrape page" -ForegroundColor Yellow
        
        # Try to clean up
        try {
            if ($Driver.WindowHandles.Count -gt 1) {
                $Driver.Close()
                $Driver.SwitchTo().Window($Driver.WindowHandles[0])
            }
        } catch {
            # Ignore
        }
    }
    
    $endTime = Get-Date
    $result.TimeSec = [Math]::Round(($endTime - $startTime).TotalSeconds, 1)
    
    return $result
}