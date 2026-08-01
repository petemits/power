<#
    PLUG & PLAY CLICKABLE SCRAPER
    Clicks first 5 search results and extracts data
    Multiple fallback selectors for reliability
#>

Clear-Host
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host "   PLUG & PLAY CLICKABLE SCRAPER" -ForegroundColor Cyan
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host ""

# Check requirements
$ScriptFolder = "C:\Users\user\power\"
$ChromeDriverExe = "$ScriptFolder\chromedriver.exe"
$WebDriverDll = "$ScriptFolder\WebDriver.dll"

if (-not (Test-Path $ChromeDriverExe)) {
    Write-Host "ERROR: ChromeDriver not found in folder!" -ForegroundColor Red
    Write-Host "Place chromedriver.exe (version 143) in: $ScriptFolder" -ForegroundColor Yellow
    Write-Host ""
    $null = [System.Console]::ReadKey($true)
    exit
}

if (-not (Test-Path $WebDriverDll)) {
    Write-Host "ERROR: WebDriver.dll not found!" -ForegroundColor Red
    Write-Host "Place WebDriver.dll in: $ScriptFolder" -ForegroundColor Yellow
    Write-Host ""
    $null = [System.Console]::ReadKey($true)
    exit
}

# Set path to force using our ChromeDriver
$env:PATH = "$ScriptFolder;$env:PATH"

# Import modules
try {
    Add-Type -Path $WebDriverDll
    Write-Host "✓ WebDriver.dll loaded" -ForegroundColor Green
} catch {
    Write-Host "ERROR loading WebDriver.dll: $($_.Exception.Message)" -ForegroundColor Red
    exit
}

try {
    Import-Module ImportExcel -ErrorAction Stop
    Write-Host "✓ ImportExcel module loaded" -ForegroundColor Green
} catch {
    Write-Host "Installing ImportExcel module..." -ForegroundColor Yellow
    Install-Module -Name ImportExcel -Force -Scope CurrentUser
    Import-Module ImportExcel
}

# Get search term
$searchTerm = Read-Host "`nEnter search term (or press Enter for 'local businesses')"
if ([string]::IsNullOrWhiteSpace($searchTerm)) {
    $searchTerm = "local businesses"
    Write-Host "Using default: 'local businesses'" -ForegroundColor Yellow
}

Write-Host "`nStarting Chrome..." -ForegroundColor Green

# Chrome setup
$options = New-Object OpenQA.Selenium.Chrome.ChromeOptions
$options.AddArgument("--disable-blink-features=AutomationControlled")
$options.AddArgument("--start-maximized")

try {
    # Start Chrome with our driver
    $service = [OpenQA.Selenium.Chrome.ChromeDriverService]::CreateDefaultService($ScriptFolder)
    $service.HideCommandPromptWindow = $true
    $driver = New-Object OpenQA.Selenium.Chrome.ChromeDriver($service, $options)
    
    Write-Host "✓ Chrome started successfully" -ForegroundColor Green
    Write-Host "  ChromeDriver version: $($service.ChromeDriverVersion)" -ForegroundColor Gray
    
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
    
    # Get search result links with MULTIPLE SELECTOR STRATEGIES
    Write-Host "`nFinding search result links..." -ForegroundColor Green
    $links = @()
    
    # Strategy 1: Try multiple CSS selectors (different Google layouts)
    $selectors = @(
        "div.g a",                     # Classic Google
        "a h3",                        # Link headings
        "div.tF2Cxc a",                # Modern Google
        "div.yuRUbf a",                # Another modern layout
        "a[jsname='UWckNb']",          # JavaScript-based links
        "h3.LC20lb"                    # Heading links
    )
    
    foreach ($selector in $selectors) {
        try {
            $elements = $driver.FindElements([OpenQA.Selenium.By]::CssSelector($selector))
            if ($elements.Count -gt 0) {
                Write-Host "  Found $($elements.Count) elements with selector: $selector" -ForegroundColor Gray
                
                foreach ($element in $elements) {
                    try {
                        $href = $element.GetAttribute("href")
                        $text = $element.Text
                        
                        # Filter valid links
                        if ($href -and $href.StartsWith("http") -and -not $href.Contains("google.com") -and $text.Trim().Length -gt 0) {
                            # Avoid duplicates
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
                
                if ($links.Count -ge 5) {
                    Write-Host "  ✓ Found enough links ($($links.Count) total)" -ForegroundColor Green
                    break
                }
            }
        } catch {
            continue
        }
    }
    
    # If no links found with CSS, try XPath
    if ($links.Count -eq 0) {
        Write-Host "  Trying XPath strategy..." -ForegroundColor Yellow
        try {
            $xpathElements = $driver.FindElements([OpenQA.Selenium.By]::XPath("//a[contains(@href, 'http') and not(contains(@href, 'google.com')) and .//h3]"))
            foreach ($element in $xpathElements) {
                $href = $element.GetAttribute("href")
                $text = $element.Text
                if ($href -and $text.Trim().Length -gt 0) {
                    $links += [PSCustomObject]@{
                        URL = $href
                        Title = $text
                    }
                }
                if ($links.Count -ge 10) { break }
            }
        } catch {
            Write-Host "  XPath also failed" -ForegroundColor Red
        }
    }
    
    if ($links.Count -eq 0) {
        Write-Host "ERROR: No search result links found!" -ForegroundColor Red
        Write-Host "The Google page structure may have changed." -ForegroundColor Yellow
        Write-Host "Taking screenshot for debugging..." -ForegroundColor Yellow
        
        $screenshotPath = "$ScriptFolder\debug_screenshot.png"
        $driver.GetScreenshot().SaveAsFile($screenshotPath)
        Write-Host "Screenshot saved to: $screenshotPath" -ForegroundColor Cyan
        
        $driver.Quit()
        Write-Host "`nPress any key to exit..." -ForegroundColor Gray
        $null = [System.Console]::ReadKey($true)
        exit
    }
    
    # Process first 5 links
    $results = @()
    $maxLinks = [Math]::Min(5, $links.Count)
    
    Write-Host "`nProcessing first $maxLinks links..." -ForegroundColor Cyan
    
    for ($i = 0; $i -lt $maxLinks; $i++) {
        $link = $links[$i]
        Write-Host "`n[$($i+1)/$maxLinks] Clicking: $($link.Title)" -ForegroundColor Cyan
        Write-Host "  URL: $($link.URL)" -ForegroundColor Gray
        
        $profileData = Extract-PageData -Driver $driver -Url $link.URL
        
        $results += [PSCustomObject]@{
            Number = $i + 1
            SearchTitle = $link.Title
            URL = $link.URL
            PageTitle = $profileData.PageTitle
            Description = $profileData.Description
            Email = $profileData.Email
            Phone = $profileData.Phone
            AddressClue = $profileData.AddressClue
            SocialLinks = $profileData.SocialLinks
            WordCount = $profileData.WordCount
            ScrapeTime = $profileData.ScrapeTime
        }
        
        Write-Host "  Extracted: $($profileData.Email -replace '^(.).+@.+$', '$1***') | $($profileData.Phone)" -ForegroundColor Green
    }
    
    # Save to Excel
    Write-Host "`nSaving results to Excel..." -ForegroundColor Green
    
    $timestamp = Get-Date -Format "yyyyMMdd_HHmm"
    $excelFile = "$env:USERPROFILE\Desktop\SearchResults_$timestamp.xlsx"
    
    try {
        $results | Export-Excel -Path $excelFile -WorksheetName "Results" -AutoSize -BoldTopRow -FreezeTopRow
        Write-Host "✓ Results saved to: $excelFile" -ForegroundColor Green
        Write-Host "  Total records: $($results.Count)" -ForegroundColor Gray
        
        # Show preview
        Write-Host "`nPreview of first result:" -ForegroundColor Cyan
        if ($results.Count -gt 0) {
            $first = $results[0]
            Write-Host "  Title: $($first.PageTitle)" -ForegroundColor Gray
            Write-Host "  Email: $($first.Email)" -ForegroundColor Gray
            Write-Host "  Phone: $($first.Phone)" -ForegroundColor Gray
        }
        
        # Ask to open
        Write-Host ""
        $openFile = Read-Host "Open Excel file now? (Y/N)"
        if ($openFile -eq "Y" -or $openFile -eq "y") {
            Start-Process $excelFile
            Write-Host "Opening Excel..." -ForegroundColor Green
        }
        
    } catch {
        Write-Host "ERROR saving Excel: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "Raw data will be shown below:" -ForegroundColor Yellow
        $results | Format-Table -AutoSize
    }
    
} catch {
    Write-Host "`nMAIN ERROR: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Check that:" -ForegroundColor Yellow
    Write-Host "1. ChromeDriver is version 143 (matching your Chrome)" -ForegroundColor Yellow
    Write-Host "2. You have internet connection" -ForegroundColor Yellow
    Write-Host "3. Google.com is accessible" -ForegroundColor Yellow
} finally {
    if ($driver -ne $null) {
        Write-Host "`nClosing browser..." -ForegroundColor Gray
        $driver.Quit()
        Write-Host "Browser closed." -ForegroundColor Gray
    }
}

Write-Host "`nScript finished." -ForegroundColor Gray
Write-Host "==========================================" -ForegroundColor Cyan

# Function to extract data from a page
function Extract-PageData {
    param([object]$Driver, [string]$Url)
    
    $result = [PSCustomObject]@{
        PageTitle = ""
        Description = ""
        Email = ""
        Phone = ""
        AddressClue = ""
        SocialLinks = ""
        WordCount = 0
        ScrapeTime = 0
    }
    
    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    
    try {
        # Open in new tab
        $originalWindow = $Driver.CurrentWindowHandle
        $Driver.ExecuteScript("window.open(arguments[0], '_blank');", $Url)
        Start-Sleep -Seconds 2
        
        # Switch to new tab
        $Driver.SwitchTo().Window($Driver.WindowHandles[-1])
        Start-Sleep -Seconds 2
        
        # Get page title
        $result.PageTitle = $Driver.Title
        
        # Get all page text
        $pageText = $Driver.FindElement([OpenQA.Selenium.By]::TagName("body")).Text
        $result.WordCount = ($pageText -split '\s+').Count
        
        # Look for email (regex pattern)
        $emailPattern = "\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b"
        $emailMatches = [regex]::Matches($pageText, $emailPattern)
        if ($emailMatches.Count -gt 0) {
            $result.Email = $emailMatches[0].Value
        }
        
        # Look for phone numbers (US pattern)
        $phonePattern = "\(?\d{3}\)?[-.\s]?\d{3}[-.\s]?\d{4}"
        $phoneMatches = [regex]::Matches($pageText, $phonePattern)
        if ($phoneMatches.Count -gt 0) {
            $result.Phone = $phoneMatches[0].Value
        }
        
        # Check for address keywords
        $addressKeywords = @("street", "avenue", "road", "boulevard", "drive", "lane", "suite", "floor", "city", "state", "zip")
        foreach ($keyword in $addressKeywords) {
            if ($pageText.ToLower().Contains($keyword)) {
                $result.AddressClue = "Contains address information"
                break
            }
        }
        
        # Look for social media links
        $socialPatterns = @("facebook.com", "twitter.com", "linkedin.com", "instagram.com")
        $socialLinks = @()
        
        $allLinks = $Driver.FindElements([OpenQA.Selenium.By]::TagName("a"))
        foreach ($link in $allLinks) {
            $href = $link.GetAttribute("href")
            if ($href) {
                foreach ($pattern in $socialPatterns) {
                    if ($href.ToLower().Contains($pattern)) {
                        $socialLinks += $pattern
                        break
                    }
                }
            }
        }
        
        if ($socialLinks.Count -gt 0) {
            $result.SocialLinks = $socialLinks -join ", "
        }
        
        # Close tab and return to original
        $Driver.Close()
        $Driver.SwitchTo().Window($originalWindow)
        
    } catch {
        Write-Host "  Warning: Error scraping $Url - $($_.Exception.Message)" -ForegroundColor Yellow
        
        # Try to return to original window
        try {
            if ($Driver.WindowHandles.Count -gt 1) {
                $Driver.Close()
            }
            $Driver.SwitchTo().Window($Driver.WindowHandles[0])
        } catch {
            # Ignore cleanup errors
        }
    }
    
    $stopwatch.Stop()
    $result.ScrapeTime = [Math]::Round($stopwatch.Elapsed.TotalSeconds, 1)
    
    return $result
}