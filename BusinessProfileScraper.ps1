<#
    ENHANCED BUSINESS PROFILE SCRAPER
    Clicks first 5 search results and extracts business/profile data
#>

Clear-Host
Write-Host "Starting Enhanced Business Profile Scraper..." -ForegroundColor Green
Write-Host ""

$ScriptFolder = "C:\Users\user\power\"
$ChromeDriverExe = "$ScriptFolder\chromedriver.exe"

# Check if ChromeDriver exists
if (-not (Test-Path $ChromeDriverExe)) {
    Write-Host "ERROR: ChromeDriver not found!" -ForegroundColor Red
    Write-Host "Place chromedriver.exe in: $ScriptFolder" -ForegroundColor Yellow
    Write-Host "Download from: https://googlechromelabs.github.io/chrome-for-testing/" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "Press any key to exit..." -ForegroundColor Gray
    $null = [System.Console]::ReadKey($true)
    exit
}

$env:PATH = "$ScriptFolder;$env:PATH"

# Import required modules
try {
    Import-Module ImportExcel
} catch {
    Write-Host "Installing ImportExcel module..." -ForegroundColor Yellow
    Install-Module -Name ImportExcel -Force -Scope CurrentUser
    Import-Module ImportExcel
}

try {
    Add-Type -Path "$ScriptFolder\WebDriver.dll"
} catch {
    Write-Host "Installing Selenium module..." -ForegroundColor Yellow
    Install-Module -Name Selenium -Force -Scope CurrentUser
    Import-Module Selenium
}

$searchTerm = Read-Host "Enter business/search term (or press Enter for 'restaurants near me')"
if ($searchTerm -eq "") {
    $searchTerm = "restaurants near me"
}

Write-Host "Starting Chrome and searching for: $searchTerm" -ForegroundColor Green

$options = New-Object OpenQA.Selenium.Chrome.ChromeOptions
$options.AddArgument("--disable-blink-features=AutomationControlled")
$options.AddArgument("--start-maximized")

try {
    $service = [OpenQA.Selenium.Chrome.ChromeDriverService]::CreateDefaultService($ScriptFolder)
    $service.HideCommandPromptWindow = $true
    $driver = New-Object OpenQA.Selenium.Chrome.ChromeDriver($service, $options)
    
    Write-Host "Chrome started!" -ForegroundColor Green
    
    # Navigate to Google
    Write-Host "Navigating to Google..." -ForegroundColor Green
    $driver.Navigate().GoToUrl("https://www.google.com")
    Start-Sleep -Seconds 2
    
    # Perform search
    Write-Host "Searching..." -ForegroundColor Green
    $searchBox = $driver.FindElement([OpenQA.Selenium.By]::Name("q"))
    $searchBox.SendKeys($searchTerm)
    $searchBox.SendKeys([OpenQA.Selenium.Keys]::Enter)
    Start-Sleep -Seconds 3
    
    # Get first 5 search result links
    Write-Host "Finding search results..." -ForegroundColor Green
    $results = @()
    
    # Get all result links
    $linkElements = $driver.FindElements([OpenQA.Selenium.By]::CssSelector("div.g a"))
    
    $counter = 0
    foreach ($linkElement in $linkElements) {
        if ($counter -ge 5) { break }
        
        try {
            $href = $linkElement.GetAttribute("href")
            
            # Skip non-http links and Google's own links
            if ($href -and $href.StartsWith("http") -and -not $href.Contains("google.com")) {
                Write-Host "`nProcessing result $($counter + 1): $href" -ForegroundColor Cyan
                
                # Store basic result
                $basicResult = [PSCustomObject]@{
                    ResultNumber = $counter + 1
                    Title = $linkElement.Text
                    URL = $href
                    BusinessName = ""
                    Description = ""
                    Email = ""
                    Phone = ""
                    Address = ""
                    SocialMedia = ""
                    Website = ""
                    ContactPage = ""
                    ProfileScore = 0
                }
                
                # Try to click and extract business profile
                $profileData = Extract-BusinessProfile -Driver $driver -Url $href
                
                if ($profileData) {
                    $basicResult.BusinessName = $profileData.BusinessName
                    $basicResult.Description = $profileData.Description
                    $basicResult.Email = $profileData.Email
                    $basicResult.Phone = $profileData.Phone
                    $basicResult.Address = $profileData.Address
                    $basicResult.SocialMedia = $profileData.SocialMedia
                    $basicResult.Website = $profileData.Website
                    $basicResult.ContactPage = $profileData.ContactPage
                    $basicResult.ProfileScore = $profileData.ProfileScore
                }
                
                $results += $basicResult
                $counter++
            }
        } catch {
            Write-Host "Skipping a link due to error: $($_.Exception.Message)" -ForegroundColor Yellow
            continue
        }
    }
    
    # Save to Excel
    $timestamp = Get-Date -Format "yyyyMMdd_HHmm"
    $excelFile = "C:\Users\user\Desktop\BusinessProfiles_$timestamp.xlsx"
    
    $results | Export-Excel -Path $excelFile -WorksheetName "Profiles" -AutoSize -BoldTopRow
    
    Write-Host ""
    Write-Host "=========================================" -ForegroundColor Cyan
    Write-Host "           SCRAPING COMPLETE!" -ForegroundColor Green
    Write-Host "=========================================" -ForegroundColor Cyan
    Write-Host "Profiles saved to: $excelFile" -ForegroundColor Yellow
    Write-Host "Total profiles extracted: $($results.Count)" -ForegroundColor Yellow
    Write-Host ""
    
    # Show summary
    foreach ($result in $results) {
        Write-Host "Result $($result.ResultNumber): $($result.Title)" -ForegroundColor Gray
        if ($result.ProfileScore -gt 0) {
            Write-Host "  Profile Score: $($result.ProfileScore)/10" -ForegroundColor Green
            if ($result.Email) { Write-Host "  Email: $($result.Email)" -ForegroundColor Gray }
            if ($result.Phone) { Write-Host "  Phone: $($result.Phone)" -ForegroundColor Gray }
        }
        Write-Host ""
    }
    
    $openFile = Read-Host "Open Excel file? (Y/N)"
    if ($openFile -eq "Y" -or $openFile -eq "y") {
        Start-Process $excelFile
    }
    
} catch {
    Write-Host "ERROR: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Full error details:" -ForegroundColor Yellow
    Write-Host $_.ScriptStackTrace -ForegroundColor Yellow
} finally {
    if ($driver -ne $null) {
        Write-Host "`nClosing browser..." -ForegroundColor Gray
        $driver.Quit()
    }
}

Write-Host "`nScript finished." -ForegroundColor Gray

# Function to extract business profile from a URL
function Extract-BusinessProfile {
    param(
        [object]$Driver,
        [string]$Url
    )
    
    $profile = [PSCustomObject]@{
        BusinessName = ""
        Description = ""
        Email = ""
        Phone = ""
        Address = ""
        SocialMedia = ""
        Website = ""
        ContactPage = ""
        ProfileScore = 0
    }
    
    $score = 0
    
    try {
        # Open URL in new tab
        $Driver.ExecuteScript("window.open(arguments[0], '_blank');", $Url)
        Start-Sleep -Seconds 2
        
        # Switch to new tab
        $Driver.SwitchTo().Window($Driver.WindowHandles[1])
        Start-Sleep -Seconds 2
        
        # Extract page title as business name
        $pageTitle = $Driver.Title
        if ($pageTitle -and $pageTitle.Length -gt 5) {
            $profile.BusinessName = $pageTitle
            $score += 1
        }
        
        # Extract meta description
        try {
            $metaDesc = $Driver.FindElement([OpenQA.Selenium.By]::CssSelector("meta[name='description']"))
            $description = $metaDesc.GetAttribute("content")
            if ($description) {
                $profile.Description = $description
                $score += 1
            }
        } catch {
            # No meta description found
        }
        
        # Extract all text from page
        $pageText = $Driver.FindElement([OpenQA.Selenium.By]::TagName("body")).Text
        
        # Look for email patterns
        $emailPattern = "\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b"
        $emailMatches = [regex]::Matches($pageText, $emailPattern)
        if ($emailMatches.Count -gt 0) {
            $profile.Email = $emailMatches[0].Value
            $score += 2
        }
        
        # Look for phone patterns
        $phonePattern = "\(?\d{3}\)?[-.\s]?\d{3}[-.\s]?\d{4}"
        $phoneMatches = [regex]::Matches($pageText, $phonePattern)
        if ($phoneMatches.Count -gt 0) {
            $profile.Phone = $phoneMatches[0].Value
            $score += 2
        }
        
        # Look for address keywords
        $addressKeywords = @("street", "avenue", "ave", "road", "rd", "boulevard", "blvd", "drive", "dr", "lane", "ln")
        foreach ($keyword in $addressKeywords) {
            if ($pageText.ToLower().Contains($keyword)) {
                $score += 1
                break
            }
        }
        
        # Look for social media links
        $socialPatterns = @("facebook.com", "twitter.com", "linkedin.com", "instagram.com", "youtube.com")
        $socialLinks = @()
        
        $allLinks = $Driver.FindElements([OpenQA.Selenium.By]::TagName("a"))
        foreach ($link in $allLinks) {
            $href = $link.GetAttribute("href")
            if ($href) {
                foreach ($pattern in $socialPatterns) {
                    if ($href.ToLower().Contains($pattern)) {
                        $socialLinks += $href
                        break
                    }
                }
            }
        }
        
        if ($socialLinks.Count -gt 0) {
            $profile.SocialMedia = $socialLinks -join ", "
            $score += 2
        }
        
        # Find contact page
        $contactKeywords = @("contact", "about", "get in touch", "reach us")
        foreach ($link in $allLinks) {
            $linkText = $link.Text.ToLower()
            foreach ($keyword in $contactKeywords) {
                if ($linkText.Contains($keyword)) {
                    $profile.ContactPage = $link.GetAttribute("href")
                    $score += 1
                    break
                }
            }
            if ($profile.ContactPage) { break }
        }
        
        # Close the tab and switch back
        $Driver.Close()
        $Driver.SwitchTo().Window($Driver.WindowHandles[0])
        
    } catch {
        Write-Host "Error extracting profile from $Url : $($_.Exception.Message)" -ForegroundColor Yellow
        
        # Make sure we're back on the main tab
        try {
            if ($Driver.WindowHandles.Count -gt 1) {
                $Driver.Close()
            }
            $Driver.SwitchTo().Window($Driver.WindowHandles[0])
        } catch {
            # Ignore cleanup errors
        }
    }
    
    $profile.ProfileScore = $score
    return $profile
}