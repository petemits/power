<#
    ENHANCED WEB SCRAPER
    Searches Google, clicks 5 results, and saves screenshots
#>

Clear-Host
Write-Host "CLICK & CAPTURE WEB SCRAPER" -ForegroundColor Cyan
Write-Host "=============================" -ForegroundColor Cyan
Write-Host ""

# Configuration
$folder = "C:\Users\user\power\"
$screenshotFolder = "$env:USERPROFILE\Desktop\WebsiteScreenshots\"
New-Item -ItemType Directory -Force -Path $screenshotFolder | Out-Null

# File checks (same as before)
if (-not (Test-Path "$folder\chromedriver.exe")) {
    Write-Host "ERROR: chromedriver.exe missing!" -ForegroundColor Red
    pause
    exit
}
if (-not (Test-Path "$folder\WebDriver.dll")) {
    Write-Host "ERROR: WebDriver.dll missing!" -ForegroundColor Red
    pause
    exit
}

# Load modules
Add-Type -Path "$folder\WebDriver.dll"
Import-Module ImportExcel -ErrorAction SilentlyContinue
if (-not (Get-Module -Name ImportExcel)) {
    Install-Module -Name ImportExcel -Force -Scope CurrentUser
    Import-Module ImportExcel
}

# Get search term
$searchTerm = Read-Host "Enter search term"
if ($searchTerm -eq "") { $searchTerm = "technology news" }

Write-Host "`nStarting Chrome..." -ForegroundColor Green

# Setup Chrome
$options = New-Object OpenQA.Selenium.Chrome.ChromeOptions
$options.AddArgument("--disable-blink-features=AutomationControlled")
$service = [OpenQA.Selenium.Chrome.ChromeDriverService]::CreateDefaultService($folder)
$service.HideCommandPromptWindow = $true
$driver = New-Object OpenQA.Selenium.Chrome.ChromeDriver($service, $options)

# Search Google
$driver.Navigate().GoToUrl("https://www.google.com")
Start-Sleep -Seconds 2
$searchBox = $driver.FindElement([OpenQA.Selenium.By]::Name("q"))
$searchBox.SendKeys($searchTerm)
$searchBox.SendKeys([OpenQA.Selenium.Keys]::Enter)
Start-Sleep -Seconds 3

# Find result links (get 5 links)
Write-Host "Finding result links..." -ForegroundColor Green
$allLinks = $driver.FindElements([OpenQA.Selenium.By]::TagName("a"))
$links = @()
$count = 0

foreach ($link in $allLinks) {
    if ($count -ge 5) { break }
    $href = $link.GetAttribute("href")
    $text = $link.Text
    if ($href -and $href.StartsWith("http") -and -not $href.Contains("google.com") -and $text.Length -gt 5) {
        $links += [PSCustomObject]@{ URL = $href; Title = $text }
        $count++
    }
}

if ($links.Count -eq 0) {
    Write-Host "No links found!" -ForegroundColor Red
    $driver.Quit()
    pause
    exit
}

# --- NEW LOGIC: Process Links & Take Screenshots ---
$results = @()
$originalWindow = $driver.CurrentWindowHandle # Store the Google tab[citation:1]

for ($i = 0; $i -lt $links.Count; $i++) {
    $link = $links[$i]
    Write-Host "`n[$($i+1)/$($links.Count)] Opening: $($link.Title)" -ForegroundColor Cyan

    # 1. Click link to open in new tab
    # (Script opens link in new tab. Your results page must have clickable links for this)
    $driver.ExecuteScript("window.open(arguments[0]);", $link.URL)
    Start-Sleep -Seconds 2

    # 2. Switch to the new tab[citation:1][citation:4]
    $allTabs = $driver.WindowHandles
    $newTab = $allTabs[-1] # Get the last opened tab (the new one)
    $driver.SwitchTo().Window($newTab)
    Start-Sleep -Seconds 3 # Wait for page to load

    # 3. Get page info and TAKE SCREENSHOT[citation:3]
    $pageTitle = $driver.Title
    $safeFileName = $pageTitle -replace '[^\w\-_\. ]', '_' -replace '\s+', '_'
    $screenshotPath = Join-Path $screenshotFolder "$($i+1)_$safeFileName.png"
    
    try {
        $screenshot = $driver.GetScreenshot()
        $screenshot.SaveAsFile($screenshotPath, [OpenQA.Selenium.ScreenshotImageFormat]::Png)
        Write-Host "  Screenshot saved: $($i+1)_$safeFileName.png" -ForegroundColor Green
    } catch {
        Write-Host "  Could not save screenshot for this page." -ForegroundColor Yellow
        $screenshotPath = "Failed"
    }

    # 4. (Optional) Extract text data like before
    $pageText = $driver.FindElement([OpenQA.Selenium.By]::TagName("body")).Text
    $email = if ($pageText -match "[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}") { $Matches[0] } else { "" }
    $phone = if ($pageText -match "\(?\d{3}\)?[-.\s]?\d{3}[-.\s]?\d{4}") { $Matches[0] } else { "" }

    # 5. Close the new tab and switch back to Google results[citation:1]
    $driver.Close()
    $driver.SwitchTo().Window($originalWindow)
    Start-Sleep -Seconds 1

    # Store results
    $results += [PSCustomObject]@{
        Number = $i + 1
        Title = $link.Title
        URL = $link.URL
        PageTitle = $pageTitle
        ScreenshotFile = Split-Path $screenshotPath -Leaf
        Email = $email
        Phone = $phone
    }
}

# Save summary to Excel (as before)
$excelFile = "$env:USERPROFILE\Desktop\ScrapedData_WithScreenshots.xlsx"
$results | Export-Excel -Path $excelFile -WorksheetName "Results" -AutoSize -BoldTopRow

Write-Host "`n================================" -ForegroundColor Green
Write-Host "PROCESS COMPLETE!" -ForegroundColor Green
Write-Host "================================" -ForegroundColor Green
Write-Host "Summary saved to: $excelFile" -ForegroundColor Yellow
Write-Host "Screenshots saved to: $screenshotFolder" -ForegroundColor Yellow
Write-Host "`nScreenshot files:" -ForegroundColor Cyan
Get-ChildItem $screenshotFolder | ForEach-Object { Write-Host "  - $($_.Name)" -ForegroundColor Gray }

# Cleanup
$driver.Quit()
Write-Host "`nBrowser closed." -ForegroundColor Gray

# Ask to open folder
Write-Host ""
$openFolder = Read-Host "Open screenshot folder? (Y/N)"
if ($openFolder -eq "Y" -or $openFolder -eq "y") {
    Start-Process $screenshotFolder
}
Write-Host "`nDone." -ForegroundColor Gray
pause