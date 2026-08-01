<#
    WEB SCRAPER WITH PAUSE AFTER EACH LINK
    Stops after processing each link, waits for user to continue
#>

Clear-Host
Write-Host "SCRAPER WITH PAUSE CONTROL" -ForegroundColor Cyan
Write-Host "============================" -ForegroundColor Cyan
Write-Host ""

# Configuration
$folder = "C:\Users\user\power\"
$screenshotFolder = "$env:USERPROFILE\Desktop\WebsiteScreenshots\"
New-Item -ItemType Directory -Force -Path $screenshotFolder | Out-Null

# File checks
if (-not (Test-Path "$folder\chromedriver.exe")) {
    Write-Host "ERROR: ChromeDriver missing!" -ForegroundColor Red; pause; exit
}
if (-not (Test-Path "$folder\WebDriver.dll")) {
    Write-Host "ERROR: WebDriver.dll missing!" -ForegroundColor Red; pause; exit
}

# Load modules
Add-Type -Path "$folder\WebDriver.dll"
Import-Module ImportExcel -ErrorAction SilentlyContinue
if (-not (Get-Module -Name ImportExcel)) {
    Install-Module -Name ImportExcel -Force -Scope CurrentUser
    Import-Module ImportExcel
}

# Get search term
$searchTerm = Read-Host "Enter search term (or press Enter for 'technology')"
if ($searchTerm -eq "") { $searchTerm = "technology" }

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

# Find 5 result links
Write-Host "Finding links..." -ForegroundColor Green
$allLinks = $driver.FindElements([OpenQA.Selenium.By]::TagName("a"))
$links = @(); $count = 0

foreach ($link in $allLinks) {
    if ($count -ge 5) { break }
    $href = $link.GetAttribute("href"); $text = $link.Text
    if ($href -and $href.StartsWith("http") -and -not $href.Contains("google.com") -and $text.Length -gt 5) {
        $links += [PSCustomObject]@{ URL = $href; Title = $text }; $count++
    }
}

if ($links.Count -eq 0) {
    Write-Host "No links found!" -ForegroundColor Red; $driver.Quit(); pause; exit
}

# Process links WITH PAUSE AFTER EACH
$results = @()
$originalWindow = $driver.CurrentWindowHandle

for ($i = 0; $i -lt $links.Count; $i++) {
    $link = $links[$i]
    Write-Host "`n[$($i+1)/$($links.Count)] Opening: $($link.Title)" -ForegroundColor Cyan
    
    # Click link, open new tab
    $driver.ExecuteScript("window.open(arguments[0]);", $link.URL)
    Start-Sleep -Seconds 2
    
    # Switch to new tab
    $driver.SwitchTo().Window($driver.WindowHandles[-1])
    Start-Sleep -Seconds 3
    
    # Get page info & screenshot
    $pageTitle = $driver.Title
    $safeName = $pageTitle -replace '[^\w\-_\. ]', '_' -replace '\s+', '_'
    $screenshotPath = Join-Path $screenshotFolder "$($i+1)_$safeName.png"
    
    try {
        $driver.GetScreenshot().SaveAsFile($screenshotPath, [OpenQA.Selenium.ScreenshotImageFormat]::Png)
        Write-Host "  Screenshot saved." -ForegroundColor Green
    } catch { Write-Host "  Could not save screenshot." -ForegroundColor Yellow; $screenshotPath = "Failed" }
    
    # Extract data
    $pageText = $driver.FindElement([OpenQA.Selenium.By]::TagName("body")).Text
    $email = if ($pageText -match "[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}") { $Matches[0] } else { "" }
    $phone = if ($pageText -match "\(?\d{3}\)?[-.\s]?\d{3}[-.\s]?\d{4}") { $Matches[0] } else { "" }
    
    # Close tab, return to Google
    $driver.Close(); $driver.SwitchTo().Window($originalWindow); Start-Sleep -Seconds 1
    
    # Add to results
    $results += [PSCustomObject]@{
        Number = $i + 1; Title = $link.Title; URL = $link.URL
        PageTitle = $pageTitle; Screenshot = Split-Path $screenshotPath -Leaf
        Email = $email; Phone = $phone
    }
    
    # ==== PAUSE AFTER PROCESSING EACH LINK ====
    if ($i -lt ($links.Count - 1)) {
        Write-Host "`nProcessed link #$($i+1). Press any key to process next link..." -ForegroundColor Yellow
        $null = [System.Console]::ReadKey($true)
    }
}

# Save to Excel
$excelFile = "$env:USERPROFILE\Desktop\ScrapedData.xlsx"
$results | Export-Excel -Path $excelFile -WorksheetName "Results" -AutoSize -BoldTopRow

Write-Host "`n=================================" -ForegroundColor Green
Write-Host "COMPLETE! Data saved to:" -ForegroundColor Green
Write-Host "- Excel: $excelFile" -ForegroundColor Yellow
Write-Host "- Screenshots: $screenshotFolder" -ForegroundColor Yellow
Write-Host "=================================" -ForegroundColor Green

$driver.Quit()
Write-Host "`nPress any key to exit..." -ForegroundColor Gray
pause