<#
    5-SITE WEB SCRAPER
    Opens 5 websites and extracts data to Excel
#>

Clear-Host
Write-Host "5-SITE WEB SCRAPER" -ForegroundColor Cyan
Write-Host "==================" -ForegroundColor Cyan
Write-Host ""

# Check files
$folder = "C:\Users\user\power\"
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

# Load WebDriver
Add-Type -Path "$folder\WebDriver.dll"

# Load ImportExcel
Import-Module ImportExcel -ErrorAction SilentlyContinue
if (-not (Get-Module -Name ImportExcel)) {
    Install-Module -Name ImportExcel -Force -Scope CurrentUser
    Import-Module ImportExcel
}

# Get search term
$searchTerm = Read-Host "Enter search term (e.g., restaurants)"
if ($searchTerm -eq "") { $searchTerm = "technology" }

Write-Host "`nStarting Chrome..." -ForegroundColor Green

# Start Chrome
$options = New-Object OpenQA.Selenium.Chrome.ChromeOptions
$options.AddArgument("--disable-blink-features=AutomationControlled")

$service = [OpenQA.Selenium.Chrome.ChromeDriverService]::CreateDefaultService($folder)
$service.HideCommandPromptWindow = $true
$driver = New-Object OpenQA.Selenium.Chrome.ChromeDriver($service, $options)

# Go to Google and search
$driver.Navigate().GoToUrl("https://www.google.com")
Start-Sleep -Seconds 2

$searchBox = $driver.FindElement([OpenQA.Selenium.By]::Name("q"))
$searchBox.SendKeys($searchTerm)
$searchBox.SendKeys([OpenQA.Selenium.Keys]::Enter)
Start-Sleep -Seconds 3

# Find links (5 instead of 3)
Write-Host "Finding results..." -ForegroundColor Green
$allLinks = $driver.FindElements([OpenQA.Selenium.By]::TagName("a"))
$links = @()
$count = 0

foreach ($link in $allLinks) {
    if ($count -ge 5) { break }  # Changed from 3 to 5
    
    $href = $link.GetAttribute("href")
    $text = $link.Text
    
    if ($href -and $href.StartsWith("http") -and -not $href.Contains("google.com") -and $text.Length -gt 5) {
        Write-Host "Found: $text" -ForegroundColor Gray
        $links += [PSCustomObject]@{
            URL = $href
            Title = $text
        }
        $count++
    }
}

if ($links.Count -eq 0) {
    Write-Host "No links found!" -ForegroundColor Red
    $driver.Quit()
    pause
    exit
}

# Process 5 links
$results = @()

for ($i = 0; $i -lt $links.Count; $i++) {
    $link = $links[$i]
    Write-Host "`n[$($i+1)/$($links.Count)] Opening: $($link.Title)" -ForegroundColor Cyan
    
    # Save current window
    $mainWindow = $driver.CurrentWindowHandle
    
    # Open new tab
    $driver.ExecuteScript("window.open('" + $link.URL + "');")
    Start-Sleep -Seconds 2
    
    # Switch to new tab
    $driver.SwitchTo().Window($driver.WindowHandles[1])
    Start-Sleep -Seconds 2
    
    # Get page data
    $pageTitle = $driver.Title
    Write-Host "Page title: $pageTitle" -ForegroundColor Gray
    
    $pageText = $driver.FindElement([OpenQA.Selenium.By]::TagName("body")).Text
    
    # Find email
    $email = ""
    if ($pageText -match "[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}") {
        $email = $Matches[0]
        Write-Host "Email found: $email" -ForegroundColor Green
    }
    
    # Find phone
    $phone = ""
    if ($pageText -match "\(?\d{3}\)?[-.\s]?\d{3}[-.\s]?\d{4}") {
        $phone = $Matches[0]
        Write-Host "Phone found: $phone" -ForegroundColor Green
    }
    
    # Add to results
    $results += [PSCustomObject]@{
        Number = $i + 1
        Title = $link.Title
        URL = $link.URL
        PageTitle = $pageTitle
        Email = $email
        Phone = $phone
    }
    
    # Close tab and return
    $driver.Close()
    $driver.SwitchTo().Window($mainWindow)
    Start-Sleep -Seconds 1
}

# Save to Excel
$excelFile = "$env:USERPROFILE\Desktop\WebData_5Sites.xlsx"
$results | Export-Excel -Path $excelFile -WorksheetName "Results" -AutoSize -BoldTopRow

Write-Host "`n========================" -ForegroundColor Green
Write-Host "SCRAPING COMPLETE!" -ForegroundColor Green
Write-Host "========================" -ForegroundColor Green
Write-Host "File saved to: $excelFile" -ForegroundColor Yellow
Write-Host "Pages processed: $($results.Count)" -ForegroundColor Yellow

# Show what we found
if ($results.Count -gt 0) {
    Write-Host "`nData extracted from 5 websites:" -ForegroundColor Cyan
    foreach ($result in $results) {
        Write-Host "$($result.Number). $($result.Title)" -ForegroundColor Gray
        if ($result.Email) { Write-Host "   Email: $($result.Email)" -ForegroundColor Gray }
        if ($result.Phone) { Write-Host "   Phone: $($result.Phone)" -ForegroundColor Gray }
    }
}

# Close browser
$driver.Quit()
Write-Host "`nBrowser closed." -ForegroundColor Gray

# Ask to open file
Write-Host ""
$openFile = Read-Host "Open Excel file? (Y/N)"
if ($openFile -eq "Y" -or $openFile -eq "y") {
    Start-Process $excelFile
    Write-Host "Opening Excel..." -ForegroundColor Green
}

Write-Host "`nDone. Press any key to exit..." -ForegroundColor Gray
pause