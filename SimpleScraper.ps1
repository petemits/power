<#
    SIMPLE WORKING SCRAPER
    No complex try-catch blocks
#>

Clear-Host
Write-Host "Starting Web Scraper..." -ForegroundColor Green
Write-Host ""

$ScriptFolder = "C:\Users\user\power\"

# Check files
if (-not (Test-Path "$ScriptFolder\chromedriver.exe")) {
    Write-Host "ERROR: chromedriver.exe missing!" -ForegroundColor Red
    pause
    exit
}

if (-not (Test-Path "$ScriptFolder\WebDriver.dll")) {
    Write-Host "ERROR: WebDriver.dll missing!" -ForegroundColor Red
    pause
    exit
}

# Load WebDriver
Add-Type -Path "$ScriptFolder\WebDriver.dll"

# Load ImportExcel
try {
    Import-Module ImportExcel
} catch {
    Install-Module -Name ImportExcel -Force -Scope CurrentUser
    Import-Module ImportExcel
}

# Get search term
$searchTerm = Read-Host "Enter search term (or press Enter for 'restaurants')"
if ($searchTerm -eq "") { $searchTerm = "restaurants" }

Write-Host "Starting Chrome..." -ForegroundColor Green

# Setup Chrome
$options = New-Object OpenQA.Selenium.Chrome.ChromeOptions
$options.AddArgument("--disable-blink-features=AutomationControlled")

# Start Chrome
$service = [OpenQA.Selenium.Chrome.ChromeDriverService]::CreateDefaultService($ScriptFolder)
$service.HideCommandPromptWindow = $true
$driver = New-Object OpenQA.Selenium.Chrome.ChromeDriver($service, $options)

Write-Host "Searching Google..." -ForegroundColor Green

# Go to Google
$driver.Navigate().GoToUrl("https://www.google.com")
Start-Sleep -Seconds 2

# Search
$searchBox = $driver.FindElement([OpenQA.Selenium.By]::Name("q"))
$searchBox.SendKeys($searchTerm)
$searchBox.SendKeys([OpenQA.Selenium.Keys]::Enter)
Start-Sleep -Seconds 3

# Find links (simple method)
$allLinks = $driver.FindElements([OpenQA.Selenium.By]::TagName("a"))
$links = @()
$count = 0

foreach ($link in $allLinks) {
    if ($count -ge 5) { break }
    
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

# Process links
$results = @()

for ($i = 0; $i -lt $links.Count; $i++) {
    $link = $links[$i]
    Write-Host "`n[$($i+1)/$links.Count] Opening: $($link.Title)" -ForegroundColor Cyan
    
    # Open in new tab
    $mainWindow = $driver.CurrentWindowHandle
    $driver.ExecuteScript("window.open('" + $link.URL + "');")
    Start-Sleep -Seconds 2
    
    # Switch to new tab
    $driver.SwitchTo().Window($driver.WindowHandles[1])
    Start-Sleep -Seconds 2
    
    # Get page data
    $pageTitle = $driver.Title
    $pageText = $driver.FindElement([OpenQA.Selenium.By]::TagName("body")).Text
    
    # Find email
    $email = ""
    if ($pageText -match "[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}") {
        $email = $Matches[0]
    }
    
    # Find phone
    $phone = ""
    if ($pageText -match "\(?\d{3}\)?[-.\s]?\d{3}[-.\s]?\d{4}") {
        $phone = $Matches[0]
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
    
    # Close tab and go back
    $driver.Close()
    $driver.SwitchTo().Window($mainWindow)
    Start-Sleep -Seconds 1
}

# Save to Excel
$excelFile = "$env:USERPROFILE\Desktop\ScrapedData.xlsx"
$results | Export-Excel -Path $excelFile -WorksheetName "Results" -AutoSize -BoldTopRow

Write-Host "`n========================================" -ForegroundColor Green
Write-Host "SCRAPING COMPLETE!" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Green
Write-Host "File saved to: $excelFile" -ForegroundColor Yellow
Write-Host "Pages processed: $($results.Count)" -ForegroundColor Yellow

# Show results
if ($results.Count -gt 0) {
    Write-Host "`nFirst result:" -ForegroundColor Cyan
    Write-Host "Title: $($results[0].Title)" -ForegroundColor Gray
    Write-Host "Email: $($results[0].Email)" -ForegroundColor Gray
    Write-Host "Phone: $($results[0].Phone)" -ForegroundColor Gray
}

# Close browser
$driver.Quit()
Write-Host "`nBrowser closed." -ForegroundColor Gray

# Ask to open file
Write-Host ""
$openFile = Read-Host "Open Excel file? (Y/N)"
if ($openFile -eq "Y" -or $openFile -eq "y") {
    Start-Process $excelFile
}

Write-Host "`nDone. Press any key to exit..." -ForegroundColor Gray
pause