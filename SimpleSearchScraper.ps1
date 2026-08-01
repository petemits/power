<#
    SIMPLE SEARCH TO EXCEL SCRAPER
    1. Searches Google
    2. Clicks first result link
    3. Copies page text
    4. Pastes to Excel
    5. Stops automatically
#>

Clear-Host
Write-Host "SIMPLE SEARCH TO EXCEL SCRAPER" -ForegroundColor Cyan
Write-Host "=================================" -ForegroundColor Cyan
Write-Host ""

# Check required files
$folder = "C:\Users\user\power\"
if (-not (Test-Path "$folder\chromedriver.exe")) {
    Write-Host "ERROR: ChromeDriver missing!" -ForegroundColor Red
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

# Get search term
$searchTerm = Read-Host "Enter your search term"
if ($searchTerm -eq "") {
    $searchTerm = "technology"
    Write-Host "Using default: 'technology'" -ForegroundColor Yellow
}

Write-Host "`nStarting Chrome..." -ForegroundColor Green

# Start Chrome
$options = New-Object OpenQA.Selenium.Chrome.ChromeOptions
$options.AddArgument("--disable-blink-features=AutomationControlled")

$service = [OpenQA.Selenium.Chrome.ChromeDriverService]::CreateDefaultService($folder)
$service.HideCommandPromptWindow = $true
$driver = New-Object OpenQA.Selenium.Chrome.ChromeDriver($service, $options)

Write-Host "Searching Google..." -ForegroundColor Green

# Go to Google
$driver.Navigate().GoToUrl("https://www.google.com")
Start-Sleep -Seconds 2

# Perform search
$searchBox = $driver.FindElement([OpenQA.Selenium.By]::Name("q"))
$searchBox.SendKeys($searchTerm)
$searchBox.SendKeys([OpenQA.Selenium.Keys]::Enter)
Start-Sleep -Seconds 3

# Find first valid result link
Write-Host "Finding first result..." -ForegroundColor Green
$firstLink = $null
$allLinks = $driver.FindElements([OpenQA.Selenium.By]::TagName("a"))

foreach ($link in $allLinks) {
    $href = $link.GetAttribute("href")
    $text = $link.Text
    
    if ($href -and $href.StartsWith("http") -and -not $href.Contains("google.com") -and $text.Length -gt 5) {
        $firstLink = [PSCustomObject]@{
            URL = $href
            Title = $text
        }
        break
    }
}

if ($firstLink -eq $null) {
    Write-Host "ERROR: No search results found!" -ForegroundColor Red
    $driver.Quit()
    pause
    exit
}

Write-Host "Found: $($firstLink.Title)" -ForegroundColor Green
Write-Host "URL: $($firstLink.URL)" -ForegroundColor Gray

# Click the link and get page text
Write-Host "`nClicking link..." -ForegroundColor Green

# Open in new tab
$originalWindow = $driver.CurrentWindowHandle
$driver.ExecuteScript("window.open(arguments[0]);", $firstLink.URL)
Start-Sleep -Seconds 2

# Switch to new tab
$driver.SwitchTo().Window($driver.WindowHandles[1])
Start-Sleep -Seconds 3

# Get page text
Write-Host "Extracting page text..." -ForegroundColor Green
$pageTitle = $driver.Title
$pageText = $driver.FindElement([OpenQA.Selenium.By]::TagName("body")).Text

Write-Host "Page title: $pageTitle" -ForegroundColor Gray
Write-Host "Text length: $($pageText.Length) characters" -ForegroundColor Gray

# Close the new tab
$driver.Close()
$driver.SwitchTo().Window($originalWindow)

# Create Excel file with the text
Write-Host "`nCreating Excel file..." -ForegroundColor Green

try {
    # Create data object
    $data = [PSCustomObject]@{
        SearchTerm = $searchTerm
        LinkTitle = $firstLink.Title
        URL = $firstLink.URL
        PageTitle = $pageTitle
        PageText = $pageText
        ScrapedDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    }
    
    # Save to Excel
    $excelFile = "$env:USERPROFILE\Desktop\ScrapedPageText.xlsx"
    $data | Export-Excel -Path $excelFile -WorksheetName "Scraped Data" -AutoSize
    
    Write-Host "=================================" -ForegroundColor Green
    Write-Host "SUCCESS!" -ForegroundColor Green
    Write-Host "=================================" -ForegroundColor Green
    Write-Host "Excel file created: $excelFile" -ForegroundColor Yellow
    
    # Show what's in the file
    Write-Host "`nFile contains:" -ForegroundColor Cyan
    Write-Host "1. Search term: $searchTerm" -ForegroundColor Gray
    Write-Host "2. Link title: $($firstLink.Title)" -ForegroundColor Gray
    Write-Host "3. URL: $($firstLink.URL)" -ForegroundColor Gray
    Write-Host "4. Page title: $pageTitle" -ForegroundColor Gray
    Write-Host "5. Page text: First 200 characters..." -ForegroundColor Gray
    Write-Host "   '$($pageText.Substring(0, [Math]::Min(200, $pageText.Length)))...'" -ForegroundColor Gray
    
} catch {
    Write-Host "ERROR creating Excel: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Displaying data instead:" -ForegroundColor Yellow
    Write-Host "Search term: $searchTerm"
    Write-Host "URL: $($firstLink.URL)"
    Write-Host "Page text (first 500 chars):"
    Write-Host $pageText.Substring(0, [Math]::Min(500, $pageText.Length))
}

# Close browser and exit
Write-Host "`nClosing browser..." -ForegroundColor Gray
$driver.Quit()

Write-Host "Script complete. Press any key to exit..." -ForegroundColor Gray
pause