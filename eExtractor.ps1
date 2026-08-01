# ============================================================
# WEBSITE DATA EXTRACTOR - SIMPLE CSV VERSION
# ============================================================
Write-Host "WEBSITE DATA EXTRACTOR" -ForegroundColor Cyan
Write-Host "=====================" -ForegroundColor Cyan
Write-Host ""

# Ask for website URL
$WebsiteUrl = Read-Host "Enter website URL (e.g., https://www.example.com)"
if ($WebsiteUrl -notmatch "^https?://") {
    $WebsiteUrl = "https://$WebsiteUrl"
}

Write-Host ""
Write-Host "Processing: $WebsiteUrl" -ForegroundColor Yellow

# Set output file (CSV instead of Excel)
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$OutputCSVPath = "BusinessData_$timestamp.csv"
$OutputHTMLPath = "PageHTML_$timestamp.txt"

Write-Host "Output files will be:" -ForegroundColor Gray
Write-Host "  Data: $OutputCSVPath" -ForegroundColor Gray
Write-Host "  HTML: $OutputHTMLPath" -ForegroundColor Gray
Write-Host ""

# Check for ChromeDriver
$ChromeDriverPath = ".\chromedriver.exe"
if (-not (Test-Path $ChromeDriverPath)) {
    Write-Host "ERROR: chromedriver.exe not found!" -ForegroundColor Red
    Write-Host "Place chromedriver.exe in: $(Get-Location)" -ForegroundColor Yellow
    exit
}

Write-Host "1. Starting browser..." -ForegroundColor Cyan

try {
    # Load Selenium assembly
    Add-Type -Path ".\WebDriver.dll" -ErrorAction Stop
    Write-Host "   Selenium loaded" -ForegroundColor Green
} catch {
    Write-Host "   ERROR: WebDriver.dll not found!" -ForegroundColor Red
    Write-Host "   Place WebDriver.dll in same folder" -ForegroundColor Yellow
    exit
}

try {
    # Create ChromeDriver service
    $service = [OpenQA.Selenium.Chrome.ChromeDriverService]::CreateDefaultService((Get-Location).Path, "chromedriver.exe")
    $service.HideCommandPromptWindow = $true
    
    # Create Chrome options
    $options = New-Object OpenQA.Selenium.Chrome.ChromeOptions
    $options.AddArgument("--headless")
    $options.AddArgument("--no-sandbox")
    
    # Create the driver
    $Driver = New-Object OpenQA.Selenium.Chrome.ChromeDriver($service, $options)
    
    Write-Host "   Browser started" -ForegroundColor Green
    
} catch {
    Write-Host "   ERROR: Failed to start browser" -ForegroundColor Red
    Write-Host "   Details: $_" -ForegroundColor Red
    exit
}

# Navigate to website
Write-Host "2. Loading website..." -ForegroundColor Cyan
try {
    $Driver.Navigate().GoToUrl($WebsiteUrl)
    Start-Sleep -Seconds 3
    
    $pageTitle = $Driver.Title
    Write-Host "   Title: $pageTitle" -ForegroundColor Green
    
} catch {
    Write-Host "   ERROR: Failed to load website" -ForegroundColor Red
    $Driver.Quit()
    exit
}

# Extract data
Write-Host "3. Extracting data..." -ForegroundColor Cyan

# Get page text and HTML
$pageText = $Driver.FindElement([OpenQA.Selenium.By]::TagName("body")).Text
$pageHTML = $Driver.PageSource

# Save HTML to file immediately (so we don't lose it)
$pageHTML | Out-File -FilePath $OutputHTMLPath -Encoding UTF8
Write-Host "   HTML saved to: $OutputHTMLPath" -ForegroundColor Green

# Find company name
$companyName = "Not found"
try {
    $h1Elements = $Driver.FindElements([OpenQA.Selenium.By]::TagName("h1"))
    if ($h1Elements.Count -gt 0) {
        $name = $h1Elements[0].Text.Trim()
        if ($name -ne "" -and $name.Length -lt 100) {
            $companyName = $name
        }
    }
} catch { }

# Find phone numbers
$phoneNumbers = @()
$phonePatterns = @(
    '\(\d{3}\) \d{3}-\d{4}',
    '\d{3}-\d{3}-\d{4}',
    '\d{3}\.\d{3}\.\d{4}'
)

foreach ($pattern in $phonePatterns) {
    $matches = [regex]::Matches($pageText, $pattern)
    foreach ($match in $matches) {
        if (-not $phoneNumbers.Contains($match.Value)) {
            $phoneNumbers += $match.Value
        }
    }
}

# Find emails
$emails = @()
$emailPattern = '\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b'
$emailMatches = [regex]::Matches($pageText, $emailPattern)

foreach ($match in $emailMatches) {
    $email = $match.Value
    if ($email -notmatch 'noreply|no-reply|donotreply') {
        if (-not $emails.Contains($email)) {
            $emails += $email
        }
    }
}

# Close browser
Write-Host "4. Closing browser..." -ForegroundColor Cyan
$Driver.Quit()
Write-Host "   Browser closed" -ForegroundColor Green

# Create data object
$extractedData = [PSCustomObject]@{
    WebsiteURL = $WebsiteUrl
    PageTitle = $pageTitle
    CompanyName = $companyName
    PhoneNumbers = if ($phoneNumbers.Count -gt 0) { $phoneNumbers -join "; " } else { "Not found" }
    EmailAddresses = if ($emails.Count -gt 0) { $emails -join "; " } else { "Not found" }
    ExtractionDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    HTMLFile = $OutputHTMLPath
}

# Save to CSV (this ALWAYS works)
Write-Host "5. Saving data to CSV..." -ForegroundColor Cyan
try {
    $extractedData | Export-Csv -Path $OutputCSVPath -NoTypeInformation -Encoding UTF8
    Write-Host "   CSV saved: $OutputCSVPath" -ForegroundColor Green
    
    # Show what was extracted
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Green
    Write-Host "✅ EXTRACTION COMPLETE!" -ForegroundColor Green
    Write-Host "========================================" -ForegroundColor Green
    Write-Host ""
    Write-Host "EXTRACTED DATA:" -ForegroundColor Cyan
    Write-Host "  URL: $WebsiteUrl" -ForegroundColor Gray
    Write-Host "  Title: $pageTitle" -ForegroundColor Gray
    Write-Host "  Company: $companyName" -ForegroundColor Gray
    Write-Host "  Phones: $(if ($phoneNumbers.Count -gt 0) { $phoneNumbers -join ', ' } else { 'None' })" -ForegroundColor Gray
    Write-Host "  Emails: $(if ($emails.Count -gt 0) { $emails -join ', ' } else { 'None' })" -ForegroundColor Gray
    Write-Host ""
    Write-Host "FILES CREATED:" -ForegroundColor Cyan
    Write-Host "  1. $OutputCSVPath (data)" -ForegroundColor Yellow
    Write-Host "  2. $OutputHTMLPath (web page HTML)" -ForegroundColor Yellow
    Write-Host ""
    
    # Ask to open files
    $openCSV = Read-Host "Open CSV file in Notepad? (Y/N)"
    if ($openCSV -eq "Y" -or $openCSV -eq "y") {
        notepad $OutputCSVPath
    }
    
    $openHTML = Read-Host "Open HTML file in browser? (Y/N)"
    if ($openHTML -eq "Y" -or $openHTML -eq "y") {
        Start-Process $OutputHTMLPath
    }
    
} catch {
    Write-Host "   ERROR saving CSV: $_" -ForegroundColor Red
    # Last resort - save as simple text
    $textPath = "Data_$timestamp.txt"
    @"
WEBSITE: $WebsiteUrl
TITLE: $pageTitle
COMPANY: $companyName
PHONES: $(if ($phoneNumbers.Count -gt 0) { $phoneNumbers -join ', ' } else { 'None' })
EMAILS: $(if ($emails.Count -gt 0) { $emails -join ', ' } else { 'None' })
DATE: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")
"@ | Out-File -FilePath $textPath
    Write-Host "   Data saved as text: $textPath" -ForegroundColor Yellow
}

Write-Host ""
Write-Host "Done! Press any key to exit..." -ForegroundColor Gray
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")