# ============================================================
# WEBSITE BUSINESS DATA EXTRACTOR - DIRECT SELENIUM VERSION
# ============================================================
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  WEBSITE BUSINESS DATA EXTRACTOR" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# Ask for website URL
$WebsiteUrl = Read-Host "Enter website URL (e.g., https://www.example.com)"
if ($WebsiteUrl -notmatch "^https?://") {
    $WebsiteUrl = "https://$WebsiteUrl"
}

Write-Host ""
Write-Host "Processing: $WebsiteUrl" -ForegroundColor Yellow

# Set output file
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$OutputExcelPath = "BusinessData_$timestamp.xlsx"

# Check for ChromeDriver
$ChromeDriverPath = ".\chromedriver.exe"
if (-not (Test-Path $ChromeDriverPath)) {
    Write-Host "ERROR: chromedriver.exe not found in current folder!" -ForegroundColor Red
    Write-Host "Please place chromedriver.exe in: $(Get-Location)" -ForegroundColor Yellow
    exit
}

Write-Host "1. Checking ChromeDriver..." -ForegroundColor Cyan
try {
    $driverInfo = & $ChromeDriverPath --version 2>&1
    Write-Host "   $driverInfo" -ForegroundColor Green
} catch {
    Write-Host "   ChromeDriver found" -ForegroundColor Yellow
}

# DIRECT SELENIUM .NET VERSION (No PowerShell module needed)
Write-Host "2. Starting browser with direct Selenium..." -ForegroundColor Cyan

try {
    # Load Selenium WebDriver .NET assemblies
    # Try to load from common locations
    $seleniumPaths = @(
        "$env:USERPROFILE\.nuget\packages\selenium.webdriver\4.16.2\lib\netstandard2.0\WebDriver.dll",
        "$env:USERPROFILE\.nuget\packages\selenium.webdriver\4.16.0\lib\netstandard2.0\WebDriver.dll",
        "$env:USERPROFILE\.nuget\packages\selenium.webdriver\4.15.0\lib\netstandard2.0\WebDriver.dll",
        ".\WebDriver.dll",
        ".\packages\Selenium.WebDriver.4.16.2\lib\netstandard2.0\WebDriver.dll"
    )
    
    $assemblyLoaded = $false
    foreach ($path in $seleniumPaths) {
        if (Test-Path $path) {
            try {
                Add-Type -Path $path
                Write-Host "   Loaded Selenium from: $path" -ForegroundColor Green
                $assemblyLoaded = $true
                break
            } catch {
                # Try next path
            }
        }
    }
    
    if (-not $assemblyLoaded) {
        # Try to load from installed module
        try {
            $modulePath = (Get-Module -Name Selenium -ListAvailable | Select-Object -First 1).Path
            $moduleDir = Split-Path $modulePath
            $dllPath = Join-Path $moduleDir "lib\WebDriver.dll"
            
            if (Test-Path $dllPath) {
                Add-Type -Path $dllPath
                Write-Host "   Loaded Selenium from module" -ForegroundColor Green
                $assemblyLoaded = $true
            }
        } catch {
            # Continue
        }
    }
    
    if (-not $assemblyLoaded) {
        Write-Host "   ERROR: Selenium .NET assembly not found!" -ForegroundColor Red
        Write-Host "   Please install Selenium with:" -ForegroundColor Yellow
        Write-Host "   1. Install NuGet package: Install-Package Selenium.WebDriver -Force" -ForegroundColor White
        Write-Host "   2. OR install module: Install-Module -Name Selenium -Force" -ForegroundColor White
        exit
    }
    
    # Create ChromeDriver service
    $service = [OpenQA.Selenium.Chrome.ChromeDriverService]::CreateDefaultService((Get-Location).Path, "chromedriver.exe")
    $service.HideCommandPromptWindow = $true
    
    # Create Chrome options using direct .NET
    $options = New-Object OpenQA.Selenium.Chrome.ChromeOptions
    $options.AddArgument("--headless")
    $options.AddArgument("--no-sandbox")
    $options.AddArgument("--disable-dev-shm-usage")
    $options.AddArgument("--disable-gpu")
    
    # Create the driver
    $Driver = New-Object OpenQA.Selenium.Chrome.ChromeDriver($service, $options)
    
    Write-Host "   Browser started successfully" -ForegroundColor Green
    
} catch {
    Write-Host "   ERROR: Failed to start browser" -ForegroundColor Red
    Write-Host "   Details: $_" -ForegroundColor Red
    Write-Host ""
    Write-Host "QUICK FIX OPTIONS:" -ForegroundColor Yellow
    Write-Host "A. Install Selenium module:" -ForegroundColor White
    Write-Host "   Install-Module -Name Selenium -Force -AllowClobber" -ForegroundColor Gray
    Write-Host "B. Use simpler script (alternative below)" -ForegroundColor White
    exit
}

# Navigate to website
Write-Host "3. Loading website..." -ForegroundColor Cyan
try {
    $Driver.Navigate().GoToUrl($WebsiteUrl)
    
    # Wait for page to load
    Start-Sleep -Seconds 3
    
    $pageTitle = $Driver.Title
    Write-Host "   Loaded: $pageTitle" -ForegroundColor Green
    
} catch {
    Write-Host "   ERROR: Failed to load website" -ForegroundColor Red
    $Driver.Quit()
    exit
}

# Extract data
Write-Host "4. Extracting business data..." -ForegroundColor Cyan

# Store results
$results = @{
    WebsiteURL = $WebsiteUrl
    CompanyName = "Not found"
    PhoneNumber = "Not found"
    EmailAddress = "Not found"
    BusinessDescription = "Not found"
    PageTitle = $pageTitle
    PageHTML = $Driver.PageSource
    Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
}

# Get page text for regex searching
$pageText = $Driver.FindElement([OpenQA.Selenium.By]::TagName("body")).Text

# Find company name
Write-Host "   Looking for company name..." -ForegroundColor Gray
try {
    # Try h1 tags first
    $h1Elements = $Driver.FindElements([OpenQA.Selenium.By]::TagName("h1"))
    if ($h1Elements.Count -gt 0) {
        $name = $h1Elements[0].Text.Trim()
        if ($name -ne "" -and $name.Length -lt 100) {
            $results.CompanyName = $name
            Write-Host "     Found in h1: $name" -ForegroundColor Green
        }
    }
} catch { }

# Find phone numbers
Write-Host "   Looking for phone numbers..." -ForegroundColor Gray
$phones = @()
$phonePatterns = @(
    '\(\d{3}\) \d{3}-\d{4}',
    '\d{3}-\d{3}-\d{4}',
    '\d{3}\.\d{3}\.\d{4}',
    '\+\d{1,3} \d{3} \d{3} \d{4}'
)

foreach ($pattern in $phonePatterns) {
    $matches = [regex]::Matches($pageText, $pattern)
    foreach ($match in $matches) {
        if (-not $phones.Contains($match.Value)) {
            $phones += $match.Value
        }
    }
}

if ($phones.Count -gt 0) {
    $results.PhoneNumber = $phones -join "; "
    Write-Host "     Found: $($results.PhoneNumber)" -ForegroundColor Green
}

# Find emails
Write-Host "   Looking for email addresses..." -ForegroundColor Gray
$emails = @()
$emailPattern = '\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b'
$emailMatches = [regex]::Matches($pageText, $emailPattern)

foreach ($match in $emailMatches) {
    $email = $match.Value
    if ($email -notmatch 'noreply|no-reply|donotreply|do-not-reply|admin') {
        if (-not $emails.Contains($email)) {
            $emails += $email
        }
    }
}

if ($emails.Count -gt 0) {
    $results.EmailAddress = $emails -join "; "
    Write-Host "     Found: $($results.EmailAddress)" -ForegroundColor Green
}

# Find description
Write-Host "   Looking for business description..." -ForegroundColor Gray
try {
    # Try meta description
    $metaElements = $Driver.FindElements([OpenQA.Selenium.By]::CssSelector("meta[name='description']"))
    if ($metaElements.Count -gt 0) {
        $desc = $metaElements[0].GetAttribute("content")
        if ($desc -ne "" -and $desc.Length -gt 20) {
            $results.BusinessDescription = $desc.Substring(0, [Math]::Min(500, $desc.Length))
            Write-Host "     Found in meta description" -ForegroundColor Green
        }
    }
} catch { }

# Close browser
Write-Host "5. Closing browser..." -ForegroundColor Cyan
$Driver.Quit()
Write-Host "   Browser closed" -ForegroundColor Green

# Save to Excel
Write-Host "6. Saving data to Excel..." -ForegroundColor Cyan

try {
    # Create object for export
    $exportData = [PSCustomObject]@{
        WebsiteURL = $results.WebsiteURL
        CompanyName = $results.CompanyName
        PhoneNumber = $results.PhoneNumber
        EmailAddress = $results.EmailAddress
        BusinessDescription = $results.BusinessDescription
        PageTitle = $results.PageTitle
        Timestamp = $results.Timestamp
    }
    
    # Check for ImportExcel module
    if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
        Write-Host "   Installing ImportExcel module..." -ForegroundColor Yellow
        Install-Module -Name ImportExcel -Force -Scope CurrentUser
    }
    
    # Export to Excel
    $exportData | Export-Excel -Path $OutputExcelPath -WorksheetName "Business Data" -AutoSize -TableName "ExtractedData"
    
    # Add HTML as second sheet
    $htmlData = [PSCustomObject]@{
        WebsiteURL = $results.WebsiteURL
        HTML_Content = $results.PageHTML
        Extraction_Date = $results.Timestamp
    }
    
    $htmlData | Export-Excel -Path $OutputExcelPath -WorksheetName "Page HTML" -AutoSize
    
    # Show results
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Green
    Write-Host "✅ EXTRACTION COMPLETE!" -ForegroundColor Green
    Write-Host "========================================" -ForegroundColor Green
    Write-Host ""
    Write-Host "EXTRACTED DATA:" -ForegroundColor Cyan
    Write-Host "  Company: $($results.CompanyName)" -ForegroundColor Gray
    Write-Host "  Phone: $($results.PhoneNumber)" -ForegroundColor Gray
    Write-Host "  Email: $($results.EmailAddress)" -ForegroundColor Gray
    Write-Host "  Description: $($results.BusinessDescription.Substring(0, [Math]::Min(100, $results.BusinessDescription.Length)))..." -ForegroundColor Gray
    Write-Host ""
    Write-Host "File saved: $OutputExcelPath" -ForegroundColor Yellow
    
    # Ask to open
    $open = Read-Host "Open Excel file? (Y/N)"
    if ($open -eq "Y" -or $open -eq "y") {
        Invoke-Item $OutputExcelPath
    }
    
} catch {
    Write-Host "   ERROR saving Excel: $_" -ForegroundColor Red
    # Save as CSV
    $csvPath = $OutputExcelPath -replace '\.xlsx$', '.csv'
    $exportData | Export-Csv -Path $csvPath -NoTypeInformation
    Write-Host "   Data saved as CSV: $csvPath" -ForegroundColor Yellow
}

Write-Host ""
Write-Host "Press any key to exit..."
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")