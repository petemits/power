# ===== CONFIGURATION - NO NEED TO CHANGE =====
$ScriptFolder = $PSScriptRoot
if ([string]::IsNullOrEmpty($ScriptFolder)) { $ScriptFolder = Get-Location }
$ChromeDriverPath = Join-Path $ScriptFolder "chromedriver.exe"
$WebDriverDllPath = Join-Path $ScriptFolder "WebDriver.dll"
# ===== END CONFIGURATION =====

# ===== FUNCTION TO SET UP STEALTHY BROWSER =====
function Start-StealthChrome {
    $ChromeOptions = New-Object OpenQA.Selenium.Chrome.ChromeOptions

    # 1. Add standard arguments to appear more human
    $ChromeOptions.AddArgument("--disable-blink-features=AutomationControlled")
    $ChromeOptions.AddArgument("--disable-dev-shm-usage")
    $ChromeOptions.AddArgument("--no-sandbox")
    $ChromeOptions.AddArgument("--disable-gpu")
    $ChromeOptions.AddArgument("start-maximized")

    # 2. Remove the "navigator.webdriver" flag (a major bot giveaway)
    $ChromeOptions.AddExcludedArgument("enable-automation")
    $ChromeOptions.AddAdditionalOption("useAutomationExtension", $false)

    # 3. Load with a realistic User-Agent and language
    $ChromeOptions.AddArgument("--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36")
    $ChromeOptions.AddArgument("--lang=en-US,en;q=0.9")
    $ChromeOptions.AddAdditionalOption("prefs", @{ "intl.accept_languages" = "en-US,en" })

    # 4. Hide popups like "Chrome is being controlled..."
    $ChromeOptions.AddAdditionalOption("excludeSwitches", [String[]]@("enable-automation", "enable-logging"))

    # 5. Start the driver
    $Driver = New-Object OpenQA.Selenium.Chrome.ChromeDriver($ChromeDriverPath, $ChromeOptions)

    # 6. Execute JavaScript to patch webdriver property (CRITICAL STEP)
    $Driver.ExecuteScript("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")

    return $Driver
}
# ===== END FUNCTION =====

# ===== MAIN SCRIPT LOGIC =====
Write-Host "`n=== CLEAN BUSINESS LIST SCRAPER ===`n" -ForegroundColor Cyan
Write-Host "Checking for required files in: $ScriptFolder" -ForegroundColor Gray

# Validate files exist
if (-not (Test-Path $WebDriverDllPath)) { Write-Host "ERROR: WebDriver.dll not found." -ForegroundColor Red; exit }
if (-not (Test-Path $ChromeDriverPath)) { Write-Host "ERROR: chromedriver.exe not found." -ForegroundColor Red; exit }
Write-Host "  -> Files found. Loading WebDriver..." -ForegroundColor Green

# Load Selenium assembly
Add-Type -Path $WebDriverDllPath
Write-Host "  -> Selenium loaded." -ForegroundColor Green

# Get user input
Write-Host "`n--- USER INPUT ---" -ForegroundColor Yellow
$targetUrl = Read-Host "Enter the FULL URL of the page to scrape"
if ([string]::IsNullOrWhiteSpace($targetUrl)) {
    Write-Host "No URL provided. Exiting." -ForegroundColor Yellow
    exit
}
Write-Host "Target URL: $targetUrl" -ForegroundColor Cyan

# Ask for element selector
Write-Host "`n--- ELEMENT SELECTION ---" -ForegroundColor Yellow
Write-Host "Tip: Open your browser's Developer Tools (F12), use 'Inspect' on an element, and copy its CSS selector."
$cssSelector = Read-Host "Enter the CSS selector for the business listing containers (e.g., '.listing', 'div.result')"
if ([string]::IsNullOrWhiteSpace($cssSelector)) {
    Write-Host "No selector provided. Exiting." -ForegroundColor Yellow
    exit
}
Write-Host "Selector: $cssSelector" -ForegroundColor Cyan

# Start the browser
Write-Host "`n--- STARTING SCRAPER ---" -ForegroundColor Yellow
$driver = $null
try {
    $driver = Start-StealthChrome
    Write-Host "  -> Browser started." -ForegroundColor Green

    # Navigate to the page
    Write-Host "  -> Navigating to page..." -ForegroundColor Yellow
    $driver.Navigate().GoToUrl($targetUrl)
    Start-Sleep -Seconds 5

    # Check for immediate blocking (common on Google)
    $pageText = $driver.PageSource
    if ($pageText -like "*Checking your browser*" -or $pageText -like "*unusual traffic*" -or $pageText -like "*captcha*") {
        Write-Host "`n!!! BLOCKED !!!" -ForegroundColor Red
        Write-Host "The website detected automation. This is especially common on Google Search." -ForegroundColor Yellow
        Write-Host "Try using a business directory website like Yellow Pages or Yelp instead of Google." -ForegroundColor Yellow
        $driver.Quit()
        exit
    }

    # Extract data
    Write-Host "  -> Extracting elements with selector: '$cssSelector'..." -ForegroundColor Yellow
    $listingElements = $driver.FindElements([OpenQA.Selenium.By]::CssSelector($cssSelector))
    $count = $listingElements.Count
    Write-Host "  -> Found $count listing elements." -ForegroundColor Green

    $results = @()
    foreach ($element in $listingElements) {
        $item = [PSCustomObject]@{
            Name = ($element.FindElement([OpenQA.Selenium.By]::CssSelector("h2, h3, .name, [class*='title']")) -replace "`n|`r", " ").Trim()
            Phone = ($element.FindElement([OpenQA.Selenium.By]::CssSelector("[href^='tel:'], .phone, .telephone")) -replace "`n|`r", " ").Trim()
            Website = $element.FindElement([OpenQA.Selenium.By]::CssSelector("a[href^='http']")).GetAttribute("href")
            Address = ($element.FindElement([OpenQA.Selenium.By]::CssSelector("address, .address, [class*='addr']")) -replace "`n|`r", " ").Trim()
        }
        $results += $item
    }

    # Save results
    if ($results.Count -gt 0) {
        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $filename = "BusinessList_$timestamp.csv"
        $fullPath = Join-Path $ScriptFolder $filename
        $results | Export-Csv -Path $fullPath -NoTypeInformation -Encoding UTF8
        Write-Host "`n--- SUCCESS ---" -ForegroundColor Green
        Write-Host "Saved $($results.Count) records to: $filename" -ForegroundColor Cyan
        Write-Host "`nPreview of first 3 items:" -ForegroundColor Yellow
        $results | Select-Object -First 3 | Format-Table
    }
    else {
        Write-Host "`n--- NO DATA ---" -ForegroundColor Yellow
        Write-Host "No data was extracted. The CSS selector might be incorrect." -ForegroundColor Yellow
    }
}
catch {
    Write-Host "`n--- ERROR ---" -ForegroundColor Red
    Write-Host "An error occurred: $($_.Exception.Message)" -ForegroundColor Red
}
finally {
    if ($driver -ne $null) {
        $driver.Quit()
        Write-Host "`nBrowser closed. Script finished." -ForegroundColor Gray
    }
}