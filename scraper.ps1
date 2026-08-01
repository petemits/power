# Business List Builder - Same Folder Setup
# Place this script in the SAME FOLDER as chromedriver.exe and Selenium.WebDriver.dll

# Get the folder where this script is located
$ScriptFolder = $PSScriptRoot
if ([string]::IsNullOrEmpty($ScriptFolder)) {
    $ScriptFolder = Get-Location
}

# Build paths relative to the script folder
$ChromeDriverPath = Join-Path $ScriptFolder "chromedriver.exe"
$WebDriverDllPath = Join-Path $ScriptFolder "Selenium.WebDriver.dll"

Write-Host "=== Business List Builder ===" -ForegroundColor Cyan
Write-Host "Script folder: $ScriptFolder" -ForegroundColor Gray

# Check if required files exist
$missingFiles = @()
if (-not (Test-Path $WebDriverDllPath)) {
    $missingFiles += "Selenium.WebDriver.dll"
    Write-Host "Missing: Selenium.WebDriver.dll" -ForegroundColor Red
}
if (-not (Test-Path $ChromeDriverPath)) {
    $missingFiles += "chromedriver.exe"
    Write-Host "Missing: chromedriver.exe" -ForegroundColor Red
}

if ($missingFiles.Count -gt 0) {
    Write-Host "`nERROR: Required files not found in the script folder!" -ForegroundColor Red
    Write-Host "Please ensure these files are in the same folder as this script:" -ForegroundColor Yellow
    Write-Host "1. chromedriver.exe" -ForegroundColor Yellow
    Write-Host "2. Selenium.WebDriver.dll" -ForegroundColor Yellow
    Write-Host "`nCurrent folder contents:" -ForegroundColor Gray
    Get-ChildItem $ScriptFolder | ForEach-Object { Write-Host "  - $($_.Name)" -ForegroundColor Gray }
    exit
}

Write-Host "All required files found." -ForegroundColor Green

# Load Selenium
try {
    Add-Type -Path $WebDriverDllPath
    Write-Host "Selenium assembly loaded." -ForegroundColor Green
} catch {
    Write-Host "ERROR: Failed to load Selenium DLL." -ForegroundColor Red
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "`nPossible issues:" -ForegroundColor Yellow
    Write-Host "1. Wrong DLL version/architecture" -ForegroundColor Yellow
    Write-Host "2. Corrupted DLL file" -ForegroundColor Yellow
    Write-Host "3. Missing dependencies" -ForegroundColor Yellow
    exit
}

# User prompt for search
$searchTerm = Read-Host "`nEnter business search (e.g., 'marketing Toronto')"
if ([string]::IsNullOrWhiteSpace($searchTerm)) {
    Write-Host "No search term. Exiting." -ForegroundColor Red
    exit
}

Write-Host "`nStarting browser automation..." -ForegroundColor Yellow

# Start browser with minimal options to avoid detection
$options = New-Object OpenQA.Selenium.Chrome.ChromeOptions

# CRITICAL: Options that help avoid immediate blocking
$options.AddArgument("--disable-blink-features=AutomationControlled")
$options.AddArgument("--disable-dev-shm-usage")
$options.AddArgument("--no-sandbox")
$options.AddArgument("--disable-gpu")

# Remove automation indicators
$options.AddExcludedArgument("enable-automation")
$options.AddAdditionalOption("useAutomationExtension", $false)

try {
    # Start ChromeDriver
    Write-Host "Launching ChromeDriver..." -ForegroundColor Yellow
    $driver = New-Object OpenQA.Selenium.Chrome.ChromeDriver($ChromeDriverPath, $options)
    
    # Remove "navigator.webdriver" property
    $driver.ExecuteScript("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
    
    Write-Host "Browser ready. Searching Google..." -ForegroundColor Green
    
    # Navigate to Google
    $driver.Navigate().GoToUrl("https://www.google.com")
    Start-Sleep -Milliseconds 1500
    
    # Check for CAPTCHA/blocking
    if ($driver.PageSource -like "*detected unusual traffic*" -or $driver.PageSource -like "*CAPTCHA*") {
        Write-Host "BLOCKED: Google detected automation." -ForegroundColor Red
        Write-Host "Manual intervention required." -ForegroundColor Yellow
        $driver.Quit()
        exit
    }
    
    # Perform search
    $searchBox = $driver.FindElement([OpenQA.Selenium.By]::Name("q"))
    $searchBox.SendKeys($searchTerm)
    $searchBox.SendKeys([OpenQA.Selenium.Keys]::Enter)
    
    # Wait for results with progressive delay
    Write-Host "Waiting for results..." -ForegroundColor Yellow
    1..5 | ForEach-Object { 
        Start-Sleep -Seconds 1
        Write-Host "." -NoNewline -ForegroundColor Gray
    }
    Write-Host ""
    
    # Extract business data
    $businesses = @()
    $results = $driver.FindElements([OpenQA.Selenium.By]::CssSelector("div.g"))
    
    Write-Host "Found $($results.Count) result blocks." -ForegroundColor Green
    
    foreach ($result in $results) {
        try {
            $business = [PSCustomObject]@{
                Name = ""
                Website = ""
                Description = ""
                Phone = "(N/A on Google)"
                Email = "(N/A on Google)"
                Address = "(N/A on Google)"
            }
            
            # Try to extract name
            try {
                $nameElement = $result.FindElement([OpenQA.Selenium.By]::CssSelector("h3"))
                $business.Name = $nameElement.Text.Trim()
            } catch { }
            
            # Try to extract link
            try {
                $linkElement = $result.FindElement([OpenQA.Selenium.By]::CssSelector("a[href^='http']"))
                $business.Website = $linkElement.GetAttribute("href")
            } catch { }
            
            # Try to extract description
            try {
                $descElement = $result.FindElement([OpenQA.Selenium.By]::CssSelector("div.VwiC3b, span.st, div.s3v9rd"))
                $business.Description = $descElement.Text.Trim()
            } catch { }
            
            # Only add if we have a name
            if (-not [string]::IsNullOrWhiteSpace($business.Name)) {
                $businesses += $business
            }
        } catch {
            # Skip problematic results
        }
    }
    
    # Save results
    if ($businesses.Count -gt 0) {
        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $outputFile = "business_list_${timestamp}.csv"
        
        $businesses | Export-Csv -Path $outputFile -NoTypeInformation -Encoding UTF8
        
        Write-Host "`n✅ SUCCESS: Saved $($businesses.Count) businesses to '$outputFile'" -ForegroundColor Green
        
        # Show preview
        Write-Host "`n--- PREVIEW (first 3 entries) ---" -ForegroundColor Cyan
        $businesses | Select-Object -First 3 | Format-Table @{
            Name="Name"; Expression={if ($_.Name.Length -gt 40) {$_.Name.Substring(0,37)+"..."} else {$_.Name}}
        }, @{
            Name="Website"; Expression={if ($_.Website.Length -gt 30) {$_.Website.Substring(0,27)+"..."} else {$_.Website}}
        } -AutoSize
        
        # Show file location
        Write-Host "`nFile saved to: $(Join-Path $ScriptFolder $outputFile)" -ForegroundColor Gray
        
    } else {
        Write-Host "`n⚠️  No business data extracted. Possible reasons:" -ForegroundColor Yellow
        Write-Host "   • Google blocked the automation" -ForegroundColor Yellow
        Write-Host "   • No results for your search term" -ForegroundColor Yellow
        Write-Host "   • Page structure changed" -ForegroundColor Yellow
    }
    
} catch {
    Write-Host "`n❌ ERROR during automation:" -ForegroundColor Red
    Write-Host "   $($_.Exception.Message)" -ForegroundColor Red
    
    if ($_.Exception.Message -like "*This version of ChromeDriver*") {
        Write-Host "`n⚠️  ChromeDriver version mismatch!" -ForegroundColor Yellow
        Write-Host "   Update ChromeDriver to match your Chrome version." -ForegroundColor Yellow
        Write-Host "   Download from: https://chromedriver.chromium.org/" -ForegroundColor Cyan
    }
} finally {
    # Clean up
    if ($driver -ne $null) {
        try {
            $driver.Quit()
            Write-Host "`nBrowser closed." -ForegroundColor Gray
        } catch { }
    }
}

Write-Host "`nScript completed." -ForegroundColor Cyan
Pause