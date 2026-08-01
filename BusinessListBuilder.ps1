# Business List Builder with Direct Selenium References
# This script prompts for a search term, scrapes Google results, and exports to CSV

# --- CONFIGURATION: SET THESE PATHS BEFORE RUNNING ---
# Path to your ChromeDriver executable
$ChromeDriverPath = "C:\Your\Actual\Path\To\chromedriver.exe"  # CHANGE THIS

# Path to your Selenium WebDriver.dll (typically in a Selenium.WebDriver folder)
$WebDriverDllPath = "C:\Your\Actual\Path\To\Selenium.WebDriver.dll"  # CHANGE THIS

# --- SCRIPT START ---
Write-Host "=== Business List Builder ===" -ForegroundColor Cyan
Write-Host "Initializing Selenium from local files..." -ForegroundColor Yellow

# Load the Selenium assembly
try {
    Add-Type -Path $WebDriverDllPath
    Write-Host "Selenium assembly loaded successfully." -ForegroundColor Green
} catch {
    Write-Host "ERROR: Failed to load Selenium assembly. Check the path: $WebDriverDllPath" -ForegroundColor Red
    Write-Host "Error details: $_" -ForegroundColor Red
    exit
}

# Prompt user for search term
$UserSearchTerm = Read-Host "`nEnter business search term (e.g., 'marketing agencies Toronto GTA')"
if ([string]::IsNullOrWhiteSpace($UserSearchTerm)) {
    Write-Host "No search term provided. Exiting." -ForegroundColor Red
    exit
}

# Set up Chrome options
$ChromeOptions = New-Object OpenQA.Selenium.Chrome.ChromeOptions
$ChromeOptions.AddArgument("--start-maximized")
# $ChromeOptions.AddArgument("--headless")  # Uncomment for headless mode

try {
    # Initialize ChromeDriver with your path
    Write-Host "Starting ChromeDriver..." -ForegroundColor Yellow
    $Driver = New-Object OpenQA.Selenium.Chrome.ChromeDriver($ChromeDriverPath, $ChromeOptions)
    Write-Host "Browser started successfully." -ForegroundColor Green
} catch {
    Write-Host "ERROR: Failed to start ChromeDriver. Check the path: $ChromeDriverPath" -ForegroundColor Red
    Write-Host "Error details: $_" -ForegroundColor Red
    exit
}

# Navigate to Google
Write-Host "Navigating to Google..." -ForegroundColor Yellow
$Driver.Navigate().GoToUrl("https://www.google.com")

# Perform search
Write-Host "Performing search: '$UserSearchTerm'" -ForegroundColor Yellow
try {
    $SearchBox = $Driver.FindElement([OpenQA.Selenium.By]::Name("q"))
    $SearchBox.SendKeys($UserSearchTerm)
    $SearchBox.SendKeys([OpenQA.Selenium.Keys]::Enter)
    Start-Sleep -Seconds 3  # Wait for results
} catch {
    Write-Host "ERROR: Could not perform search. Google page structure may have changed." -ForegroundColor Red
    $Driver.Quit()
    exit
}

# Extract business listings
Write-Host "Extracting business listings..." -ForegroundColor Yellow
$BusinessList = @()

try {
    # Get all search result elements (Google's result blocks)
    $Results = $Driver.FindElements([OpenQA.Selenium.By]::CssSelector("div.g"))
    
    Write-Host "Found $($Results.Count) search results." -ForegroundColor Green
    
    foreach ($Result in $Results) {
        try {
            $BusinessInfo = [PSCustomObject]@{
                Name = ""
                Website = ""
                Phone = ""
                Email = ""
                Description = ""
                Address = ""
            }
            
            # Extract business name
            try {
                $NameElement = $Result.FindElement([OpenQA.Selenium.By]::CssSelector("h3"))
                $BusinessInfo.Name = $NameElement.Text
            } catch { }
            
            # Extract website link
            try {
                $LinkElement = $Result.FindElement([OpenQA.Selenium.By]::CssSelector("a"))
                $BusinessInfo.Website = $LinkElement.GetAttribute("href")
            } catch { }
            
            # Extract description
            try {
                $DescElement = $Result.FindElement([OpenQA.Selenium.By]::CssSelector("div.VwiC3b"))
                $BusinessInfo.Description = $DescElement.Text
            } catch { }
            
            # IMPORTANT: Google search results don't typically show phone/email/address
            # To get these, you'd need to visit each website and scrape contact pages
            
            # Only add to list if we have at least a name
            if (-not [string]::IsNullOrWhiteSpace($BusinessInfo.Name)) {
                $BusinessList += $BusinessInfo
            }
        } catch {
            Write-Verbose "Skipping a result due to extraction error"
        }
    }
} catch {
    Write-Host "ERROR: Could not extract search results. Google may be blocking automated access." -ForegroundColor Red
}

# Save to CSV
if ($BusinessList.Count -gt 0) {
    $Timestamp = Get-Date -Format "yyyyMMdd_HHmm"
    $OutputFile = "BusinessList_$Timestamp.csv"
    
    try {
        $BusinessList | Export-Csv -Path $OutputFile -NoTypeInformation -Encoding UTF8
        Write-Host "`nSUCCESS: Saved $($BusinessList.Count) businesses to $OutputFile" -ForegroundColor Green
        
        # Show a preview
        Write-Host "`nFirst 3 results preview:" -ForegroundColor Cyan
        $BusinessList | Select-Object -First 3 | Format-Table Name, Website
    } catch {
        Write-Host "ERROR: Could not save CSV file." -ForegroundColor Red
    }
} else {
    Write-Host "WARNING: No business data was extracted." -ForegroundColor Yellow
    Write-Host "This could be because:" -ForegroundColor Yellow
    Write-Host "1. Google is blocking automated access (most likely)" -ForegroundColor Yellow
    Write-Host "2. The search returned no results" -ForegroundColor Yellow
    Write-Host "3. The CSS selectors need updating (Google changed their page structure)" -ForegroundColor Yellow
}

# Clean up
Write-Host "`nCleaning up..." -ForegroundColor Yellow
$Driver.Quit()
Write-Host "Script completed." -ForegroundColor Cyan