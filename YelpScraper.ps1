# YELP GTA BUSINESS SCRAPER
# This script searches Yelp for businesses in Greater Toronto Area

# Working folder check
$folder = Get-Location
Write-Host "Working in folder: $folder"

# Check for required files
$chromeDriver = "chromedriver.exe"
$webDriverDll = "WebDriver.dll"

if (-not (Test-Path $chromeDriver)) {
    Write-Host "ERROR: chromedriver.exe not found in current folder" -ForegroundColor Red
    Write-Host "Please download chromedriver.exe and place it here: $folder" -ForegroundColor Yellow
    exit
}

if (-not (Test-Path $webDriverDll)) {
    Write-Host "ERROR: WebDriver.dll not found in current folder" -ForegroundColor Red
    Write-Host "Please ensure WebDriver.dll is in this folder" -ForegroundColor Yellow
    exit
}

Write-Host "Files found: chromedriver.exe and WebDriver.dll" -ForegroundColor Green

# Load the Selenium DLL
try {
    Add-Type -Path $webDriverDll
    Write-Host "Selenium DLL loaded successfully" -ForegroundColor Green
}
catch {
    Write-Host "ERROR: Failed to load WebDriver.dll" -ForegroundColor Red
    Write-Host "Error details: $_" -ForegroundColor Red
    exit
}

# Get user search input
Write-Host ""
Write-Host "YELP GTA BUSINESS SEARCH" -ForegroundColor Cyan
Write-Host ""
Write-Host "What type of business do you want to find?"
Write-Host "Examples: marketing agencies, restaurants, coffee shops, plumbers"
Write-Host ""
$searchTerm = Read-Host "Enter business type"

if (-not $searchTerm) {
    Write-Host "No search term entered. Exiting." -ForegroundColor Yellow
    exit
}

# Create Yelp search URL for GTA
$cleanTerm = $searchTerm.Replace(" ", "+")
$yelpUrl = "https://www.yelp.ca/search?find_desc=$cleanTerm&find_loc=Greater+Toronto+Area"

Write-Host ""
Write-Host "Searching for: $searchTerm" -ForegroundColor Yellow
Write-Host "URL: $yelpUrl" -ForegroundColor Gray
Write-Host ""

# Start the browser
Write-Host "Starting browser..." -ForegroundColor Yellow

try {
    # Start Chrome with chromedriver from current folder
    
    # Method 1: Set environment path to current folder
    $envPath = [Environment]::GetEnvironmentVariable("PATH", "Process")
    [Environment]::SetEnvironmentVariable("PATH", "$folder;$envPath", "Process")
    
    # Create Chrome options
    $options = New-Object OpenQA.Selenium.Chrome.ChromeOptions
    $options.AddArgument("start-maximized")
    
    # Start Chrome driver
    $driver = New-Object OpenQA.Selenium.Chrome.ChromeDriver($options)
    
    Write-Host "Browser started successfully" -ForegroundColor Green
    
    # Navigate to Yelp
    Write-Host "Loading Yelp page..." -ForegroundColor Yellow
    $driver.Navigate().GoToUrl($yelpUrl)
    
    # Wait for page to load
    Start-Sleep -Seconds 5
    
    # Check if blocked
    $pageText = $driver.PageSource
    if ($pageText.Contains("unusual traffic") -or $pageText.Contains("blocked") -or $pageText.Contains("CAPTCHA")) {
        Write-Host "YELP BLOCKED THE REQUEST" -ForegroundColor Red
        Write-Host "Yelp detects automated browsing and blocks it" -ForegroundColor Yellow
        Write-Host ""
        Write-Host "Try this instead:" -ForegroundColor Yellow
        Write-Host "1. Open this URL manually in your browser:" -ForegroundColor Cyan
        Write-Host "    $yelpUrl" -ForegroundColor Cyan
        Write-Host "2. Save the page as HTML" -ForegroundColor Cyan
        Write-Host "3. Use a different approach for data extraction" -ForegroundColor Cyan
        $driver.Quit()
        exit
    }
    
    # Extract business information
    Write-Host "Extracting business listings..." -ForegroundColor Yellow
    
    # Create list for results
    $businessList = @()
    $count = 0
    
    # Multiple approaches to find listings
    
    # Approach 1: Look for business containers
    $containers = $driver.FindElements([OpenQA.Selenium.By]::ClassName("container"))
    
    if ($containers.Count -eq 0) {
        # Approach 2: Look for any divs with business-like content
        $allDivs = $driver.FindElements([OpenQA.Selenium.By]::TagName("div"))
        $containers = @()
        
        foreach ($div in $allDivs) {
            $text = $div.Text
            if ($text -and $text.Length -gt 30) {
                if ($text -match "business" -or $text -match "restaurant" -or $text -match "service") {
                    $containers += $div
                }
            }
            if ($containers.Count -gt 20) { break }
        }
    }
    
    # Process found containers
    foreach ($container in $containers) {
        $count++
        if ($count -gt 15) { break }
        
        try {
            $text = $container.Text
            if (-not $text -or $text.Length -lt 10) { continue }
            
            # Extract business name (first line)
            $lines = $text -split "`n"
            $name = "Unknown Business $count"
            if ($lines.Count -gt 0) {
                $name = $lines[0].Trim()
                if ($name.Length -gt 50) {
                    $name = $name.Substring(0, 47) + "..."
                }
            }
            
            # Extract phone number
            $phone = "Not found"
            if ($text -match "(\d{3}) \d{3}-\d{4}") {
                $phone = $matches[0]
            }
            elseif ($text -match "\d{3}-\d{3}-\d{4}") {
                $phone = $matches[0]
            }
            
            # Extract address (look for Toronto area)
            $address = "Not found"
            $gtaCities = "Toronto,Mississauga,Brampton,Markham,Vaughan,Oakville,Burlington,Hamilton"
            foreach ($city in $gtaCities -split ",") {
                if ($text -match $city) {
                    # Find the line with city
                    foreach ($line in $lines) {
                        if ($line -match $city -and $line.Length -gt 5 -and $line.Length -lt 100) {
                            $address = $line.Trim()
                            break
                        }
                    }
                    if ($address -eq "Not found") {
                        $address = "$city, ON"
                    }
                    break
                }
            }
            
            # Create business record
            $business = [PSCustomObject]@{
                Name = $name
                Phone = $phone
                Address = $address
                SearchTerm = $searchTerm
                Source = "Yelp"
                DateScraped = Get-Date -Format "yyyy-MM-dd"
            }
            
            $businessList += $business
            Write-Host "Found: $name" -ForegroundColor Gray
        }
        catch {
            Write-Host "Error processing item $count" -ForegroundColor DarkYellow
        }
    }
    
    # Save results
    if ($businessList.Count -gt 0) {
        # Create filename
        $filename = "Yelp_GTA_$searchTerm_$(Get-Date -Format 'yyyyMMdd').csv"
        # Remove invalid characters
        $invalidChars = [IO.Path]::GetInvalidFileNameChars()
        foreach ($char in $invalidChars) {
            $filename = $filename.Replace($char, "")
        }
        
        # Save to CSV
        $businessList | Export-Csv -Path $filename -NoTypeInformation
        
        Write-Host ""
        Write-Host "SUCCESS" -ForegroundColor Green
        Write-Host "Saved $($businessList.Count) businesses to $filename" -ForegroundColor Cyan
        
        # Show sample
        Write-Host ""
        Write-Host "Sample of results:" -ForegroundColor Yellow
        $businessList | Select-Object -First 5 | Format-Table Name, Phone, Address -AutoSize
    }
    else {
        Write-Host ""
        Write-Host "No business data could be extracted" -ForegroundColor Red
        Write-Host ""
        Write-Host "This usually means:" -ForegroundColor Yellow
        Write-Host "1. Yelp blocked the automated request" -ForegroundColor Yellow
        Write-Host "2. The page structure changed" -ForegroundColor Yellow
        Write-Host "3. No results for your search term" -ForegroundColor Yellow
        Write-Host ""
        Write-Host "Try visiting manually: $yelpUrl" -ForegroundColor Cyan
    }
    
    # Close browser
    Write-Host ""
    Write-Host "Closing browser..." -ForegroundColor Gray
    $driver.Quit()
}
catch {
    Write-Host ""
    Write-Host "ERROR: $($_.Exception.Message)" -ForegroundColor Red
    
    if ($_.Exception.Message -contains "chromedriver" -or $_.Exception.Message -contains "driver") {
        Write-Host ""
        Write-Host "ChromeDriver troubleshooting:" -ForegroundColor Yellow
        Write-Host "1. Make sure chromedriver.exe is in: $folder" -ForegroundColor Yellow
        Write-Host "2. Close all Chrome windows and try again" -ForegroundColor Yellow
        Write-Host "3. Download matching ChromeDriver from chromedriver.chromium.org" -ForegroundColor Yellow
    }
}

Write-Host ""
Write-Host "Script finished" -ForegroundColor Cyan