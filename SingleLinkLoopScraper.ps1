<#
    SINGLE LINK SCRAPER WITH LOOP CONTROL
    Processes one link at a time, waits for your command
#>

# Function to process a single link
function Process-SingleLink {
    param(
        [object]$Driver,
        [object]$Link,
        [string]$ScreenshotFolder,
        [int]$LinkNumber
    )
    
    $result = $null
    
    try {
        Write-Host "`n[$LinkNumber] Opening: $($Link.Title)" -ForegroundColor Cyan
        
        # Save original window
        $originalWindow = $Driver.CurrentWindowHandle
        
        # Open link in new tab
        $Driver.ExecuteScript("window.open(arguments[0]);", $Link.URL)
        Start-Sleep -Seconds 2
        
        # Switch to new tab
        $Driver.SwitchTo().Window($Driver.WindowHandles[-1])
        Start-Sleep -Seconds 3
        
        # Get page info
        $pageTitle = $Driver.Title
        Write-Host "  Page title: $pageTitle" -ForegroundColor Gray
        
        # Take screenshot
        $safeFileName = $pageTitle -replace '[^\w\-_\. ]', '_' -replace '\s+', '_'
        $screenshotPath = Join-Path $ScreenshotFolder "${LinkNumber}_${safeFileName}.png"
        
        try {
            $screenshot = $Driver.GetScreenshot()
            $screenshot.SaveAsFile($screenshotPath, [OpenQA.Selenium.ScreenshotImageFormat]::Png)
            Write-Host "  Screenshot saved." -ForegroundColor Green
        } catch {
            Write-Host "  Could not save screenshot." -ForegroundColor Yellow
            $screenshotPath = "Failed"
        }
        
        # Extract data
        $pageText = $Driver.FindElement([OpenQA.Selenium.By]::TagName("body")).Text
        
        $email = ""
        if ($pageText -match "[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}") {
            $email = $Matches[0]
            Write-Host "  Email found." -ForegroundColor Green
        }
        
        $phone = ""
        if ($pageText -match "\(?\d{3}\)?[-.\s]?\d{3}[-.\s]?\d{4}") {
            $phone = $Matches[0]
            Write-Host "  Phone found." -ForegroundColor Green
        }
        
        # Close tab and return
        $Driver.Close()
        $Driver.SwitchTo().Window($originalWindow)
        
        # Create result object
        $result = [PSCustomObject]@{
            Number = $LinkNumber
            Title = $Link.Title
            URL = $Link.URL
            PageTitle = $pageTitle
            ScreenshotFile = Split-Path $screenshotPath -Leaf
            Email = $email
            Phone = $phone
        }
        
    } catch {
        Write-Host "  Error processing link: $($_.Exception.Message)" -ForegroundColor Red
        
        # Try to return to original window
        try {
            if ($Driver.WindowHandles.Count -gt 1) {
                $Driver.Close()
                $Driver.SwitchTo().Window($Driver.WindowHandles[0])
            }
        } catch {
            # Ignore cleanup errors
        }
    }
    
    return $result
}

# Main script
Clear-Host
Write-Host "SINGLE LINK SCRAPER WITH LOOP" -ForegroundColor Cyan
Write-Host "==============================" -ForegroundColor Cyan
Write-Host ""

# Configuration
$folder = "C:\Users\user\power\"
$screenshotFolder = "$env:USERPROFILE\Desktop\WebsiteScreenshots\"
New-Item -ItemType Directory -Force -Path $screenshotFolder | Out-Null

# File checks
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

# Load modules
Add-Type -Path "$folder\WebDriver.dll"

Import-Module ImportExcel -ErrorAction SilentlyContinue
if (-not (Get-Module -Name ImportExcel)) {
    Install-Module -Name ImportExcel -Force -Scope CurrentUser
    Import-Module ImportExcel
}

# Get search term
$searchTerm = Read-Host "Enter search term (or press Enter for 'technology')"
if ($searchTerm -eq "") {
    $searchTerm = "technology"
}

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

# Find all result links
Write-Host "Finding search results..." -ForegroundColor Green
$allLinks = $driver.FindElements([OpenQA.Selenium.By]::TagName("a"))
$availableLinks = @()

foreach ($link in $allLinks) {
    $href = $link.GetAttribute("href")
    $text = $link.Text
    
    if ($href -and $href.StartsWith("http") -and -not $href.Contains("google.com") -and $text.Length -gt 5) {
        $availableLinks += [PSCustomObject]@{
            URL = $href
            Title = $text
            Processed = $false
        }
    }
}

if ($availableLinks.Count -eq 0) {
    Write-Host "No links found!" -ForegroundColor Red
    $driver.Quit()
    pause
    exit
}

Write-Host "Found $($availableLinks.Count) links." -ForegroundColor Green

# Main processing loop
$results = @()
$continueProcessing = $true
$currentIndex = 0

while ($continueProcessing) {
    # Find next unprocessed link
    $nextLink = $null
    $nextIndex = -1
    
    for ($i = 0; $i -lt $availableLinks.Count; $i++) {
        if (-not $availableLinks[$i].Processed) {
            $nextLink = $availableLinks[$i]
            $nextIndex = $i
            break
        }
    }
    
    if ($nextLink -eq $null) {
        Write-Host "`nAll links have been processed!" -ForegroundColor Cyan
        $continueProcessing = $false
        break
    }
    
    # Process the link
    $result = Process-SingleLink -Driver $driver -Link $nextLink -ScreenshotFolder $screenshotFolder -LinkNumber ($nextIndex + 1)
    
    if ($result -ne $null) {
        $results += $result
        $availableLinks[$nextIndex].Processed = $true
        
        # Show what we found
        Write-Host "`nExtracted from this link:" -ForegroundColor Gray
        if ($result.Email) { Write-Host "  Email: $($result.Email)" -ForegroundColor Gray }
        if ($result.Phone) { Write-Host "  Phone: $($result.Phone)" -ForegroundColor Gray }
    } else {
        Write-Host "Failed to process this link. Marking as processed." -ForegroundColor Yellow
        $availableLinks[$nextIndex].Processed = $true
    }
    
    # Ask user what to do next
    if ($continueProcessing) {
        Write-Host "`n========================================" -ForegroundColor Cyan
        Write-Host "What would you like to do next?" -ForegroundColor Cyan
        Write-Host "  1) Process next link"
        Write-Host "  2) Process ALL remaining links automatically"
        Write-Host "  3) Save results and exit"
        Write-Host "  4) Exit without saving"
        Write-Host "========================================" -ForegroundColor Cyan
        
        $choice = Read-Host "`nEnter choice (1-4)"
        
        switch ($choice) {
            "1" {
                # Continue with single link mode
                Write-Host "Continuing with next link..." -ForegroundColor Green
            }
            "2" {
                # Process all remaining links automatically
                Write-Host "`nProcessing all remaining links automatically..." -ForegroundColor Green
                
                for ($i = 0; $i -lt $availableLinks.Count; $i++) {
                    if (-not $availableLinks[$i].Processed) {
                        $result = Process-SingleLink -Driver $driver -Link $availableLinks[$i] -ScreenshotFolder $screenshotFolder -LinkNumber ($i + 1)
                        
                        if ($result -ne $null) {
                            $results += $result
                        }
                        
                        $availableLinks[$i].Processed = $true
                        Write-Host "  Processed link $($i+1)/$($availableLinks.Count)" -ForegroundColor Gray
                    }
                }
                
                $continueProcessing = $false
            }
            "3" {
                $continueProcessing = $false
            }
            "4" {
                Write-Host "Exiting without saving..." -ForegroundColor Yellow
                $driver.Quit()
                exit
            }
            default {
                Write-Host "Invalid choice. Exiting..." -ForegroundColor Red
                $continueProcessing = $false
            }
        }
    }
}

# Save results to Excel if we have any
if ($results.Count -gt 0) {
    Write-Host "`nSaving results to Excel..." -ForegroundColor Green
    
    $timestamp = Get-Date -Format "yyyyMMdd_HHmm"
    $excelFile = "$env:USERPROFILE\Desktop\ScrapedData_$timestamp.xlsx"
    
    try {
        $results | Export-Excel -Path $excelFile -WorksheetName "Results" -AutoSize -BoldTopRow
        Write-Host "✓ Results saved to: $excelFile" -ForegroundColor Green
        
        # Show summary
        Write-Host "`n========================================" -ForegroundColor Green
        Write-Host "           PROCESSING SUMMARY" -ForegroundColor Green
        Write-Host "========================================" -ForegroundColor Green
        Write-Host "Total links processed: $($results.Count)" -ForegroundColor Gray
        Write-Host "Screenshots saved to: $screenshotFolder" -ForegroundColor Gray
        Write-Host "Excel file: $excelFile" -ForegroundColor Gray
        
        # Ask to open files
        Write-Host ""
        $openExcel = Read-Host "Open Excel file? (Y/N)"
        if ($openExcel -eq "Y" -or $openExcel -eq "y") {
            Start-Process $excelFile
        }
        
        $openFolder = Read-Host "Open screenshot folder? (Y/N)"
        if ($openFolder -eq "Y" -or $openFolder -eq "y") {
            Start-Process $screenshotFolder
        }
        
    } catch {
        Write-Host "ERROR saving Excel: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "Displaying results instead:" -ForegroundColor Yellow
        $results | Format-Table -AutoSize
    }
} else {
    Write-Host "No results to save." -ForegroundColor Yellow
}

# Cleanup
$driver.Quit()
Write-Host "`nBrowser closed. Script complete." -ForegroundColor Gray
Write-Host "Press any key to exit..." -ForegroundColor Gray
pause