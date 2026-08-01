# BUSINESS PROFILE EXTRACTOR - COMPLETE VERSION
Write-Host "BUSINESS PROFILE EXTRACTOR" -ForegroundColor Cyan
Write-Host "==========================" -ForegroundColor Cyan
Write-Host ""

$WebsiteUrl = Read-Host "Enter website URL (e.g., https://www.example.com)"
if ($WebsiteUrl -notmatch "^https?://") {
    $WebsiteUrl = "https://$WebsiteUrl"
}

Write-Host ""
Write-Host "Processing website: $WebsiteUrl" -ForegroundColor Yellow
Write-Host ""

$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$OutputExcelPath = "BusinessProfile_$timestamp.xlsx"
$ScreenshotPath = "screenshot_$timestamp.png"

Write-Host "Output file: $OutputExcelPath" -ForegroundColor Gray
Write-Host ""

$ChromeDriverPath = ".\chromedriver.exe"
if (-not (Test-Path $ChromeDriverPath)) {
    Write-Host "ERROR: chromedriver.exe not found in current folder!" -ForegroundColor Red
    Write-Host "Please place chromedriver.exe in: $(Get-Location)" -ForegroundColor Yellow
    exit
}

Write-Host "1. Starting browser..." -ForegroundColor Cyan
try {
    Add-Type -Path ".\WebDriver.dll"
    
    $service = [OpenQA.Selenium.Chrome.ChromeDriverService]::CreateDefaultService((Get-Location).Path, "chromedriver.exe")
    $service.HideCommandPromptWindow = $true
    
    $options = New-Object OpenQA.Selenium.Chrome.ChromeOptions
    $options.AddArgument("--headless")
    $options.AddArgument("--no-sandbox")
    $options.AddArgument("--window-size=1920,1080")
    
    $Driver = New-Object OpenQA.Selenium.Chrome.ChromeDriver($service, $options)
    
    Write-Host "   Browser started successfully" -ForegroundColor Green
    
} catch {
    Write-Host "   ERROR: Failed to start browser" -ForegroundColor Red
    Write-Host "   Details: $_" -ForegroundColor Red
    exit
}

Write-Host "2. Loading website..." -ForegroundColor Cyan
try {
    $Driver.Navigate().GoToUrl($WebsiteUrl)
    Start-Sleep -Seconds 3
    
    $pageTitle = $Driver.Title
    Write-Host "   Website loaded: $pageTitle" -ForegroundColor Green
    
} catch {
    Write-Host "   ERROR: Failed to load website" -ForegroundColor Red
    $Driver.Quit()
    exit
}

Write-Host "3. Taking screenshot..." -ForegroundColor Cyan
try {
    $screenshot = $Driver.GetScreenshot()
    $screenshot.SaveAsFile($ScreenshotPath, [OpenQA.Selenium.ScreenshotImageFormat]::Png)
    Write-Host "   Screenshot saved: $ScreenshotPath" -ForegroundColor Green
} catch {
    Write-Host "   WARNING: Could not take screenshot" -ForegroundColor Yellow
    $ScreenshotPath = $null
}

Write-Host "4. Extracting business data..." -ForegroundColor Cyan

$pageText = $Driver.FindElement([OpenQA.Selenium.By]::TagName("body")).Text
$pageHTML = $Driver.PageSource

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

$phoneNumbers = @()
$phonePatterns = @('\(\d{3}\) \d{3}-\d{4}', '\d{3}-\d{3}-\d{4}', '\d{3}\.\d{3}\.\d{4}')
foreach ($pattern in $phonePatterns) {
    $matches = [regex]::Matches($pageText, $pattern)
    foreach ($match in $matches) {
        if (-not $phoneNumbers.Contains($match.Value)) {
            $phoneNumbers += $match.Value
        }
    }
}

$emails = @()
$emailPattern = '\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b'
$emailMatches = [regex]::Matches($pageText, $emailPattern)
foreach ($match in $emailMatches) {
    $email = $match.Value
    if ($email -notmatch 'noreply|no-reply|donotreply|do-not-reply') {
        if (-not $emails.Contains($email)) {
            $emails += $email
        }
    }
}

$businessDescription = "Not found"
try {
    $metaElements = $Driver.FindElements([OpenQA.Selenium.By]::CssSelector("meta[name='description']"))
    if ($metaElements.Count -gt 0) {
        $desc = $metaElements[0].GetAttribute("content")
        if ($desc -ne "" -and $desc.Length -gt 20) {
            $businessDescription = $desc.Substring(0, [Math]::Min(500, $desc.Length))
        }
    }
} catch { }

Write-Host "5. Closing browser..." -ForegroundColor Cyan
$Driver.Quit()
Write-Host "   Browser closed" -ForegroundColor Green

Write-Host "6. Creating Excel file with embedded content..." -ForegroundColor Cyan
try {
    $businessData = [PSCustomObject]@{
        WebsiteURL = $WebsiteUrl
        CompanyName = $companyName
        PhoneNumbers = if ($phoneNumbers.Count -gt 0) { $phoneNumbers -join "; " } else { "Not found" }
        EmailAddresses = if ($emails.Count -gt 0) { $emails -join "; " } else { "Not found" }
        BusinessDescription = $businessDescription
        PageTitle = $pageTitle
        ExtractionDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    }

    $businessData | Export-Excel -Path $OutputExcelPath -WorksheetName "Business Profile" -AutoSize -TableName "BusinessData"
    Write-Host "   Excel file created with business data" -ForegroundColor Green

    if ($ScreenshotPath -and (Test-Path $ScreenshotPath)) {
        try {
            $excel = New-Object -ComObject Excel.Application
            $excel.Visible = $false
            $workbook = $excel.Workbooks.Open((Get-Item $OutputExcelPath).FullName)
            $worksheet = $workbook.Worksheets.Item("Business Profile")
            
            $lastRow = $worksheet.UsedRange.Rows.Count + 2
            $worksheet.Cells.Item($lastRow, 1) = "WEB PAGE SCREENSHOT:"
            
            $picture = $worksheet.Shapes.AddPicture(
                (Get-Item $ScreenshotPath).FullName,
                $false, $true,
                $worksheet.Cells.Item($lastRow, 2).Left,
                $worksheet.Cells.Item($lastRow, 2).Top,
                800, 450
            )
            
            $lastRow = $lastRow + 30
            $worksheet.Cells.Item($lastRow, 1) = "FULL PAGE HTML:"
            $worksheet.Cells.Item($lastRow, 2) = $pageHTML
            
            $worksheet.Columns.Item(2).ColumnWidth = 100
            $worksheet.Rows.Item($lastRow).RowHeight = 300
            
            $workbook.Save()
            $workbook.Close()
            $excel.Quit()
            
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($picture) | Out-Null
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($worksheet) | Out-Null
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
            
            Write-Host "   Screenshot embedded in Excel" -ForegroundColor Green
            Write-Host "   Full HTML embedded in Excel" -ForegroundColor Green
        } catch {
            Write-Host "   WARNING: Could not embed screenshot (Excel COM issue)" -ForegroundColor Yellow
            Write-Host "   Screenshot saved separately: $ScreenshotPath" -ForegroundColor Gray
        }
    }

    Write-Host ""
    Write-Host "========================================" -ForegroundColor Green
    Write-Host "BUSINESS PROFILE EXTRACTION COMPLETE!" -ForegroundColor Green
    Write-Host "========================================" -ForegroundColor Green
    Write-Host ""
    Write-Host "EXTRACTED BUSINESS DATA:" -ForegroundColor Cyan
    Write-Host "  Website: $WebsiteUrl" -ForegroundColor Gray
    Write-Host "  Company Name: $companyName" -ForegroundColor Gray
    Write-Host "  Phone Numbers: $(if ($phoneNumbers.Count -gt 0) { $phoneNumbers -join ', ' } else { 'Not found' })" -ForegroundColor Gray
    Write-Host "  Email Addresses: $(if ($emails.Count -gt 0) { $emails -join ', ' } else { 'Not found' })" -ForegroundColor Gray
    Write-Host "  Description: $($businessDescription.Substring(0, [Math]::Min(100, $businessDescription.Length)))..." -ForegroundColor Gray
    Write-Host "  Page Title: $pageTitle" -ForegroundColor Gray
    Write-Host ""
    Write-Host "OUTPUT FILES:" -ForegroundColor Cyan
    Write-Host "  1. $OutputExcelPath (Excel with embedded content)" -ForegroundColor Yellow
    if ($ScreenshotPath -and (Test-Path $ScreenshotPath)) {
        Write-Host "  2. $ScreenshotPath (web page screenshot)" -ForegroundColor Yellow
    }
    Write-Host ""

    $openExcel = Read-Host "Open Excel file now? (Y/N)"
    if ($openExcel -eq "Y" -or $openExcel -eq "y") {
        Invoke-Item $OutputExcelPath
        Write-Host "Opening Excel file..." -ForegroundColor Green
    }

} catch {
    Write-Host "   ERROR creating Excel: $_" -ForegroundColor Red
    
    $csvPath = "BusinessProfile_$timestamp.csv"
    $businessData | Export-Csv -Path $csvPath -NoTypeInformation
    Write-Host "   Data saved as CSV: $csvPath" -ForegroundColor Yellow
    
    if ($ScreenshotPath -and (Test-Path $ScreenshotPath)) {
        Write-Host "   Screenshot saved: $ScreenshotPath" -ForegroundColor Yellow
    }
    
    $htmlPath = "webpage_$timestamp.html"
    $pageHTML | Out-File -FilePath $htmlPath -Encoding UTF8
    Write-Host "   HTML saved: $htmlPath" -ForegroundColor Yellow
}

Write-Host ""
Write-Host "Extraction complete. Press any key to exit..." -ForegroundColor Gray
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")