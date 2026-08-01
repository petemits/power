# FIXED WEBSITE EXTRACTOR WITH PROPER FILE CREATION
Write-Host "WEBSITE EXTRACTOR" -ForegroundColor Cyan
Write-Host "=================" -ForegroundColor Cyan
Write-Host ""

$WebsiteUrl = Read-Host "Enter website URL"
if ($WebsiteUrl -notmatch "^https?://") {
    $WebsiteUrl = "https://$WebsiteUrl"
}

Write-Host ""
Write-Host "Processing: $WebsiteUrl" -ForegroundColor Yellow
Write-Host ""

$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$OutputExcelPath = "WebsiteData_$timestamp.xlsx"
$TextFilePath = "WebsiteText_$timestamp.html"
$ImageFolder = "WebsiteImages_$timestamp"
$OutputFolder = "ExtractedData_$timestamp"

Write-Host "Creating output in folder: $OutputFolder" -ForegroundColor Gray
New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null

$OutputExcelPath = Join-Path $OutputFolder "WebsiteData.xlsx"
$TextFilePath = Join-Path $OutputFolder "WebsiteText.html"

Write-Host "Excel: $OutputExcelPath" -ForegroundColor Gray
Write-Host "Text: $TextFilePath" -ForegroundColor Gray
Write-Host "Images: $ImageFolder" -ForegroundColor Gray
Write-Host ""

$ChromeDriverPath = ".\chromedriver.exe"
if (-not (Test-Path $ChromeDriverPath)) {
    Write-Host "ERROR: chromedriver.exe not found!" -ForegroundColor Red
    Write-Host "Place chromedriver.exe in current folder" -ForegroundColor Yellow
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
    
    $Driver = New-Object OpenQA.Selenium.Chrome.ChromeDriver($service, $options)
    
    Write-Host "   Browser ready" -ForegroundColor Green
    
} catch {
    Write-Host "   ERROR: Cannot start browser" -ForegroundColor Red
    Write-Host "   Error: $_" -ForegroundColor Red
    exit
}

Write-Host "2. Loading website..." -ForegroundColor Cyan
try {
    $Driver.Navigate().GoToUrl($WebsiteUrl)
    Start-Sleep -Seconds 3
    
    $pageTitle = $Driver.Title
    Write-Host "   Loaded: $pageTitle" -ForegroundColor Green
    
} catch {
    Write-Host "   ERROR: Cannot load website" -ForegroundColor Red
    $Driver.Quit()
    exit
}

Write-Host "3. Extracting website text..." -ForegroundColor Cyan
$pageText = ""
$allText = @()

try {
    $bodyText = $Driver.FindElement([OpenQA.Selenium.By]::TagName("body")).Text
    $pageText = $bodyText
    
    $paragraphs = $Driver.FindElements([OpenQA.Selenium.By]::TagName("p"))
    foreach ($p in $paragraphs) {
        $text = $p.Text.Trim()
        if ($text.Length -gt 10) {
            $allText += $text
        }
    }
    
    Write-Host "   Text extracted" -ForegroundColor Green
    
} catch {
    Write-Host "   WARNING: Text extraction issue" -ForegroundColor Yellow
}

Write-Host "4. Saving text to HTML file..." -ForegroundColor Cyan
try {
    $htmlContent = @"
<!DOCTYPE html>
<html>
<head>
    <title>Website Text: $pageTitle</title>
    <meta charset="UTF-8">
    <style>
        body { font-family: Arial, sans-serif; margin: 40px; line-height: 1.6; }
        h1 { color: #333; border-bottom: 2px solid #0066cc; padding-bottom: 10px; }
        h2 { color: #555; margin-top: 30px; }
        .info { background: #f5f5f5; padding: 15px; border-left: 4px solid #0066cc; margin: 20px 0; }
        .text-content { white-space: pre-wrap; background: white; padding: 20px; border: 1px solid #ddd; }
        .url { color: #0066cc; font-weight: bold; }
        .timestamp { color: #666; font-size: 0.9em; }
    </style>
</head>
<body>
    <h1>Website Text Extraction</h1>
    
    <div class="info">
        <p><strong>Website URL:</strong> <span class="url">$WebsiteUrl</span></p>
        <p><strong>Page Title:</strong> $pageTitle</p>
        <p><strong>Extraction Date:</strong> <span class="timestamp">$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</span></p>
    </div>
    
    <h2>Full Page Text Content:</h2>
    <div class="text-content">
$([System.Web.HttpUtility]::HtmlEncode($pageText) -replace "`r`n", "<br>" -replace "`n", "<br>" -replace "`r", "<br>")
    </div>
</body>
</html>
"@

    $htmlContent | Out-File -FilePath $TextFilePath -Encoding UTF8
    Write-Host "   HTML file saved" -ForegroundColor Green
    
} catch {
    Write-Host "   ERROR: Cannot save HTML file" -ForegroundColor Red
}

Write-Host "5. Extracting contact information..." -ForegroundColor Cyan
$contactData = @()

try {
    $fullPageText = $Driver.FindElement([OpenQA.Selenium.By]::TagName("body")).Text
    
    $phonePatterns = @('\(\d{3}\) \d{3}-\d{4}', '\d{3}-\d{3}-\d{4}', '\d{3}\.\d{3}\.\d{4}', '\+\d{1,3} \d{3} \d{3} \d{4}')
    $phones = @()
    
    foreach ($pattern in $phonePatterns) {
        $matches = [regex]::Matches($fullPageText, $pattern)
        foreach ($match in $matches) {
            $phone = $match.Value
            if (-not $phones.Contains($phone)) {
                $phones += $phone
            }
        }
    }
    
    if ($phones.Count -gt 0) {
        foreach ($phone in $phones) {
            $contactData += [PSCustomObject]@{
                ContactType = "Phone"
                ContactValue = $phone
                Source = "Page text"
            }
        }
    }
    
    $emailPattern = '\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b'
    $emailMatches = [regex]::Matches($fullPageText, $emailPattern)
    $emails = @()
    
    foreach ($match in $emailMatches) {
        $email = $match.Value
        if ($email -notmatch 'noreply|no-reply|donotreply|do-not-reply|admin|webmaster') {
            if (-not $emails.Contains($email)) {
                $emails += $email
            }
        }
    }
    
    if ($emails.Count -gt 0) {
        foreach ($email in $emails) {
            $contactData += [PSCustomObject]@{
                ContactType = "Email"
                ContactValue = $email
                Source = "Page text"
            }
        }
    }
    
    $addressKeywords = @("address", "location", "office", "street", "avenue", "road", "city", "state", "zip")
    $addressLines = $fullPageText -split "`r`n|`n"
    
    foreach ($line in $addressLines) {
        $line = $line.Trim()
        if ($line.Length -gt 20 -and $line.Length -lt 200) {
            $hasAddressWord = $false
            foreach ($word in $addressKeywords) {
                if ($line -imatch $word) {
                    $hasAddressWord = $true
                    break
                }
            }
            
            if ($hasAddressWord -or ($line -match '\d+ [A-Za-z]+ (Street|St|Avenue|Ave|Road|Rd|Boulevard|Blvd)')) {
                $contactData += [PSCustomObject]@{
                    ContactType = "Address"
                    ContactValue = $line
                    Source = "Page text"
                }
            }
        }
    }
    
    Write-Host "   Found $($contactData.Count) contact items" -ForegroundColor Green
    
} catch {
    Write-Host "   WARNING: Contact extraction issue" -ForegroundColor Yellow
}

Write-Host "6. Extracting website images..." -ForegroundColor Cyan
$imagesData = @()

try {
    New-Item -ItemType Directory -Path (Join-Path $OutputFolder "images") -Force | Out-Null
    
    $allImages = $Driver.FindElements([OpenQA.Selenium.By]::TagName("img"))
    $imageCount = 0
    
    foreach ($image in $allImages) {
        try {
            $imageSrc = $image.GetAttribute("src")
            $imageAlt = $image.GetAttribute("alt")
            
            if ($imageSrc -ne "" -and $imageSrc -notmatch '^data:image') {
                $imageCount++
                
                $imageName = "image_$imageCount.jpg"
                if ($imageAlt -ne "") {
                    $cleanName = $imageAlt -replace '[^\w]', '_'
                    if ($cleanName.Length -gt 5) {
                        $imageName = "$cleanName.jpg"
                    }
                }
                
                $imagesData += [PSCustomObject]@{
                    ImageNumber = $imageCount
                    ImageSource = $imageSrc
                    AltText = $imageAlt
                    FileName = $imageName
                }
            }
        } catch { }
    }
    
    Write-Host "   Found $imageCount images" -ForegroundColor Green
    
} catch {
    Write-Host "   WARNING: Image extraction issue" -ForegroundColor Yellow
}

Write-Host "7. Extracting company information..." -ForegroundColor Cyan
$companyInfo = @()

try {
    $h1Elements = $Driver.FindElements([OpenQA.Selenium.By]::TagName("h1"))
    if ($h1Elements.Count -gt 0) {
        foreach ($h1 in $h1Elements) {
            $text = $h1.Text.Trim()
            if ($text -ne "") {
                $companyInfo += [PSCustomObject]@{
                    InfoType = "Company Name"
                    InfoValue = $text
                }
            }
        }
    }
    
    $metaElements = $Driver.FindElements([OpenQA.Selenium.By]::CssSelector("meta[name='description']"))
    if ($metaElements.Count -gt 0) {
        foreach ($meta in $metaElements) {
            $content = $meta.GetAttribute("content")
            if ($content -ne "") {
                $companyInfo += [PSCustomObject]@{
                    InfoType = "Description"
                    InfoValue = $content
                }
            }
        }
    }
    
    Write-Host "   Company info extracted" -ForegroundColor Green
    
} catch {
    Write-Host "   WARNING: Company info extraction issue" -ForegroundColor Yellow
}

Write-Host "8. Closing browser..." -ForegroundColor Cyan
$Driver.Quit()
Write-Host "   Browser closed" -ForegroundColor Green

Write-Host "9. Creating Excel file..." -ForegroundColor Cyan
try {
    $summaryData = [PSCustomObject]@{
        WebsiteURL = $WebsiteUrl
        PageTitle = $pageTitle
        ExtractionDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        TotalContacts = $contactData.Count
        TotalImages = $imagesData.Count
        TextFile = "WebsiteText.html"
        ImageFolder = "images"
    }

    if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
        Write-Host "   Installing ImportExcel module..." -ForegroundColor Yellow
        Install-Module -Name ImportExcel -Force -Scope CurrentUser
    }

    if ($imagesData.Count -eq 0) {
        $imagesData = @([PSCustomObject]@{
            Status = "No images found on this webpage"
        })
    }

    $summaryData | Export-Excel -Path $OutputExcelPath -WorksheetName "Summary" -AutoSize
    Write-Host "   Created Summary worksheet" -ForegroundColor Green

    if ($contactData.Count -gt 0) {
        $contactData | Export-Excel -Path $OutputExcelPath -WorksheetName "Contact Information" -AutoSize
        Write-Host "   Created Contact Information worksheet" -ForegroundColor Green
    }

    $imagesData | Export-Excel -Path $OutputExcelPath -WorksheetName "Website Images" -AutoSize
    Write-Host "   Created Website Images worksheet" -ForegroundColor Green

    if ($companyInfo.Count -gt 0) {
        $companyInfo | Export-Excel -Path $OutputExcelPath -WorksheetName "Company Information" -AutoSize
        Write-Host "   Created Company Information worksheet" -ForegroundColor Green
    }

    $websiteInfo = [PSCustomObject]@{
        URL = $WebsiteUrl
        Title = $pageTitle
        TextFile = "WebsiteText.html"
        TextPreview = if ($pageText.Length -gt 500) { $pageText.Substring(0, 500) + "..." } else { $pageText }
    }
    $websiteInfo | Export-Excel -Path $OutputExcelPath -WorksheetName "Website Details" -AutoSize
    Write-Host "   Created Website Details worksheet" -ForegroundColor Green

    Write-Host ""
    Write-Host "========================================" -ForegroundColor Green
    Write-Host "EXTRACTION COMPLETE" -ForegroundColor Green
    Write-Host "========================================" -ForegroundColor Green
    Write-Host ""
    Write-Host "ALL FILES CREATED IN FOLDER:" -ForegroundColor Cyan
    Write-Host "  Folder: $OutputFolder" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "Files inside folder:" -ForegroundColor Cyan
    Write-Host "  1. WebsiteData.xlsx (Excel with all data)" -ForegroundColor Gray
    Write-Host "  2. WebsiteText.html (Text content in Chrome)" -ForegroundColor Gray
    Write-Host "  3. images\ (Folder for website images)" -ForegroundColor Gray
    Write-Host ""
    Write-Host "Excel worksheets created:" -ForegroundColor Cyan
    Write-Host "  - Summary" -ForegroundColor Gray
    if ($contactData.Count -gt 0) { Write-Host "  - Contact Information" -ForegroundColor Gray }
    Write-Host "  - Website Images" -ForegroundColor Gray
    if ($companyInfo.Count -gt 0) { Write-Host "  - Company Information" -ForegroundColor Gray }
    Write-Host "  - Website Details" -ForegroundColor Gray
    Write-Host ""
    Write-Host "Data extracted:" -ForegroundColor Cyan
    Write-Host "  Contact items: $($contactData.Count)" -ForegroundColor Gray
    Write-Host "  Website images: $($imagesData.Count)" -ForegroundColor Gray
    Write-Host "  Company info items: $($companyInfo.Count)" -ForegroundColor Gray
    Write-Host ""

    $openFolder = Read-Host "Open output folder in File Explorer? (Y/N)"
    if ($openFolder -eq "Y" -or $openFolder -eq "y") {
        explorer $OutputFolder
        Write-Host "   Opening folder..." -ForegroundColor Green
    }

    $openExcel = Read-Host "Open Excel file? (Y/N)"
    if ($openExcel -eq "Y" -or $openExcel -eq "y") {
        Invoke-Item $OutputExcelPath
        Write-Host "   Opening Excel file..." -ForegroundColor Green
    }

    $openText = Read-Host "Open text file in Chrome? (Y/N)"
    if ($openText -eq "Y" -or $openText -eq "y") {
        Start-Process "chrome.exe" -ArgumentList $TextFilePath
        Write-Host "   Opening text in Chrome..." -ForegroundColor Green
    }
    
} catch {
    Write-Host "   ERROR creating Excel: $_" -ForegroundColor Red
    
    $csvPath = Join-Path $OutputFolder "WebsiteData.csv"
    $summaryData | Export-Csv -Path $csvPath -NoTypeInformation
    Write-Host "   Summary saved as CSV: $csvPath" -ForegroundColor Yellow
    
    if ($contactData.Count -gt 0) {
        $contactCsv = Join-Path $OutputFolder "Contacts.csv"
        $contactData | Export-Csv -Path $contactCsv -NoTypeInformation
        Write-Host "   Contacts saved as CSV: $contactCsv" -ForegroundColor Yellow
    }
}

Write-Host ""
Write-Host "Your files are saved in: $OutputFolder" -ForegroundColor Yellow
Write-Host "Check that folder for all extracted data." -ForegroundColor Yellow
Write-Host ""
Write-Host "Done!" -ForegroundColor Gray