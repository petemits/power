# COMPLETE WEBSITE EXTRACTOR
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
$TextFilePath = "WebsiteText_$timestamp.txt"
$ImageFolder = "WebsiteImages_$timestamp"

Write-Host "Creating output files..." -ForegroundColor Gray
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

Write-Host "4. Saving text to file..." -ForegroundColor Cyan
try {
    $textContent = "WEBSITE: $WebsiteUrl`r`n"
    $textContent += "TITLE: $pageTitle`r`n"
    $textContent += "DATE: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`r`n"
    $textContent += "`r`n" + ("=" * 80) + "`r`n"
    $textContent += "FULL PAGE TEXT:`r`n"
    $textContent += ("=" * 80) + "`r`n`r`n"
    $textContent += $pageText
    
    $textContent | Out-File -FilePath $TextFilePath -Encoding UTF8
    Write-Host "   Text saved to: $TextFilePath" -ForegroundColor Green
    
} catch {
    Write-Host "   ERROR: Cannot save text file" -ForegroundColor Red
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
    New-Item -ItemType Directory -Path $ImageFolder -Force | Out-Null
    
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
        TextFile = $TextFilePath
        ImageFolder = $ImageFolder
    }

    $excelPackage = $summaryData | Export-Excel -Path $OutputExcelPath -WorksheetName "Summary" -AutoSize -PassThru
    
    if ($contactData.Count -gt 0) {
        $contactData | Export-Excel -ExcelPackage $excelPackage -WorksheetName "Contact Information" -AutoSize
    }
    
    if ($imagesData.Count -gt 0) {
        $imagesData | Export-Excel -ExcelPackage $excelPackage -WorksheetName "Website Images" -AutoSize
    }
    
    if ($companyInfo.Count -gt 0) {
        $companyInfo | Export-Excel -ExcelPackage $excelPackage -WorksheetName "Company Information" -AutoSize
    }
    
    $websiteInfo = [PSCustomObject]@{
        URL = $WebsiteUrl
        Title = $pageTitle
        TextFilePath = $TextFilePath
        TextPreview = $pageText.Substring(0, [Math]::Min(500, $pageText.Length))
    }
    $websiteInfo | Export-Excel -ExcelPackage $excelPackage -WorksheetName "Website Details" -AutoSize
    
    Close-ExcelPackage $excelPackage
    
    Write-Host "   Excel file created" -ForegroundColor Green
    
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Green
    Write-Host "EXTRACTION COMPLETE" -ForegroundColor Green
    Write-Host "========================================" -ForegroundColor Green
    Write-Host ""
    Write-Host "FILES CREATED:" -ForegroundColor Cyan
    Write-Host "1. $OutputExcelPath" -ForegroundColor Yellow
    Write-Host "   Worksheets: Summary, Contact Information, Website Images, Company Information, Website Details" -ForegroundColor Gray
    Write-Host "2. $TextFilePath" -ForegroundColor Yellow
    Write-Host "   Contains all website text" -ForegroundColor Gray
    Write-Host "3. $ImageFolder\" -ForegroundColor Yellow
    Write-Host "   Folder for website images" -ForegroundColor Gray
    Write-Host ""
    Write-Host "DATA EXTRACTED:" -ForegroundColor Cyan
    Write-Host "  Contact items: $($contactData.Count)" -ForegroundColor Gray
    Write-Host "  Website images: $($imagesData.Count)" -ForegroundColor Gray
    Write-Host "  Company info items: $($companyInfo.Count)" -ForegroundColor Gray
    Write-Host ""
    
    $openText = Read-Host "Open text file in Notepad? (Y/N)"
    if ($openText -eq "Y" -or $openText -eq "y") {
        notepad $TextFilePath
    }
    
    $openExcel = Read-Host "Open Excel file? (Y/N)"
    if ($openExcel -eq "Y" -or $openExcel -eq "y") {
        Invoke-Item $OutputExcelPath
    }
    
} catch {
    Write-Host "   ERROR creating Excel: $_" -ForegroundColor Red
    
    $csvPath = "WebsiteData_$timestamp.csv"
    $summaryData | Export-Csv -Path $csvPath -NoTypeInformation
    Write-Host "   Data saved as CSV: $csvPath" -ForegroundColor Yellow
}

Write-Host ""
Write-Host "Done!" -ForegroundColor Gray