# ENHANCED BUSINESS PROFILE EXTRACTOR
# Extracts: Business data, menu navigation, images, text content
Write-Host "ENHANCED BUSINESS PROFILE EXTRACTOR" -ForegroundColor Cyan
Write-Host "===================================" -ForegroundColor Cyan
Write-Host ""

$WebsiteUrl = Read-Host "Enter website URL (e.g., https://www.example.com)"
if ($WebsiteUrl -notmatch "^https?://") {
    $WebsiteUrl = "https://$WebsiteUrl"
}

Write-Host ""
Write-Host "Processing website: $WebsiteUrl" -ForegroundColor Yellow
Write-Host ""

$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$OutputExcelPath = "EnhancedProfile_$timestamp.xlsx"
$ScreenshotPath = "screenshot_$timestamp.png"
$ImageFolder = "website_images_$timestamp"

Write-Host "Output file: $OutputExcelPath" -ForegroundColor Gray
Write-Host ""

$ChromeDriverPath = ".\chromedriver.exe"
if (-not (Test-Path $ChromeDriverPath)) {
    Write-Host "ERROR: chromedriver.exe not found!" -ForegroundColor Red
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
    
    Write-Host "   Browser started" -ForegroundColor Green
    
} catch {
    Write-Host "   ERROR: $_" -ForegroundColor Red
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

Write-Host "3. Taking homepage screenshot..." -ForegroundColor Cyan
try {
    $screenshot = $Driver.GetScreenshot()
    $screenshot.SaveAsFile($ScreenshotPath, [OpenQA.Selenium.ScreenshotImageFormat]::Png)
    Write-Host "   Screenshot saved: $ScreenshotPath" -ForegroundColor Green
} catch {
    Write-Host "   WARNING: Could not take screenshot" -ForegroundColor Yellow
    $ScreenshotPath = $null
}

Write-Host "4. Extracting menu buttons and navigation..." -ForegroundColor Cyan
$menuData = @()
$buttonsData = @()
$linksData = @()

try {
    $menuSelectors = @("nav", "ul.nav", "div.nav", "header", ".menu", ".navigation", "#menu", "#nav")
    
    foreach ($selector in $menuSelectors) {
        try {
            $menus = $Driver.FindElements([OpenQA.Selenium.By]::CssSelector($selector))
            if ($menus.Count -gt 0) {
                Write-Host "   Found menu with selector: $selector" -ForegroundColor Gray
                break
            }
        } catch { }
    }
    
    $allButtons = $Driver.FindElements([OpenQA.Selenium.By]::TagName("button"))
    $allLinks = $Driver.FindElements([OpenQA.Selenium.By]::TagName("a"))
    $allInputs = $Driver.FindElements([OpenQA.Selenium.By]::TagName("input"))
    
    $buttonCount = 0
    foreach ($button in $allButtons) {
        try {
            $buttonText = $button.Text.Trim()
            $buttonType = $button.GetAttribute("type")
            $buttonClass = $button.GetAttribute("class")
            
            if ($buttonText -ne "" -or $buttonType -ne "") {
                $buttonCount++
                $buttonsData += [PSCustomObject]@{
                    ElementType = "Button"
                    Text = $buttonText
                    Type = $buttonType
                    Class = $buttonClass
                    Order = $buttonCount
                }
            }
        } catch { }
    }
    
    $linkCount = 0
    foreach ($link in $allLinks) {
        try {
            $linkText = $link.Text.Trim()
            $linkHref = $link.GetAttribute("href")
            $linkClass = $link.GetAttribute("class")
            
            if ($linkText -ne "" -and $linkHref -ne "") {
                $linkCount++
                $linksData += [PSCustomObject]@{
                    ElementType = "Link"
                    Text = $linkText
                    URL = $linkHref
                    Class = $linkClass
                    Order = $linkCount
                }
            }
        } catch { }
    }
    
    Write-Host "   Found $buttonCount buttons and $linkCount links" -ForegroundColor Green
    
} catch {
    Write-Host "   WARNING: Could not extract menu data" -ForegroundColor Yellow
}

Write-Host "5. Extracting images..." -ForegroundColor Cyan
$imagesData = @()

try {
    New-Item -ItemType Directory -Path $ImageFolder -Force | Out-Null
    
    $allImages = $Driver.FindElements([OpenQA.Selenium.By]::TagName("img"))
    $imageCount = 0
    
    foreach ($image in $allImages) {
        try {
            $imageSrc = $image.GetAttribute("src")
            $imageAlt = $image.GetAttribute("alt")
            $imageClass = $image.GetAttribute("class")
            
            if ($imageSrc -ne "") {
                $imageCount++
                
                $imageName = "image_$imageCount"
                if ($imageAlt -ne "") {
                    $imageName = $imageAlt -replace '[^\w]', '_'
                }
                
                $imageExtension = ".jpg"
                if ($imageSrc -match '\.(png|gif|svg|webp)$') {
                    $imageExtension = ".$($matches[1])"
                }
                
                $imageFileName = "$imageName$imageExtension"
                $imageFilePath = Join-Path $ImageFolder $imageFileName
                
                try {
                    if ($imageSrc.StartsWith("data:image")) {
                        $imagesData += [PSCustomObject]@{
                            ImageNumber = $imageCount
                            AltText = $imageAlt
                            SourceType = "Data URI"
                            FilePath = "Embedded in HTML"
                        }
                    } else {
                        $imagesData += [PSCustomObject]@{
                            ImageNumber = $imageCount
                            AltText = $imageAlt
                            Source = $imageSrc
                            Class = $imageClass
                            FilePath = $imageFilePath
                        }
                    }
                } catch { }
            }
        } catch { }
    }
    
    Write-Host "   Found $imageCount images" -ForegroundColor Green
    
} catch {
    Write-Host "   WARNING: Could not extract images" -ForegroundColor Yellow
}

Write-Host "6. Extracting text content sections..." -ForegroundColor Cyan
$textSections = @()

try {
    $sectionSelectors = @("section", "article", "div.content", "div.section", "main", ".text-content", "#content")
    
    $sectionCount = 0
    foreach ($selector in $sectionSelectors) {
        try {
            $sections = $Driver.FindElements([OpenQA.Selenium.By]::CssSelector($selector))
            foreach ($section in $sections) {
                try {
                    $sectionText = $section.Text.Trim()
                    if ($sectionText.Length -gt 50 -and $sectionText.Length -lt 5000) {
                        $sectionCount++
                        $sectionClass = $section.GetAttribute("class")
                        $sectionId = $section.GetAttribute("id")
                        
                        $textSections += [PSCustomObject]@{
                            SectionNumber = $sectionCount
                            Selector = $selector
                            Class = $sectionClass
                            ID = $sectionId
                            TextContent = $sectionText.Substring(0, [Math]::Min(1000, $sectionText.Length))
                            CharacterCount = $sectionText.Length
                        }
                    }
                } catch { }
            }
        } catch { }
    }
    
    Write-Host "   Found $sectionCount text sections" -ForegroundColor Green
    
} catch {
    Write-Host "   WARNING: Could not extract text sections" -ForegroundColor Yellow
}

Write-Host "7. Extracting basic business data..." -ForegroundColor Cyan
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
    if ($email -notmatch 'noreply|no-reply') {
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

Write-Host "8. Closing browser..." -ForegroundColor Cyan
$Driver.Quit()
Write-Host "   Browser closed" -ForegroundColor Green

Write-Host "9. Creating enhanced Excel report..." -ForegroundColor Cyan
try {
    $businessData = [PSCustomObject]@{
        WebsiteURL = $WebsiteUrl
        CompanyName = $companyName
        PhoneNumbers = if ($phoneNumbers.Count -gt 0) { $phoneNumbers -join "; " } else { "Not found" }
        EmailAddresses = if ($emails.Count -gt 0) { $emails -join "; " } else { "Not found" }
        BusinessDescription = $businessDescription
        PageTitle = $pageTitle
        ExtractionDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        TotalButtonsFound = $buttonsData.Count
        TotalLinksFound = $linksData.Count
        TotalImagesFound = $imagesData.Count
        TotalSectionsFound = $textSections.Count
    }

    $excelPackage = $businessData | Export-Excel -Path $OutputExcelPath -WorksheetName "Summary" -AutoSize -PassThru
    
    if ($buttonsData.Count -gt 0) {
        $buttonsData | Export-Excel -ExcelPackage $excelPackage -WorksheetName "Buttons" -AutoSize
    }
    
    if ($linksData.Count -gt 0) {
        $linksData | Export-Excel -ExcelPackage $excelPackage -WorksheetName "Links" -AutoSize
    }
    
    if ($imagesData.Count -gt 0) {
        $imagesData | Export-Excel -ExcelPackage $excelPackage -WorksheetName "Images" -AutoSize
    }
    
    if ($textSections.Count -gt 0) {
        $textSections | Export-Excel -ExcelPackage $excelPackage -WorksheetName "Text Sections" -AutoSize
    }
    
    $htmlSheet = [PSCustomObject]@{
        WebsiteURL = $WebsiteUrl
        HTML_Content = $pageHTML
    }
    $htmlSheet | Export-Excel -ExcelPackage $excelPackage -WorksheetName "Full HTML" -AutoSize
    
    Close-ExcelPackage $excelPackage
    
    Write-Host "   Excel report created with multiple worksheets" -ForegroundColor Green
    
    if ($ScreenshotPath -and (Test-Path $ScreenshotPath)) {
        try {
            $excelApp = New-Object -ComObject Excel.Application
            $excelApp.Visible = $false
            $workbook = $excelApp.Workbooks.Open((Get-Item $OutputExcelPath).FullName)
            $worksheet = $workbook.Worksheets.Item("Summary")
            
            $lastRow = $worksheet.UsedRange.Rows.Count + 2
            $worksheet.Cells.Item($lastRow, 1) = "HOMEPAGE SCREENSHOT:"
            
            $picture = $worksheet.Shapes.AddPicture(
                (Get-Item $ScreenshotPath).FullName,
                $false, $true,
                $worksheet.Cells.Item($lastRow, 2).Left,
                $worksheet.Cells.Item($lastRow, 2).Top,
                600, 400
            )
            
            $workbook.Save()
            $workbook.Close()
            $excelApp.Quit()
            
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($picture) | Out-Null
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($worksheet) | Out-Null
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excelApp) | Out-Null
            
            Write-Host "   Screenshot embedded in Excel" -ForegroundColor Green
        } catch {
            Write-Host "   WARNING: Could not embed screenshot" -ForegroundColor Yellow
        }
    }

    Write-Host ""
    Write-Host "========================================" -ForegroundColor Green
    Write-Host "ENHANCED EXTRACTION COMPLETE!" -ForegroundColor Green
    Write-Host "========================================" -ForegroundColor Green
    Write-Host ""
    Write-Host "EXTRACTION SUMMARY:" -ForegroundColor Cyan
    Write-Host "  Website: $WebsiteUrl" -ForegroundColor Gray
    Write-Host "  Company: $companyName" -ForegroundColor Gray
    Write-Host "  Buttons found: $($buttonsData.Count)" -ForegroundColor Gray
    Write-Host "  Links found: $($linksData.Count)" -ForegroundColor Gray
    Write-Host "  Images found: $($imagesData.Count)" -ForegroundColor Gray
    Write-Host "  Text sections: $($textSections.Count)" -ForegroundColor Gray
    Write-Host ""
    Write-Host "EXCEL WORKSHEETS:" -ForegroundColor Cyan
    Write-Host "  1. Summary - Business overview" -ForegroundColor Yellow
    Write-Host "  2. Buttons - All button elements" -ForegroundColor Yellow
    Write-Host "  3. Links - All hyperlinks found" -ForegroundColor Yellow
    WriteHost "  4. Images - Image information" -ForegroundColor Yellow
    Write-Host "  5. Text Sections - Content areas" -ForegroundColor Yellow
    Write-Host "  6. Full HTML - Complete page source" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "OUTPUT FILES:" -ForegroundColor Cyan
    Write-Host "  1. $OutputExcelPath (Complete Excel report)" -ForegroundColor Yellow
    if ($ScreenshotPath -and (Test-Path $ScreenshotPath)) {
        Write-Host "  2. $ScreenshotPath (Homepage screenshot)" -ForegroundColor Yellow
    }
    if (Test-Path $ImageFolder) {
        Write-Host "  3. $ImageFolder\ (Image references folder)" -ForegroundColor Yellow
    }
    Write-Host ""

    $openExcel = Read-Host "Open Excel report now? (Y/N)"
    if ($openExcel -eq "Y" -or $openExcel -eq "y") {
        Invoke-Item $OutputExcelPath
        Write-Host "Opening Excel report..." -ForegroundColor Green
    }

} catch {
    Write-Host "   ERROR creating Excel: $_" -ForegroundColor Red
    
    $csvPath = "EnhancedData_$timestamp.csv"
    $businessData | Export-Csv -Path $csvPath -NoTypeInformation
    Write-Host "   Summary saved as CSV: $csvPath" -ForegroundColor Yellow
}

Write-Host ""
Write-Host "Extraction complete. Press any key to exit..." -ForegroundColor Gray
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")