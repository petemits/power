# ======================================================================
# PuriLegal-Chatbot-Master.ps1 - CORRECTED VERSION
# Complete Legal Research Chatbot System with All Modules
# ======================================================================

#region: MODULE IMPORTS & DEPENDENCY CHECK
# ======================================================================

Write-Host "Initializing Puri Legal Services Chatbot System..." -ForegroundColor Cyan
Write-Host "Checking system dependencies..." -ForegroundColor Yellow

# Check for required modules
$requiredModules = @("ImportExcel", "PSSelenium")
foreach ($module in $requiredModules) {
    if (-not (Get-Module -ListAvailable -Name $module)) {
        Write-Host "Installing module: $module" -ForegroundColor Yellow
        Install-Module -Name $module -Force -Scope CurrentUser -SkipPublisherCheck
    }
    Import-Module $module -Force
}

# Check external dependencies
function Test-ExternalDependencies {
    $dependencies = @{
        "Tesseract OCR" = { 
            try { 
                $null = tesseract --version 2>$null
                $true 
            } catch { $false }
        }
        "Ghostscript" = { 
            Test-Path (Get-Command gswin64c.exe -ErrorAction SilentlyContinue).Source 
        }
        "FFmpeg" = { 
            Test-Path (Get-Command ffmpeg.exe -ErrorAction SilentlyContinue).Source 
        }
    }
    
    foreach ($dep in $dependencies.Keys) {
        if (& $dependencies[$dep]) {
            Write-Host "[✓] $dep installed" -ForegroundColor Green
        } else {
            Write-Host "[✗] $dep not found - some features limited" -ForegroundColor Red
        }
    }
}

Test-ExternalDependencies
#endregion

#region: MODULE 1 - INTENT RECOGNITION ENGINE
# ======================================================================
function Get-LegalIntent {
    [CmdletBinding()]
    param([string]$UserQuery)
    
    $intentPatterns = @{
        "TRAFFIC_TICKET" = @("speeding", "ticket", "traffic", "careless driving", "HTA")
        "SMALL_CLAIMS" = @("small claims", "owe", "debt", "contractor", '\$[0-9,]+')
        "LANDLORD_TENANT" = @("landlord", "tenant", "rent", "evict", "LTB")
        "IMMIGRATION" = @("visa", "sponsorship", "immigration", "super visa", "work permit")
        "RESEARCH" = @("find", "search", "look up", "case law", "precedent")
        "DOCUMENT_PROCESS" = @("pdf", "document", "scan", "image", "extract text")
        "SAVE_DATA" = @("save", "export", "excel", "download", "record")
    }
    
    $detectedIntent = "GENERAL_QUERY"
    $confidence = 0
    $extractedEntities = @()
    
    # Convert to lowercase for matching
    $lowerQuery = $UserQuery.ToLower()
    
    foreach ($intent in $intentPatterns.Keys) {
        foreach ($pattern in $intentPatterns[$intent]) {
            if ($lowerQuery -match $pattern) {
                $detectedIntent = $intent
                $confidence += 0.3
                # Extract potential entities
                if ($matches[0]) {
                    $extractedEntities += [PSCustomObject]@{
                        Type = "KEYWORD"
                        Value = $matches[0]
                        Source = "User Query"
                    }
                }
            }
        }
    }
    
    # Extract monetary amounts (FIXED regex)
    if ($UserQuery -match '\$?(\d{1,3}(?:,\d{3})*(?:\.\d{2})?)') {
        $extractedEntities += [PSCustomObject]@{
            Type = "MONETARY_AMOUNT"
            Value = $matches[1]
            Source = "User Query"
        }
    }
    
    # Extract dates (simple pattern)
    if ($UserQuery -match '\b(\d{1,2}[/-]\d{1,2}[/-]\d{2,4})\b') {
        $extractedEntities += [PSCustomObject]@{
            Type = "DATE"
            Value = $matches[1]
            Source = "User Query"
        }
    }
    
    return [PSCustomObject]@{
        Intent = $detectedIntent
        Confidence = [Math]::Min($confidence, 1.0)
        Entities = $extractedEntities
        OriginalQuery = $UserQuery
    }
}
#endregion

#region: MODULE 2 - WEB SCRAPING ENGINE WITH SELENIUM
# ======================================================================
class LegalWebScraper {
    [PSObject]$Driver
    [string]$DataPath = "C:\PuriLegalAssistant\Data"
    
    LegalWebScraper() {
        # Initialize Selenium WebDriver
        try {
            $this.Driver = Start-SeChrome -Headless
            Write-Host "[WebScraper] Chrome driver initialized" -ForegroundColor Green
        } catch {
            Write-Host "[WebScraper] Selenium unavailable, using basic scraping" -ForegroundColor Yellow
        }
    }
    
    [PSObject] SearchLegalWeb([string]$Query, [string[]]$Sources) {
        $results = @()
        
        # Define legal source URLs
        $legalSources = @{
            "CanLII" = "https://www.canlii.org/en/#search/text="
            "OntarioLaws" = "https://www.ontario.ca/laws/search?search="
            "JusticeLaws" = "https://laws-lois.justice.gc.ca/eng/search/search.html?txtS="
            "GoogleScholar" = "https://scholar.google.com/scholar?q="
        }
        
        foreach ($source in $Sources) {
            if ($legalSources.ContainsKey($source)) {
                $url = $legalSources[$source] + [System.Web.HttpUtility]::UrlEncode($Query)
                Write-Host "[WebScraper] Searching $source : $url" -ForegroundColor Cyan
                
                $sourceResults = $this.ScrapeSource($url, $source)
                $results += $sourceResults
            }
        }
        
        return $results
    }
    
    [PSObject[]] ScrapeSource([string]$Url, [string]$SourceName) {
        $scrapedData = @()
        
        try {
            if ($this.Driver) {
                # Use Selenium for JavaScript-heavy sites
                Enter-SeUrl -Driver $this.Driver -Url $Url
                Start-Sleep -Seconds 3
                
                $pageSource = Get-SeElement -Driver $this.Driver -TagName "body"
                $htmlContent = $pageSource.Text
            } else {
                # Fallback to Invoke-WebRequest
                $response = Invoke-WebRequest -Uri $Url -UseBasicParsing
                $htmlContent = $response.Content
            }
            
            # Parse content based on source
            switch ($SourceName) {
                "CanLII" {
                    # Extract case information (simplified for now)
                    $scrapedData += [PSCustomObject]@{
                        Source = "CanLII"
                        Title = "Search results for legal cases"
                        Content = $htmlContent.Substring(0, [Math]::Min(500, $htmlContent.Length))
                        Url = $Url
                        DateScraped = Get-Date
                        RelevanceScore = $this.CalculateRelevance($htmlContent, $Query)
                    }
                }
                default {
                    # Generic extraction
                    $scrapedData += [PSCustomObject]@{
                        Source = $SourceName
                        Content = ($htmlContent -replace '<[^>]+>', ' ').Substring(0, [Math]::Min(500, $htmlContent.Length))
                        Url = $Url
                        DateScraped = Get-Date
                        RelevanceScore = 0.5
                    }
                }
            }
            
        } catch {
            Write-Host "[WebScraper] Error scraping $SourceName : $_" -ForegroundColor Red
        }
        
        return $scrapedData
    }
    
    [double] CalculateRelevance([string]$Text, [string]$Query) {
        $queryWords = $Query.ToLower() -split '\s+'
        $textLower = $Text.ToLower()
        $score = 0
        
        foreach ($word in $queryWords) {
            if ($textLower.Contains($word)) {
                $score += 0.1
            }
        }
        
        return [Math]::Min($score, 1.0)
    }
    
    [void] Close() {
        if ($this.Driver) {
            Stop-SeDriver -Driver $this.Driver
        }
    }
}
#endregion

#region: MODULE 3 - DOCUMENT PROCESSING ENGINE
# ======================================================================
class DocumentProcessor {
    [string]$TempPath = "C:\Temp\PuriLegal"
    
    DocumentProcessor() {
        # Create temp directory
        if (-not (Test-Path $this.TempPath)) {
            New-Item -ItemType Directory -Path $this.TempPath -Force
        }
    }
    
    [PSObject] ProcessDocument([string]$FilePath) {
        $fileExt = [System.IO.Path]::GetExtension($FilePath).ToLower()
        $results = [PSCustomObject]@{
            FilePath = $FilePath
            FileType = $fileExt
            Content = ""
            Metadata = @{}
            Processed = $false
        }
        
        try {
            switch ($fileExt) {
                ".pdf" {
                    $results.Content = $this.ExtractFromPdf($FilePath)
                }
                {$_ -in ".xlsx", ".xls"} {
                    $results.Content = $this.ExtractFromExcel($FilePath)
                }
                {$_ -in ".jpg", ".jpeg", ".png", ".bmp"} {
                    $results.Content = $this.ExtractFromImage($FilePath)
                }
                {$_ -in ".mp4", ".avi", ".mov"} {
                    $results.Content = $this.ExtractFromVideo($FilePath)
                }
                {$_ -in ".txt", ".doc", ".docx", ".rtf"} {
                    $results.Content = Get-Content -Path $FilePath -Raw
                }
                default {
                    Write-Host "[DocumentProcessor] Unsupported file type: $fileExt" -ForegroundColor Yellow
                    return $results
                }
            }
            
            $results.Processed = $true
            return $results
            
        } catch {
            Write-Host "[DocumentProcessor] Error processing $FilePath : $_" -ForegroundColor Red
            return $results
        }
    }
    
    [string] ExtractFromPdf([string]$PdfPath) {
        # Try to extract text using simple method first
        try {
            # Using a simple text extraction approach
            $pdfText = ""
            
            # Check if we have pdftotext command (from Poppler or Xpdf)
            if (Get-Command pdftotext.exe -ErrorAction SilentlyContinue) {
                $tempText = Join-Path $this.TempPath "temp.txt"
                & pdftotext.exe $PdfPath $tempText
                if (Test-Path $tempText) {
                    $pdfText = Get-Content $tempText -Raw
                    Remove-Item $tempText
                }
            }
            
            if (-not [string]::IsNullOrEmpty($pdfText)) {
                return $pdfText
            }
            
        } catch {
            Write-Host "[DocumentProcessor] PDF text extraction failed: $_" -ForegroundColor Yellow
        }
        
        return "[PDF content extraction requires additional tools]"
    }
    
    [string] ExtractFromImage([string]$ImagePath) {
        try {
            if (Get-Command tesseract.exe -ErrorAction SilentlyContinue) {
                $tempText = Join-Path $this.TempPath "ocr_output.txt"
                & tesseract.exe $ImagePath $tempText.Replace(".txt", "") 2>$null
                if (Test-Path $tempText) {
                    $text = Get-Content $tempText -Raw
                    Remove-Item $tempText
                    return $text
                }
            }
            return "[Image OCR requires Tesseract installation]"
        } catch {
            return "[Image OCR error]"
        }
    }
    
    [string] ExtractFromVideo([string]$VideoPath) {
        try {
            if (Get-Command ffprobe.exe -ErrorAction SilentlyContinue) {
                $duration = & ffprobe.exe -v error -show_entries format=duration -of default=noprint_wrappers=1:nokey=1 $VideoPath 2>$null
                return "Video Duration: $duration seconds. [Speech-to-text requires additional service integration]"
            }
            return "[Video processing requires FFmpeg]"
        } catch {
            return "[Video processing error]"
        }
    }
    
    [string] ExtractFromExcel([string]$ExcelPath) {
        try {
            # Try using ImportExcel module
            if (Get-Module -Name ImportExcel -ListAvailable) {
                $data = Import-Excel -Path $ExcelPath
                $summary = @()
                $summary += "Excel File: $(Split-Path $ExcelPath -Leaf)"
                
                # Get worksheet names
                try {
                    $excel = New-Object -ComObject Excel.Application
                    $workbook = $excel.Workbooks.Open($ExcelPath)
                    $sheetNames = @()
                    foreach ($sheet in $workbook.Sheets) {
                        $sheetNames += $sheet.Name
                    }
                    $workbook.Close()
                    $excel.Quit()
                    $summary += "Sheets: $($sheetNames -join ', ')"
                } catch {
                    $summary += "Sheets: [Could not read sheet names]"
                }
                
                return $summary -join "`n"
            }
            return "[Excel processing requires ImportExcel module]"
        } catch {
            return "[Excel processing error: $($_.Exception.Message)]"
        }
    }
}
#endregion

#region: MODULE 4 - SOCIAL MEDIA & API INTEGRATION
# ======================================================================
class SocialMediaCollector {
    [hashtable]$ApiKeys = @{}
    
    SocialMediaCollector() {
        # Load API keys from secure storage
        $keyPath = "C:\PuriLegalAssistant\Config\api_keys.json"
        if (Test-Path $keyPath) {
            $this.ApiKeys = Get-Content $keyPath | ConvertFrom-Json -AsHashtable
        }
    }
    
    [PSObject] GetYouTubeInfo([string]$VideoId) {
        try {
            $apiKey = $this.ApiKeys["YouTube"]
            if (-not $apiKey) {
                return [PSCustomObject]@{ Error = "YouTube API key not configured" }
            }
            
            # FIXED: Escape ampersands in URL
            $url = "https://www.googleapis.com/youtube/v3/videos?id=$VideoId`&key=$apiKey`&part=snippet,contentDetails,statistics"
            $response = Invoke-RestMethod -Uri $url -Method Get
            
            if ($response.items.Count -gt 0) {
                $video = $response.items[0]
                return [PSCustomObject]@{
                    Source = "YouTube"
                    VideoId = $VideoId
                    Title = $video.snippet.title
                    Description = $video.snippet.description
                    Channel = $video.snippet.channelTitle
                    PublishedAt = $video.snippet.publishedAt
                    Duration = $video.contentDetails.duration
                    ViewCount = $video.statistics.viewCount
                    LikeCount = $video.statistics.likeCount
                    CollectedAt = Get-Date
                }
            }
        } catch {
            Write-Host "[SocialMedia] YouTube API error: $_" -ForegroundColor Red
        }
        
        return $null
    }
    
    [PSObject[]] SearchSocialMedia([string]$Query, [string]$Platform) {
        $results = @()
        
        switch ($Platform.ToLower()) {
            "twitter" {
                # Twitter API v2 integration
                $bearerToken = $this.ApiKeys["TwitterBearer"]
                if ($bearerToken) {
                    # FIXED: Escape ampersand in URL
                    $url = "https://api.twitter.com/2/tweets/search/recent?query=`"$Query`"`&max_results=10"
                    $headers = @{ "Authorization" = "Bearer $bearerToken" }
                    
                    try {
                        $response = Invoke-RestMethod -Uri $url -Headers $headers -Method Get
                        foreach ($tweet in $response.data) {
                            $results += [PSCustomObject]@{
                                Platform = "Twitter"
                                Content = $tweet.text
                                CreatedAt = $tweet.created_at
                                TweetId = $tweet.id
                                Query = $Query
                            }
                        }
                    } catch {
                        Write-Host "[SocialMedia] Twitter API error: $_" -ForegroundColor Yellow
                    }
                }
            }
        }
        
        return $results
    }
}
#endregion

#region: MODULE 5 - DATA MANAGEMENT & EXPORT
# ======================================================================
class DataManager {
    [string]$DataPath = "C:\PuriLegalAssistant\Data"
    [System.Collections.ArrayList]$ResearchData
    
    DataManager() {
        # Initialize data directory
        if (-not (Test-Path $this.DataPath)) {
            New-Item -ItemType Directory -Path $this.DataPath -Force
        }
        
        # Initialize array list for research data
        $this.ResearchData = [System.Collections.ArrayList]::new()
    }
    
    [void] AddResearchData([PSObject]$Data) {
        $item = [PSCustomObject]@{
            ID = $this.ResearchData.Count + 1
            SourceType = $Data.Source
            Content = if ($Data.Content) { $Data.Content.ToString().Substring(0, [Math]::Min(1000, $Data.Content.Length)) } else { "" }
            URL = if ($Data.Url) { $Data.Url } else { "" }
            FilePath = if ($Data.FilePath) { $Data.FilePath } else { "" }
            DateCollected = Get-Date
            RelevanceScore = if ($Data.RelevanceScore) { $Data.RelevanceScore } else { 0.5 }
            IntentCategory = if ($Data.IntentCategory) { $Data.IntentCategory } else { "UNCATEGORIZED" }
            Tags = if ($Data.Tags) { $Data.Tags -join "," } else { "" }
        }
        
        $this.ResearchData.Add($item) | Out-Null
    }
    
    [void] ExportToExcel([string]$FileName = "LegalResearch_$(Get-Date -Format 'yyyyMMdd_HHmmss').xlsx") {
        $filePath = Join-Path $this.DataPath $FileName
        
        try {
            # Use ImportExcel module for robust export
            if (Get-Module -Name ImportExcel -ListAvailable) {
                $this.ResearchData | Export-Excel `
                    -Path $filePath `
                    -WorksheetName "LegalResearch" `
                    -TableName "ResearchData" `
                    -AutoSize `
                    -FreezeTopRow `
                    -BoldTopRow `
                    -Show
            
                Write-Host "[DataManager] Data exported to: $filePath" -ForegroundColor Green
                Write-Host "[DataManager] Total records: $($this.ResearchData.Count)" -ForegroundColor Cyan
                
                # Also export as CSV for compatibility
                $csvPath = $filePath -replace '\.xlsx$', '.csv'
                $this.ResearchData | Export-Csv -Path $csvPath -NoTypeInformation
            } else {
                # Fallback to CSV only
                $this.ResearchData | Export-Csv -Path (Join-Path $this.DataPath "LegalResearch_$(Get-Date -Format 'yyyyMMdd').csv") -NoTypeInformation
                Write-Host "[DataManager] CSV exported (Install ImportExcel module for Excel export)" -ForegroundColor Yellow
            }
            
        } catch {
            Write-Host "[DataManager] Export error: $($_.Exception.Message)" -ForegroundColor Red
            # Simple fallback
            $this.ResearchData | Out-File -FilePath (Join-Path $this.DataPath "LegalResearch_$(Get-Date -Format 'yyyyMMdd').txt")
        }
    }
    
    [PSObject[]] SearchData([string]$Keyword) {
        $results = @()
        foreach ($item in $this.ResearchData) {
            if ($item.Content -match $Keyword -or $item.Tags -match $Keyword) {
                $results += [PSCustomObject]@{
                    ID = $item.ID
                    Source = $item.SourceType
                    Excerpt = if ($item.Content) { $item.Content.Substring(0, [Math]::Min(200, $item.Content.Length)) } else { "" }
                    Date = $item.DateCollected
                    Relevance = $item.RelevanceScore
                }
            }
        }
        
        return $results | Sort-Object Relevance -Descending
    }
}
#endregion

#region: MODULE 6 - CHATBOT INTERFACE
# ======================================================================
class PuriLegalChatbot {
    [LegalWebScraper]$WebScraper
    [DocumentProcessor]$DocProcessor
    [SocialMediaCollector]$SocialCollector
    [DataManager]$DataManager
    [hashtable]$ConversationHistory
    [string]$UserName
    
    PuriLegalChatbot([string]$UserName = "Client") {
        $this.UserName = $UserName
        $this.WebScraper = [LegalWebScraper]::new()
        $this.DocProcessor = [DocumentProcessor]::new()
        $this.SocialCollector = [SocialMediaCollector]::new()
        $this.DataManager = [DataManager]::new()
        $this.ConversationHistory = @{
            Messages = @()
            StartTime = Get-Date
            IntentsDetected = @()
        }
        
        Write-Host "`n===========================================" -ForegroundColor Cyan
        Write-Host "   PURI LEGAL SERVICES AI ASSISTANT" -ForegroundColor White
        Write-Host "   Ontario Paralegal Chatbot System" -ForegroundColor White
        Write-Host "===========================================`n" -ForegroundColor Cyan
    }
    
    [PSObject] ProcessQuery([string]$UserQuery) {
        # Add to conversation history
        $this.ConversationHistory.Messages += @{
            Time = Get-Date
            Sender = "User"
            Message = $UserQuery
        }
        
        # Step 1: Detect intent
        $intent = Get-LegalIntent -UserQuery $UserQuery
        $this.ConversationHistory.IntentsDetected += $intent.Intent
        
        # Step 2: Generate response based on intent
        $response = $this.GenerateResponse($intent, $UserQuery)
        
        # Step 3: Store in conversation history
        $this.ConversationHistory.Messages += @{
            Time = Get-Date
            Sender = "Bot"
            Message = $response.Message
            Data = $response.Data
        }
        
        # Step 4: Export conversation if requested
        if ($UserQuery -match "export|save|download") {
            $this.ExportConversation()
        }
        
        return $response
    }
    
    [PSObject] GenerateResponse([PSObject]$Intent, [string]$UserQuery) {
        $response = [PSCustomObject]@{
            Message = ""
            Data = $null
            Actions = @()
        }
        
        switch ($Intent.Intent) {
            "TRAFFIC_TICKET" {
                $response.Message = "I can help with traffic tickets. Let me search for recent Ontario traffic law precedents and success rates."
                $webResults = $this.WebScraper.SearchLegalWeb("Ontario traffic ticket precedent 2024", @("CanLII", "GoogleScholar"))
                $response.Data = $webResults
                $response.Actions += "SEARCHED_LEGAL_DATABASE"
                
                # Add to data manager
                foreach ($result in $webResults) {
                    $result | Add-Member -NotePropertyName "IntentCategory" -NotePropertyValue "TRAFFIC_TICKET"
                    $this.DataManager.AddResearchData($result)
                }
            }
            
            "SMALL_CLAIMS" {
                $response.Message = "For small claims matters, I'll research recent Ontario cases and success statistics."
                
                # Extract monetary amount if mentioned
                $amount = ""
                if ($UserQuery -match '\$?(\d{1,3}(?:,\d{3})*(?:\.\d{2})?)') {
                    $amount = "under $$($matches[1]) "
                }
                
                $webResults = $this.WebScraper.SearchLegalWeb("Ontario small claims court $amount", @("CanLII", "OntarioLaws"))
                $response.Data = $webResults
                $response.Actions += "SEARCHED_CASE_LAW"
                
                foreach ($result in $webResults) {
                    $result | Add-Member -NotePropertyName "IntentCategory" -NotePropertyValue "SMALL_CLAIMS"
                    $this.DataManager.AddResearchData($result)
                }
            }
            
            "RESEARCH" {
                $response.Message = "I'll conduct comprehensive legal research for you."
                
                # Extract search terms
                $searchTerms = $UserQuery -replace "find|search|look up|for|me", ""
                $webResults = $this.WebScraper.SearchLegalWeb($searchTerms, @("CanLII", "JusticeLaws", "GoogleScholar"))
                $response.Data = $webResults
                $response.Actions += "CONDUCTED_RESEARCH"
                
                foreach ($result in $webResults) {
                    $result | Add-Member -NotePropertyName "IntentCategory" -NotePropertyValue "RESEARCH"
                    $this.DataManager.AddResearchData($result)
                }
            }
            
            "DOCUMENT_PROCESS" {
                # Check if user mentioned a file path
                if ($UserQuery -match 'C:\\[^"]+\.(pdf|docx?|xlsx?|jpg|png)') {
                    $filePath = $matches[0]
                    if (Test-Path $filePath) {
                        $docResult = $this.DocProcessor.ProcessDocument($filePath)
                        $response.Message = "Processed document: $(Split-Path $filePath -Leaf). Extracted $($docResult.Content.Length) characters."
                        $response.Data = $docResult
                        $response.Actions += "PROCESSED_DOCUMENT"
                        
                        $docResult | Add-Member -NotePropertyName "IntentCategory" -NotePropertyValue "DOCUMENT_PROCESS"
                        $this.DataManager.AddResearchData($docResult)
                    } else {
                        $response.Message = "File not found: $filePath"
                    }
                } else {
                    $response.Message = "Please specify a document path to process (e.g., C:\Documents\case.pdf)"
                }
            }
            
            "SAVE_DATA" {
                $fileName = "PuriLegal_Research_$(Get-Date -Format 'yyyyMMdd_HHmm').xlsx"
                $this.DataManager.ExportToExcel($fileName)
                $response.Message = "All research data has been exported to Excel. Total records: $($this.DataManager.ResearchData.Count)"
                $response.Actions += "EXPORTED_DATA"
            }
            
            default {
                $response.Message = "I can help with traffic tickets, small claims, landlord-tenant issues, immigration, legal research, and document processing. What specific legal matter do you need assistance with?"
            }
        }
        
        # Add success rate information for legal matters
        if ($Intent.Intent -in @("TRAFFIC_TICKET", "SMALL_CLAIMS", "LANDLORD_TENANT")) {
            $successRates = @{
                "TRAFFIC_TICKET" = "60-65% favorable outcomes"
                "SMALL_CLAIMS" = "~70% successful resolutions"
                "LANDLORD_TENANT" = "~60% favorable LTB outcomes"
                "IMMIGRATION" = "~75% success with refused applications"
            }
            
            if ($successRates.ContainsKey($Intent.Intent)) {
                $response.Message += "`n`nBased on our experience, we typically see $($successRates[$Intent.Intent]) with proper representation."
            }
        }
        
        return $response
    }
    
    [void] ExportConversation() {
        $exportData = @()
        
        foreach ($msg in $this.ConversationHistory.Messages) {
            $exportData += [PSCustomObject]@{
                Timestamp = $msg.Time
                Speaker = $msg.Sender
                Message = $msg.Message
                Intent = if ($msg.Sender -eq "User") { (Get-LegalIntent -UserQuery $msg.Message).Intent } else { "" }
            }
        }
        
        $conversationFile = Join-Path $this.DataManager.DataPath "Conversation_$(Get-Date -Format 'yyyyMMdd_HHmm').xlsx"
        
        if (Get-Module -Name ImportExcel -ListAvailable) {
            $exportData | Export-Excel -Path $conversationFile -WorksheetName "ChatLog" -AutoSize
            Write-Host "[Chatbot] Conversation exported to: $conversationFile" -ForegroundColor Green
        } else {
            $csvFile = $conversationFile -replace '\.xlsx$', '.csv'
            $exportData | Export-Csv -Path $csvFile -NoTypeInformation
            Write-Host "[Chatbot] Conversation exported to: $csvFile" -ForegroundColor Green
        }
    }
    
    [void] Cleanup() {
        $this.WebScraper.Close()
        Write-Host "`nChatbot session completed. Duration: $((Get-Date) - $this.ConversationHistory.StartTime)" -ForegroundColor Cyan
        Write-Host "Total intents detected: $($this.ConversationHistory.IntentsDetected.Count)" -ForegroundColor Cyan
        Write-Host "Research records collected: $($this.DataManager.ResearchData.Count)" -ForegroundColor Cyan
    }
}
#endregion

#region: MAIN EXECUTION - INTERACTIVE CHAT INTERFACE
# ======================================================================
function Start-LegalChatbot {
    param([string]$UserName)
    
    # Initialize chatbot
    $chatbot = [PuriLegalChatbot]::new($UserName)
    
    Write-Host "Hello $UserName! I'm your Puri Legal Services assistant." -ForegroundColor Green
    Write-Host "I can help with:" -ForegroundColor Yellow
    Write-Host "  • Traffic tickets and Ontario HTA violations" -ForegroundColor White
    Write-Host "  • Small Claims Court matters (up to `$35,000)" -ForegroundColor White
    Write-Host "  • Landlord & Tenant Board disputes" -ForegroundColor White
    Write-Host "  • Immigration application support" -ForegroundColor White
    Write-Host "  • Legal research and case law lookup" -ForegroundColor White
    Write-Host "  • Document processing (PDF, Excel, images)" -ForegroundColor White
    Write-Host "  • Data collection and Excel export" -ForegroundColor White
    Write-Host ""
    Write-Host "Type 'exit' to end, 'export' to save data, or 'help' for options.`n" -ForegroundColor Cyan
    
    # Main conversation loop
    while ($true) {
        $userInput = Read-Host "`n$UserName"
        
        if ($userInput -eq "exit" -or $userInput -eq "quit") {
            $chatbot.Cleanup()
            break
        }
        
        if ($userInput -eq "help") {
            Write-Host "`nAvailable commands:" -ForegroundColor Yellow
            Write-Host "  • [legal question] - Ask about any legal matter" -ForegroundColor White
            Write-Host "  • process [filepath] - Process a document (PDF, Excel, image)" -ForegroundColor White
            Write-Host "  • search [query] - Search legal databases" -ForegroundColor White
            Write-Host "  • export - Save all data to Excel" -ForegroundColor White
            Write-Host "  • history - Show conversation history" -ForegroundColor White
            Write-Host "  • exit - End session" -ForegroundColor White
            continue
        }
        
        if ($userInput -eq "history") {
            Write-Host "`nConversation History:" -ForegroundColor Cyan
            foreach ($msg in $chatbot.ConversationHistory.Messages) {
                $time = $msg.Time.ToString("HH:mm:ss")
                $color = if ($msg.Sender -eq "User") { "Green" } else { "Cyan" }
                Write-Host "[$time $($msg.Sender)]: $($msg.Message)" -ForegroundColor $color
            }
            continue
        }
        
        if ([string]::IsNullOrWhiteSpace($userInput)) {
            continue
        }
        
        # Process the user query
        $response = $chatbot.ProcessQuery($userInput)
        
        # Display the response
        Write-Host ""
        Write-Host "[Assistant]: $($response.Message)" -ForegroundColor Cyan
        
        # Show data preview if available
        if ($response.Data -and $response.Data.Count -gt 0) {
            Write-Host "`nFound $($response.Data.Count) results:" -ForegroundColor Yellow
            for ($i = 0; $i -lt [Math]::Min(3, $response.Data.Count); $i++) {
                $result = $response.Data[$i]
                $preview = if ($result.Title) { $result.Title } else { $result.Content.Substring(0, [Math]::Min(50, $result.Content.Length)) }
                Write-Host "  • [$($result.Source)] $preview..." -ForegroundColor White
            }
            
            if ($response.Data.Count -gt 3) {
                Write-Host "  ... and $($response.Data.Count - 3) more" -ForegroundColor Gray
            }
        }
    }
}
#endregion

# ======================================================================
# EXECUTION START
# ======================================================================

# Check if running in interactive mode or with parameters
if ($MyInvocation.InvocationName -ne ".") {
    if ($args.Count -gt 0) {
        # Command-line mode
        switch ($args[0]) {
            "--search" {
                $scraper = [LegalWebScraper]::new()
                $results = $scraper.SearchLegalWeb($args[1], @("CanLII", "GoogleScholar"))
                $results | Format-Table -AutoSize
                $scraper.Close()
            }
            "--process" {
                $processor = [DocumentProcessor]::new()
                $result = $processor.ProcessDocument($args[1])
                Write-Host "Extracted text ($($result.Content.Length) chars):"
                Write-Host $result.Content
            }
            "--export" {
                $manager = [DataManager]::new()
                $manager.ExportToExcel()
            }
            default {
                # Start interactive chatbot
                $name = if ($args[0]) { $args[0] } else { "Client" }
                Start-LegalChatbot -UserName $name
            }
        }
    } else {
        # Start interactive chatbot
        Start-LegalChatbot -UserName "Client"
    }
}