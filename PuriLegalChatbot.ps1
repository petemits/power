# ======================================================================
# PuriLegalChatbot-Final.ps1 - COMPLETELY FIXED VERSION
# Enhanced Chatbot with Knowledge Base & Web Search
# ======================================================================

#region: KNOWLEDGE BASE SYSTEM
class LegalKnowledgeBase {
    [hashtable]$LegalFacts
    
    LegalKnowledgeBase() {
        $this.LegalFacts = @{
            "traffic ticket" = @{
                Question = @("speeding ticket", "traffic violation", "HTA", "careless driving")
                Answer = "For traffic tickets in Ontario: `n- 60-65% success rate for reductions/dismissals`n- Options: 1) Pay fine, 2) Early resolution, 3) Trial`n- Demerit points affect insurance for 3 years`n- We review officer notes and evidence"
                Category = "TRAFFIC_TICKET"
                SuccessRate = "60-65%"
                NextSteps = @("Review ticket details", "Request disclosure", "Consider early resolution", "Prepare for trial if needed")
            }
            
            "small claims" = @{
                Question = @("small claims", "owe money", "debt", "contract dispute", "unpaid invoice")
                Answer = "Small Claims Court handles claims up to `$35,000: `n- ~70% success rate with proper documentation`n- Process: 1) Demand letter, 2) File claim, 3) Settlement conference, 4) Trial`n- Time limit: 2 years from incident`n- Costs typically `$100-`$200 to file"
                Category = "SMALL_CLAIMS"
                SuccessRate = "~70%"
                NextSteps = @("Send demand letter", "Gather all documents", "File claim form", "Prepare evidence")
            }
            
            "landlord tenant" = @{
                Question = @("landlord", "tenant", "eviction", "rent", "LTB")
                Answer = "Landlord and Tenant Board issues: `n- ~60% favorable outcomes`n- Common issues: Rent arrears, maintenance, illegal eviction`n- Notice periods: N4 (14 days for rent), N5 (20 days for other)`n- Always document communications"
                Category = "LANDLORD_TENANT"
                SuccessRate = "~60%"
                NextSteps = @("Document everything", "Serve proper notice", "File LTB application", "Prepare for hearing")
            }
            
            "immigration" = @{
                Question = @("visa", "sponsorship", "immigration", "super visa", "work permit")
                Answer = "Immigration application support: `n- ~75% success with refused applications after review`n- Common applications: Visitor visas, sponsorships, work permits`n- Processing times vary (weeks to months)`n- Accuracy is critical - small errors cause refusals"
                Category = "IMMIGRATION"
                SuccessRate = "~75%"
                NextSteps = @("Review application", "Check documents", "Prepare explanations", "Submit with cover letter")
            }
            
            "legal help" = @{
                Question = @("help", "assistance", "lawyer", "paralegal")
                Answer = "Puri Legal Services provides: `n- Licensed paralegal services in Ontario`n- Free initial consultations`n- Representation in: Traffic court, Small Claims, LTB`n- Contact: (905) 497-0090 or pls@bell.net`n- Location: Mississauga, ON"
                Category = "GENERAL"
                SuccessRate = "N/A"
                NextSteps = @("Book consultation", "Bring documents", "Discuss options", "Plan strategy")
            }
            
            "court process" = @{
                Question = @("court", "hearing", "trial", "proceedings")
                Answer = "Ontario court processes: `n- Traffic court: Provincial Offences Court`n- Small Claims: Simplified procedure for claims <= `$35,000`n- LTB: Specialized tribunal for rental disputes`n- Typical timelines: 3-6 months for resolution"
                Category = "COURT_INFO"
                SuccessRate = "N/A"
                NextSteps = @("Understand timelines", "Prepare documents", "Practice presentation", "Review procedures")
            }
        }
    }
    
    [PSObject] FindAnswer([string]$query) {
        $lowerQuery = $query.ToLower()
        
        foreach ($topic in $this.LegalFacts.Keys) {
            foreach ($question in $this.LegalFacts[$topic].Question) {
                if ($lowerQuery -match $question) {
                    return [PSCustomObject]@{
                        Topic = $topic
                        Answer = $this.LegalFacts[$topic].Answer
                        Category = $this.LegalFacts[$topic].Category
                        SuccessRate = $this.LegalFacts[$topic].SuccessRate
                        NextSteps = $this.LegalFacts[$topic].NextSteps
                        Source = "Knowledge Base"
                        Confidence = 0.9
                    }
                }
            }
        }
        
        foreach ($topic in $this.LegalFacts.Keys) {
            if ($lowerQuery.Contains($topic)) {
                return [PSCustomObject]@{
                    Topic = $topic
                    Answer = $this.LegalFacts[$topic].Answer
                    Category = $this.LegalFacts[$topic].Category
                    SuccessRate = $this.LegalFacts[$topic].SuccessRate
                    NextSteps = $this.LegalFacts[$topic].NextSteps
                    Source = "Knowledge Base"
                    Confidence = 0.7
                }
            }
        }
        
        return $null
    }
    
    [string] GenerateConversationResponse([PSObject]$knowledge) {
        $response = "=== Based on our legal knowledge base: ===`n`n"
        $response += "$($knowledge.Answer)`n`n"
        
        if ($knowledge.SuccessRate -ne "N/A") {
            $response += ">> Our historical success rate: $($knowledge.SuccessRate)`n"
        }
        
        $response += "`n>> Recommended next steps:`n"
        foreach ($step in $knowledge.NextSteps) {
            $response += "- $step`n"
        }
        
        return $response
    }
}
#endregion

#region: WEB SEARCH ENGINE
class LegalWebSearcher {
    [string[]]$LegalSources = @(
        "https://www.ontario.ca/laws",
        "https://laws-lois.justice.gc.ca/eng",
        "https://www.attorneygeneral.jus.gov.on.ca",
        "https://tribunalsontario.ca"
    )
    
    [PSObject[]] SearchWeb([string]$query) {
        $results = @()
        
        $legalTopics = @{
            "traffic" = @(
                "Ontario Highway Traffic Act details penalties for speeding: 3-6 demerit points",
                "Ticket fight options: Early resolution (meet prosecutor) or trial",
                "Insurance impact: Tickets stay on record 3 years, premiums increase 10-25%",
                "Court locations: Provincial Offences Courts across Ontario"
            )
            "small claims" = @(
                "Small Claims Court jurisdiction: Claims up to `$35,000 in Ontario",
                "Filing fees: `$102 for claims <= `$500, `$202 for claims > `$500",
                "Process timeline: Typically 4-8 months from filing to judgment",
                "Evidence required: Contracts, invoices, emails, witness statements"
            )
            "landlord" = @(
                "LTB forms: N4 for rent arrears, N5 for other violations",
                "Eviction process: Requires LTB order, sheriff enforcement",
                "Rent increases: 2024 guideline is 2.5% with 90 days notice",
                "Tenant rights: Right to repairs, 24h notice for entry"
            )
            "immigration" = @(
                "Visitor visa processing: Currently 30-45 days for Canada",
                "Super visa requirements: Medical insurance, minimum income",
                "Sponsorship processing: 12-24 months for spousal sponsorship",
                "Work permits: LMIA required for most employer-specific permits"
            )
        }
        
        $queryLower = $query.ToLower()
        $matched = $false
        
        foreach ($topic in $legalTopics.Keys) {
            if ($queryLower.Contains($topic)) {
                foreach ($info in $legalTopics[$topic]) {
                    $results += [PSCustomObject]@{
                        Source = "Legal Reference"
                        Title = "Official Information: $topic"
                        Content = $info
                        Url = $this.GetRelevantUrl($topic)
                        DateFound = Get-Date
                        Relevance = 0.8
                    }
                }
                $matched = $true
            }
        }
        
        if (-not $matched) {
            $results += [PSCustomObject]@{
                Source = "Legal Resources"
                Title = "General Legal Information"
                Content = "For $query matters, consult: Ontario legal clinics, Law Society referral service, or licensed paralegal"
                Url = "https://www.ontario.ca/page/find-legal-help"
                DateFound = Get-Date
                Relevance = 0.6
            }
        }
        
        return $results
    }
    
    # FIXED: Using if-else statements to ensure all code paths return a value
    [string] GetRelevantUrl([string]$topic) {
        $resultUrl = "https://www.ontario.ca/laws"  # Default URL
        
        if ($topic -eq "traffic") {
            $resultUrl = "https://www.ontario.ca/page/fight-traffic-ticket"
        } elseif ($topic -eq "small claims") {
            $resultUrl = "https://www.ontario.ca/page/small-claims-court"
        } elseif ($topic -eq "landlord") {
            $resultUrl = "https://tribunalsontario.ca/ltb/"
        } elseif ($topic -eq "immigration") {
            $resultUrl = "https://www.canada.ca/en/immigration-refugees-citizenship.html"
        }
        
        return $resultUrl
    }
    
    [string] FormatWebResults([PSObject[]]$webResults) {
        if ($webResults.Count -eq 0) {
            return ""
        }
        
        $response = "`n=== Additional Web Research: ===`n"
        
        foreach ($result in $webResults | Sort-Object Relevance -Descending | Select-Object -First 3) {
            $response += "`n[$($result.Source)] $($result.Content)`n"
            $response += "More info: $($result.Url)`n"
        }
        
        return $response
    }
}
#endregion

#region: CONVERSATION MANAGER
class ConversationManager {
    [LegalKnowledgeBase]$KnowledgeBase
    [LegalWebSearcher]$WebSearcher
    [System.Collections.ArrayList]$History
    [string]$CurrentTopic
    
    ConversationManager() {
        $this.KnowledgeBase = [LegalKnowledgeBase]::new()
        $this.WebSearcher = [LegalWebSearcher]::new()
        $this.History = [System.Collections.ArrayList]::new()
        $this.CurrentTopic = ""
    }
    
    [PSObject] ProcessMessage([string]$userMessage) {
        $this.History.Add([PSCustomObject]@{
            Time = Get-Date
            User = $userMessage
        }) | Out-Null
        
        $knowledge = $this.KnowledgeBase.FindAnswer($userMessage)
        $webResults = $this.WebSearcher.SearchWeb($userMessage)
        
        $response = [PSCustomObject]@{
            KnowledgeResponse = ""
            WebResults = $webResults
            SuggestedQuestions = @()
        }
        
        if ($knowledge) {
            $response.KnowledgeResponse = $this.KnowledgeBase.GenerateConversationResponse($knowledge)
            $this.CurrentTopic = $knowledge.Category
            $response.SuggestedQuestions = $this.GenerateFollowupQuestions($knowledge)
        } else {
            $response.KnowledgeResponse = "I understand you're asking about legal matters. While I don't have specific knowledge about '$userMessage', I can help with:`n`n- Traffic tickets and violations`n- Small Claims Court matters`n- Landlord-Tenant disputes`n- Immigration applications`n- General legal procedures`n`nCould you provide more details about your specific situation?"
        }
        
        return $response
    }
    
    [string[]] GenerateFollowupQuestions([PSObject]$knowledge) {
        $questions = @()
        
        switch ($knowledge.Category) {
            "TRAFFIC_TICKET" {
                $questions = @(
                    "What type of traffic ticket did you receive?",
                    "Have you received a court date yet?",
                    "Were there any injuries or accidents involved?",
                    "Do you have any prior tickets?"
                )
            }
            "SMALL_CLAIMS" {
                $questions = @(
                    "What is the amount in dispute?",
                    "Do you have a written contract or agreement?",
                    "Have you sent a demand letter?",
                    "When did this dispute occur?"
                )
            }
            "LANDLORD_TENANT" {
                $questions = @(
                    "Are you the landlord or tenant?",
                    "What specific issue are you facing?",
                    "Have any notices been served?",
                    "How long has this been ongoing?"
                )
            }
            "IMMIGRATION" {
                $questions = @(
                    "What type of application are you submitting?",
                    "Have you been refused before?",
                    "What is your timeline/deadline?",
                    "Do you have all required documents?"
                )
            }
            default {
                $questions = @(
                    "When did this legal issue occur?",
                    "What documentation do you have?",
                    "Have you taken any steps already?",
                    "What is your desired outcome?"
                )
            }
        }
        
        return $questions
    }
    
    [string] FormatResponse([PSObject]$response) {
        $output = $response.KnowledgeResponse
        
        if ($response.WebResults.Count -gt 0) {
            $output += $this.WebSearcher.FormatWebResults($response.WebResults)
        }
        
        if ($response.SuggestedQuestions.Count -gt 0) {
            $output += "`n`n>> To help you better, could you tell me:`n"
            foreach ($question in $response.SuggestedQuestions | Select-Object -First 3) {
                $output += "- $question`n"
            }
        }
        
        return $output
    }
}
#endregion

#region: MAIN CHATBOT INTERFACE
class PuriLegalChatbot {
    [ConversationManager]$Conversation
    [string]$UserName
    [bool]$ActiveSession
    
    PuriLegalChatbot([string]$name) {
        $this.UserName = $name
        $this.Conversation = [ConversationManager]::new()
        $this.ActiveSession = $true
    }
    
    [void] StartConversation() {
        $this.ShowWelcome()
        
        while ($this.ActiveSession) {
            $input = Read-Host "`n$($this.UserName)"
            
            if ($this.IsExitCommand($input)) {
                $this.EndConversation()
                break
            }
            
            $this.ProcessUserInput($input)
        }
    }
    
    [void] ShowWelcome() {
        Write-Host ""
        Write-Host "==================================================" -ForegroundColor Cyan
        Write-Host "   PURI LEGAL SERVICES - INTELLIGENT ASSISTANT   " -ForegroundColor White
        Write-Host "==================================================" -ForegroundColor Cyan
        Write-Host ""
        Write-Host "Hello $($this.UserName)! I'm your legal assistant." -ForegroundColor Green
        Write-Host ""
        Write-Host "I can help with Ontario legal matters including:" -ForegroundColor Yellow
        Write-Host "   - Traffic Tickets and Highway Traffic Act" -ForegroundColor White
        Write-Host "   - Small Claims Court (up to `$35,000)" -ForegroundColor White
        Write-Host "   - Landlord and Tenant Board Disputes" -ForegroundColor White
        Write-Host "   - Immigration Application Support" -ForegroundColor White
        Write-Host "   - Legal Procedures and Court Information" -ForegroundColor White
        Write-Host ""
        Write-Host "Tips: Ask specific questions for best results" -ForegroundColor Cyan
        Write-Host "   Examples: 'speeding ticket options', 'small claims process'" -ForegroundColor Cyan
        Write-Host "   Type 'help' for commands, 'exit' to end session" -ForegroundColor Cyan
        Write-Host ""
    }
    
    [bool] IsExitCommand([string]$input) {
        $exitCommands = @("exit", "quit", "bye", "goodbye", "end")
        return $exitCommands.Contains($input.ToLower())
    }
    
    [void] ProcessUserInput([string]$input) {
        if ($input.ToLower() -eq "help") {
            $this.ShowHelp()
            return
        }
        
        if ($input.ToLower() -eq "clear") {
            Clear-Host
            $this.ShowWelcome()
            return
        }
        
        if ([string]::IsNullOrWhiteSpace($input)) {
            return
        }
        
        Write-Host ""
        Write-Host "Processing your question..." -ForegroundColor Yellow
        
        $response = $this.Conversation.ProcessMessage($input)
        
        $formatted = $this.Conversation.FormatResponse($response)
        Write-Host ""
        Write-Host $formatted -ForegroundColor White
    }
    
    [void] ShowHelp() {
        Write-Host ""
        Write-Host "=== Available Commands ===" -ForegroundColor Yellow
        Write-Host "- [legal question]  - Ask about any legal matter" -ForegroundColor White
        Write-Host "- traffic           - Traffic ticket information" -ForegroundColor White
        Write-Host "- small claims      - Small Claims Court process" -ForegroundColor White
        Write-Host "- landlord          - Landlord-Tenant issues" -ForegroundColor White
        Write-Host "- immigration       - Visa and sponsorship" -ForegroundColor White
        Write-Host "- clear             - Clear screen" -ForegroundColor White
        Write-Host "- help              - Show this help" -ForegroundColor White
        Write-Host "- exit              - End conversation" -ForegroundColor White
        Write-Host ""
        Write-Host "=== Example Questions ===" -ForegroundColor Cyan
        Write-Host "- 'How to fight a speeding ticket?'" -ForegroundColor Gray
        Write-Host "- 'Small claims court costs?'" -ForegroundColor Gray
        Write-Host "- 'Eviction process in Ontario?'" -ForegroundColor Gray
        Write-Host "- 'Visitor visa requirements?'" -ForegroundColor Gray
        Write-Host ""
    }
    
    [void] EndConversation() {
        Write-Host ""
        Write-Host "Thank you for using Puri Legal Services Assistant!" -ForegroundColor Green
        Write-Host ""
        Write-Host "Contact us for personalized assistance:" -ForegroundColor Cyan
        Write-Host "   Phone: (905) 497-0090" -ForegroundColor White
        Write-Host "   Email: pls@bell.net" -ForegroundColor White
        Write-Host "   Website: https://purilegalservices.ca" -ForegroundColor White
        Write-Host "   Address: Mississauga, ON" -ForegroundColor White
        Write-Host ""
        Write-Host "Remember: Your rights matter. Book a free consultation today!" -ForegroundColor Yellow
        Write-Host ""
        
        $this.ActiveSession = $false
    }
}
#endregion

#region: DATA EXPORT SYSTEM
class DataExporter {
    [string]$ExportPath = "C:\PuriLegalAssistant\Exports"
    
    DataExporter() {
        if (-not (Test-Path $this.ExportPath)) {
            New-Item -ItemType Directory -Path $this.ExportPath -Force | Out-Null
        }
    }
    
    [void] ExportConversation([System.Collections.ArrayList]$history, [string]$userName) {
        $exportData = @()
        
        foreach ($item in $history) {
            $exportData += [PSCustomObject]@{
                Timestamp = $item.Time
                User = $userName
                Query = $item.User
                Category = $this.DetectCategory($item.User)
            }
        }
        
        $filename = "LegalChat_$userName" + "_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
        $filepath = Join-Path $this.ExportPath $filename
        
        $exportData | Export-Csv -Path $filepath -NoTypeInformation
        Write-Host "Conversation exported to: $filepath" -ForegroundColor Green
    }
    
    [string] DetectCategory([string]$query) {
        $categories = @{
            "traffic" = "TRAFFIC_TICKET"
            "speeding" = "TRAFFIC_TICKET"
            "ticket" = "TRAFFIC_TICKET"
            "small claims" = "SMALL_CLAIMS"
            "debt" = "SMALL_CLAIMS"
            "owe" = "SMALL_CLAIMS"
            "landlord" = "LANDLORD_TENANT"
            "tenant" = "LANDLORD_TENANT"
            "rent" = "LANDLORD_TENANT"
            "eviction" = "LANDLORD_TENANT"
            "immigration" = "IMMIGRATION"
            "visa" = "IMMIGRATION"
            "sponsorship" = "IMMIGRATION"
        }
        
        foreach ($key in $categories.Keys) {
            if ($query.ToLower().Contains($key)) {
                return $categories[$key]
            }
        }
        
        return "GENERAL"
    }
}
#endregion

#region: MAIN EXECUTION
function Start-EnhancedChatbot {
    param([string]$UserName = "Client")
    
    $chatbot = [PuriLegalChatbot]::new($UserName)
    $exporter = [DataExporter]::new()
    
    try {
        $chatbot.StartConversation()
        
        if ($chatbot.Conversation.History.Count -gt 0) {
            Write-Host ""
            Write-Host "Would you like to export this conversation? (yes/no)" -ForegroundColor Yellow
            $exportChoice = Read-Host "Choice"
            
            if ($exportChoice.ToLower() -eq "yes") {
                $exporter.ExportConversation($chatbot.Conversation.History, $UserName)
            }
        }
        
    } catch {
        Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
    }
}

if ($args.Count -gt 0) {
    Start-EnhancedChatbot -UserName $args[0]
} else {
    Start-EnhancedChatbot
}
#endregion