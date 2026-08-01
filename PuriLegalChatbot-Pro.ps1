# ======================================================================
# PuriLegalChatbot.ps1 - Simple Legal Assistant
# ======================================================================

# Configuration
$Config = @{
    CompanyName = "Puri Legal Services"
    CompanyPhone = "(905) 497-0090"
    CompanyEmail = "pls@bell.net"
    CompanyAddress = "808 Britannia Rd W, Unit 207, Mississauga, ON"
}

# Initialize system
function Initialize-System {
    Clear-Host
    Write-Host "================================================" -ForegroundColor Cyan
    Write-Host "       PURI LEGAL SERVICES CHATBOT" -ForegroundColor White
    Write-Host "================================================" -ForegroundColor Cyan
    Write-Host ""
}

# Knowledge base
$KnowledgeBase = @{
    greeting = @(
        "Hello! Welcome to Puri Legal Services. How can I help you today?",
        "Hi there! Thanks for contacting us. What brings you here?",
        "Welcome! I'm here to assist with Ontario legal matters."
    )
    
    traffic = @(
        "We help with traffic tickets. 60-65% success rate. Tell me about your ticket.",
        "Traffic tickets affect insurance for 3+ years. We can help reduce or dismiss charges.",
        "For traffic matters, we review officer notes and negotiate reductions."
    )
    
    smallclaims = @(
        "Small claims court handles disputes up to $35,000. 70% success rate.",
        "For money recovery, we draft claims and organize evidence.",
        "Contract disputes? We can help recover what you're owed."
    )
    
    landlord = @(
        "Landlord-tenant issues require proper LTB procedures. 60% success rate.",
        "We handle rent disputes, evictions, and maintenance issues.",
        "LTB matters need proper forms and representation."
    )
    
    immigration = @(
        "Immigration applications need precision. 75% success rate.",
        "We help with visas, sponsorships, and work permits.",
        "We fix immigration errors that cause refusals."
    )
    
    appointment = @(
        "I can book a FREE consultation for you. Would you like to schedule?",
        "Let's schedule a consultation to review your case.",
        "Book a free initial assessment today."
    )
    
    contact = @(
        "Contact: Phone (905) 497-0090, Email pls@bell.net",
        "Call us at (905) 497-0090 or email pls@bell.net",
        "Reach us at (905) 497-0090"
    )
    
    help = @(
        "I help with: Traffic Tickets, Small Claims, Landlord-Tenant, Immigration.",
        "Services: Traffic tickets, Small claims, LTB disputes, Immigration.",
        "We offer legal representation and document preparation."
    )
}

# Simple response generator
function Get-Response {
    param([string]$inputText)
    
    $text = $inputText.ToLower()
    
    # Check for keywords
    if ($text -match 'hello|hi|hey|good morning|good afternoon') {
        $responses = $KnowledgeBase.greeting
    }
    elseif ($text -match 'traffic|speeding|ticket|HTA|demerit') {
        $responses = $KnowledgeBase.traffic
    }
    elseif ($text -match 'small claims|owe money|debt|contract|lawsuit') {
        $responses = $KnowledgeBase.smallclaims
    }
    elseif ($text -match 'landlord|tenant|rent|eviction|LTB') {
        $responses = $KnowledgeBase.landlord
    }
    elseif ($text -match 'immigration|visa|sponsorship|work permit|PR') {
        $responses = $KnowledgeBase.immigration
    }
    elseif ($text -match 'appointment|consultation|meeting|schedule|book') {
        $responses = $KnowledgeBase.appointment
    }
    elseif ($text -match 'contact|phone|email|address|location') {
        $responses = $KnowledgeBase.contact
    }
    elseif ($text -match 'help|what can you do|services') {
        $responses = $KnowledgeBase.help
    }
    elseif ($text -match 'bye|goodbye|exit|quit') {
        return "Thank you for contacting Puri Legal Services. Call us at (905) 497-0090 for assistance."
    }
    else {
        return "I understand you need legal help. Could you tell me more about your situation?"
    }
    
    # Return random response from category
    $randomIndex = Get-Random -Minimum 0 -Maximum $responses.Count
    return $responses[$randomIndex]
}

# Main chatbot function
function Start-Chatbot {
    Initialize-System
    
    Write-Host "Welcome to $($Config.CompanyName)!" -ForegroundColor Green
    Write-Host "Type 'exit' to end, 'help' for services, or ask your question." -ForegroundColor Cyan
    Write-Host "--------------------------------------------------------------" -ForegroundColor Gray
    
    $conversationCount = 0
    
    while ($true) {
        Write-Host ""
        $userInput = Read-Host "You"
        
        if ($userInput.ToLower() -eq 'exit') {
            break
        }
        
        if ($userInput.ToLower() -eq 'help') {
            Write-Host ""
            Write-Host "Available services:" -ForegroundColor Yellow
            Write-Host "- Traffic tickets" -ForegroundColor White
            Write-Host "- Small claims (up to $35,000)" -ForegroundColor White
            Write-Host "- Landlord-tenant disputes" -ForegroundColor White
            Write-Host "- Immigration applications" -ForegroundColor White
            Write-Host "- Free consultations" -ForegroundColor White
            Write-Host ""
            continue
        }
        
        if ([string]::IsNullOrWhiteSpace($userInput)) {
            continue
        }
        
        $response = Get-Response $userInput
        
        Write-Host ""
        Write-Host "Assistant: $response" -ForegroundColor Cyan
        
        $conversationCount++
        
        # Suggest appointment after a few messages
        if ($conversationCount -eq 3) {
            Write-Host ""
            Write-Host "Would you like to book a FREE consultation? (yes/no)" -ForegroundColor Yellow
            $choice = Read-Host "Choice"
            
            if ($choice.ToLower() -eq 'yes') {
                Write-Host ""
                Write-Host "Great! Please provide:" -ForegroundColor Green
                $name = Read-Host "Your name"
                Write-Host "Thank you $name. We'll contact you to schedule." -ForegroundColor Green
                Write-Host "Or call us directly at (905) 497-0090" -ForegroundColor White
                break
            }
        }
    }
    
    # End conversation
    Write-Host ""
    Write-Host "================================================" -ForegroundColor Cyan
    Write-Host "Thank you for using $($Config.CompanyName)!" -ForegroundColor Green
    Write-Host "Contact us: $($Config.CompanyPhone)" -ForegroundColor White
    Write-Host "Email: $($Config.CompanyEmail)" -ForegroundColor White
    Write-Host "Address: $($Config.CompanyAddress)" -ForegroundColor White
    Write-Host "================================================" -ForegroundColor Cyan
}

# Start the chatbot
Start-Chatbot