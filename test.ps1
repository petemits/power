# SmartChatBot.ps1 - A conversational agent with web search capabilities

param(
    [switch]$Headless = $false # Use -Headless to run browser in background
)

# Import the Selenium module[citation:1][citation:2]
Import-Module -Name Selenium

# Function to perform a web search and extract key info
function Get-WebAnswer {
    param([string]$Query)
    
    try {
        # Step 1: Navigate to a search engine
        $Driver.Navigate().GoToUrl("https://www.bing.com/search?q=$Query")
        Start-Sleep -Seconds 2  # Brief wait for page load
        
        # Step 2: Extract text from the first few search results
        # This targets the main content areas of Bing's results
        $resultElements = $Driver.FindElementsByCssSelector('.b_algo, .b_caption p')
        $textContent = @()
        
        foreach ($elem in $resultElements) {
            if ($textContent.Count -gt 5) { break } # Limit content length
            $text = $elem.Text.Trim()
            if ($text -and $text.Length -gt 50) {
                $textContent += $text
            }
        }
        
        # Step 3: Summarize the extracted information
        if ($textContent.Count -eq 0) {
            return "I searched but couldn't find a clear answer on that. Could you ask in a different way?"
        }
        
        $combinedContent = $textContent -join " "
        # Simple summarization: take the first 3 sentences
        $summary = ($combinedContent -split '\. ')[0..2] -join '. '
        return "Based on my search: $summary"
        
    } catch {
        return "I encountered an issue while searching. Please try again."
    }
}

# Function to generate a response using chat logic and web search
function Get-ChatResponse {
    param([string]$UserInput)
    
    # Simple pattern matching for greetings and farewells
    switch -regex ($UserInput.ToLower()) {
        '^(hi|hello|hey|greetings)' { 
            $greetings = @("Hello!", "Hi there!", "Greetings!", "Hey!") 
            return $greetings | Get-Random
        }
        '^(bye|exit|quit|goodbye)' { 
            $farewells = @("Goodbye!", "Talk to you later!", "Have a great day!") 
            return $farewells | Get-Random
        }
        'how are you' { 
            return "I'm functioning well, thank you for asking! What would you like to discuss?" 
        }
        'your name' { 
            return "I'm your PowerShell Chat Assistant. I can search the web to help answer your questions." 
        }
        default {
            # For other queries, perform a web search
            Write-Host "[Bot is searching the web...]" -ForegroundColor Yellow
            return Get-WebAnswer -Query $UserInput
        }
    }
}

# Main chat loop
Write-Host "`n=== PowerShell Smart Chat Assistant ===" -ForegroundColor Cyan
Write-Host "Type 'quit', 'exit', or 'bye' to end the conversation.`n"

try {
    # Initialize the Chrome browser[citation:1][citation:2]
    $ChromeArgs = @{ WebDriverDirectory = '.' }
    if ($Headless) { $ChromeArgs.Headless = $true }
    $Driver = Start-SeChrome @ChromeArgs
    
    # Main conversation loop
    do {
        $userInput = Read-Host "`nYou"
        
        if ($userInput.Trim() -notmatch '^(quit|exit|bye)$') {
            $response = Get-ChatResponse -UserInput $userInput
            Write-Host "Bot: $response" -ForegroundColor Green
        }
        
    } while ($userInput.Trim() -notmatch '^(quit|exit|bye)$')
    
} finally {
    # Always close the browser to free resources[citation:1][citation:4]
    if ($Driver) {
        $Driver.Quit()
        Write-Host "`nBrowser session closed. Goodbye!" -ForegroundColor Cyan
    }
}