# GoogleBot.ps1 - Complete Google Search Chatbot
# Save this as: GoogleBot.ps1
# Run from PowerShell: .\GoogleBot.ps1

# Setup
$ProjectPath = "C:\GoogleBot"
if (-not (Test-Path $ProjectPath)) {
    New-Item -Path $ProjectPath -ItemType Directory -Force | Out-Null
}
Set-Location $ProjectPath

Write-Host "=== GOOGLE SEARCH BOT ===" -ForegroundColor Cyan
Write-Host "Setting up... This may take a few minutes." -ForegroundColor Yellow

# 1. Install required modules
Write-Host "`n[1/10] Checking dependencies..." -ForegroundColor Green
try {
    # Install Selenium if needed
    if (-not (Get-Module -ListAvailable -Name Selenium)) {
        Write-Host "  Installing Selenium module..." -ForegroundColor Yellow
        Install-Module -Name Selenium -Force -Scope CurrentUser -Confirm:$false -ErrorAction SilentlyContinue
    }
    
    # Check Python
    $pythonCheck = python --version 2>&1
    if ($LASTEXITCODE -ne 0) {
        Write-Host "  ERROR: Python not found!" -ForegroundColor Red
        Write-Host "  Please install Python from python.org" -ForegroundColor Red
        exit 1
    }
    
    # Check Rasa
    $rasaCheck = rasa --version 2>&1
    if ($LASTEXITCODE -ne 0) {
        Write-Host "  Installing Rasa..." -ForegroundColor Yellow
        pip install rasa --quiet
    }
    
    # Install required Python packages
    Write-Host "  Installing Python packages..." -ForegroundColor Yellow
    pip install selenium flask --quiet
    
} catch {
    Write-Host "  Dependency check failed: $_" -ForegroundColor Red
}

# 2. Create all necessary files
Write-Host "`n[2/10] Creating configuration files..." -ForegroundColor Green

# config.yml
@"
language: en
pipeline:
- name: WhitespaceTokenizer
- name: RegexFeaturizer
- name: LexicalSyntacticFeaturizer
- name: CountVectorsFeaturizer
  analyzer: "word"
- name: DIETClassifier
  epochs: 50
  entity_recognition: true
policies:
- name: MemoizationPolicy
  max_history: 5
- name: RulePolicy
- name: TEDPolicy
  max_history: 5
  epochs: 50
"@ | Out-File -FilePath "config.yml" -Encoding UTF8

# domain.yml
@"
version: "3.1"
intents:
  - greet
  - goodbye
  - search
  - help

responses:
  utter_greet:
  - text: "🔍 Hello! I'm Google Search Bot. Ask me to search for anything!"
  utter_goodbye:
  - text: "👋 Goodbye! Come back if you need more searches."
  utter_help:
  - text: "🤖 I can search Google for you. Just say: search [something], find [something], or look up [something]"

actions:
  - action_google_search
  - utter_greet
  - utter_goodbye
  - utter_help
"@ | Out-File -FilePath "domain.yml" -Encoding UTF8

# Create data folder if not exists
New-Item -Path "data" -ItemType Directory -Force -ErrorAction SilentlyContinue | Out-Null

# nlu.yml
@"
version: "3.1"
nlu:
- intent: greet
  examples: |
    - hi
    - hello
    - hey
    - hi there
    - hello there
    - good morning
    - hey bot

- intent: goodbye
  examples: |
    - bye
    - goodbye
    - see you
    - goodbye bot
    - exit
    - quit

- intent: search
  examples: |
    - search for python
    - search python
    - find weather
    - look up news
    - google search
    - search nba
    - find nba scores
    - look up time
    - search for time
    - find information
    - look up tutorials
    - search google
    - google it
    - search the web
    - find results
    - look something up

- intent: help
  examples: |
    - help
    - what can you do
    - how does this work
    - what is this
    - help me
    - what can i ask
    - instructions
"@ | Out-File -FilePath "data\nlu.yml" -Encoding UTF8

# rules.yml
@"
version: "3.1"
rules:
- rule: Say greet
  steps:
  - intent: greet
  - action: utter_greet

- rule: Say goodbye
  steps:
  - intent: goodbye
  - action: utter_goodbye

- rule: Provide help
  steps:
  - intent: help
  - action: utter_help

- rule: Google search
  steps:
  - intent: search
  - action: action_google_search
"@ | Out-File -FilePath "data\rules.yml" -Encoding UTF8

# endpoints.yml
@"
action_endpoint:
  url: "http://localhost:5055/webhook"
"@ | Out-File -FilePath "endpoints.yml" -Encoding UTF8

# 3. Create Google search PowerShell script
Write-Host "`n[3/10] Creating Google search script..." -ForegroundColor Green

@'
# google_search.ps1 - Google Search Script
param([string]$Query)

# Function to clean query
function Clean-Query {
    param([string]$text)
    $text = $text -replace '^search for ', ''
    $text = $text -replace '^search ', ''
    $text = $text -replace '^find ', ''
    $text = $text -replace '^look up ', ''
    $text = $text -replace '^google ', ''
    return $text.Trim()
}

try {
    # Load Selenium module
    Import-Module Selenium -ErrorAction Stop
    
    # Clean the query
    $cleanQuery = Clean-Query $Query
    
    if ([string]::IsNullOrWhiteSpace($cleanQuery)) {
        return @{error="Empty query"} | ConvertTo-Json
    }
    
    Write-Host "Searching Google for: $cleanQuery" -ForegroundColor Yellow
    
    # Setup Chrome options
    $chromeOptions = New-Object OpenQA.Selenium.Chrome.ChromeOptions
    $chromeOptions.AddArgument("--headless")  # Run in background
    $chromeOptions.AddArgument("--no-sandbox")
    $chromeOptions.AddArgument("--disable-dev-shm-usage")
    $chromeOptions.AddArgument("--disable-gpu")
    $chromeOptions.AddArgument("--window-size=1920,1080")
    
    # Start Chrome
    $driver = Start-SeChrome -Options $chromeOptions -ErrorAction Stop
    
    # Encode query for URL
    $encodedQuery = [System.Web.HttpUtility]::UrlEncode($cleanQuery)
    $searchUrl = "https://www.google.com/search?q=$encodedQuery"
    
    # Navigate to Google
    $driver.Navigate().GoToUrl($searchUrl)
    Start-Sleep -Seconds 2
    
    # Initialize result
    $result = @{
        query = $cleanQuery
        found = $false
        title = ""
        link = ""
        snippet = ""
    }
    
    # Try to get first result
    try {
        # Get first title
        $firstTitle = $driver.FindElements([OpenQA.Selenium.By]::CssSelector("h3")) | Select-Object -First 1
        if ($firstTitle) {
            $result.title = $firstTitle.Text
            $result.found = $true
        }
        
        # Get parent link of the title
        if ($firstTitle) {
            $parentDiv = $firstTitle.FindElement([OpenQA.Selenium.By]::XPath("./ancestor::a[1]"))
            if ($parentDiv) {
                $result.link = $parentDiv.GetAttribute("href")
            }
        }
        
        # Try to get snippet
        try {
            $snippet = $driver.FindElements([OpenQA.Selenium.By]::CssSelector("div[data-sncf]")) | Select-Object -First 1
            if ($snippet) {
                $result.snippet = $snippet.Text.Substring(0, [Math]::Min(150, $snippet.Text.Length)) + "..."
            }
        } catch { }
        
    } catch {
        # Alternative method
        try {
            $links = $driver.FindElements([OpenQA.Selenium.By]::CssSelector("a h3"))
            if ($links.Count -gt 0) {
                $result.title = $links[0].Text
                $result.found = $true
                
                $linkElement = $links[0].FindElement([OpenQA.Selenium.By]::XPath("./ancestor::a[1]"))
                $result.link = $linkElement.GetAttribute("href")
            }
        } catch { }
    }
    
    # Close browser
    $driver.Quit()
    
    # Return result
    $result | ConvertTo-Json
    
} catch {
    $errorMsg = $_.Exception.Message
    Write-Host "Search error: $errorMsg" -ForegroundColor Red
    
    # Return error
    @{
        query = $Query
        found = $false
        error = "Search failed: $errorMsg"
        title = ""
        link = ""
        snippet = ""
    } | ConvertTo-Json
}
'@ | Out-File -FilePath "google_search.ps1" -Encoding UTF8

# 4. Create actions folder and Python action
Write-Host "`n[4/10] Creating actions..." -ForegroundColor Green
New-Item -Path "actions" -ItemType Directory -Force -ErrorAction SilentlyContinue | Out-Null

# actions/actions.py
@'
# actions.py - Google Search Action
import subprocess
import json
import os
import re
from rasa_sdk import Action, Tracker
from rasa_sdk.executor import CollectingDispatcher

class ActionGoogleSearch(Action):
    def name(self) -> str:
        return "action_google_search"
    
    def run(self, dispatcher: CollectingDispatcher,
            tracker: Tracker,
            domain: dict) -> list:
        
        # Get the user's message
        user_message = tracker.latest_message.get('text', '').strip()
        
        # Clean the query - remove common prefixes
        query = user_message.lower()
        
        # List of prefixes to remove
        prefixes = [
            r'^search for ',
            r'^search ',
            r'^find ',
            r'^look up ',
            r'^google ',
            r'^search the web for ',
            r'^search google for ',
        ]
        
        for prefix in prefixes:
            if re.match(prefix, query):
                query = re.sub(prefix, '', query)
                break
        
        # If query is too short or just "search", ask for clarification
        if not query or len(query) < 2 or query in ['search', 'find', 'look', 'google']:
            dispatcher.utter_message(text="🔍 What would you like me to search for?")
            return []
        
        # Show searching message
        dispatcher.utter_message(text=f"🔍 Searching for: **{query}**")
        
        # Path to PowerShell script
        script_path = os.path.join(os.path.dirname(os.path.dirname(__file__)), "google_search.ps1")
        
        try:
            # Run PowerShell script
            result = subprocess.run(
                ["powershell", "-ExecutionPolicy", "Bypass", "-File", script_path, "-Query", query],
                capture_output=True,
                text=True,
                timeout=30,
                encoding='utf-8'
            )
            
            # Debug: Print output
            print(f"PowerShell Output: {result.stdout}")
            print(f"PowerShell Error: {result.stderr}")
            
            if result.returncode == 0 and result.stdout.strip():
                data = json.loads(result.stdout)
                
                if "error" in data:
                    dispatcher.utter_message(text=f"❌ Error: {data['error']}")
                elif data.get("found"):
                    title = data.get("title", "No title")
                    link = data.get("link", "No link")
                    snippet = data.get("snippet", "")
                    
                    response = f"✅ **Found:** {title}\n"
                    if snippet:
                        response += f"📝 {snippet}\n"
                    response += f"🔗 {link}"
                    
                    dispatcher.utter_message(text=response)
                else:
                    dispatcher.utter_message(text=f"⚠️ No results found for '{query}'")
                    
            else:
                error_msg = result.stderr if result.stderr else "Unknown error"
                dispatcher.utter_message(text=f"❌ Search failed: {error_msg}")
                
        except json.JSONDecodeError as e:
            dispatcher.utter_message(text=f"❌ Error parsing search results: {str(e)}")
        except subprocess.TimeoutExpired:
            dispatcher.utter_message(text="❌ Search timed out. Please try again.")
        except Exception as e:
            dispatcher.utter_message(text=f"❌ Error: {str(e)}")
        
        return []
'@ | Out-File -FilePath "actions\actions.py" -Encoding UTF8

# Create __init__.py in actions folder
@'
# This file makes the actions directory a Python package
'@ | Out-File -FilePath "actions\__init__.py" -Encoding UTF8

# 5. Train the Rasa model
Write-Host "`n[5/10] Training the AI model..." -ForegroundColor Green
Write-Host "This may take 1-2 minutes..." -ForegroundColor Yellow

try {
    rasa train --quiet
    Write-Host "✓ Model trained successfully!" -ForegroundColor Green
} catch {
    Write-Host "✗ Model training failed: $_" -ForegroundColor Red
    # Try alternative training
    python -m rasa train --quiet
}

# 6. Create HTML chat interface
Write-Host "`n[6/10] Creating chat interface..." -ForegroundColor Green

@'
<!DOCTYPE html>
<html>
<head>
    <title>Google Search Bot 🤖</title>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <style>
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
        }
        
        body {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            min-height: 100vh;
            display: flex;
            justify-content: center;
            align-items: center;
            padding: 20px;
        }
        
        .container {
            width: 100%;
            max-width: 800px;
            background: white;
            border-radius: 20px;
            box-shadow: 0 20px 60px rgba(0,0,0,0.3);
            overflow: hidden;
        }
        
        .header {
            background: linear-gradient(135deg, #4285f4 0%, #34a853 100%);
            color: white;
            padding: 25px;
            text-align: center;
        }
        
        .header h1 {
            font-size: 28px;
            margin-bottom: 10px;
            display: flex;
            align-items: center;
            justify-content: center;
            gap: 10px;
        }
        
        .header p {
            opacity: 0.9;
            font-size: 16px;
        }
        
        .chat-container {
            height: 500px;
            padding: 20px;
            overflow-y: auto;
            display: flex;
            flex-direction: column;
            gap: 15px;
        }
        
        .message {
            max-width: 80%;
            padding: 15px;
            border-radius: 18px;
            line-height: 1.4;
            animation: fadeIn 0.3s ease;
        }
        
        @keyframes fadeIn {
            from { opacity: 0; transform: translateY(10px); }
            to { opacity: 1; transform: translateY(0); }
        }
        
        .user-message {
            background: #4285f4;
            color: white;
            align-self: flex-end;
            border-bottom-right-radius: 5px;
        }
        
        .bot-message {
            background: #f1f3f4;
            color: #202124;
            align-self: flex-start;
            border-bottom-left-radius: 5px;
        }
        
        .input-container {
            padding: 20px;
            border-top: 1px solid #e0e0e0;
            display: flex;
            gap: 10px;
        }
        
        #messageInput {
            flex: 1;
            padding: 15px 20px;
            border: 2px solid #e0e0e0;
            border-radius: 50px;
            font-size: 16px;
            outline: none;
            transition: border-color 0.3s;
        }
        
        #messageInput:focus {
            border-color: #4285f4;
        }
        
        #sendButton {
            background: #4285f4;
            color: white;
            border: none;
            border-radius: 50%;
            width: 56px;
            height: 56px;
            cursor: pointer;
            font-size: 20px;
            transition: background 0.3s;
            display: flex;
            align-items: center;
            justify-content: center;
        }
        
        #sendButton:hover {
            background: #3367d6;
        }
        
        .examples {
            padding: 15px 25px;
            background: #f8f9fa;
            border-top: 1px solid #e0e0e0;
        }
        
        .examples h3 {
            margin-bottom: 10px;
            color: #5f6368;
            font-size: 14px;
            text-transform: uppercase;
            letter-spacing: 1px;
        }
        
        .example-buttons {
            display: flex;
            flex-wrap: wrap;
            gap: 10px;
        }
        
        .example-button {
            background: white;
            border: 1px solid #dadce0;
            border-radius: 20px;
            padding: 8px 16px;
            font-size: 14px;
            cursor: pointer;
            transition: all 0.3s;
        }
        
        .example-button:hover {
            background: #f1f3f4;
            border-color: #4285f4;
        }
        
        .typing-indicator {
            display: none;
            padding: 15px;
            align-self: flex-start;
        }
        
        .typing-dots {
            display: flex;
            gap: 4px;
        }
        
        .typing-dots span {
            width: 8px;
            height: 8px;
            background: #5f6368;
            border-radius: 50%;
            animation: typing 1.4s infinite;
        }
        
        .typing-dots span:nth-child(2) { animation-delay: 0.2s; }
        .typing-dots span:nth-child(3) { animation-delay: 0.4s; }
        
        @keyframes typing {
            0%, 100% { opacity: 0.4; transform: translateY(0); }
            50% { opacity: 1; transform: translateY(-5px); }
        }
        
        .bot-icon {
            font-size: 24px;
        }
        
        .search-icon {
            margin-right: 5px;
        }
        
        .timestamp {
            font-size: 12px;
            opacity: 0.6;
            margin-top: 5px;
            text-align: right;
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1><span class="bot-icon">🤖</span> Google Search Bot</h1>
            <p>Ask me to search for anything on Google!</p>
        </div>
        
        <div class="examples">
            <h3>Try these:</h3>
            <div class="example-buttons">
                <button class="example-button" onclick="sendExample('search for python tutorials')">Python tutorials</button>
                <button class="example-button" onclick="sendExample('find weather in london')">London weather</button>
                <button class="example-button" onclick="sendExample('look up nba scores')">NBA scores</button>
                <button class="example-button" onclick="sendExample('search latest news')">Latest news</button>
                <button class="example-button" onclick="sendExample('help')">Help</button>
            </div>
        </div>
        
        <div class="chat-container" id="chatContainer">
            <div class="message bot-message">
                <span class="search-icon">🔍</span> Hello! I'm your Google Search Bot. I can search for anything on the web!<br><br>
                Try saying: <strong>search for python</strong> or <strong>find weather</strong>
            </div>
        </div>
        
        <div class="typing-indicator" id="typingIndicator">
            <div class="typing-dots">
                <span></span>
                <span></span>
                <span></span>
            </div>
        </div>
        
        <div class="input-container">
            <input type="text" id="messageInput" placeholder="Type your search query..." autocomplete="off">
            <button id="sendButton" onclick="sendMessage()">➤</button>
        </div>
    </div>

    <script>
        const chatContainer = document.getElementById('chatContainer');
        const messageInput = document.getElementById('messageInput');
        const typingIndicator = document.getElementById('typingIndicator');
        
        // Get current time for timestamps
        function getCurrentTime() {
            const now = new Date();
            return now.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
        }
        
        // Add message to chat
        function addMessage(text, isUser = false) {
            const messageDiv = document.createElement('div');
            messageDiv.className = isUser ? 'message user-message' : 'message bot-message';
            
            const timeSpan = document.createElement('div');
            timeSpan.className = 'timestamp';
            timeSpan.textContent = getCurrentTime();
            
            messageDiv.innerHTML = text;
            messageDiv.appendChild(timeSpan);
            
            chatContainer.appendChild(messageDiv);
            chatContainer.scrollTop = chatContainer.scrollHeight;
        }
        
        // Show/hide typing indicator
        function showTyping(show) {
            typingIndicator.style.display = show ? 'flex' : 'none';
            chatContainer.scrollTop = chatContainer.scrollHeight;
        }
        
        // Send message to bot
        async function sendMessage() {
            const text = messageInput.value.trim();
            if (!text) return;
            
            // Clear input and add user message
            messageInput.value = '';
            addMessage(text, true);
            
            // Show typing indicator
            showTyping(true);
            
            try {
                const response = await fetch('http://localhost:5005/webhooks/rest/webhook', {
                    method: 'POST',
                    headers: {
                        'Content-Type': 'application/json',
                    },
                    body: JSON.stringify({
                        sender: 'user',
                        message: text
                    })
                });
                
                const data = await response.json();
                
                // Hide typing indicator
                showTyping(false);
                
                // Add bot response
                if (data && data.length > 0) {
                    data.forEach(msg => {
                        if (msg.text) {
                            addMessage(msg.text, false);
                        }
                    });
                } else {
                    addMessage("I'm here! What would you like me to search for?", false);
                }
                
            } catch (error) {
                showTyping(false);
                addMessage("Oops! Connection error. Make sure the bot server is running.", false);
                console.error('Error:', error);
            }
        }
        
        // Send example message
        function sendExample(text) {
            messageInput.value = text;
            sendMessage();
        }
        
        // Send on Enter key
        messageInput.addEventListener('keypress', (e) => {
            if (e.key === 'Enter') {
                sendMessage();
            }
        });
        
        // Focus input on load
        messageInput.focus();
        
        // Auto-resize input
        messageInput.addEventListener('input', function() {
            this.style.height = 'auto';
            this.style.height = (this.scrollHeight) + 'px';
        });
    </script>
</body>
</html>
'@ | Out-File -FilePath "chat.html" -Encoding UTF8

# 7. Create a simple test script
Write-Host "`n[7/10] Creating test script..." -ForegroundColor Green

@'
# test-bot.ps1 - Test the bot
cd C:\GoogleBot

Write-Host "Testing Google Search Bot..." -ForegroundColor Cyan
Write-Host "`nTest 1: Testing Google search script..." -ForegroundColor Yellow

# Test the PowerShell search script directly
$testQuery = "Python programming"
.\google_search.ps1 -Query $testQuery

Write-Host "`nTest 2: Testing Rasa NLU..." -ForegroundColor Yellow

# Test messages
$testMessages = @(
    "hello",
    "search for python",
    "find weather",
    "help",
    "bye"
)

foreach ($msg in $testMessages) {
    Write-Host "`nTesting: '$msg'" -ForegroundColor Green
    
    $body = @{
        text = $msg
    } | ConvertTo-Json
    
    # Test with rasa shell nlu
    $body | Out-File -FilePath "temp_test.json" -Encoding UTF8
    rasa shell nlu --model models --quiet < "temp_test.json"
    Remove-Item "temp_test.json" -Force
}

Write-Host "`n✓ Tests complete!" -ForegroundColor Green
Write-Host "Run .\start-bot.ps1 to start the full bot" -ForegroundColor Cyan
'@ | Out-File -FilePath "test-bot.ps1" -Encoding UTF8

# 8. Create start script
Write-Host "`n[8/10] Creating start script..." -ForegroundColor Green

@'
# start-bot.ps1 - Start the Google Search Bot
cd C:\GoogleBot

Write-Host "=== GOOGLE SEARCH BOT ===" -ForegroundColor Cyan
Write-Host "Starting servers..." -ForegroundColor Yellow

# Kill any existing Rasa processes
Get-Process python -ErrorAction SilentlyContinue | Where-Object { $_.CommandLine -like "*rasa*" } | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 2

# Function to check if port is in use
function Test-Port {
    param([int]$Port)
    try {
        $socket = New-Object System.Net.Sockets.TcpClient
        $socket.Connect("localhost", $Port)
        $socket.Close()
        return $true
    } catch {
        return $false
    }
}

# Check if ports are available
if (Test-Port -Port 5055) {
    Write-Host "⚠️ Port 5055 is in use. Trying port 5056..." -ForegroundColor Yellow
    $actionPort = 5056
} else {
    $actionPort = 5055
}

if (Test-Port -Port 5005) {
    Write-Host "⚠️ Port 5005 is in use. Trying port 5006..." -ForegroundColor Yellow
    $mainPort = 5006
} else {
    $mainPort = 5005
}

# Update endpoints.yml with correct port
"action_endpoint:
  url: 'http://localhost:$actionPort/webhook'" | Out-File -FilePath "endpoints.yml" -Encoding UTF8

# Start actions server
Write-Host "Starting actions server on port $actionPort..." -ForegroundColor Green
Start-Job -Name "RasaActions" -ScriptBlock {
    param($port)
    cd C:\GoogleBot
    rasa run actions --port $port --cors "*"
} -ArgumentList $actionPort

Start-Sleep -Seconds 5

# Start main server
Write-Host "Starting main server on port $mainPort..." -ForegroundColor Green
Start-Job -Name "RasaMain" -ScriptBlock {
    param($port)
    cd C:\GoogleBot
    rasa run --enable-api --cors "*" --port $port --model models
} -ArgumentList $mainPort

Start-Sleep -Seconds 5

# Open browser
Write-Host "Opening chat interface..." -ForegroundColor Green
Start-Process "chat.html"

Write-Host "`n✅ BOT IS RUNNING!" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Chat Interface: Open chat.html in browser" -ForegroundColor White
Write-Host "Try these commands:" -ForegroundColor Yellow
Write-Host "  • 'search for python'" -ForegroundColor White
Write-Host "  • 'find weather'" -ForegroundColor White
Write-Host "  • 'look up news'" -ForegroundColor White
Write-Host "  • 'help' - for help" -ForegroundColor White
Write-Host "  • 'bye' - to exit" -ForegroundColor White
Write-Host "`nPress Ctrl+C to stop the bot" -ForegroundColor Red
Write-Host "========================================" -ForegroundColor Cyan

# Wait for Ctrl+C
try {
    while ($true) {
        Start-Sleep -Seconds 1
    }
} finally {
    Write-Host "`nStopping bot..." -ForegroundColor Yellow
    
    # Stop jobs
    Get-Job | Stop-Job -Force
    Get-Job | Remove-Job -Force
    
    # Kill any remaining processes
    Get-Process python -ErrorAction SilentlyContinue | Where-Object { $_.CommandLine -like "*rasa*" } | Stop-Process -Force -ErrorAction SilentlyContinue
    
    Write-Host "Bot stopped." -ForegroundColor Green
}
'@ | Out-File -FilePath "start-bot.ps1" -Encoding UTF8

# 9. Create cleanup script
Write-Host "`n[9/10] Creating cleanup script..." -ForegroundColor Green

@'
# cleanup.ps1 - Clean up and reset
cd C:\GoogleBot

Write-Host "Cleaning up..." -ForegroundColor Yellow

# Stop all jobs
Get-Job | Stop-Job -Force -ErrorAction SilentlyContinue
Get-Job | Remove-Job -Force -ErrorAction SilentlyContinue

# Kill Python processes
Get-Process python -ErrorAction SilentlyContinue | Where-Object { $_.CommandLine -like "*rasa*" } | Stop-Process -Force -ErrorAction SilentlyContinue

# Clean cache
Remove-Item -Path ".rasa\cache" -Force -Recurse -ErrorAction SilentlyContinue
Remove-Item -Path "models\*" -Force -Recurse -ErrorAction SilentlyContinue

# Remove __pycache__
Remove-Item -Path "actions\__pycache__" -Force -Recurse -ErrorAction SilentlyContinue

Write-Host "✓ Cleanup complete!" -ForegroundColor Green
Write-Host "Run .\GoogleBot.ps1 to setup again" -ForegroundColor Cyan
'@ | Out-File -FilePath "cleanup.ps1" -Encoding UTF8

# 10. Create README file
Write-Host "`n[10/10] Creating documentation..." -ForegroundColor Green

@'
# 🤖 Google Search Bot

A complete Google search chatbot powered by Rasa AI and PowerShell.

## 📋 Files Created:

1. **GoogleBot.ps1** - Main setup script (this file)
2. **start-bot.ps1** - Start the bot
3. **test-bot.ps1** - Test the bot
4. **cleanup.ps1** - Clean up and reset
5. **google_search.ps1** - PowerShell Google search script
6. **chat.html** - Beautiful chat interface
7. **config.yml** - Rasa configuration
8. **domain.yml** - Rasa domain
9. **data/nlu.yml** - Training data
10. **data/rules.yml** - Conversation rules
11. **actions/actions.py** - Python actions
12. **endpoints.yml** - Server endpoints

## 🚀 Quick Start:

1. **First time setup:**
   ```powershell
   .\GoogleBot.ps1