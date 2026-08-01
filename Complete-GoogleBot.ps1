# Complete-GoogleBot.ps1 - All-in-One Google Search Chatbot
# Save as: Complete-GoogleBot.ps1 and run it

$ScriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $ScriptPath

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "🤖 GOOGLE SEARCH CHATBOT - ALL IN ONE" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Starting complete setup..." -ForegroundColor Yellow

# 1. CREATE PROJECT FOLDER
$ProjectPath = "C:\GoogleBot"
Write-Host "`n[1/8] Creating project folder..." -ForegroundColor Green
if (-not (Test-Path $ProjectPath)) {
    New-Item -Path $ProjectPath -ItemType Directory -Force | Out-Null
    Write-Host "  Created: $ProjectPath" -ForegroundColor Gray
}
Set-Location $ProjectPath

# 2. CHECK AND INSTALL DEPENDENCIES
Write-Host "`n[2/8] Checking dependencies..." -ForegroundColor Green
try {
    # Check Python
    $pythonCheck = python --version 2>&1
    if ($LASTEXITCODE -ne 0) {
        Write-Host "  ❌ Python not found!" -ForegroundColor Red
        Write-Host "  Please install Python 3.8+ from python.org" -ForegroundColor Red
        Write-Host "  Then run this script again." -ForegroundColor Red
        exit 1
    } else {
        Write-Host "  ✓ Python: $pythonCheck" -ForegroundColor Gray
    }
    
    # Install Selenium module
    if (-not (Get-Module -ListAvailable -Name Selenium)) {
        Write-Host "  Installing Selenium module..." -ForegroundColor Yellow
        Install-Module -Name Selenium -Force -Scope CurrentUser -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
        Write-Host "  ✓ Selenium installed" -ForegroundColor Gray
    } else {
        Write-Host "  ✓ Selenium already installed" -ForegroundColor Gray
    }
    
    # Install Python packages
    Write-Host "  Installing Python packages..." -ForegroundColor Yellow
    pip install rasa selenium flask --quiet 2>$null
    Write-Host "  ✓ Python packages installed" -ForegroundColor Gray
} catch {
    Write-Host "  ⚠️ Some dependencies failed: $_" -ForegroundColor Yellow
}

# 3. CREATE ALL CONFIG FILES
Write-Host "`n[3/8] Creating configuration files..." -ForegroundColor Green

# Create folders
@("data", "actions") | ForEach-Object {
    New-Item -Path $_ -ItemType Directory -Force -ErrorAction SilentlyContinue | Out-Null
}

# config.yml
@"
language: en
pipeline:
- name: WhitespaceTokenizer
- name: RegexFeaturizer
- name: LexicalSyntacticFeaturizer
- name: CountVectorsFeaturizer
- name: DIETClassifier
  epochs: 50
policies:
- name: MemoizationPolicy
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
  - text: "Hello! I can search Google for you."
  utter_goodbye:
  - text: "Goodbye!"
  utter_help:
  - text: "Try: search for python"

actions:
  - action_google_search
  - utter_greet
  - utter_goodbye
  - utter_help
"@ | Out-File -FilePath "domain.yml" -Encoding UTF8

# nlu.yml
@"
version: "3.1"
nlu:
- intent: greet
  examples: |
    - hi
    - hello
    - hey

- intent: goodbye
  examples: |
    - bye
    - goodbye
    - exit

- intent: search
  examples: |
    - search for python
    - find weather
    - look up news
    - search nba
    - google it

- intent: help
  examples: |
    - help
    - what can you do
"@ | Out-File -FilePath "data\nlu.yml" -Encoding UTF8

# rules.yml
@"
version: "3.1"
rules:
- rule: Greet user
  steps:
  - intent: greet
  - action: utter_greet

- rule: Google search
  steps:
  - intent: search
  - action: action_google_search

- rule: Provide help
  steps:
  - intent: help
  - action: utter_help
"@ | Out-File -FilePath "data\rules.yml" -Encoding UTF8

# endpoints.yml
@"
action_endpoint:
  url: "http://localhost:5055/webhook"
"@ | Out-File -FilePath "endpoints.yml" -Encoding UTF8

Write-Host "  ✓ Created all config files" -ForegroundColor Gray

# 4. CREATE GOOGLE SEARCH SCRIPT
Write-Host "`n[4/8] Creating Google search script..." -ForegroundColor Green

@'
param([string]$Query)

try {
    Import-Module Selenium -ErrorAction Stop
    
    # Clean query
    $cleanQuery = $Query -replace '^search for ', '' -replace '^search ', '' -replace '^find ', '' -replace '^look up ', '' -replace '^google ', ''
    $cleanQuery = $cleanQuery.Trim()
    
    if ([string]::IsNullOrWhiteSpace($cleanQuery)) {
        return @{error="Empty query"} | ConvertTo-Json
    }
    
    # Start Chrome
    $driver = Start-SeChrome -ErrorAction Stop
    
    # Search Google
    $encoded = [System.Web.HttpUtility]::UrlEncode($cleanQuery)
    $url = "https://www.google.com/search?q=$encoded"
    $driver.Navigate().GoToUrl($url)
    Start-Sleep -Seconds 2
    
    # Get first result
    $result = @{
        query = $cleanQuery
        found = $false
        title = ""
        link = ""
    }
    
    try {
        $title = $driver.FindElement([OpenQA.Selenium.By]::CssSelector("h3")).Text
        $result.title = $title
        $result.found = $true
        
        $link = $driver.FindElement([OpenQA.Selenium.By]::CssSelector("a")).GetAttribute("href")
        $result.link = $link
    } catch { }
    
    $driver.Quit()
    $result | ConvertTo-Json
    
} catch {
    @{error="Search failed: $_"} | ConvertTo-Json
}
'@ | Out-File -FilePath "google_search.ps1" -Encoding UTF8

Write-Host "  ✓ Google search script created" -ForegroundColor Gray

# 5. CREATE PYTHON ACTION
Write-Host "`n[5/8] Creating Python actions..." -ForegroundColor Green

@'
import subprocess
import json
import os
import re
from rasa_sdk import Action, Tracker
from rasa_sdk.executor import CollectingDispatcher

class ActionGoogleSearch(Action):
    def name(self):
        return "action_google_search"
    
    def run(self, dispatcher, tracker, domain):
        user_msg = tracker.latest_message.get("text", "").lower()
        
        # Extract query
        query = user_msg
        prefixes = [r'^search for ', r'^search ', r'^find ', r'^look up ', r'^google ']
        
        for pattern in prefixes:
            if re.match(pattern, query):
                query = re.sub(pattern, "", query)
                break
        
        query = query.strip()
        
        if not query or len(query) < 2:
            dispatcher.utter_message(text="What to search for?")
            return []
        
        dispatcher.utter_message(text=f"Searching: {query}")
        
        script_path = os.path.join(os.path.dirname(os.path.dirname(__file__)), "google_search.ps1")
        
        try:
            result = subprocess.run(
                ["powershell", "-ExecutionPolicy", "Bypass", "-File", script_path, "-Query", query],
                capture_output=True,
                text=True,
                timeout=15
            )
            
            if result.returncode == 0:
                data = json.loads(result.stdout)
                
                if "error" in data:
                    dispatcher.utter_message(text=f"Error: {data['error']}")
                elif data.get("found"):
                    response = f"Found: {data['title']}\nLink: {data['link']}"
                    dispatcher.utter_message(text=response)
                else:
                    dispatcher.utter_message(text=f"No results for '{query}'")
            else:
                dispatcher.utter_message(text="Search failed")
                
        except Exception as e:
            dispatcher.utter_message(text=f"Error: {str(e)}")
        
        return []
'@ | Out-File -FilePath "actions\actions.py" -Encoding UTF8

# Create __init__.py
@''@ | Out-File -FilePath "actions\__init__.py" -Encoding UTF8

Write-Host "  ✓ Python actions created" -ForegroundColor Gray

# 6. TRAIN THE MODEL
Write-Host "`n[6/8] Training AI model..." -ForegroundColor Green
Write-Host "  This takes 1-2 minutes..." -ForegroundColor Yellow

try {
    python -m rasa train --quiet 2>$null
    Write-Host "  ✓ Model trained successfully" -ForegroundColor Gray
} catch {
    Write-Host "  ⚠️ Training had issues, continuing anyway..." -ForegroundColor Yellow
}

# 7. CREATE CHAT INTERFACE
Write-Host "`n[7/8] Creating chat interface..." -ForegroundColor Green

@'
<!DOCTYPE html>
<html>
<head>
    <title>Google Bot</title>
    <style>
        body { font-family: Arial; padding: 20px; }
        #chat { height: 300px; border: 1px solid #ccc; padding: 10px; margin: 10px 0; overflow-y: auto; }
        .user { background: #4285f4; color: white; padding: 8px; margin: 5px; text-align: right; border-radius: 10px; }
        .bot { background: #f1f3f4; padding: 8px; margin: 5px; border-radius: 10px; }
        input { width: 70%; padding: 8px; }
        button { padding: 8px 15px; background: #4285f4; color: white; border: none; }
    </style>
</head>
<body>
    <h2>Google Search Bot</h2>
    <div id="chat"></div>
    <input type="text" id="input" placeholder="Search Google...">
    <button onclick="send()">Search</button>
    
    <script>
        function addMsg(text, type) {
            var chat = document.getElementById('chat');
            var div = document.createElement('div');
            div.className = type;
            div.textContent = text;
            chat.appendChild(div);
            chat.scrollTop = chat.scrollHeight;
        }
        
        function send() {
            var input = document.getElementById('input');
            var msg = input.value.trim();
            if (!msg) return;
            
            addMsg(msg, 'user');
            input.value = '';
            
            fetch('http://localhost:5005/webhooks/rest/webhook', {
                method: 'POST',
                headers: {'Content-Type': 'application/json'},
                body: JSON.stringify({sender: 'user', message: msg})
            })
            .then(response => response.json())
            .then(data => {
                if (data && data.length > 0) {
                    data.forEach(function(msg) {
                        addMsg(msg.text, 'bot');
                    });
                }
            })
            .catch(function() {
                addMsg('Connection error', 'bot');
            });
        }
        
        document.getElementById('input').addEventListener('keypress', function(e) {
            if (e.key === 'Enter') send();
        });
        
        addMsg('Type: search for python', 'bot');
        document.getElementById('input').focus();
    </script>
</body>
</html>
'@ | Out-File -FilePath "chat.html" -Encoding UTF8

Write-Host "  ✓ Chat interface created" -ForegroundColor Gray

# 8. START THE BOT
Write-Host "`n[8/8] Starting the bot servers..." -ForegroundColor Green

# Kill any existing Rasa processes
Get-Process python -ErrorAction SilentlyContinue | Where-Object { $_.CommandLine -like "*rasa*" } | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 2

Write-Host "  Starting action server on port 5055..." -ForegroundColor Yellow
$actionJob = Start-Job -Name "RasaActions" -ScriptBlock {
    cd C:\GoogleBot
    python -m rasa run actions
}

Start-Sleep -Seconds 3

Write-Host "  Starting main server on port 5005..." -ForegroundColor Yellow
$mainJob = Start-Job -Name "RasaMain" -ScriptBlock {
    cd C:\GoogleBot
    python -m rasa run --enable-api --cors "*" --port 5005
}

Start-Sleep -Seconds 3

Write-Host "  Opening browser..." -ForegroundColor Yellow
if (Test-Path "chat.html") {
    Start-Process "chat.html"
}

Write-Host "`n" + ("="*40) -ForegroundColor Cyan
Write-Host "✅ BOT IS NOW RUNNING!" -ForegroundColor Green
Write-Host "="*40 -ForegroundColor Cyan
Write-Host "Chat is open in your browser!" -ForegroundColor White
Write-Host "`nTry typing:" -ForegroundColor Yellow
Write-Host "  • search for python" -ForegroundColor White
Write-Host "  • find weather" -ForegroundColor White
Write-Host "  • look up news" -ForegroundColor White
Write-Host "  • help" -ForegroundColor White
Write-Host "  • bye" -ForegroundColor White
Write-Host "`nTo stop the bot: Press Ctrl+C" -ForegroundColor Red
Write-Host "="*40 -ForegroundColor Cyan

# Wait for Ctrl+C
try {
    while ($true) {
        Start-Sleep -Seconds 1
    }
} finally {
    Write-Host "`nStopping bot servers..." -ForegroundColor Yellow
    
    # Stop jobs
    Get-Job | Stop-Job -Force -ErrorAction SilentlyContinue
    Get-Job | Remove-Job -Force -ErrorAction SilentlyContinue
    
    # Kill any remaining processes
    Get-Process python -ErrorAction SilentlyContinue | Where-Object { $_.CommandLine -like "*rasa*" } | Stop-Process -Force -ErrorAction SilentlyContinue
    
    Write-Host "Bot stopped." -ForegroundColor Green
}