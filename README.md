# GPTforGDocs
## AI Assistant for Google Docs
This is a menu extension for Google Docs that integrates with OpenAI and X.AI models. Select text in your document and use AI to analyze, generate, or transform it in various ways.

## Installation
1. Open a Google Doc and go to *Extensions->Apps Script*
2. Paste the script into the editor and save
3. Reload your document
4. Set up your API key(s) using the AI-settings menu

You can also deploy this as an add-on to make it available in all your Google docs.

## Available AI Providers
- OpenAI Models:
  - GPT-4o Mini 
  - GPT-4o 
  - O1 Mini 
  - O1 
- X.AI Models:
  - Grok-2

## Features
The extension adds two menus to your Google Docs:

### AI-Writer Menu
#### Content Generation
- *Continue this* - Seamlessly continue writing in the same style and tone
- *Create 5 LinkedIn Posts* - Generate professional social media content
- *Create 5 Twitter Posts* - Create engaging tweets
- *Write Email* - Generate professional emails
- *Generate Response* - Create well-structured responses

#### Analysis
- *Summarize* - Create a concise summary
- *Extract Keywords* - List key terms by relevance
- *Generate Key Points* - Create comprehensive bullet points
- *Competitive Analysis*
  - Generate Competitive Analysis
  - Generate Structured Profile
  - Create SWOT Analysis
  - Generate Comparison Analysis

#### Market Research
- *Analyze Market Segments*
- *Analyze Market Trends*
- *Generate Survey Questions*
- *Industry Analysis*
  - Porter's Five Forces Analysis
  - BCG Matrix Analysis

#### Reporting
- *Summarize Research Document*
- *Generate Recommendations*

#### Custom
- *Custom Instruction* - Use your own AI instructions

### AI-settings Menu
- Display current active model
- Set API keys (OpenAI and/or X.AI)
- Select AI model
- Delete API keys

## Setup
1. Click the AI-settings menu
2. Set your API key(s):
   - OpenAI API key for OpenAI models
   - X.AI API key for X.AI models
3. Select your preferred AI model
4. Start using the AI-Writer features!

## Usage
1. Select text in your document
2. Choose an action from the AI-Writer menu
3. Generated content will be inserted after your selection

## Credits
Original by J. Grant, 2023
Forked and modified by Or Zelig, 2025

## Links
- Original: https://github.com/jedediahg/GPTforGDocs
- Fork: https://github.com/orzelig/GDocs-OpenAI

## License
Apache 2.0
