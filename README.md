# GPTforGDocs
## Open AI assistant for Google Docs
This is a simple menu extension for Google Docs. With it you can select elements in the docs that contain text and have GPT take action on it. 

## Installation
To install, simply open a Google Doc and go to *Extensions->Apps Script*. Paste the script into the editor and save. Reload your document. You can also deploy this as an add-on to make it available in all your Google docs.

## Models and Pricing
The following models are available:
- GPT-4o Mini 
- GPT-4o 
- O1 Mini 
- O1 

## Usage
Once installed, you will find a ChatGPT menu with several items organized into categories:

### Content Generation
- *Generate Essay* - Write an essay about the topic of the selected text
- *Continue Story* - Continue telling a story based on the initial selected text
- *Continue Report* - Continue writing a report based on the initial selected text

### Analysis
- *Summarize* - Create a summary paragraph of the selected text
- *Extract Keywords* - Create a list of keywords in the selected text
- *Generate Key Points* - Generate a set of bullet points based on the selected text

### Social Media
- *LinkedIn Posts* - Generate LinkedIn posts that could be used to promote the idea in the selected text
- *Twitter Posts* - Generate Tweets that could be used to promote the idea in the selected text

### Business
- *Write Email* - Generate a professional email based on the selected text
- *Generate Response* - Create a response to the selected text

### Other Features
- *Translate to English* - Translate the selected text to English
- *Set API Key* - Configure your OpenAI API key
- *Set AI Model* - Choose between available AI models
- *Help* - View usage instructions

Generated text is inserted immediately after the selected text. If this doesn't happen, check the start of the document.

## Setup
1. Set your OpenAI API key using the "Set API Key" menu option
2. (Optional) Choose your preferred model using "Set AI Model" - GPT-4o Mini is the default

## Credits
Original by J. Grant, 2023
Forked and modified by Or Zelig, 2025

## Links
- Original: https://github.com/jedediahg/GPTforGDocs
- Fork: https://github.com/orzelig/GDocs-OpenAI

## License
Apache 2.0
