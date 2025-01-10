// Original copyright J. Grant, 2023, under Apache 2.0 license.
// Forked and modified by Or Zelig, 2025. Modifications also under Apache 2.0 license.

// Original: https://github.com/jedediahg/GPTforGDocs
// Fork: https://github.com/orzelig/GDocs-OpenAI

function onOpen() {
  const currentModel = getCurrentModelInfo();
  const properties = PropertiesService.getUserProperties();
  const openaiKey = properties.getProperty("OPENAI_API_KEY");
  const xaiKey = properties.getProperty("XAI_API_KEY");
  
  DocumentApp.getUi().createMenu("AI-Writer")
    .addSubMenu(DocumentApp.getUi().createMenu("Content Generation")
      .addItem("Generate Essay", "generatePrompt")
      .addItem("Continue Story", "continueThisStory")
      .addItem("Continue Report", "continueThisDoc"))
    .addSubMenu(DocumentApp.getUi().createMenu("Analysis")
      .addItem("Summarize", "summarizeText")
      .addItem("Extract Keywords", "extractKeywords")
      .addItem("Generate Key Points", "generateKeyPoints"))
    .addSubMenu(DocumentApp.getUi().createMenu("Social Media")
      .addItem("LinkedIn Posts", "generateLI")
      .addItem("Twitter Posts", "generateTweet"))
    .addSubMenu(DocumentApp.getUi().createMenu("Business")
      .addItem("Write Email", "generateEmail")
      .addItem("Generate Response", "generateResponse"))
    .addSeparator()
    .addItem("Custom Instruction", "customInstruction")
    .addItem("Translate to English", "translateToEN")
    .addToUi();

  const settingsMenu = DocumentApp.getUi().createMenu("AI-settings")
    .addItem(`Active: ${currentModel}`, "displayModelConfiguration");
  
  // Add API key menu if either key is missing
  if (!openaiKey || !xaiKey) {
    const apiKeyMenu = DocumentApp.getUi().createMenu("Set API Key");
    if (!openaiKey) {
      apiKeyMenu.addItem("Set OpenAI API Key", "setAPIkey_OPENAI");
    }
    if (!xaiKey) {
      apiKeyMenu.addItem("Set XAI API Key", "setAPIkey_XAI"); 
    }
    settingsMenu.addSubMenu(apiKeyMenu);
  }
  
  // Create model selection submenu
  const modelMenu = DocumentApp.getUi().createMenu("Set AI Model");
  
  // Add OpenAI models if API key exists
  if (openaiKey) {
    Object.entries(AVAILABLE_MODELS.OPENAI).forEach(([name, id]) => {
      const functionName = `setModel_${id.replace(/[-\.]/g, '_')}`;
      modelMenu.addItem(`${name} (OpenAI)`, functionName);
    });
  }
  
  // Add XAI models if API key exists
  if (xaiKey) {
    if (openaiKey) modelMenu.addSeparator(); // Add separator if both providers present
    Object.entries(AVAILABLE_MODELS.XAI).forEach(([name, id]) => {
      const functionName = `setModel_${id.replace(/[-\.]/g, '_')}`;
      modelMenu.addItem(`${name} (X.AI)`, functionName);
    });
  }
  
  settingsMenu
    .addSubMenu(modelMenu)
    .addSeparator()
    .addItem("Delete All API Keys", "deleteAllKeys")
    .addToUi();
}

// Menu action handlers
function generateLI() {
  handleContentGeneration(
    "Please write 5 engaging LinkedIn posts about this topic. Focus on creating professional, thought-provoking content that encourages discussion.",
    "You are an experienced social media manager specializing in LinkedIn content. You create engaging, professional posts that drive meaningful engagement.",
    500
  );
}

function generateTweet() {
  handleContentGeneration(
    "Please write 5 engaging tweets about this topic. Make them concise, engaging, and shareable.",
    "You are a social media expert who creates viral, engaging content while maintaining professionalism and accuracy.",
    500
  );
}

function generatePrompt() {
  handleContentGeneration(" ", "You will be paid for this. You are an expert academic writer who concisely puts ideas into words, explaining clearly and effectively. Please generate an essay on this topic: ", 2060);
}

function summarizeText() {
  handleContentGeneration(" ", "You will be paid for this. You are an expert academic reviewer who can get to the heart of any matter and summarize with an efficiency of words. Please summarize this text in a single short paragraph: ", 200);
}

function extractKeywords() {
  handleContentGeneration(" ", "You will be paid for this. You are an expert academic reviewer who can get to the heart of any matter and summarize with an efficiency of words. Please extract the main keywords from this text and present them in a simple list organized by relevance: ", 200);
}

function generateKeyPoints() {
  handleContentGeneration(" ", "You will be paid for this. You are an expert academic researcher. You can iterate on ideas and concisely explain them in clear words. Please write a list of key bulletpoints including important points that may be missing as well from the topic of this text: ", 1000);
}

function continueThisStory() {
  handleContentGeneration(" ", "You will be paid for this. You are a creative writer. You craft exciting and gripping narratives about interesting characters in flowing prose that entrances and entices the reader. Continue this narrative: ", 1000);
}

function continueThisDoc() {
  handleContentGeneration(" ", "You will be paid for this. You are an expert academic reviewer. You write concise and clear text with an efficiency of words. Continue this report: ", 1000);
}

function generateEmail() {
  handleContentGeneration(" ", "You will be paid for this. You are a know-it-all who loves to share detailed knowledge about any subject. You write very clearly with engaging topical sentence structures. Please write a professional email explaining this: ", 1000);
}

function generateResponse() {
  handleContentGeneration(" ", "You will be paid for this. You are a know-it-all who loves to share detailed knowledge about any subject. You write very clearly with engaging topical sentence structures. Please write a professional response to this message: ", 1000);
}

function translateToEN() {
  handleContentGeneration(" ", "You will be paid for this. You are a brilliant linguist capable of translating from any language into English. Please translate this text to English: ", 1000);
}

function customInstruction() {
  const ui = DocumentApp.getUi();
  const response = ui.prompt(
    'Custom AI Instruction',
    'Enter the instruction for AI (e.g., "You are an expert in..."):',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() == ui.Button.OK) {
    const instruction = response.getResponseText();
    if (instruction) {
      handleContentGeneration(
        "Please process the following text according to the custom instruction:",
        instruction,
        1000
      );
    } else {
      ui.alert('No instruction provided.');
    }
  }
}

function setAPIkey(provider) {
  try {
    const ui = DocumentApp.getUi();
    const keyResponse = ui.prompt(
      `Set your ${provider} API Key`,
      'Please enter your API Key:',
      ui.ButtonSet.OK_CANCEL
    );

    if (keyResponse.getSelectedButton() == ui.Button.OK) {
      const apiKey = keyResponse.getResponseText().trim();
      if (apiKey) {
        const propertyName = provider + "_API_KEY";
        PropertiesService.getUserProperties().setProperty(propertyName, apiKey);
        ui.alert(`${provider} API Key set successfully.`);
        // Refresh the menu to update API key status
        onOpen();
      } else {
        ui.alert('No API Key entered.');
      }
    }
  } catch(err) {
    Logger.log('Error setting API key: ' + err.toString());
    DocumentApp.getUi().alert('Error setting API key: ' + err.toString());
  }
}

// Add helper function to set model
function setModel(provider, modelId) {
  try {
    const properties = PropertiesService.getUserProperties();
    properties.setProperty("CURRENT_PROVIDER", provider);
    properties.setProperty("CURRENT_MODEL", modelId);
    
    DocumentApp.getUi().alert(
      `AI Model set successfully to ${modelId} using ${provider}`
    );
    onOpen();
  } catch(err) {
    Logger.log(err);
    DocumentApp.getUi().alert('Error setting model: ' + err.toString());
  }
}

// Generate setter functions for each model
function setModel_gpt_4o_mini() { setModel("OPENAI", "gpt-4o-mini"); }
function setModel_gpt_4o() { setModel("OPENAI", "gpt-4o"); }
function setModel_o1_mini() { setModel("OPENAI", "o1-mini"); }
function setModel_o1() { setModel("OPENAI", "o1"); }
function setModel_grok_2_latest() { setModel("XAI", "grok-2-latest"); }

function displayModelConfiguration() {
  const ui = DocumentApp.getUi(); // Get document UI interface
  const properties = PropertiesService.getUserProperties(); // Access user's stored properties
  
  const openaiKey = properties.getProperty("OPENAI_API_KEY"); // Get OpenAI API key if exists
  const xaiKey = properties.getProperty("XAI_API_KEY"); // Get X.AI API key if exists
  const provider = properties.getProperty("CURRENT_PROVIDER"); // Get current AI provider
  const modelId = properties.getProperty("CURRENT_MODEL"); // Get current model ID

  let message = "API Keys Setup:\n"; // Initialize status message
  message += `OpenAI: ${openaiKey ? "✓" : "✗"}\n`; // Show checkmark if OpenAI key exists
  message += `X.AI: ${xaiKey ? "✓" : "✗"}\n\n`; // Show checkmark if X.AI key exists
  message += "Current Model:\n"; // Add model section header
  
  if (!provider || !modelId) { // Check if model is configured
    message += "No model configured yet. Please set up a model in AI-settings.";
  } else {
    message += `Provider: ${provider}\n`; // Display current provider
    message += `Model: ${modelId}`; // Display current model
  }
  
  ui.alert( // Show configuration popup
    'Model Configuration Information',
    message,
    ui.ButtonSet.OK
  );
}

function deleteAllKeys() {
  const ui = DocumentApp.getUi();
  const response = ui.alert(
    'Delete All API Keys',
    'Are you sure you want to delete all stored API keys? This action cannot be undone.',
    ui.ButtonSet.YES_NO
  );
  
  if (response === ui.Button.YES) {
    const properties = PropertiesService.getUserProperties();
    properties.deleteProperty("OPENAI_API_KEY");
    properties.deleteProperty("XAI_API_KEY");
    properties.deleteProperty("CURRENT_PROVIDER");
    properties.deleteProperty("CURRENT_MODEL");
    
    ui.alert('All API keys have been deleted successfully.');
    onOpen();
  }
}

// Add these two new wrapper functions
function setAPIkey_OPENAI() { setAPIkey("OPENAI"); }
function setAPIkey_XAI() { setAPIkey("XAI"); }

/////// only utils below ///////

/**
 * @fileoverview Utility functions and constants for GPTforGDocs.
 * This file contains core functionality for API communication and content generation.
 */

/** @const {Object.<string, string>} API endpoints for different providers */
const ENDPOINTS = {
  OPENAI: "https://api.openai.com/v1/chat/completions",
  XAI: "https://api.x.ai/v1/chat/completions"
};

/** @const {Object.<string, Object.<string, string>>} Available AI models for each provider */
const AVAILABLE_MODELS = {
  OPENAI: {
    "GPT-4o Mini": "gpt-4o-mini",
    "GPT-4o": "gpt-4o",
    "O1 Mini": "o1-mini",
    "O1": "o1"
  },
  XAI: {
    "Grok-2": "grok-2-latest"
  }
};

/**
 * Makes an API call to get content from the AI provider.
 * @param {Array<Object>} messages - Array of message objects for the conversation
 * @param {number} maxTokens - Maximum number of tokens in the response
 * @returns {string|null} Generated content or null if an error occurs
 */
function getGPTcontent(messages, maxTokens) {
  try {
    const currentProvider = PropertiesService.getUserProperties().getProperty("CURRENT_PROVIDER") || "OPENAI";
    const currentModel = PropertiesService.getUserProperties().getProperty("CURRENT_MODEL") || "gpt-4o";
    
    const apiKey = PropertiesService.getUserProperties().getProperty(`${currentProvider}_API_KEY`);
    
    if (!apiKey) {
      throw new Error(`API Key for ${currentProvider} is not set`);
    }

    const endpoint = ENDPOINTS[currentProvider];
    
    let retries = 3;
    while (retries > 0) {
      try {
        const requestBody = {
          model: currentModel,
          messages: messages,
          temperature: 0.7,
          max_tokens: maxTokens,
          top_p: 1,
          frequency_penalty: 0,
          presence_penalty: 0
        };
        
        const requestOptions = {
          method: "POST",
          headers: {
            "Content-Type": "application/json",
            Authorization: "Bearer " + apiKey,
          },
          payload: JSON.stringify(requestBody),
          muteHttpExceptions: true
        };
        
        const response = UrlFetchApp.fetch(endpoint, requestOptions);
        const responseCode = response.getResponseCode();
        const responseText = response.getContentText();
        
        if (responseCode === 429) {
          Utilities.sleep(2000);
          retries--;
          continue;
        }
        
        if (responseCode !== 200) {
          throw new Error(`API returned status ${responseCode}: ${responseText}`);
        }

        const json = JSON.parse(responseText);
        return json.choices[0].message.content;
      } catch (error) {
        if (retries === 0) throw error;
        retries--;
        Utilities.sleep(1000);
      }
    }
  } catch (error) {
    DocumentApp.getUi().alert("API Error: " + error.message);
    return null;
  }
}

/**
 * Gets the display name of the current AI model.
 * @returns {string} Display name in format "Model Name (Provider)"
 */
function getCurrentModelInfo() {
  const properties = PropertiesService.getUserProperties();
  const provider = properties.getProperty("CURRENT_PROVIDER");
  const modelId = properties.getProperty("CURRENT_MODEL");

  // If no provider or model configured yet
  if (!provider || !modelId) {
    return "No model configured";
  }

  // Verify provider exists in available models
  if (!AVAILABLE_MODELS[provider]) {
    return "Invalid provider configured";
  }

  let modelName = modelId;
  // Look up friendly model name
  for (const [name, id] of Object.entries(AVAILABLE_MODELS[provider])) {
    if (id === modelId) {
      modelName = name;
      break;
    }
  }

  return `${modelName} (${provider})`;
}

/**
 * Handles the generation of content using selected text and AI.
 * @param {string} promptTemplate - Template for the prompt
 * @param {string} newInstruction - System instruction for the AI
 * @param {number} maxTokens - Maximum number of tokens in the response
 */
function handleContentGeneration(promptTemplate, newInstruction, maxTokens) {
  try {
    const doc = DocumentApp.getActiveDocument();
    const selection = doc.getSelection();
    if (!selection) {
      DocumentApp.getUi().alert("Please select some text in the document.");
      return;
    }

    const selectedText = selection.getRangeElements()
      .map(element => element.getElement().asText().getText())
      .join(" ")
      .trim();

    const messages = [
      {
        role: "system",
        content: newInstruction || "You are a helpful content writer."
      },
      {
        role: "user",
        content: `${promptTemplate}\n\nText to process:\n${selectedText}`
      }
    ];

    const generatedText = getGPTcontent(messages, maxTokens);
    if (!generatedText) return;

    const lastElement = selection.getRangeElements().slice(-1)[0];
    const parent = lastElement.getElement().getParent();
    const insertIndex = parent.getChildIndex(lastElement.getElement()) + 1;
    doc.getBody().insertParagraph(insertIndex, generatedText);
  } catch (error) {
    Logger.log("Error: " + error.toString());
    DocumentApp.getUi().alert("An error occurred: " + error.message);
  }
} 