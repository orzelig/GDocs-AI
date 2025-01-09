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
  const provider = PropertiesService.getUserProperties().getProperty("CURRENT_PROVIDER") || "OPENAI";
  const modelId = PropertiesService.getUserProperties().getProperty("CURRENT_MODEL") || "gpt-4o-mini";
  
  let modelName = modelId;
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