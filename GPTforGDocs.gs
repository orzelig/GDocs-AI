// Original copyright J. Grant, 2023, under Apache 2.0 license.
// Forked and modified by OOr Zelig, 2025 . Modifications also under Apache 2.0 license.

// Original: https://github.com/jedediahg/GPTforGDocs
// Fork: https://github.com/orzelig/GDocs-OpenAI
// 

// Globals 
const ENDPOINTS = {
  OPENAI: "https://api.openai.com/v1/chat/completions",
  XAI: "https://api.x.ai/v1/chat/completions"
};
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

// Add provider tracking
var currentProvider = "openai"; // or "xai"
var aiModel = "gpt-4o-mini";

// Creates a custom menu in Google Docs
function onOpen() {
  const currentModel = getCurrentModelInfo();
  
  DocumentApp.getUi().createMenu("AI-Writer")
    .addItem(`Active: ${currentModel}`, "showModelInfo")
    .addSeparator()
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
    .addItem("Translate to English", "translateToEN")
    .addToUi();

  DocumentApp.getUi().createMenu(`Active: ${currentModel}`)
    .addItem("Set API Key", "setApiKey")
    .addItem("Set AI Model", "setAIModel")
    .addToUi();

}

// Functions to prompt based on menu items
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

function describeText() {
  handleContentGeneration(" ", "You will be paid for this. You are an expert academic writer gifeted at explaining complex subjects in easy to understand terms. Please generate a descriptive paragraph explaining this text. Order the response chronologially and use a positive tone. Here is the text to describe:", 500);
}

function generateKeyPoints() {
  handleContentGeneration(" ", "You will be paid for this. You are an expert academic researcher. You can itereate on ideas and concisely explain them in clear words. Please write a list of key bulletpoints including important points that may be missing as well from the topic of this text: ", 1000);
  }

function expandOnThis() {
  handleContentGeneration(" ", "You will be paid for this. You are a know-it-all who loves to share detailed knowledge about any subject. You write very clearly with engaging topical sentence structures. Expand on this text: ", 1000);
  }

function translateToEN() {
  handleContentGeneration(" ", "You will be paid for this. You are a brilliant linquist capable of translating from any language into English. Please translate this text to English: ", 1000);
  }

function continueThisStory() {
  handleContentGeneration(" ", "You will be paid for this. You are a creative writer. You craft exciting and gripping narratives about interesting characters in flowing prose that entrances and entices the reader. Continue this narrative: ", 1000);
  }

function continueThisDoc() {
  handleContentGeneration(" ", "You will be paid for this. You are an expert academic reviewer. You write concise and clear text  with an efficiency of words. Continue this report: ", 1000);
  }

function generateEmail() {
  handleContentGeneration(" ", "You will be paid for this. You are a know-it-all who loves to share detailed knowledge about any subject. You write very clearly with engaging topical sentence structures. Please write a professional email explaining this: ", 1000);
  }

function generateResponse() {
  handleContentGeneration(" ", "You will be paid for this. You are a know-it-all who loves to share detailed knowledge about any subject. You write very clearly with engaging topical sentence structures. Please write a professional response to this message: ", 1000);
  }



// Function for setting the API Key
function setApiKey() {
  try {
    const ui = DocumentApp.getUi();
    const response = ui.alert('API Key Setup',
      'Would you like to set up OpenAI or X.AI key?\n\nClick YES for OpenAI\nClick NO for X.AI',
      ui.ButtonSet.YES_NO_CANCEL);

    if (response === ui.Button.CANCEL) return;
    
    const provider = response === ui.Button.YES ? "OPENAI" : "XAI";
    const keyResponse = ui.prompt(`Set your ${provider} API Key`, 'Please enter your API Key:', ui.ButtonSet.OK_CANCEL);

    if (keyResponse.getSelectedButton() == ui.Button.OK) {
      const apiKey = keyResponse.getResponseText();
      if (apiKey) {
        PropertiesService.getUserProperties().setProperty(`${provider}_API_KEY`, apiKey);
        ui.alert(`${provider} API Key set successfully.`);
      } else {
        ui.alert('No API Key entered.');
      }
    }
  } catch(err) {
    Logger.log(err);
  }
}

// Function for setting the AI Model 
function setAIModel() {
  try {
    const ui = DocumentApp.getUi();
    
    // Check if both API keys are set
    const openaiKey = PropertiesService.getUserProperties().getProperty("OPENAI_API_KEY");
    const xaiKey = PropertiesService.getUserProperties().getProperty("XAI_API_KEY");
    
    let modelOptions = "";
    let counter = 1;
    let allModels = [];  // Add this array to track all available models
    
    if (openaiKey) {
      Object.entries(AVAILABLE_MODELS.OPENAI).forEach(([name, id]) => {
        modelOptions += `${counter}. ${name} (OpenAI)\n`;
        allModels.push({ provider: "OPENAI", id: id });  // Store model info
        counter++;
      });
    }
    
    if (xaiKey) {
      Object.entries(AVAILABLE_MODELS.XAI).forEach(([name, id]) => {
        modelOptions += `${counter}. ${name} (X.AI)\n`;
        allModels.push({ provider: "XAI", id: id });  // Store model info
        counter++;
      });
    }
    
    if (!modelOptions) {
      ui.alert('Please set up at least one API key first.');
      return;
    }

    const response = ui.prompt(
      'Select AI Model',
      'Enter model number:\n' + modelOptions,
      ui.ButtonSet.OK_CANCEL
    );

    if (response.getSelectedButton() == ui.Button.OK) {
      const selection = parseInt(response.getResponseText()) - 1;  // Convert to 0-based index
      
      // Validate selection
      if (selection >= 0 && selection < allModels.length) {
        const selectedModel = allModels[selection];
        aiModel = selectedModel.id;  // Set the actual model ID
        PropertiesService.getUserProperties().setProperty("CURRENT_PROVIDER", selectedModel.provider);
        PropertiesService.getUserProperties().setProperty("CURRENT_MODEL", selectedModel.id);
        ui.alert(`AI Model set successfully to ${selectedModel.id} using ${selectedModel.provider}`);
      } else {
        ui.alert('Invalid selection.');
      }
    }
  } catch(err) {
    Logger.log(err);
    DocumentApp.getUi().alert('Error setting model: ' + err.toString());
  }
}

// Function to handle the generation of content
function handleContentGeneration(promptTemplate, newInstruction, maxTokens) {
  try {
    const doc = DocumentApp.getActiveDocument();
    const selection = doc.getSelection();
    if (!selection) {
      DocumentApp.getUi().alert("Please select some text in the document.");
      return;
    }

    // Improved text selection handling
    const selectedText = selection.getRangeElements()
      .map(element => element.getElement().asText().getText())
      .join(" ")
      .trim();

    // Create a more structured conversation
    const messages = [
      {
        role: "system",
        content: newInstruction || "You are a helpful document writer."
      },
      {
        role: "user",
        content: `${promptTemplate}\n\nText to process:\n${selectedText}`
      }
    ];

    const generatedText = getGPTcontent(messages, maxTokens);
    if (!generatedText) return;

    // Insert the response after the selection
    const lastElement = selection.getRangeElements().slice(-1)[0];
    const parent = lastElement.getElement().getParent();
    const insertIndex = parent.getChildIndex(lastElement.getElement()) + 1;
    doc.getBody().insertParagraph(insertIndex, generatedText);
  } catch (error) {
    Logger.log("Error: " + error.toString());
    DocumentApp.getUi().alert("An error occurred: " + error.message);
  }
}


// Add debug logging function
function debugLog(message, data) {
  Logger.log(`DEBUG: ${message}`);
  if (data) Logger.log(JSON.stringify(data, null, 2));
}

// Add validation for X.AI requests
function validateXAIRequest(requestBody) {
  if (!requestBody.model.startsWith('grok-')) {
    throw new Error('Invalid model for X.AI. Must use a Grok model.');
  }
  
  // Validate message format
  if (!Array.isArray(requestBody.messages)) {
    throw new Error('Messages must be an array');
  }
  
  requestBody.messages.forEach((msg, index) => {
    if (!msg.role || !msg.content) {
      throw new Error(`Invalid message format at index ${index}`);
    }
    if (!['system', 'user', 'assistant'].includes(msg.role)) {
      throw new Error(`Invalid role "${msg.role}" at index ${index}`);
    }
  });
  
  return requestBody;
}

// Update getGPTcontent with better debugging
function getGPTcontent(messages, maxTokens) {
  try {
    // Get the current provider and model from stored properties
    const currentProvider = PropertiesService.getUserProperties().getProperty("CURRENT_PROVIDER") || "OPENAI";
    const currentModel = PropertiesService.getUserProperties().getProperty("CURRENT_MODEL") || "gpt-4o";
    
    debugLog(`Current provider: ${currentProvider}`);
    debugLog(`Current model: ${currentModel}`);
    
    const apiKey = PropertiesService.getUserProperties().getProperty(`${currentProvider}_API_KEY`);
    debugLog(`API Key exists: ${Boolean(apiKey)}`);
    
    if (!apiKey) {
      throw new Error(`API Key for ${currentProvider} is not set`);
    }

    const endpoint = ENDPOINTS[currentProvider];
    debugLog(`Using endpoint: ${endpoint}`);
    
    let retries = 3;
    while (retries > 0) {
      try {
        const requestBody = {
          model: currentModel,  // Use the stored model instead of aiModel global
          messages: messages,
          temperature: 0.7,
          max_tokens: maxTokens,
          top_p: 1,
          frequency_penalty: 0,
          presence_penalty: 0
        };
        
        debugLog('Request body:', requestBody);

        const requestOptions = {
          method: "POST",
          headers: {
            "Content-Type": "application/json",
            Authorization: "Bearer " + apiKey,
          },
          payload: JSON.stringify(requestBody),
          muteHttpExceptions: true
        };
        
        debugLog('Making API request...');
        const response = UrlFetchApp.fetch(endpoint, requestOptions);
        const responseCode = response.getResponseCode();
        const responseText = response.getContentText();
        
        debugLog(`Response code: ${responseCode}`);
        debugLog('Response body:', responseText);
        
        if (responseCode === 429) {
          debugLog('Rate limit hit, retrying...');
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
        debugLog(`Error in API call: ${error.toString()}`);
        if (retries === 0) throw error;
        retries--;
        Utilities.sleep(1000);
      }
    }
  } catch (error) {
    debugLog(`Fatal error: ${error.toString()}`);
    DocumentApp.getUi().alert("API Error: " + error.message);
    return null;
  }
}

// Add function to get current model display name
function getCurrentModelInfo() {
  const provider = PropertiesService.getUserProperties().getProperty("CURRENT_PROVIDER") || "OPENAI";
  const modelId = PropertiesService.getUserProperties().getProperty("CURRENT_MODEL") || "gpt-4o-mini";
  
  // Find the display name for the current model
  let modelName = modelId;
  for (const [name, id] of Object.entries(AVAILABLE_MODELS[provider])) {
    if (id === modelId) {
      modelName = name;
      break;
    }
  }
  
  return `${modelName} (${provider})`;
}

// Add function to show model info
function showModelInfo() {
  const ui = DocumentApp.getUi();
  const provider = PropertiesService.getUserProperties().getProperty("CURRENT_PROVIDER") || "OPENAI";
  const modelId = PropertiesService.getUserProperties().getProperty("CURRENT_MODEL") || "gpt-4o-mini";
  
  ui.alert(
    'Current Model Information',
    `Provider: ${provider}\nModel: ${modelId}`,
    ui.ButtonSet.OK
  );
}


