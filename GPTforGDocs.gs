// Original copyright J. Grant, 2023, under Apache 2.0 license.
// Forked and modified by Or Zelig, 2025. Modifications also under Apache 2.0 license.

// Original: https://github.com/jedediahg/GPTforGDocs
// Fork: https://github.com/orzelig/GDocs-OpenAI

function onOpen() {
  const currentModel = getCurrentModelInfo();
  
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

  DocumentApp.getUi().createMenu("AI-settings")
    .addItem(`Active: ${currentModel}`, "showModelInfo")
    .addItem("Set API Key", "setApiKey")
    .addItem("Set AI Model", "setAIModel")
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

function setAIModel() {
  try {
    const ui = DocumentApp.getUi();
    
    const openaiKey = PropertiesService.getUserProperties().getProperty("OPENAI_API_KEY");
    const xaiKey = PropertiesService.getUserProperties().getProperty("XAI_API_KEY");
    
    let modelOptions = "";
    let counter = 1;
    let allModels = [];
    
    if (openaiKey) {
      Object.entries(AVAILABLE_MODELS.OPENAI).forEach(([name, id]) => {
        modelOptions += `${counter}. ${name} (OpenAI)\n`;
        allModels.push({ provider: "OPENAI", id: id });
        counter++;
      });
    }
    
    if (xaiKey) {
      Object.entries(AVAILABLE_MODELS.XAI).forEach(([name, id]) => {
        modelOptions += `${counter}. ${name} (X.AI)\n`;
        allModels.push({ provider: "XAI", id: id });
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
      const selection = parseInt(response.getResponseText()) - 1;
      
      if (selection >= 0 && selection < allModels.length) {
        const selectedModel = allModels[selection];
        PropertiesService.getUserProperties().setProperty("CURRENT_PROVIDER", selectedModel.provider);
        PropertiesService.getUserProperties().setProperty("CURRENT_MODEL", selectedModel.id);
        ui.alert(`AI Model set successfully to ${selectedModel.id} using ${selectedModel.provider}`);
        onOpen();
      } else {
        ui.alert('Invalid selection.');
      }
    }
  } catch(err) {
    Logger.log(err);
    DocumentApp.getUi().alert('Error setting model: ' + err.toString());
  }
}

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

