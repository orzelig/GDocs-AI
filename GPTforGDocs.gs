// Original copyright J. Grant, 2023, under Apache 2.0 license.
// Forked and modified by Or Zelig, 2025. Modifications also under Apache 2.0 license.

// Original: https://github.com/jedediahg/GPTforGDocs
// Fork: https://github.com/orzelig/GDocs-OpenAI

// Constants and Configuration
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
    "Grok-2": "grok-2"
  }
};

// 1. Core Setup and Menu Creation
function onOpen() {
  const currentModel = getCurrentModelInfo();
  const properties = PropertiesService.getUserProperties();
  const openaiKey = properties.getProperty("OPENAI_API_KEY");
  const xaiKey = properties.getProperty("XAI_API_KEY");

  // Create AI-Writer menu
  DocumentApp.getUi().createMenu("AI-Writer")
    .addSubMenu(DocumentApp.getUi().createMenu("Content Generation")
      .addItem("Continue this", "continueThis")
      .addItem("Create 5 LinkedIn Posts", "generateLI")
      .addItem("Create 5 Twitter Posts", "generateTweet")
      .addItem("Write Email", "generateEmail")
      .addItem("Generate Response", "generateResponse"))
    .addSubMenu(DocumentApp.getUi().createMenu("Analysis")
      .addItem("Summarize", "summarizeText")
      .addItem("Extract Keywords", "extractKeywords")
      .addItem("Generate Key Points", "generateKeyPoints")
      .addSubMenu(DocumentApp.getUi().createMenu("Competitive Analysis")
        .addItem("Generate Competitive Analysis", "generateCompetitiveAnalysis")
        .addItem("Generate Structured Profile", "generateStructuredProfile")
        .addItem("Create SWOT Analysis", "createSWOT")
        .addItem("Generate Comparison Analysis", "generateComparison")))
    .addSubMenu(DocumentApp.getUi().createMenu("Market Research")
      .addItem("Analyze Market Segments", "analyzeMarketSegments")
      .addItem("Analyze Market Trends", "analyzeMarketTrends")
      .addItem("Generate Survey Questions", "generateSurveyQuestions")
      .addSubMenu(DocumentApp.getUi().createMenu("Industry Analysis")
        .addItem("Porter's Five Forces Analysis", "generatePortersAnalysis")
        .addItem("BCG Matrix Analysis", "generateBCGMatrix")))
    .addSubMenu(DocumentApp.getUi().createMenu("Reporting")
      .addItem("Summarize Research Document", "summarizeResearch")
      .addItem("Generate Recommendations", "generateRecommendations"))
    .addSeparator()
    .addItem("Custom Instruction", "customInstruction")
    .addToUi();

  const ui = DocumentApp.getUi();
  const settingsMenu = ui.createMenu("AI-settings");

  function addApiKeySubMenu() {
    const apiKeyMenu = ui.createMenu("Set API Key");
    if (!openaiKey) apiKeyMenu.addItem("Set OpenAI API Key", "setAPIkey_OPENAI");
    if (!xaiKey) apiKeyMenu.addItem("Set XAI API Key", "setAPIkey_XAI");
    settingsMenu.addSubMenu(apiKeyMenu);
  }

  function addModelSubMenu() {
    const modelMenu = ui.createMenu("Set AI Model");
    if (openaiKey) {
      Object.entries(AVAILABLE_MODELS.OPENAI).forEach(([name, id]) => {
        modelMenu.addItem(`${name} (OpenAI)`, `setModel_${id.replace(/[-\.]/g, '_')}`);
      });
    }
    if (xaiKey) {
      if (openaiKey) modelMenu.addSeparator();
      Object.entries(AVAILABLE_MODELS.XAI).forEach(([name, id]) => {
        modelMenu.addItem(`${name} (X.AI)`, `setModel_${id.replace(/[-\.]/g, '_')}`);
      });
    }
    settingsMenu.addSubMenu(modelMenu);
  }

  settingsMenu.addItem(`Active: ${currentModel}`, "displayModelConfiguration");

  if (!openaiKey || !xaiKey) addApiKeySubMenu();
  if (openaiKey || xaiKey) addModelSubMenu();

  settingsMenu.addSeparator()
              .addItem("Delete All API Keys", "deleteAllKeys")
              .addToUi();
}

// 2. API Key Management Functions
function setAPIkey(provider) {
  const ui = DocumentApp.getUi();
  const keyResponse = ui.prompt(`Set your ${provider} API Key`, 'Please enter your API Key:', ui.ButtonSet.OK_CANCEL);

  if (keyResponse.getSelectedButton() !== ui.Button.OK) return;

  const apiKey = keyResponse.getResponseText().trim();
  if (!apiKey) return ui.alert('No API Key entered.');

  PropertiesService.getUserProperties().setProperty(provider + "_API_KEY", apiKey);
  ui.alert(`${provider} API Key set successfully.`);
  onOpen();
}

function setAPIkey_OPENAI() { setAPIkey("OPENAI"); }
function setAPIkey_XAI() { setAPIkey("XAI"); }
function deleteAllKeys() {
  const ui = DocumentApp.getUi();
  if (ui.alert('Delete All API Keys', 
               'Are you sure you want to delete all stored API keys? This action cannot be undone.',
               ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  
  const props = PropertiesService.getUserProperties();
  ['OPENAI_API_KEY', 'XAI_API_KEY', 'CURRENT_PROVIDER', 'CURRENT_MODEL']
    .forEach(key => props.deleteProperty(key));
  
  ui.alert('All API keys have been deleted successfully.');
  onOpen();
}

// 3. Model Management Functions
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
    DocumentApp.getUi().alert('Error setting model: ' + err.toString());
  }
}

function setModel_gpt_4o_mini() { setModel("OPENAI", "gpt-4o-mini"); }
function setModel_gpt_4o() { setModel("OPENAI", "gpt-4o"); }
function setModel_o1_mini() { setModel("OPENAI", "o1-mini"); }
function setModel_o1() { setModel("OPENAI", "o1"); }
function setModel_grok_2() { setModel("XAI", "grok-2-latest"); }

function displayModelConfiguration() {
  const ui = DocumentApp.getUi();
  const props = PropertiesService.getUserProperties();
  
  const hasOpenAI = !!props.getProperty("OPENAI_API_KEY");
  const hasXAI = !!props.getProperty("XAI_API_KEY");
  const provider = props.getProperty("CURRENT_PROVIDER");
  const model = props.getProperty("CURRENT_MODEL");

  const message = [
    "API Keys Setup:",
    `OpenAI: ${hasOpenAI ? "✓" : "✗"}`,
    `X.AI: ${hasXAI ? "✓" : "✗"}`,
    "",
    "Current Model:",
    (!provider || !model) 
      ? "No model configured yet. Please set up a model in AI-settings."
      : `Provider: ${provider}\nModel: ${model}`
  ].join("\n");
  
  ui.alert('Model Configuration Information', message, ui.ButtonSet.OK);
}

// 4. Content Generation Functions
function continueThis() {
  handleContentGeneration(" ",
    "You are an expert writer who can seamlessly continue any type of content while maintaining consistency in style, tone, and context.",
    1000
  );
}

function generateLI() {
  handleContentGeneration(" ",
    "You are an experienced social media manager specializing in LinkedIn content. You create engaging, professional posts that drive meaningful engagement. Please write 5 engaging LinkedIn posts about this topic. Focus on creating professional, thought-provoking content that encourages discussion.",
    500
  );
}

function generateTweet() {
  handleContentGeneration(" ",
    "You are a social media expert who creates viral, engaging content while maintaining professionalism and accuracy. Please write 5 engaging tweets about this topic. Make them concise, engaging, and shareable.",
    500
  );
}

function generateEmail() {
  handleContentGeneration(" ", "You are an expert business communicator who writes clear, engaging emails that effectively convey information. Please write a professional email explaining this:", 1000);
}

function generateResponse() {
  handleContentGeneration(" ", "You are an expert communicator who crafts clear, thoughtful, and well-structured responses. Please write a professional response to this message:", 1000);
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
      handleContentGeneration(" ", instruction, 1000);
    } else {
      ui.alert('No instruction provided.');
    }
  }
}

// 5. Analysis Functions
function summarizeText() {
  handleContentGeneration(" ",
    "You are an expert academic reviewer who can get to the heart of any matter and summarize with an efficiency of words. Please summarize this text in a single short paragraph.",
    200
  );
}

function extractKeywords() {
  handleContentGeneration(" ",
    "You are an expert academic reviewer who can get to the heart of any matter and summarize with an efficiency of words. Please extract the main keywords from this text and present them in a simple list organized by relevance.",
    200
  );
}

function generateKeyPoints() {
  handleContentGeneration(" ",
    "You are an expert academic researcher who can iterate on ideas and concisely explain them in clear words. Please write a list of key bulletpoints including important points that may be missing from the topic of this text.",
    1000
  );
}

// 6. Competitive Analysis Functions
function generateCompetitiveAnalysis() {
  handleContentGeneration(" ",
    "You are a strategic business analyst specializing in competitive analysis. Create a comprehensive analysis that identifies key competitive factors, market positioning, and strategic implications based on this information.",
    2000
  );
}

function generateStructuredProfile() {
  handleContentGeneration(" ",
    "You are a business analyst who excels at organizing and presenting information in clear, structured profiles. Create a comprehensive profile that highlights key characteristics, strengths, and notable features based on this data.",
    1500
  );
}

function createSWOT() {
  handleContentGeneration(" ",
    "You are a strategic analyst specializing in SWOT analysis. Create a comprehensive analysis of Strengths, Weaknesses, Opportunities, and Threats, providing detailed insights for each category based on this information.",
    2000
  );
}

function generateComparison() {
  handleContentGeneration(" ",
    "You are an analytical expert who excels at comparative analysis. Create a detailed, structured comparison that highlights key differences and similarities across multiple dimensions based on this information.",
    2000
  );
}

// 7. Market Research Functions
function analyzeMarketSegments() {
  handleContentGeneration(" ",
    "You are a market research expert specializing in market segmentation. Analyze the data to identify distinct market segments and their characteristics.",
    2000
  );
}

function analyzeMarketTrends() {
  handleContentGeneration(" ",
    "You are a market analyst expert in identifying and analyzing trends. Examine the data to identify key trends, patterns, and their implications.",
    2000
  );
}

function generateSurveyQuestions() {
  handleContentGeneration(" ",
    "You are a market research expert specializing in survey design. Create clear, unbiased, and effective questions that will gather the needed information based on this research objective.",
    1500
  );
}

// 8. Industry Analysis Functions
function generatePortersAnalysis() {
  handleContentGeneration(" ",
    "You are a strategic analyst specializing in Porter's Five Forces framework. Create a comprehensive analysis of competitive forces affecting the industry based on this information.",
    2000
  );
}

function generateBCGMatrix() {
  handleContentGeneration(" ",
    "You are a portfolio strategy expert specializing in BCG matrix analysis. Analyze the products/services and categorize them appropriately with detailed justification based on this portfolio information.",
    2000
  );
}

// 9. Reporting Functions
function summarizeResearch() {
  handleContentGeneration(" ",
    "You are a research analyst who excels at distilling complex documents into clear, concise summaries while retaining key information and insights. Please create a concise summary of this research document.",
    1500
  );
}

function generateRecommendations() {
  handleContentGeneration(" ",
    "You are a strategic consultant who excels at developing actionable recommendations based on analytical insights. Create clear, justified recommendations with supporting rationale based on this analysis.",
    2000
  );
}

// 10. Core Utility Functions
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

function handleContentGeneration(_, instruction, maxTokens) {
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
        content: instruction || "You are a helpful content writer."
      },
      {
        role: "user",
        content: `Text to process:\n${selectedText}`
      }
    ];

    const generatedText = getGPTcontent(messages, maxTokens);
    if (!generatedText) return;

    const lastElement = selection.getRangeElements().slice(-1)[0];
    const parent = lastElement.getElement().getParent();
    const insertIndex = parent.getChildIndex(lastElement.getElement()) + 1;
    doc.getBody().insertParagraph(insertIndex, generatedText);
  } catch (error) {
    DocumentApp.getUi().alert("An error occurred: " + error.message);
  }
} 