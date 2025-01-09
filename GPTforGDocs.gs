// Original copyright J. Grant, 2023, under Apache 2.0 license.
// Forked and modified by Or Zelig, 2025. Modifications also under Apache 2.0 license.
// Original: https://github.com/jedediahg/GPTforGDocs
// Fork: https://github.com/orzelig/GDocs-OpenAI
// 

// Globals 
var aiModel = "gpt-4o-mini"; // Updated default model
const API_ENDPOINT = "https://api.openai.com/v1/chat/completions";
const AVAILABLE_MODELS = {
  "GPT-4o Mini ($0.15)": "gpt-4o-mini",
  "GPT-4o ($2.50)": "gpt-4o",
  "O1 Mini ($3.00)": "o1-mini",
  "O1 ($15.00)": "o1"
};

// Creates a custom menu in Google Docs
function onOpen() {
  DocumentApp.getUi().createMenu("ChatGPT")
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
    .addSeparator()
    .addItem("Set API Key", "setApiKey")
    .addItem("Set AI Model", "setAIModel")
    .addItem("Help", "help")
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


//Help
function help() {
  const ui = DocumentApp.getUi();
  const response = ui.alert('Using this add-on is easy. Just highlight the text you want to work on and select the menu option that you want to do. Be sure to select the text from top to bottom so the cursor is at the bottom. The new text will appear after the cursor.\nRemember to set your API key and model first.', ui.ButtonSet.OK_CANCEL);
}

// Function for setting the API Key
function setApiKey() {
  try {
  const ui = DocumentApp.getUi();
  const response = ui.prompt('Set your OpenAI API Key', 'Please enter your API Key:', ui.ButtonSet.OK_CANCEL);

  if (response.getSelectedButton() == ui.Button.OK) {
    const apiKey = response.getResponseText();
    if (apiKey) {
      PropertiesService.getUserProperties().setProperty("OPENAI_API_KEY", apiKey);
      ui.alert('API Key set successfully.');
    } else {
      ui.alert('No API Key entered.');
    }
  }
  }
  catch(err){
    Logger.log(err);
    }
}

// Function for setting the AI Model 
function setAIModel() {
  try {
    const ui = DocumentApp.getUi();
    const response = ui.prompt(
      'Select AI Model',
      'Enter model number:\n' +
      '1. GPT-4O Mini ($0.15/1k tokens) - Good balance of performance and cost\n' +
      '2. GPT-4O ($2.50/1k tokens) - Excellent performance\n' +
      '3. O1 Mini ($3.00/1k tokens) - Advanced capabilities\n' +
      '4. O1 ($15.00/1k tokens) - Most powerful model',
      ui.ButtonSet.OK_CANCEL
    );

    if (response.getSelectedButton() == ui.Button.OK) {
      const selection = response.getResponseText();
      switch(selection) {
        case "1":
          aiModel = "gpt-4o-mini";
          break;
        case "2":
          aiModel = "gpt-4o";
          break;
        case "3":
          aiModel = "o1-mini";
          break;
        case "4":
          aiModel = "o1";
          break;
        default:
          ui.alert('Invalid selection. Using default GPT-4O Mini.');
          aiModel = "gpt-4o-mini";
      }
      PropertiesService.getUserProperties().setProperty("OPENAI_Model", aiModel);
      ui.alert('AI Model set successfully to ' + aiModel);
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


// Sends the prompt and gets the response from the API
function getGPTcontent(messages, maxTokens) {
  try {
    const apiKey = PropertiesService.getUserProperties().getProperty("OPENAI_API_KEY");
    if (!apiKey) {
      throw new Error("API Key is not set");
    }

    let retries = 3;
    while (retries > 0) {
      try {
        const requestBody = {
          model: aiModel,
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

        const response = UrlFetchApp.fetch(API_ENDPOINT, requestOptions);
        const responseCode = response.getResponseCode();
        
        if (responseCode === 429) {
          Utilities.sleep(2000);
          retries--;
          continue;
        }
        
        if (responseCode !== 200) {
          throw new Error(`API returned status ${responseCode}: ${response.getContentText()}`);
        }

        const json = JSON.parse(response.getContentText());
        return json.choices[0].message.content;
      } catch (error) {
        if (retries === 0) throw error;
        retries--;
        Utilities.sleep(1000);
      }
    }
  } catch (error) {
    Logger.log("Error: " + error.toString());
    DocumentApp.getUi().alert("API Error: " + error.message);
    return null;
  }
}


