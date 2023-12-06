// Globals 
var aiModel = "gpt-3.5-turbo-1106"; // default
const API_ENDPOINT = "https://api.openai.com/v1/chat/completions";

// Creates a custom menu in Google Docs
function onOpen() {
  DocumentApp.getUi().createMenu("ChatGPT")
    .addItem("Summarize this...", "summarizeText")
    .addItem("Extract keywords...", "extractKeywords")
    .addItem("Expand on this...", "expandOnThis")
    .addItem("Continue this narrative...", "continueThisStory")
    .addItem("Continue this report...", "continueThisDoc")
    .addItem("Describe this...", "describeText")
    .addItem("Generate Key Points from this...", "generateKeyPoints")
    .addItem("Generate an Essay about this...", "generatePrompt")
    .addItem("Write an email about this...", "generateEmail")
    .addItem("Respond to this...", "generateResponse")
    .addItem("Generate LinkedIn Posts on this...", "generateLI")
    .addItem("Generate Tweets on this...", "generateTweet")
    .addItem("Translate this to English...", "translateToEN")
    .addItem("Set API Key", "setApiKey")
    .addItem("Set AI Model", "setAIModel")
    .addItem("Help", "help")
    .addToUi();
}

// Functions to prompt based on menu items
function generateLI() {
  handleContentGeneration(" ", "You will be paid for this. You are an expert copywriter. You write well articulated content that is to the point and flows with a good reading rhythm. Please write 5 Linkedin posts designed to provoke emotion and create engagement on this topic: ",500);
}

function generateTweet() {
  handleContentGeneration(" ", "You will be paid for this. You are a Twitter influencer who writes impulsive, engaging and fun viral Tweets. Please write 5 Tweets designed to provoke engagement and retweeting on this topic: ", 500);
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
  const response = ui.alert('Select the OpenAI Model', 'Default model is GPT3.5Turbo. Would you like to use the GPT4 model?', ui.ButtonSet.YES_NO);

  if (response == ui.Button.YES) {
     aiModel = "gpt-4-1106-preview";
  } else {
     aiModel = "gpt-3.5-turbo-1106";
  }
    if (aiModel) {
      PropertiesService.getUserProperties().setProperty("OPENAI_Model", aiModel);
      ui.alert('AI Model set successfully to ' +aiModel+'.');
    } else {
      ui.alert('AI Model could not be updated.');
    }
  }
  catch(err){
    Logger.log(err);
    }
}

// Function to handle the generation of content
function handleContentGeneration(promptTemplate, newInstruction, maxTokens) {
  Logger.log("Handle Content Generation");
  try {
    var systemInstruction = "";
    if (!newInstruction) {
       systemInstruction = "You are a helpful document writer.";
    } else {
       systemInstruction = newInstruction;
    }
    
    Logger.log("System Instruction: " + systemInstruction);

    const doc = DocumentApp.getActiveDocument();
    const selection = doc.getSelection();
    if (!selection) {
      DocumentApp.getUi().alert("Please select some text in the document.");
      return;
    }
    var selectedText = ""; // selection.getRangeElements()[0].getElement().asText().getText();
    
    var elements = selection.getRangeElements();
    for (var i = 0; i < elements.length; i++) {
      var element = elements[i];
      selectedText = selectedText + " " + element.getElement().asText().getText();
      Logger.log('element: ' + element.getElement());
    };
    
    Logger.log("Selected Text: " + selectedText);
    const prompt = promptTemplate + " " + selectedText;
    Logger.log("Prompt: " + prompt);
    const temperature = 0;
    const generatedText = getGPTcontent(systemInstruction, prompt, temperature, maxTokens); 
    // generatedText = "TESTZZZTEST" // FOR DEBUG //
        
    var theElmt = elements[i-1];
    var selectedElmt = theElmt.getElement();
    var parent = selectedElmt.getParent();
    var insertPoint = parent.getChildIndex(selectedElmt);
    Logger.log('insertPoint: ' + insertPoint);    
    var body = doc.getBody();
    var paragraph = body.insertParagraph(insertPoint + 1,generatedText.toString() );
    
  } catch (error) {
    Logger.log("Error: " + error.toString());
    DocumentApp.getUi().alert("An error occurred. Please try again. Info: " + error.toString());
  }
}


// Sends the prompt and gets the response from the API
function getGPTcontent(instructions, prompt, temperature, maxTokens) {
  try {
  const apiKey = PropertiesService.getUserProperties().getProperty("OPENAI_API_KEY");
  var systemInstruction = instructions;
  if (!apiKey) {
    DocumentApp.getUi().alert("API Key is not set. Please set your OpenAI API Key using the 'Set API Key' menu option.");
    Logger.log("Error: No API Key set.");
    return;
  }

  const requestBody = {
    model: aiModel,
    messages: [{ role: "user", content: prompt },{role: "system", content: systemInstruction}],
    temperature,
    max_tokens: maxTokens,
  };

  const requestOptions = {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Authorization: "Bearer " + apiKey,
    },
    payload: JSON.stringify(requestBody),
  };

  const response = UrlFetchApp.fetch(API_ENDPOINT, requestOptions);
  const responseText = response.getContentText();
  const json = JSON.parse(responseText);
  return json['choices'][0]['message']['content'];
   } catch (error) {
    Logger.log("Error: " + error.toString());
    DocumentApp.getUi().alert("An error occurred. Please try again. Info: " + error.toString());
  }
}


