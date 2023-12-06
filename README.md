# GPTforGDocs
## Open AI assistant for Google Docs
This is a simple menu extension for Google Docs. With it you can select elements in the docs that contain text and have GPT take action on it. 
## Installation
To install, simply open a Google Doc and go to *Extensions->Apps Script*. Paste the script into the editor and save. Reload your document.  You can also deploy this as an add-on to make it available in all your Google docs.
## Usage
Once installed, you will fid a new ChatGPT menu with several items.
|Command|Action|
|-----|-----|
|*Summarize this...* | Create a summary paragraph of the selected text.| 
|*Extract Keywords...* | Create a list of keywords in the selected text.| 
|*Expand on this...* | Add several more paragraphs expanding on the idea of the selected text.|
|*Continue this narrative...* | Continue telling a story based on the initial selected text. The profile is that of a creative writer.|
|*Continue this report...* | Continue writing a report based on the initial selected text. The profile is that of an academic expert.|
|*Describe this...* | Write a descriptive text based on the selected elements. Great for describing tables.|
|*Generate Key Points from this...* | Generate a set of bullet points based on the selected text.|
|*Generate an Essay about this...* | Write an essay about the topic of the selected text.|
|*Generate LinkedIn Posts on this...* | Generate LinkedIn posts that could be used to promote the idea in the selected text.|
|*Generate Tweets on this...* | Generate Tweets that could be used to promote the idea in the selected text.|
|*Translate this to English...* | Translate the selected text to English.|

Normally, the generated text is inserted immediately after the selected text. If this doesn't happen, check the start of the document.

You will have to set your OpenAI API key using the menu option to do so.

GPT-3.5 is the default model. If you wish to use GPT4, you can enable it with the menu item to do so.
