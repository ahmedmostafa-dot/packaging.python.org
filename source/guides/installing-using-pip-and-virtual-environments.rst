// Replace this with the Verification Token from your Slack App (found in Basic Information)
const SLACK_VERIFICATION_TOKEN = 'xapp-1-A0ANNDE2GV9-10776875600530-d24e16d8e8ee8a40a916246655c5dea2c56d9b30292a162adc25d748cd114069xapp-1-A0ANNDE2GV9-10776875600530-d24e16d8e8ee8a40a916246655c5dea2c56d9b30292a162adc25d748cd114069'; 

function doPost(e) {
  // Ensure the request exists
  if (typeof e !== 'undefined') {
    const params = e.parameter;

    // Verify the request is actually coming from your Slack app
    if (params.token !== SLACK_VERIFICATION_TOKEN) {
      return ContentService.createTextOutput("Invalid Token").setMimeType(ContentService.MimeType.TEXT);
    }

    // Get the text the user typed after the slash command
    const query = params.text.trim().toLowerCase(); 
    
    // If they just typed the command with no text
    if (!query) {
      return sendToSlack("Please provide a keyword to search for! Example: `/ask-sheet wifi`");
    }

    // Search the sheet for the answer
    const responseText = searchSheet(query);

    return sendToSlack(responseText);
  }
}

function searchSheet(query) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();

  // Loop through rows (Starting at index 1 to skip the header row)
  for (let i = 1; i < data.length; i++) {
    let keyword = String(data[i][0]).toLowerCase();
    
    // Check if the cell exactly matches or contains the query
    if (keyword === query || keyword.includes(query)) {
      return `*Here is what I found:* \n${data[i][1]}`; // Returns the answer from Column B
    }
  }

  return "Sorry, I couldn't find an answer for that in the sheet.";
}

function sendToSlack(message) {
  const payload = {
    "response_type": "in_channel", // Change to "ephemeral" if you only want the user who asked to see the answer
    "text": message
  };

  return ContentService.createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}
