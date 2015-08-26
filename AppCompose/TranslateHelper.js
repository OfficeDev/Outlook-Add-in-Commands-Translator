// Helper function to generate an API request
// URL to the Yandex translator service
function generateRequestUrl(text) {
  // Split the selected data into individual lines
  var tempLines = text.split(/\r\n|\r|\n/g);
  var lines = [];
  
  // Add non-empty lines to the data to translate
  for (var i = 0; i < tempLines.length; i++)
    if (tempLines[i] != "")
        lines.push(tempLines[i]);

  // Add each line as a 'text' query parameter
  var encodedText = "";
  for (var i = 0; i < (lines.length) ; i++) {
    encodedText += "&text=" + encodeURI(lines[i].replace(/ /g, "+"));
  }

  // API Key for the yandex service
  // Get one at https://translate.yandex.com/developers
  var apiKey = "YOUR API KEY HERE";
  
  return "https://translate.yandex.net/api/v1.5/tr.json/translate?key=" + apiKey + "&lang=en-ru" + encodedText;
}