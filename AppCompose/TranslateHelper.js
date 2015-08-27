// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.

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

// MIT License: 
 
// Permission is hereby granted, free of charge, to any person obtaining 
// a copy of this software and associated documentation files (the 
// ""Software""), to deal in the Software without restriction, including 
// without limitation the rights to use, copy, modify, merge, publish, 
// distribute, sublicense, and/or sell copies of the Software, and to 
// permit persons to whom the Software is furnished to do so, subject to 
// the following conditions: 
 
// The above copyright notice and this permission notice shall be 
// included in all copies or substantial portions of the Software. 
 
// THE SOFTWARE IS PROVIDED ""AS IS"", WITHOUT WARRANTY OF ANY KIND, 
// EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF 
// MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND 
// NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE 
// LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION 
// OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION 
// WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.