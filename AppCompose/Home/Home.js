// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.

/// <reference path="../App.js" />

(function () {
  'use strict';

  // The initialize function must be run each time a new page is loaded
  Office.initialize = function (reason) {
    $(document).ready(function () {
      app.initialize();

      $('#translate').click(translateSelection);
    });
  };

  function translateSelection() {
    Office.context.mailbox.item.getSelectedDataAsync("text", function (ar) {

      // Make sure there is a selection
      if (ar === undefined || ar === null ||
          ar.value === undefined || ar.value === null ||
          ar.value.data === undefined || ar.value.data === null) {
        // Display an error message
        app.showNotification("No text selected!", "Please select text to translate and try again.");
        return;
      }
      
      // Generate the API call URL
      var requestUrl = generateRequestUrl(ar.value.data);

      $.ajax({
        url: requestUrl,
        jsonp: "callback",
        dataType: "jsonp",
        success: function (response) {
          var translatedText = response.text;
          var textToWrite = "";
          
          // The response is an array of one or more translated lines.
          // Append them together with <br/> tags.
          for (var i = 0; i < translatedText.length; i++)
            textToWrite += translatedText[i] + "<br/><br/>";
            
          // Replace the selected text with the translated version
          Office.context.mailbox.item.setSelectedDataAsync(textToWrite, { coercionType: "html" }, function(asyncResult){
            app.showNotification("Success", "Translated successfully");    
          });
        }
      });
    });
  }
})();

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