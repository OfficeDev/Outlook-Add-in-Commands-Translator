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