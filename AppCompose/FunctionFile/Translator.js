Office.initialize = function () {
}

function translate(event) {
  Office.context.mailbox.item.getSelectedDataAsync("text", function (ar) {

    // Make sure there is a selection
    if (ar === undefined || ar === null ||
        ar.value === undefined || ar.value === null ||
        ar.value.data === undefined || ar.value.data === null) {
      // Display an error message
      Office.context.mailbox.item.notificationMessages.addAsync("noSelectionError", {
        type: "errorMessage",
        message: "No text selected! Please select text to translate and try again."
      });
      
      // Signal that we are done.
      event.completed();
      return;
    }
    
    try {
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
            textToWrite += translatedText[i] + "<br/>";

          // Replace the selected text with the translated version
          Office.context.mailbox.item.setSelectedDataAsync(textToWrite, { coercionType: "html" }, function (asyncResult) {
            Office.context.mailbox.item.notificationMessages.addAsync("success", {
              type: "informationalMessage",
              icon: "icon1_16x16",
              message: "Translated successfully",
              persistent: false
            });

            // Signal that we are done.
            event.completed();
          });
        }
      });
    }
    catch (err) {
      Office.context.mailbox.item.notificationMessages.addAsync("translateError", {
        type: "errorMessage",
        message: "ERROR translating message: " + err.message
      });
      
      // Signal that we are done.
      event.completed();
    }
  });
}