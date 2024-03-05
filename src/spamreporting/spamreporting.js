// Handles the SpamReporting event to process a reported message.
function onSpamReport(event) {
    console.log("SPAM REPORTED !")
    // Do nothing but report event completed.
    event.completed({
        onErrorDeleteItem: true,
        showPostProcessingDialog: {
          title: "Contoso Spam Reporting",
          description: "Thank you for reporting this message.",
        },
      });
  }
  

Office.actions.associate("onSpamReport", onSpamReport);