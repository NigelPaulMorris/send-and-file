/* Send & File - event handlers. Runs when you press Send, and when you start a new email. */

const ASK_TEXT =
  "Where should this email be filed once it's sent? Choose a folder, " +
  "or press Send Anyway to leave it in Sent Items.";

function onMessageSendHandler(event) {
  const item = Office.context.mailbox.item;
  item.sessionData.getAsync("fileTo", (result) => {
    const chosen = result.status === Office.AsyncResultStatus.Succeeded && result.value;
    if (chosen) {
      // A folder (or "don't file") was picked in the pane - let it send.
      event.completed({ allowEvent: true });
    } else {
      // Pause the send and offer the folder picker.
      event.completed({
        allowEvent: false,
        errorMessage: ASK_TEXT,
        cancelLabel: "Choose folder",
        commandId: "msgComposeOpenPaneButton",
      });
    }
  });
}

// Whenever you start a new email/reply, quietly file anything sent earlier.
async function onNewMessageComposeHandler(event) {
  try {
    if (typeof sweepSentItems === "function") {
      await Promise.race([sweepSentItems(), new Promise((r) => setTimeout(r, 25000))]);
    }
  } catch (err) {
    console.warn("Send & File: background filing skipped", err);
  }
  event.completed();
}

Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
Office.actions.associate("onNewMessageComposeHandler", onNewMessageComposeHandler);
