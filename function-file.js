Office.onReady(() => {
  // Required for event-based activation
});

async function onMessageSend(event) {
  try {
    const item = Office.context.mailbox.item;
    const internalDomain = "itworks.co.nz";

    const recipients = [];

    if (item.to) recipients.push(...item.to.map(r => r.emailAddress));
    if (item.cc) recipients.push(...item.cc.map(r => r.emailAddress));
    if (item.bcc) recipients.push(...item.bcc.map(r => r.emailAddress));

    const externalRecipients = recipients.filter(email =>
      !email.toLowerCase().endsWith(`@${internalDomain}`)
    );

    if (externalRecipients.length > 0) {
      event.completed({
        allowEvent: false,
        errorMessage:
          "This message is addressed to recipients outside your organization. Please review before sending."
      });
      return;
    }

    event.completed({ allowEvent: true });

  } catch (e) {
    // Fail open
    event.completed({ allowEvent: true });
  }
}

// Mandatory export
if (typeof window !== "undefined") {
  window.onMessageSend = onMessageSend;
}
