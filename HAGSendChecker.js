Office.initialize = function() {
    // Add event listener to the send button
    document.getElementById("sendButton").addEventListener("click", confirmRecipients);
};

function confirmRecipients() {
    // Get the current item (email) being composed
    var item = Office.context.mailbox.item;

    // Get the recipients of the email
    var recipients = item.to.concat(item.cc).concat(item.bcc);

    // Prompt for confirmation
    var confirmMessage = "Are you sure you want to send this email to the following recipients?\n\n" +
                         recipients.join(", ");
    if (confirm(confirmMessage)) {
        // User confirmed, proceed with sending the email
        item.send();
    } else {
        // User cancelled, do not send the email
    }
}
