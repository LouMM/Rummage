

Office.onReady(() => {
    // If needed, Office.js is ready to be called
});

/**
 * Shows a notification when the add-in command is executed.
 * @param event
 */
function action(event: Office.AddinCommands.Event) {
    const message: Office.NotificationMessageDetails = {
        type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
        message: "Performed action.",
        icon: "Icon.80x80",
        persistent: true,
    };

    // Show a notification message
    if (Office.context.mailbox.item) {
        Office.context.mailbox.item.notificationMessages.replaceAsync("action", message);
    } else {
        console.error("Office.context.mailbox.item is undefined.");
    }

    // Be sure to indicate when the add-in command function is complete
    event.completed();
}

function getGlobal() {
    return typeof self !== "undefined"
        ? self
        : typeof window !== "undefined"
            ? window
            : typeof global !== "undefined"
                ? global
                : undefined;
}

const g = getGlobal() as any;

// The add-in command functions need to be available in global scope
g.action = action;
