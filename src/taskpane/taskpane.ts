Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        const sideloadMsg = document.getElementById("sideload-msg");
        const appBody = document.getElementById("app-body");
        const runButton = document.getElementById("run");

        if (sideloadMsg && appBody && runButton) {
            sideloadMsg.style.display = "none";
            appBody.style.display = "flex";
            runButton.onclick = run;
        } else {
            console.error("One or more required elements not found.");
        }
        // Set up ItemChanged event
        Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, itemChanged);

        UpdateTaskPaneUI(Office.context.mailbox.item);
    }
});

function OnAppointmentSend(event: Office.AddinCommands.Event) {
    const item = Office.context.mailbox.item;
    if (item) {
        item.body.getAsync(
            "text",
            { asyncContext: event },
            getBodyCallback
        );
    }
}

function getBodyCallback(asyncResult: Office.AsyncResult<string>) {
    const emailBody = asyncResult.value;
    console.log(emailBody);
}

function itemChanged(eventArgs: any) {
    // Update UI based on the new current item
    UpdateTaskPaneUI(Office.context.mailbox.item);
}


// Example implementation
function UpdateTaskPaneUI(item: Office.Item & Office.ItemCompose & Office.ItemRead & Office.Message & Office.MessageCompose & Office.MessageRead & Office.Appointment & Office.AppointmentCompose & Office.AppointmentRead | undefined) {
    // Assuming that item is always a read item (instead of a compose item).
    if (item != null) console.log(item.subject);
}

export async function run() {
    /**
     * Insert your Outlook code here
     */

    console.log("Calling run()...");
}