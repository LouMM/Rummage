
export function OnAppointmentTimeChanged(event:any) {
    setSubject(event);
}

export function setSubject(event: any) {
    if (Office.context.mailbox.item) {
        Office.context.mailbox.item.subject.setAsync(
            "Set by an event-based add-in!",
            {
                "asyncContext": event
            },
            function (asyncResult) {
                // Handle success or error.
                if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
                    console.error("Failed to set subject: " + JSON.stringify(asyncResult.error));
                }

                // Call event.completed() to signal to the Outlook client that the add-in has completed processing the event.
                asyncResult.asyncContext.completed();
            });
    } else {
        console.error("Office.context.mailbox.item is undefined.");
    }
}
  
// IMPORTANT: To ensure your add-in is supported in the Outlook client on Windows, remember to map the event handler name specified in the manifest to its JavaScript counterpart.
if (Office.context.platform === Office.PlatformType.PC || Office.context.platform == null) {
    Office.actions.associate("OnAppointmentTimeChanged", OnAppointmentTimeChanged);}

