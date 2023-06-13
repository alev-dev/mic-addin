// 1. How to construct online meeting details.
// Not shown: How to get the meeting organizer's ID and other details from your service.
const newBody = '<br>' +
    '<a href="https://events-staging.onlive.site/event/71edc0d8-c99e-4b30-8c05-ebd5b5a71248" target="_blank">Join Onlive.site meeting</a>' +
    '<br><br>';

const location = 'https://events-staging.onlive.site/event/71edc0d8-c99e-4b30-8c05-ebd5b5a71248';

let mailboxItem;

// Office is ready.
Office.onReady(function () {
        mailboxItem = Office.context.mailbox.item;
    }
);

// 2. How to define and register a function command named `insertOnliveMeeting` (referenced in the manifest)
//    to update the meeting body with the online meeting details.
function insertOnliveMeeting(event) {
    // Get HTML body from the client.
    mailboxItem.location.getAsync(
        function callback(asyncResult) {
        if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
            updateBody(asyncResult.context, asyncResult.value);
        } else {
            console.error("Failed to get Location.");
            return;
        }
    });
    /*mailboxItem.body.getAsync("html",
        { asyncContext: event },
        function (getBodyResult) {
            if (getBodyResult.status === Office.AsyncResultStatus.Succeeded) {
                updateBody(getBodyResult.asyncContext, getBodyResult.value);
            } else {
                console.error("Failed to get HTML body.");
                getBodyResult.asyncContext.completed({ allowEvent: false });
            }
        }
    );*/
}
// Register the function.
Office.actions.associate("insertOnliveMeeting", insertOnliveMeeting);

// 3. How to implement a supporting function `updateBody`
//    that appends the online meeting details to the current body of the meeting.
function updateBody(event, existingBody) {
    // Append new body to the existing body.
    mailboxItem.location.setAsync(location, (result) => {
        if (result.status !== Office.AsyncResultStatus.Succeeded) {
          console.error(`Action failed with message ${result.error.message}`);
          return;
        }
        console.log(`Successfully set location to ${location}`);
    });
    /*mailboxItem.body.setAsync(existingBody + newBody,
        { asyncContext: event, coercionType: "html" },
        function (setBodyResult) {
            if (setBodyResult.status === Office.AsyncResultStatus.Succeeded) {
                setBodyResult.asyncContext.completed({ allowEvent: true });
            } else {
                console.error("Failed to set HTML body.");
                setBodyResult.asyncContext.completed({ allowEvent: false });
            }
        }
    );*/
}