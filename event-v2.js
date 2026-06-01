/* Neutralny event-v2 - nic nie wstawia. */
function onNewMessageComposeHandlerV2(event) {
  try {
    if (event && typeof event.completed === "function") {
      event.completed();
    }
  } catch (e) {}
}
try {
  Office.actions.associate("onNewMessageComposeHandlerV2", onNewMessageComposeHandlerV2);
} catch (e) {}
