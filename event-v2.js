// Stopka Familijna - event-v2.js
// Neutralny plik zabezpieczajacy na wypadek starego cache Outlooka.
// Niczego nie wstawia.

function onNewMessageComposeHandlerV2(event) {
  try {
    if (event && event.completed) event.completed();
  } catch (e) {
  }
}

Office.actions.associate(
  "onNewMessageComposeHandlerV2",
  onNewMessageComposeHandlerV2
);
