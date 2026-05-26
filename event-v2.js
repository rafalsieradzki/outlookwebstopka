// event.js - WYLACZONE automatyczne dzialanie
// Ten plik celowo nic nie wstawia przy tworzeniu nowej wiadomosci.

function onNewMessageComposeHandler(event) {
  try {
    event.completed();
  } catch (e) {}
}

function onNewMessageComposeHandlerV2(event) {
  try {
    event.completed();
  } catch (e) {}
}

Office.actions.associate("onNewMessageComposeHandler", onNewMessageComposeHandler);
Office.actions.associate("onNewMessageComposeHandlerV2", onNewMessageComposeHandlerV2);
